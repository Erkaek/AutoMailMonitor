/* Mail Monitor — preload shim
 * Reconstitue une API simple côté page atop chrome.webview (WebView2).
 * Le host C# (WebBridge.cs) répond à des messages JSON {id, method, args}.
 */
(function () {
  if (!window.chrome || !window.chrome.webview) {
    console.error('WebView2 non détecté — page chargée hors host ?');
    return;
  }

  const wv = window.chrome.webview;
  const pending = new Map();
  let seq = 1;
  const eventListeners = new Map();

  wv.addEventListener('message', (ev) => {
    let msg = ev.data;
    if (typeof msg === 'string') {
      try { msg = JSON.parse(msg); } catch { return; }
    }
    if (!msg) return;

    if (msg.id !== undefined && pending.has(msg.id)) {
      const entry = pending.get(msg.id);
      pending.delete(msg.id);
      // Annule le timeout pour éviter d'accumuler des timers (review Copilot)
      if (entry.timeout) clearTimeout(entry.timeout);
      if (msg.ok) entry.resolve(msg.result);
      else entry.reject(new Error(msg.error || 'RPC error'));
      return;
    }
    if (msg.event) {
      const fns = eventListeners.get(msg.event);
      if (fns) for (const fn of fns) {
        try { fn(msg.data); } catch (e) { console.error(e); }
      }
    }
  });

  function call(method, ...args) {
    const id = seq++;
    return new Promise((resolve, reject) => {
      // Timeout auto-cleanup (10s) — l'id du timer est stocké dans pending
      // pour pouvoir l'annuler dès qu'une réponse arrive (review Copilot)
      const timeout = setTimeout(() => {
        if (pending.has(id)) {
          pending.delete(id);
          reject(new Error(`[MailMonitor] RPC ${method}(${id}) timeout`));
        }
      }, 10000);
      pending.set(id, { resolve, reject, timeout });
      try {
        wv.postMessage({ id, method, args });
      } catch (e) {
        pending.delete(id);
        clearTimeout(timeout);
        reject(e);
      }
    });
  }

  // Purge pending map sur navigation/reload
  wv.addEventListener('navigationStarting', () => {
    for (const entry of pending.values()) {
      if (entry.timeout) clearTimeout(entry.timeout);
      try { entry.reject(new Error('[MailMonitor] navigation: pending RPC annulé')); } catch {}
    }
    pending.clear();
  });

  function on(event, fn) {
    if (!eventListeners.has(event)) eventListeners.set(event, new Set());
    eventListeners.get(event).add(fn);
    return () => eventListeners.get(event)?.delete(fn);
  }

  // API publique
  window.api = {
    // app
    version:           () => call('app.version'),
    autostartGet:      () => call('app.autostart.get'),
    autostartSet:      (v) => call('app.autostart.set', !!v),
    checkUpdates:      () => call('app.check-updates'),
    applyUpdate:       () => call('app.apply-update'),

    // état
    monitoringStatus:  () => call('monitoring.status'),
    outlookStatus:     () => call('outlook.status'),

    // outlook discovery
    outlookListStores:  () => call('outlook.list-stores'),
    outlookListFolders: (storeId) => call('outlook.list-folders', storeId),

    // dossiers monitorés
    foldersList:    () => call('folders.list-monitored'),
    folderAdd:      (payload) => call('folders.add', payload),
    folderRemove:   (entryId) => call('folders.remove', entryId),

    // stats
    statsSummary:    () => call('stats.summary'),
    statsWeekly:     (weeks=12) => call('stats.weekly', weeks),
    statsByCategory: () => call('stats.by-category'),

    // emails / logs
    emailsRecent: (limit=50) => call('emails.recent', limit),
    logsRecent:   () => call('logs.recent'),

    // commentaires hebdo
    weeklyCommentsList: (year, week) => call('weekly-comments.list', year, week),
    weeklyCommentsAdd:  (payload)    => call('weekly-comments.add', payload),

    // window
    windowMinimize: () => call('window.minimize'),
    windowClose:    () => call('window.close'),

    // events
    on,
  };

  // Compat layer : recrée window.electronAPI utilisé par les scripts legacy.
  // Couvre l'ensemble des canaux IPC de l'ancien préload Electron.
  const c = (method, ...args) => call(method, ...args);
  window.electronAPI = {
    // --- Application / fenêtre ---
    getAppVersion:       () => c('app.version'),
    minimize:            () => c('window.minimize'),
    close:               () => c('window.close'),

    // --- Outlook / monitoring statut ---
    outlookStatus:        () => c('outlook.status'),
    getMonitoringStatus:  () => c('monitoring.status'),

    // --- Outlook discovery (stores / dossiers) ---
    olStores:              () => c('outlook.list-stores'),
    olFoldersShallow:      (storeId, parentId) => c('outlook.folders-shallow', storeId, parentId),
    ewsTopLevel:           () => c('outlook.list-stores'),
    ewsChildren:           (storeId, parentId) => c('outlook.folders-shallow', storeId, parentId),
    getSubFolders:         (storeId, parentId) => c('outlook.folders-shallow', storeId, parentId),
    getFolderStructure:    () => c('outlook.folders-tree'),
    getFolderTreeFromRoot: (rootPath, maxDepth) => c('outlook.folder-tree-from-path', rootPath, maxDepth ?? 4),
    getFoldersTree:        () => c('folders.tree'),

    // --- Dossiers monitorés ---
    loadFoldersConfig:        () => c('folders.list-monitored'),
    saveFoldersConfig:        (list) => c('folders.save-config', list),
    addFolderToMonitoring:    (payload) => c('folders.add', payload),
    addFoldersToMonitoringBulk: (list) => c('folders.add-bulk', list),
    removeFolderFromMonitoring: (entryId) => c('folders.remove', entryId),
    updateFolderCategory:     (entryId, category) => c('folders.update-category', entryId, category),
    getFolderStats:           (entryId) => c('folders.stats', entryId),

    // --- Stats / Dashboard / Emails ---
    getStatsSummary:    () => c('stats.summary'),
    getStatsByCategory: () => c('stats.by-category'),
    getRecentEmails:    (opts) => c('emails.recent', opts || {}),

    // --- Suivi hebdomadaire ---
    listWeeklyComments:   (year, week) => c('weekly.comments-list', year, week),
    addWeeklyComment:     (payload) => c('weekly.comments-add', payload),
    updateWeeklyComment:  (payload) => c('weekly.comments-update', payload),
    deleteWeeklyComment:  (id) => c('weekly.comments-delete', id),
    listWeeksForComments: () => c('weekly.weeks-list'),

    // --- VBA / xlsb stats ---
    getVBAMetricsSummary:    () => c('vba.metrics-summary'),
    getVBAFolderDistribution:() => c('vba.folder-distribution'),
    getVBAWeeklyEvolution:   (weeks) => c('vba.weekly-evolution', weeks || 12),

    // --- Import Activité (XLSB) ---
    openXlsbFile:         () => c('xlsb.pick-file'),
    importActivityPreview:(path) => c('xlsb.preview', path),
    importActivityRun:    (path) => c('xlsb.import', path),

    // --- DB lecteur brut ---
    getDbTables:        () => c('db.tables'),
    getDbTablePreview:  (a, b) => (a && typeof a === 'object') ? c('db.table-preview', a) : c('db.table-preview', a, b || 100),

    // --- Settings ---
    loadAppSettings:    () => c('settings.get-all'),
    saveAppSettings:    (settings) => c('settings.set-all', settings),

    // --- Logs ---
    openLogsFolder:     () => c('logs.open-folder'),

    // --- Invocation générique (canaux 'api-xxx') ---
    invoke: (channel, ...args) => c(channel, ...args),

    // --- Events realtime ---
    onLogEntry:                 (fn) => on('logs.entry', fn),
    onStatsUpdate:              (fn) => on('stats.update', fn),
    onStatsCacheInvalidated:    (fn) => on('stats.cache-invalidated', fn),
    onWeeklyStatsUpdated:       (fn) => on('weekly.stats-updated', fn),
    onEmailUpdate:              (fn) => on('email.update', fn),
    onNewEmail:                 (fn) => on('email.new', fn),
    onRealtimeEmailUpdate:      (fn) => on('email.realtime-update', fn),
    onRealtimeNewEmail:         (fn) => on('email.realtime-new', fn),
    onMonitoringStatus:         (fn) => on('monitoring.status', fn),
    onMonitoringCycleComplete:  (fn) => on('monitoring.cycle-complete', fn),
    onFolderCountUpdated:       (fn) => on('folders.count-updated', fn),
    onCOMListeningStarted:      (fn) => on('com.listening-started', fn),
    onCOMListeningFailed:       (fn) => on('com.listening-failed', fn),
    onUpdateAvailable:          (fn) => on('update.available', fn),
    onUpdateNotAvailable:       (fn) => on('update.not-available', fn),
    onUpdateChecking:           (fn) => on('update.checking', fn),
    onUpdateDownloadProgress:   (fn) => on('update.download-progress', fn),
    onUpdateError:              (fn) => on('update.error', fn),
    onUpdatePendingRestart:     (fn) => on('update.pending-restart', fn),

    on,
  };
})();

