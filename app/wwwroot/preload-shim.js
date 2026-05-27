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
      const { resolve, reject } = pending.get(msg.id);
      pending.delete(msg.id);
      if (msg.ok) resolve(msg.result);
      else reject(new Error(msg.error || 'RPC error'));
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
      pending.set(id, { resolve, reject });
      try {
        wv.postMessage({ id, method, args });
      } catch (e) {
        pending.delete(id);
        reject(e);
      }
    });
  }

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

  // Compat layer : recrée window.electronAPI utilisé par certains scripts legacy.
  window.electronAPI = {
    getAppVersion:       window.api.version,
    getMonitoringStatus: window.api.monitoringStatus,
    outlookStatus:       window.api.outlookStatus,
    getStatsSummary:     window.api.statsSummary,
    getStatsByCategory:  window.api.statsByCategory,
    getRecentEmails:     window.api.emailsRecent,
    listWeeklyComments:  window.api.weeklyCommentsList,
    addWeeklyComment:    window.api.weeklyCommentsAdd,
    minimize:            window.api.windowMinimize,
    close:               window.api.windowClose,
    invoke: (method, ...args) => call(method, ...args),
    on,
  };
})();
