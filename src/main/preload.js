/**
 * Preload script pour Mail Monitor
 * Expose l'API IPC de manière sécurisée au frontend
 */

const { contextBridge, ipcRenderer } = require('electron');

// Exposer l'API de manière sécurisée
contextBridge.exposeInMainWorld('electronAPI', {
  // API Outlook
  outlookStatus: () => ipcRenderer.invoke('api-outlook-status'),
  getMailboxes: () => ipcRenderer.invoke('api-outlook-mailboxes'),
  getFolderStructure: (storeId) => ipcRenderer.invoke('api-outlook-folder-structure', storeId),
  getFolderTreeFromRoot: (rootPath, maxDepth) => ipcRenderer.invoke('api-outlook-folder-tree-from-path', { rootPath, maxDepth }),
  getSubFolders: (payload) => ipcRenderer.invoke('api-outlook-subfolders', payload),
  listFoldersRecursive: (storeId, maxDepth) => ipcRenderer.invoke('api-outlook-folders-recursive', { storeId, maxDepth }),
  // EWS fast enumeration
  ewsTopLevel: (mailbox) => ipcRenderer.invoke('api-ews-top-level', { mailbox }),
  ewsChildren: (mailbox, parentId) => ipcRenderer.invoke('api-ews-children', { mailbox, parentId }),
  // COM fast enumeration
  olStores: () => ipcRenderer.invoke('api-ol-stores'),
  olFoldersShallow: (storeId, parentEntryId) => ipcRenderer.invoke('api-ol-folders-shallow', { storeId, parentEntryId }),
  
  // API Stats
  getStatsSummary: () => ipcRenderer.invoke('api-stats-summary'),
  getStatsByCategory: () => ipcRenderer.invoke('api-stats-by-category'),
  
  // API VBA Metrics (nouvelles métriques compatibles macros VBA)
  getVBAMetricsSummary: () => ipcRenderer.invoke('api-vba-metrics-summary'),
  getVBAFolderDistribution: () => ipcRenderer.invoke('api-vba-folder-distribution'),
  getVBAWeeklyEvolution: () => ipcRenderer.invoke('api-vba-weekly-evolution'),
  
  // API Emails
  getRecentEmails: () => ipcRenderer.invoke('api-recent-emails'),
  
  // API Settings
  saveFoldersConfig: (data) => ipcRenderer.invoke('api-settings-folders', data),
  loadFoldersConfig: () => ipcRenderer.invoke('api-settings-folders-load'),
  dumpFoldersConfig: () => ipcRenderer.invoke('api-settings-folders-dump'),
  
  // API Gestion hiérarchique des dossiers (nouveau)
  getFoldersTree: () => ipcRenderer.invoke('api-folders-tree'),
  addFolderToMonitoring: (data) => ipcRenderer.invoke('api-folders-add', data),
  addFoldersToMonitoringBulk: (data) => ipcRenderer.invoke('api-folders-add-bulk', data),
  updateFolderCategory: (data) => ipcRenderer.invoke('api-folders-update-category', data),
  removeFolderFromMonitoring: (data) => ipcRenderer.invoke('api-folders-remove', data),
  
  // Alias pour compatibilité avec le FoldersTreeManager
  invoke: (channel, data) => ipcRenderer.invoke(channel, data),
  
  // API Paramètres de l'application
  saveAppSettings: (settings) => ipcRenderer.invoke('api-app-settings-save', settings),
  loadAppSettings: () => ipcRenderer.invoke('api-app-settings-load'),

  // Version applicative (source unique)
  getAppVersion: () => ipcRenderer.invoke('app-get-version'),
  checkUpdatesNow: () => ipcRenderer.invoke('app-check-updates-now'),

  // Logs API
  getLogs: (opts) => ipcRenderer.invoke('api-logs-list', opts || {}),
  // Compat: ancien export => nouvel export filtré (par défaut: tout)
  exportLogs: () => ipcRenderer.invoke('api-export-log-history', { level: 'ALL', category: 'ALL', search: '', limit: 2000 }),
  openLogsFolder: () => ipcRenderer.invoke('api-logs-open-folder'),
  onLogEntry: (callback) => ipcRenderer.on('log-entry', (event, ...args) => callback(...args)),

  // API Commentaires hebdomadaires
  addWeeklyComment: (payload) => ipcRenderer.invoke('api-weekly-comments-add', payload),
  listWeeklyComments: (weekIdentifier) => ipcRenderer.invoke('api-weekly-comments-list', { week_identifier: weekIdentifier }),
  updateWeeklyComment: (id, comment_text, category) => ipcRenderer.invoke('api-weekly-comments-update', { id, comment_text, category }),
  deleteWeeklyComment: (id) => ipcRenderer.invoke('api-weekly-comments-delete', { id }),
  listWeeksForComments: (limit) => ipcRenderer.invoke('api-weekly-weeks-list', { limit }),
  
  // API Base de données
  getCategoryStats: () => ipcRenderer.invoke('api-database-category-stats'),
  getFolderStats: () => ipcRenderer.invoke('api-database-folder-stats'),
  
  // APIs de monitoring
  // start/stop manuels retirés de l'UI; on conserve uniquement la lecture du statut
  getMonitoringStatus: () => ipcRenderer.invoke('api-monitoring-status'),
  
  // Listeners pour les événements du processus principal
  onForceSync: (callback) => ipcRenderer.on('force-sync', callback),
  onSettingsChanged: (callback) => ipcRenderer.on('settings-changed', callback),
  
  // Événements de monitoring en temps réel
  onStatsUpdate: (callback) => ipcRenderer.on('stats-update', (event, ...args) => callback(...args)),
  onEmailUpdate: (callback) => ipcRenderer.on('email-update', (event, ...args) => callback(...args)),
  onNewEmail: (callback) => ipcRenderer.on('new-email', (event, ...args) => callback(...args)),
  onMonitoringCycleComplete: (callback) => ipcRenderer.on('monitoring-cycle-complete', (event, ...args) => callback(...args)),
  onMonitoringStatus: (callback) => ipcRenderer.on('monitoring-status', (event, ...args) => callback(...args)),
  
  // NOUVEAU: Événements COM Outlook temps réel
  onCOMListeningStarted: (callback) => ipcRenderer.on('com-listening-started', (event, ...args) => callback(...args)),
  onCOMListeningFailed: (callback) => ipcRenderer.on('com-listening-failed', (event, ...args) => callback(...args)),
  onRealtimeEmailUpdate: (callback) => ipcRenderer.on('realtime-email-update', (event, ...args) => callback(...args)),
  onRealtimeNewEmail: (callback) => ipcRenderer.on('realtime-new-email', (event, ...args) => callback(...args)),
  // Événement de mise à jour des stats hebdo (import, ajustements)
  onWeeklyStatsUpdated: (callback) => ipcRenderer.on('weekly-stats-updated', (event, ...args) => callback(...args)),
  
  // Événements de mise à jour automatique
  onUpdateChecking: (callback) => ipcRenderer.on('update-checking', (event, ...args) => callback(...args)),
  onUpdateAvailable: (callback) => ipcRenderer.on('update-available', (event, ...args) => callback(...args)),
  onUpdateNotAvailable: (callback) => ipcRenderer.on('update-not-available', (event, ...args) => callback(...args)),
  onUpdateError: (callback) => ipcRenderer.on('update-error', (event, ...args) => callback(...args)),
  onUpdateDownloadProgress: (callback) => ipcRenderer.on('update-download-progress', (event, ...args) => callback(...args)),
  onUpdatePendingRestart: (callback) => ipcRenderer.on('update-pending-restart', (event, ...args) => callback(...args)),
  
  // Loading events - pour la page de chargement
  onLoadingProgress: (callback) => ipcRenderer.on('loading-progress', (event, ...args) => callback(...args)),
  onLoadingComplete: (callback) => ipcRenderer.on('loading-complete', (event, ...args) => callback(...args)),
  onLoadingError: (callback) => ipcRenderer.on('loading-error', (event, ...args) => callback(...args)),
  onPowerShellAnalysisProgress: (callback) => ipcRenderer.on('analysis-progress', (event, ...args) => callback(...args)),
  
  // Actions loading
  loadingComplete: () => ipcRenderer.send('loading-page-complete'),
  retryLoading: () => ipcRenderer.send('loading-retry'),
  
  // Informations système
  getVersion: () => process.versions.electron,
  getPlatform: () => process.platform,
  
  // Gestion de la fenêtre
  minimizeWindow: () => ipcRenderer.invoke('window-minimize'),
  maximizeWindow: () => ipcRenderer.invoke('window-maximize'),
  closeWindow: () => ipcRenderer.invoke('window-close'),
  
  // Gestion des notifications
  showNotification: (title, body, options) => 
    ipcRenderer.invoke('show-notification', { title, body, options }),
  
  // Gestion des fichiers
  saveFile: (data, defaultPath) => ipcRenderer.invoke('dialog-save-file', { data, defaultPath }),
  openXlsbFile: () => ipcRenderer.invoke('dialog-open-xlsb'),
  importActivityPreview: (filePath, weeks) => ipcRenderer.invoke('api-activity-import-preview', { filePath, weeks }),
  importActivityRun: (filePath, weeks, outCsv) => ipcRenderer.invoke('api-activity-import-run', { filePath, weeks, outCsv }),
  
  // Contrôles de fenêtre
  minimize: () => ipcRenderer.invoke('window-minimize'),
  close: () => ipcRenderer.invoke('window-close'),
  
  // Nettoyage des listeners
  removeAllListeners: (channel) => ipcRenderer.removeAllListeners(channel),
  
  // Utilitaires
  log: (message) => console.log(`[Frontend] ${message}`)
});

console.log('🔧 Preload script chargé - API Electron disponible');