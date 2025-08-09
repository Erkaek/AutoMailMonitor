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
  
  // API Gestion hiérarchique des dossiers (nouveau)
  getFoldersTree: () => ipcRenderer.invoke('api-folders-tree'),
  addFolderToMonitoring: (data) => ipcRenderer.invoke('api-folders-add', data),
  updateFolderCategory: (data) => ipcRenderer.invoke('api-folders-update-category', data),
  removeFolderFromMonitoring: (data) => ipcRenderer.invoke('api-folders-remove', data),
  
  // Alias pour compatibilité avec le FoldersTreeManager
  invoke: (channel, data) => ipcRenderer.invoke(channel, data),
  
  // API Paramètres de l'application
  saveAppSettings: (settings) => ipcRenderer.invoke('api-app-settings-save', settings),
  loadAppSettings: () => ipcRenderer.invoke('api-app-settings-load'),
  
  // API Base de données
  getCategoryStats: () => ipcRenderer.invoke('api-database-category-stats'),
  getFolderStats: () => ipcRenderer.invoke('api-database-folder-stats'),
  
  // APIs de monitoring
  startMonitoring: () => ipcRenderer.invoke('api-monitoring-start'),
  stopMonitoring: () => ipcRenderer.invoke('api-monitoring-stop'),
  getMonitoringStatus: () => ipcRenderer.invoke('api-monitoring-status'),
  
  // Listeners pour les événements du processus principal
  onForceSync: (callback) => ipcRenderer.on('force-sync', callback),
  onSettingsChanged: (callback) => ipcRenderer.on('settings-changed', callback),
  
  // Événements de monitoring en temps réel
  onStatsUpdate: (callback) => ipcRenderer.on('stats-update', (event, ...args) => callback(...args)),
  onEmailUpdate: (callback) => ipcRenderer.on('email-update', (event, ...args) => callback(...args)),
  onNewEmail: (callback) => ipcRenderer.on('new-email', (event, ...args) => callback(...args)),
  onMonitoringCycleComplete: (callback) => ipcRenderer.on('monitoring-cycle-complete', (event, ...args) => callback(...args)),
  
  // NOUVEAU: Événements COM Outlook temps réel
  onCOMListeningStarted: (callback) => ipcRenderer.on('com-listening-started', (event, ...args) => callback(...args)),
  onCOMListeningFailed: (callback) => ipcRenderer.on('com-listening-failed', (event, ...args) => callback(...args)),
  onRealtimeEmailUpdate: (callback) => ipcRenderer.on('realtime-email-update', (event, ...args) => callback(...args)),
  onRealtimeNewEmail: (callback) => ipcRenderer.on('realtime-new-email', (event, ...args) => callback(...args)),
  
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
  
  // Contrôles de fenêtre
  minimize: () => ipcRenderer.invoke('window-minimize'),
  close: () => ipcRenderer.invoke('window-close'),
  
  // Nettoyage des listeners
  removeAllListeners: (channel) => ipcRenderer.removeAllListeners(channel),
  
  // Utilitaires
  log: (message) => console.log(`[Frontend] ${message}`)
});

console.log('🔧 Preload script chargé - API Electron disponible');