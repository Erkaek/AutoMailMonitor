const { contextBridge, ipcRenderer } = require('electron');

// API pour la page de chargement
contextBridge.exposeInMainWorld('electronAPI', {
  // Ecouter les mises à jour de progression
  onLoadingProgress: (callback) => {
    ipcRenderer.on('loading-progress', (event, data) => callback(data));
  },

  // Ecouter la completion du chargement
  onLoadingComplete: (callback) => {
    ipcRenderer.on('loading-complete', () => callback());
  },

  // Ecouter les erreurs de chargement
  onLoadingError: (callback) => {
    ipcRenderer.on('loading-error', (event, error) => callback(error));
  },
  
  // Écouter la progression de l'analyse PowerShell
  onPowerShellAnalysisProgress: (callback) => {
    ipcRenderer.on('powershell-analysis-progress', (event, data) => callback(data));
  },

  // Signaler que la page de chargement est prete à se fermer
  loadingComplete: () => {
    ipcRenderer.send('loading-page-complete');
  },

  // Demander un retry
  retryLoading: () => {
    ipcRenderer.send('loading-retry');
  }
});
