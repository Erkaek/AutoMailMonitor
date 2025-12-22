/**
 * Gestionnaire de mises à jour automatiques amélioré
 * Gère les mises à jour avec retry, timeout, et logs détaillés
 */

const { autoUpdater } = require('electron-updater');
const { dialog, app } = require('electron');
const logService = require('../services/logService');

class UpdateManager {
  constructor() {
    this.config = {
      requestTimeout: 30000, // 30 secondes
      maxRetries: 3,
      retryDelay: 5000, // 5 secondes entre chaque retry
      checkInterval: 2 * 60 * 60 * 1000, // 2 heures
      minCheckInterval: 5 * 60 * 1000 // 5 minutes minimum entre vérifications
    };
    
    this.updateCheckAttempts = 0;
    this.lastUpdateCheck = null;
    this.mainWindow = null;
    this.isSetup = false;

    // Si une MAJ échoue pour signature (Windows), on évite de re-télécharger en boucle.
    this.disableAutoDownloadForSession = false;

    this.releaseUrl = process.env.UPDATE_RELEASES_URL || 'https://github.com/Erkaek/AutoMailMonitor/releases/latest';
  }

  isWindowsSignatureError(message) {
    const msg = String(message || '').toLowerCase();
    return (
      msg.includes('not signed by the application owner') ||
      msg.includes('publishernames') ||
      msg.includes('certificat racine') ||
      msg.includes("n'est pas approuvé") ||
      msg.includes('untrustedroot') ||
      msg.includes('chain') && msg.includes('not trusted')
    );
  }

  /**
   * Configure l'auto-updater
   */
  configure() {
    // Par défaut en mode “manuel” (certificat interne non approuvé sur les postes)
    // Opt-in possible via env pour les postes où la chaîne est approuvée.
    autoUpdater.autoDownload = String(process.env.AUTO_UPDATE_AUTO_DOWNLOAD || 'false').toLowerCase() === 'true';
    autoUpdater.autoInstallOnAppQuit = true;
    autoUpdater.allowPrerelease = String(process.env.ALLOW_PRERELEASE || 'true').toLowerCase() !== 'false';
    
    // Headers pour éviter le cache
    autoUpdater.requestHeaders = autoUpdater.requestHeaders || {};
    autoUpdater.requestHeaders['Cache-Control'] = 'no-cache';
    
    logService.info('INIT', 'Auto-updater configuré');
  }

  /**
   * Configure le token GitHub pour dépôts privés
   */
  configureGitHubToken() {
    try {
      let bundledToken = null;
      try {
        bundledToken = require('../main/updaterToken');
      } catch {}
      
      const envToken = process.env.GH_TOKEN || process.env.UPDATER_TOKEN || process.env.ELECTRON_UPDATER_TOKEN;
      const updaterToken = (typeof bundledToken === 'string' && bundledToken.trim()) ? bundledToken.trim() : envToken;
      
      if (updaterToken) {
        autoUpdater.requestHeaders.Authorization = `token ${updaterToken}`;
        logService.info('INIT', 'Token GitHub configuré pour mises à jour (repo privé supporté)');
        return true;
      }
    } catch (e) {
      logService.warn('INIT', 'Erreur configuration token GitHub', e.message);
    }
    return false;
  }

  /**
   * Configure tous les événements de l'auto-updater
   */
  setupEvents() {
    if (this.isSetup) return;
    
    autoUpdater.on('checking-for-update', () => {
      this.updateCheckAttempts++;
      logService.info('INIT', `Vérification de mise à jour démarrée (tentative ${this.updateCheckAttempts})`);
      this.notifyWindow('update-checking');
    });
    
    autoUpdater.on('error', (err) => {
      const errorMsg = err?.message || String(err);
      logService.error('INIT', 'Erreur lors de la vérification de mise à jour', errorMsg);
      this.notifyWindow('update-error', { error: errorMsg });

      // Erreur de signature Windows: ne pas retry, et désactiver l’auto-download pour la session.
      if (this.isWindowsSignatureError(errorMsg)) {
        this.disableAutoDownloadForSession = true;
        try { autoUpdater.autoDownload = false; } catch {}
        logService.warn(
          'INIT',
          'Mise à jour rejetée (signature Windows). Désactivation auto-download pour la session.',
          'Vérifie le certificat de signature (éditeur) et la chaîne de confiance sur les postes.'
        );
        this.notifyWindow('update-error', {
          error: errorMsg,
          kind: 'SIGNATURE',
          action: 'Signature non approuvée sur ce poste. Installer manuellement depuis la page de release (lien) ou déployer la chaîne de confiance (si possible).',
          releaseUrl: this.releaseUrl
        });
        // Stopper les retries automatiques dans ce cas.
        this.updateCheckAttempts = 0;
        return;
      }

      this.handleUpdateError(errorMsg);
    });
    
    autoUpdater.on('update-available', (info) => {
      const currentVersion = app.getVersion();
      logService.success('INIT', `Mise à jour disponible: v${info.version} (actuelle: v${currentVersion})`);
      this.notifyWindow('update-available', {
        version: info.version,
        currentVersion,
        releaseDate: info.releaseDate,
        releaseNotes: info.releaseNotes,
        releaseUrl: this.releaseUrl,
        autoDownload: Boolean(autoUpdater.autoDownload)
      });
      this.updateCheckAttempts = 0;
    });
    
    autoUpdater.on('update-not-available', () => {
      const currentVersion = app.getVersion();
      logService.info('INIT', `Aucune mise à jour disponible (version actuelle: v${currentVersion})`);
      this.notifyWindow('update-not-available');
      this.updateCheckAttempts = 0;
    });
    
    autoUpdater.on('download-progress', (progressObj) => {
      const percent = Math.floor(progressObj.percent);
      const downloadedMB = (progressObj.transferred / 1024 / 1024).toFixed(2);
      const totalMB = (progressObj.total / 1024 / 1024).toFixed(2);
      const speedMBps = (progressObj.bytesPerSecond / 1024 / 1024).toFixed(2);
      
      logService.info('INIT', `Téléchargement: ${percent}% (${downloadedMB}/${totalMB} Mo, ${speedMBps} Mo/s)`);
      this.notifyWindow('update-download-progress', {
        percent,
        transferred: progressObj.transferred,
        total: progressObj.total,
        bytesPerSecond: progressObj.bytesPerSecond
      });
    });
    
    autoUpdater.on('update-downloaded', async (info) => {
      logService.success('INIT', `Mise à jour v${info.version} téléchargée et prête à installer`);
      await this.handleUpdateDownloaded(info);
    });
    
    this.isSetup = true;
    logService.success('INIT', 'Événements auto-updater configurés');
  }

  /**
   * Gère les erreurs de mise à jour avec retry automatique
   */
  handleUpdateError(errorMsg = '') {
    if (this.updateCheckAttempts < this.config.maxRetries) {
      // Éviter les retries à 0s (ex: erreur après update-available où attempts peut être retombé à 0)
      const attempts = Math.max(1, Number(this.updateCheckAttempts) || 0);
      const retryDelay = this.config.retryDelay * attempts;
      logService.info('INIT', `Nouvelle tentative dans ${retryDelay / 1000}s...`);
      
      setTimeout(() => {
        this.checkForUpdates().catch(e => {
          logService.error('INIT', 'Échec retry vérification MAJ', e.message);
        });
      }, retryDelay);
    } else {
      logService.warn('INIT', `Abandon après ${this.updateCheckAttempts} tentatives`);
      this.updateCheckAttempts = 0;
    }
  }

  /**
   * Gère la mise à jour téléchargée
   */
  async handleUpdateDownloaded(info) {
    try {
      // Créer un message détaillé
      let message = `La version ${info.version} a été téléchargée avec succès.\n\n`;
      
      if (info.releaseNotes) {
        const notes = typeof info.releaseNotes === 'string' 
          ? info.releaseNotes.substring(0, 200) 
          : '';
        if (notes) {
          message += `Nouveautés:\n${notes}${notes.length >= 200 ? '...' : ''}\n\n`;
        }
      }
      
      message += 'Voulez-vous redémarrer maintenant pour l\'installer ?';
      
      const result = await dialog.showMessageBox({
        type: 'info',
        title: 'Mise à jour prête',
        message,
        detail: 'L\'application se fermera et la mise à jour sera installée automatiquement.',
        buttons: ['Redémarrer maintenant', 'Plus tard'],
        cancelId: 1,
        defaultId: 0,
        noLink: true
      });
      
      if (result.response === 0) {
        logService.info('INIT', 'Installation de la mise à jour...');
        setImmediate(() => {
          autoUpdater.quitAndInstall(false, true);
        });
      } else {
        logService.info('INIT', 'Installation reportée - sera installée au prochain démarrage');
        this.notifyWindow('update-pending-restart', { version: info.version });
      }
    } catch (e) {
      logService.error('INIT', 'Erreur dialogue mise à jour', e.message);
    }
  }

  /**
   * Vérifie les mises à jour
   */
  async checkForUpdates() {
    try {
      if (this.disableAutoDownloadForSession) {
        logService.warn('INIT', 'Auto-download désactivé pour la session (erreur signature précédente)');
      }

      // Vérifier qu'on n'a pas vérifié trop récemment
      const now = Date.now();
      if (this.lastUpdateCheck && (now - this.lastUpdateCheck) < this.config.minCheckInterval) {
        logService.debug('INIT', 'Vérification ignorée (vérification récente)');
        return null;
      }
      
      this.lastUpdateCheck = now;
      this.updateCheckAttempts = 0;
      
      // Timeout pour ne pas bloquer indéfiniment
      const checkPromise = autoUpdater.checkForUpdates();
      const timeoutPromise = new Promise((_, reject) => 
        setTimeout(() => reject(new Error('Timeout vérification MAJ')), this.config.requestTimeout)
      );
      
      return await Promise.race([checkPromise, timeoutPromise]);
    } catch (e) {
      const errorMsg = e?.message || String(e);
      logService.error('INIT', 'Erreur vérification MAJ', errorMsg);
      throw e;
    }
  }

  /**
   * Vérifie les mises à jour manuellement (pour IPC)
   */
  async checkManually() {
    try {
      logService.info('IPC', 'Vérification manuelle des mises à jour demandée');
      this.updateCheckAttempts = 0;
      
      const res = await this.checkForUpdates();
      const info = res?.updateInfo || null;
      
      if (info) {
        logService.success('IPC', `Vérification réussie - Version distante: v${info.version}`);
      } else {
        logService.info('IPC', 'Aucune mise à jour disponible');
      }
      
      return {
        success: true,
        updateInfo: info,
        downloading: !!res?.downloadPromise,
        currentVersion: app.getVersion()
      };
    } catch (e) {
      const errorMsg = e?.message || String(e);
      logService.error('IPC', 'Erreur vérification manuelle MAJ', errorMsg);
      return { 
        success: false, 
        error: errorMsg,
        currentVersion: app.getVersion()
      };
    }
  }

  /**
   * Démarre la vérification périodique
   */
  startPeriodicCheck() {
    setInterval(() => {
      logService.info('INIT', 'Vérification périodique des mises à jour...');
      this.checkForUpdates().catch(e => {
        logService.warn('INIT', 'Échec vérification périodique MAJ', e.message);
      });
    }, this.config.checkInterval);
    
    logService.info('INIT', `Vérification périodique activée (toutes les ${this.config.checkInterval / 1000 / 60} minutes)`);
  }

  /**
   * Définit la fenêtre principale pour les notifications
   */
  setMainWindow(window) {
    this.mainWindow = window;
  }

  /**
   * Envoie une notification à la fenêtre principale
   */
  notifyWindow(event, data = {}) {
    if (this.mainWindow && !this.mainWindow.isDestroyed()) {
      this.mainWindow.webContents.send(event, data);
    }
  }

  /**
   * Initialise complètement le gestionnaire
   */
  initialize() {
    this.configure();
    this.configureGitHubToken();
    this.setupEvents();
    logService.success('INIT', 'UpdateManager initialisé');
  }
}

// Singleton
const updateManager = new UpdateManager();

module.exports = updateManager;
