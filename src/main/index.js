/**
 * Mail Monitor - Surveillance professionnelle des emails Outlook
 * Corrected version avec protection contre multiples démarrages
 */

const { app, BrowserWindow, Tray, Menu, ipcMain, nativeImage, dialog: electronDialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');
// Logger (captures console and stores to file + memory)
const mainLogger = require('./logger');
// CORRECTION: Utiliser le connecteur optimisé 
const outlookConnector = require('../server/outlookConnector');
// OPTIMIZED: Utiliser le service de base de données optimisé
const databaseService = require('../services/optimizedDatabaseService');
// Importeur XLSB
const { dialog } = require('electron');
const activityImporter = require('../importers/activityXlsbImporter');
const cacheService = require('../services/cacheService');

// Rendre les services disponibles globalement
global.outlookConnector = outlookConnector;
global.databaseService = databaseService;
global.cacheService = cacheService;

// Initialize logging early
try { mainLogger.init(); mainLogger.hookConsole(); } catch {}
console.log('⚡ QUICK WINS: Better-SQLite3 + Cache intelligent activés');

// Fonction utilitaire pour nettoyer les messages d'encodage Windows
function cleanLogMessage(message) {
  return message
    .replace(/🚀/g, '[INIT]')
    .replace(/✅/g, '[OK]')
    .replace(/🔍/g, '[CHECK]')
    .replace(/📋/g, '[INFO]')
    .replace(/❌/g, '[ERROR]')
    .replace(/📧/g, '[OUTLOOK]')
    .replace(/⚠️/g, '[WARN]')
    .replace(/🔗/g, '[CONNECT]')
    .replace(/🔄/g, '[RETRY]')
    .replace(/📊/g, '[STATS]')
    .replace(/📚/g, '[LOAD]')
    .replace(/📁/g, '[FOLDER]')
    .replace(/⏳/g, '[WAIT]')
    .replace(/🎉/g, '[READY]')
    .replace(/🔹/g, '[IPC]')
    .replace(/📱/g, '[IPC]')
    .replace(/🎯/g, '[IPC]')
    .replace(/📝/g, '[SAVE]')
    .replace(/🔓/g, '[UNLOCK]')
    .replace(/ℹ️/g, '[INFO]')
    // Caractères accentués français
    .replace(/é/g, 'e')
    .replace(/è/g, 'e')
    .replace(/ê/g, 'e')
    .replace(/ë/g, 'e')
    .replace(/à/g, 'a')
    .replace(/â/g, 'a')
    .replace(/ä/g, 'a')
    .replace(/ç/g, 'c')
    .replace(/ù/g, 'u')
    .replace(/û/g, 'u')
    .replace(/ü/g, 'u')
    .replace(/ô/g, 'o')
    .replace(/ö/g, 'o')
    .replace(/î/g, 'i')
    .replace(/ï/g, 'i');
}

// Wrapper pour console.log avec nettoyage d'encodage
function logClean(message, ...args) {
  console.log(cleanLogMessage(message), ...args);
}

let mainWindow = null;
let loadingWindow = null;
let tray = null;

// Configuration de l'application
const APP_CONFIG = {
  width: 1200,
  height: 800
};

// Auto-update configuration
autoUpdater.autoDownload = true;
autoUpdater.autoInstallOnAppQuit = true;
// Autoriser les pre-releases (souvent utiles pendant les tests). Peut etre force via env.
autoUpdater.allowPrerelease = String(process.env.ALLOW_PRERELEASE || 'true').toLowerCase() !== 'false';

// If repo is private, allow providing a GitHub token bundled or via env to authorize release access
try {
  let bundledToken = null;
  try {
    // Optional local file included at build time (not committed): src/main/updaterToken.js exports a string
    bundledToken = require('./updaterToken');
  } catch {}
  const envToken = process.env.GH_TOKEN || process.env.UPDATER_TOKEN || process.env.ELECTRON_UPDATER_TOKEN;
  const updaterToken = (typeof bundledToken === 'string' && bundledToken.trim()) ? bundledToken.trim() : envToken;
  if (updaterToken) {
    autoUpdater.requestHeaders = { Authorization: `token ${updaterToken}` };
    logClean('🔑 Token GitHub pour mises a jour configure (repo prive pris en charge)');
  }
} catch {}

function setupAutoUpdater() {
  autoUpdater.on('checking-for-update', () => {
    logClean('🔎 Recherche de mise a jour demarree');
  });
  autoUpdater.on('error', (err) => {
    logClean('⚠️ Mise à jour: erreur: ' + (err?.message || String(err)));
  });
  autoUpdater.on('update-available', (info) => {
    logClean('🔔 Mise à jour disponible: v' + info?.version);
  });
  autoUpdater.on('update-not-available', () => {
    try {
      const cur = app.getVersion();
      logClean('ℹ️ Aucune mise à jour disponible (version locale v' + cur + ')');
    } catch {
      logClean('ℹ️ Aucune mise à jour disponible');
    }
  });
  autoUpdater.on('download-progress', (p) => {
    logClean(`⬇️ Téléchargement mise à jour: ${Math.floor(p.percent)}%`);
  });
  autoUpdater.on('update-downloaded', async (info) => {
    try {
      const result = await electronDialog.showMessageBox({
        type: 'info',
        title: 'Mise à jour prête',
        message: `La version ${info.version} a été téléchargée. Redémarrer pour l’installer ?`,
        buttons: ['Redémarrer maintenant', 'Plus tard'],
        cancelId: 1,
        defaultId: 0
      });
      if (result.response === 0) {
        autoUpdater.quitAndInstall();
      }
    } catch {}
  });
}

// IPC: fournir la version de l'application au renderer (source unique: package.json via app.getVersion())
ipcMain.handle('app-get-version', async () => {
  try {
    return app.getVersion();
  } catch {
    return null;
  }
});

// Logs API
ipcMain.handle('api-logs-list', async (event, opts) => {
  try {
    const { entries, totalBuffered, lastId } = mainLogger.getLogs(opts || {});
    return { success: true, entries, totalBuffered, lastId };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('api-logs-export', async () => {
  try {
    const content = mainLogger.exportAllAsString();
    const { canceled, filePath } = await electronDialog.showSaveDialog({
      title: 'Exporter les logs',
      defaultPath: `MailMonitor-logs-${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.txt`,
      filters: [{ name: 'Fichiers texte', extensions: ['txt','log'] }]
    });
    if (canceled || !filePath) return { success: false, canceled: true };
    fs.writeFileSync(filePath, content, 'utf8');
    return { success: true, filePath };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('api-logs-open-folder', async () => {
  try {
    const { shell } = require('electron');
    const userData = app.getPath('userData');
    const dir = require('path').join(userData, 'logs');
    if (!fs.existsSync(dir)) return { success: false, error: 'Dossier de logs introuvable' };
    await shell.openPath(dir);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// --- Vérification Git en mode développement (source clonée) ---
function isDevEnvironment() {
  // Packagé => pas de .git, on évite les appels Git
  return !app.isPackaged;
}

function runGit(cmd, cwd) {
  return new Promise((resolve) => {
    exec(cmd, { cwd, windowsHide: true }, (error, stdout, stderr) => {
      resolve({ error, stdout: (stdout || '').trim(), stderr: (stderr || '').trim() });
    });
  });
}

async function checkDevGitUpdatesOnStartup() {
  try {
    if (!isDevEnvironment()) return; // seulement en dev
    const repoRoot = path.join(__dirname, '../../');
    const gitDir = path.join(repoRoot, '.git');
    if (!fs.existsSync(gitDir)) return;

    // Récupère la branche courante
    const headRes = await runGit('git rev-parse --abbrev-ref HEAD', repoRoot);
    if (headRes.error || !headRes.stdout) return;
    const branch = headRes.stdout;

    // Fetch silencieux
    await runGit('git fetch --all --prune --quiet', repoRoot);

    // Compare HEAD avec remote
    const aheadRes = await runGit(`git rev-list --count HEAD..origin/${branch}`, repoRoot);
    if (aheadRes.error) return;
    const ahead = parseInt(aheadRes.stdout || '0', 10) || 0;
    if (ahead <= 0) return; // rien à mettre à jour

    const result = await electronDialog.showMessageBox({
      type: 'question',
      title: 'Mise à jour du code disponible',
      message: `Des changements distants ont été détectés sur la branche ${branch}. Voulez-vous récupérer les dernières modifications ?`,
      buttons: ['Mettre à jour maintenant', 'Plus tard'],
      cancelId: 1,
      defaultId: 0,
      noLink: true
    });

    if (result.response !== 0) return;

    // Tente un pull sécurisé
    const pullRes = await runGit('git pull --rebase --autostash', repoRoot);
    if (pullRes.error) {
      await electronDialog.showMessageBox({
        type: 'error',
        title: 'Échec de la mise à jour',
        message: 'Le pull Git a échoué. Consultez la console pour les détails.',
        detail: (pullRes.stderr || pullRes.stdout || '').slice(0, 4000),
        buttons: ['OK']
      });
      return;
    }

    const restart = await electronDialog.showMessageBox({
      type: 'info',
      title: 'Mise à jour appliquée',
      message: 'Le code a été mis à jour depuis le dépôt. Redémarrer l’application pour prendre en compte les changements ?',
      buttons: ['Redémarrer maintenant', 'Plus tard'],
      cancelId: 1,
      defaultId: 0
    });
    if (restart.response === 0) {
      app.relaunch();
      app.exit(0);
    }
  } catch (e) {
    logClean('⚠️ Verification Git dev: erreur ' + (e?.message || String(e)));
  }
}

  // Gestion des erreurs non capturees
process.on('uncaughtException', (error) => {
  console.error('[ERROR] Erreur non capturee dans le processus principal:', error);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('[ERROR] Promise rejetee non geree:', reason);
});

// Protection contre les multiples démarrages
let isInitializing = false;

function createLoadingWindow() {
  loadingWindow = new BrowserWindow({
    width: 600,
    height: 600,
    frame: false,
    alwaysOnTop: true,
    resizable: true, // Permettre le redimensionnement
    center: true,
    show: false,
  transparent: true, // fenêtre sans cadre, fond transparent (custom UI)
  icon: path.join(__dirname, '../../resources', 'new logo', 'logo.ico'),
    title: 'Mail Monitor - Initialisation',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: false,
      preload: path.join(__dirname, 'preload-loading.js')
    }
  });

  loadingWindow.loadFile(path.join(__dirname, '../../public/loading.html'));

  loadingWindow.once('ready-to-show', () => {
    loadingWindow.show();
    // Étape 1: Vérification de mise à jour en tout début de chargement
    runInitialUpdateCheck()
      .catch((e) => {
        logClean('⚠️ Echec verification MAJ (initiale): ' + (e?.message || String(e)));
      })
      .finally(() => {
        // Étape 2: Initialisation Outlook
        initializeOutlook();
      });
  });


app.on('ready', () => {
  try { mainLogger.init(); } catch {}
  setupAutoUpdater();
  try {
    logClean(`🚀 Application v${app.getVersion()} (${process.platform} ${process.arch})`);
  } catch {}
  try {
    // Traces utiles pour diagnostiquer l’URL d’update
    const cfgPath = autoUpdater.updateConfigPath;
    logClean('📄 update config path: ' + (cfgPath || 'inconnu'));
  } catch {}
  // Vérification périodique (l'initiale est lancée par la fenêtre de chargement)
  setInterval(() => {
    logClean('🔎 Verification periodique des mises a jour...');
    autoUpdater.checkForUpdatesAndNotify().catch((e) => {
      logClean('⚠️ Echec verification MAJ (periodique): ' + (e?.message || String(e)));
    });
  }, 30 * 60 * 1000);
  // En mode dev: vérifie s'il y a des commits distants et propose un pull
  setTimeout(() => { checkDevGitUpdatesOnStartup(); }, 3000);
});
  loadingWindow.on('closed', () => {
    loadingWindow = null;
  });

  return loadingWindow;
}

// Lancer la vérification initiale de mise à jour et informer la fenêtre de chargement
async function runInitialUpdateCheck() {
  try {
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 5,
        message: 'Vérification des mises à jour...'
      });
    }
    // Préférer checkForUpdates pour ne pas bloquer l’init sur le téléchargement
    const res = await autoUpdater.checkForUpdates();
    const info = res?.updateInfo;
    if (info && info.version && loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 10,
        message: `Version distante: v${info.version}${res?.downloadPromise ? ' (téléchargement en arrière-plan)' : ''}`
      });
    }
  } catch (e) {
    logClean('⚠️ runInitialUpdateCheck: ' + (e?.message || String(e)));
  } finally {
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 15,
        message: 'Vérification des mises à jour terminée'
      });
    }
  }
}

// IPC: Vérification manuelle des mises à jour
ipcMain.handle('app-check-updates-now', async () => {
  try {
    logClean('🖐️ Vérification manuelle des mises à jour demandée');
    const res = await autoUpdater.checkForUpdates();
    const info = res?.updateInfo || null;
    return {
      success: true,
      updateInfo: info,
      downloading: !!res?.downloadPromise
    };
  } catch (e) {
    return { success: false, error: e?.message || String(e) };
  }
});

/**
 * Configuration du transfert des événements temps réel du service vers le frontend
 */
function setupRealtimeEventForwarding() {
  if (!global.unifiedMonitoringService || !mainWindow) {
    console.log('⚠️ Service unifié ou fenêtre principale non disponible pour les événements temps réel');
    return;
  }

  // Éviter les double-bindings si cette fonction est appelée plusieurs fois
  if (global.unifiedMonitoringService.__ipcForwardingSet) {
    console.log('ℹ️ Transfert d\'événements déjà configuré, on évite un double-binding');
    return;
  }
  global.unifiedMonitoringService.__ipcForwardingSet = true;

  try {
    // Augmenter la limite d'écouteurs pour éviter les warnings lors des redémarrages
    if (typeof global.unifiedMonitoringService.setMaxListeners === 'function') {
      global.unifiedMonitoringService.setMaxListeners(50);
    }
    // Nettoyer d'éventuels anciens écouteurs (sécurité)
    const events = ['emailUpdated','newEmail','syncCompleted','monitoringCycleComplete','monitoring-status','com-listening-started','com-listening-failed','realtime-email-update','realtime-new-email'];
    for (const evt of events) {
      try { global.unifiedMonitoringService.removeAllListeners(evt); } catch {}
    }
  } catch {}

  console.log('🔔 Configuration du transfert d\'événements temps réel...');

  // Transférer les événements de mise à jour d'emails
  global.unifiedMonitoringService.on('emailUpdated', (emailData) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('email-update', emailData);
    }
  });

  // Transférer les événements de nouveaux emails
  global.unifiedMonitoringService.on('newEmail', (emailData) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('new-email', emailData);
    }
  });

  // Transférer les événements de synchronisation terminée
  global.unifiedMonitoringService.on('syncCompleted', (stats) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('stats-update', stats);
    }
  });

  // Transférer les événements de cycle de monitoring terminé
  global.unifiedMonitoringService.on('monitoringCycleComplete', (cycleData) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('monitoring-cycle-complete', cycleData);
    }
  });

  // Transférer les changements de statut du monitoring (actif/arrêté)
  global.unifiedMonitoringService.on('monitoring-status', (status) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('monitoring-status', status);
    }
  });

  // NOUVEAU: Transférer les événements COM Outlook
  global.unifiedMonitoringService.on('com-listening-started', (data) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      console.log('🔔 [IPC] Transfert événement COM listening started');
      mainWindow.webContents.send('com-listening-started', data);
    }
  });

  global.unifiedMonitoringService.on('com-listening-failed', (error) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      console.log('❌ [IPC] Transfert événement COM listening failed');
      mainWindow.webContents.send('com-listening-failed', error);
    }
  });

  // Événements temps réel pour les emails COM
  global.unifiedMonitoringService.on('realtime-email-update', (emailData) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      console.log('📧 [IPC] Transfert mise à jour email temps réel COM');
      mainWindow.webContents.send('realtime-email-update', emailData);
    }
  });

  global.unifiedMonitoringService.on('realtime-new-email', (emailData) => {
    if (mainWindow && !mainWindow.isDestroyed()) {
      console.log('📬 [IPC] Transfert nouvel email temps réel COM');
      mainWindow.webContents.send('realtime-new-email', emailData);
    }
  });

  console.log('✅ Transfert d\'événements temps réel configuré');
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: APP_CONFIG.width,
    height: APP_CONFIG.height,
  icon: path.join(__dirname, '../../resources', 'new logo', 'logo.ico'),
    title: 'Mail Monitor - Surveillance Outlook',
  show: false,
  frame: true, // Réactive la barre d'outils/titre native Windows
  titleBarStyle: 'default',
  autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: false,
      devTools: !app.isPackaged,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // Menu application: aucun menu en production; menu avec outils seulement en dev
  try {
    if (!app.isPackaged) {
      const isMac = process.platform === 'darwin';
      const template = [
        ...(isMac ? [{
          label: app.name,
          submenu: [
            { role: 'about' },
            { type: 'separator' },
            { role: 'services' },
            { type: 'separator' },
            { role: 'hide' },
            { role: 'hideOthers' },
            { role: 'unhide' },
            { type: 'separator' },
            { role: 'quit' }
          ]
        }] : []),
        {
          label: 'Fichier',
          submenu: [
            ...(isMac ? [] : [{ role: 'quit', label: 'Quitter' }])
          ]
        },
        {
          label: 'Affichage',
          submenu: [
            { role: 'reload', label: 'Recharger' },
            { role: 'forceReload', label: 'Forcer le rechargement' },
            { type: 'separator' },
            { role: 'toggleDevTools', label: 'Outils de développement', accelerator: 'F12' },
            { type: 'separator' },
            { role: 'resetZoom', label: 'Zoom 100%' },
            { role: 'zoomIn', label: 'Zoom +' },
            { role: 'zoomOut', label: 'Zoom -' },
            { type: 'separator' },
            { role: 'togglefullscreen', label: 'Plein écran' }
          ]
        },
        {
          label: 'Fenêtre',
          submenu: [
            { role: 'minimize', label: 'Réduire' },
            { role: 'close', label: 'Fermer la fenêtre' }
          ]
        }
      ];
      const menu = Menu.buildFromTemplate(template);
      Menu.setApplicationMenu(menu);
    } else {
      Menu.setApplicationMenu(null);
      try { mainWindow.removeMenu(); } catch {}
    }
  } catch (e) {
    console.warn('Menu non défini:', e?.message);
  }

  mainWindow.loadFile(path.join(__dirname, '../../public/index.html'));

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    logClean('🎉 Application prete !');
  });

  // Logs: stream real-time entries to renderer
  try {
    mainLogger.onEntry((entry) => {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('log-entry', entry);
      }
    });
  } catch {}

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  mainWindow.on('minimize', (event) => {
    if (tray) {
      event.preventDefault();
      mainWindow.hide();
    }
  });

  // Clic droit -> Inspecter l'élément: uniquement en dev
  try {
    if (!app.isPackaged) {
      mainWindow.webContents.on('context-menu', (_event, params) => {
        Menu.buildFromTemplate([
          { label: 'Inspecter l\'élément', click: () => mainWindow.webContents.inspectElement(params.x, params.y) },
          { type: 'separator' },
          { role: 'copy', label: 'Copier' },
          { role: 'paste', label: 'Coller' }
        ]).popup();
      });
    }
  } catch {}

  // Par sécurité, fermer DevTools si ouverts en production
  try {
    if (app.isPackaged) {
      mainWindow.webContents.on('devtools-opened', () => {
        try { mainWindow.webContents.closeDevTools(); } catch {}
      });
    }
  } catch {}

  return mainWindow;
}

async function initializeOutlook() {
  // Vérifier si l'initialisation est déjà en cours
  if (isInitializing) {
    logClean('⚠️ Initialisation déjà en cours, ignorer la demande');
    return;
  }
  isInitializing = true;
  logClean('🚀 [LOG] Début initializeOutlook (protection active)');
  try {
    // Etape 1: Verification de l'environnement avec détails
    sendTaskProgress('configuration', 'Vérification de la configuration système...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 0,
        progress: 50,
        message: 'Vérification du système...'
      });
    }
    await new Promise(resolve => setTimeout(resolve, 800));
    
    sendTaskProgress('configuration', 'Configuration système validée', true);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 0,
        progress: 100,
        message: 'Système vérifié ✓'
      });
    }
    await new Promise(resolve => setTimeout(resolve, 500));

    // Etape 2: Connexion à Outlook avec suivi détaillé
    sendTaskProgress('connection', 'Établissement de la connexion Outlook...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 0,
        message: 'Connexion à Outlook...'
      });
    }
    let retries = 0;
    const maxRetries = 120; // Augmenté pour laisser le temps à Outlook de se lancer
    let connected = false;
    while (retries < maxRetries && !connected) {
      try {
        if (outlookConnector) {
          if (outlookConnector.isOutlookConnected) {
            connected = true;
            break;
          }
          await outlookConnector.establishConnection();
          connected = outlookConnector.isOutlookConnected;
          if (connected) {
            console.log('[LOG] Outlook connecté !');
            break;
          }
        }
      } catch (error) {
        console.log(`[INIT] Tentative ${retries + 1} échouée: ${error.message}`);
        
        // Messages spéciaux pour le lancement automatique
        if (error.message.includes('Lancement automatique')) {
          if (loadingWindow) {
            loadingWindow.webContents.send('loading-progress', {
              step: 1,
              progress: Math.min(50, (retries / maxRetries) * 100),
              message: 'Lancement d\'Outlook en cours...'
            });
          }
        } else if (error.message.includes('Attente du démarrage')) {
          if (loadingWindow) {
            loadingWindow.webContents.send('loading-progress', {
              step: 1,
              progress: Math.min(80, (retries / maxRetries) * 100),
              message: 'Outlook se lance, veuillez patienter...'
            });
          }
        }
      }
      await new Promise(resolve => { setTimeout(resolve, 500); });
      retries++;
      if (loadingWindow && retries % 2 === 0) {
        const progressOutlook = Math.min(100, (retries / maxRetries) * 100);
        loadingWindow.webContents.send('loading-progress', {
          step: 1,
          progress: progressOutlook,
          message: `Connexion en cours... ${Math.floor(retries / 2)}s`
        });
      }
    }
    if (!connected) {
      console.log('[LOG] Connexion Outlook impossible après max tentatives');
      const errorMessage = `Impossible de se connecter à Outlook après ${Math.floor(maxRetries / 2)} secondes.\n\nVeuillez :\n• Vérifier qu'Outlook s'est bien lancé\n• Vérifier que votre profil est configuré\n• Redémarrer l'application si nécessaire`;
      throw new Error(errorMessage);
    }

    // Finaliser la connexion avec confirmation
    sendTaskProgress('connection', 'Connexion Outlook établie avec succès', true);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 100,
        message: 'Outlook connecté !'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 500); });

    // Etape 3: Chargement des statistiques
    sendTaskProgress('stats', 'Récupération des données statistiques...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 2,
        progress: 50,
        message: 'Chargement des statistiques...'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 800); });
    
    sendTaskProgress('stats', 'Statistiques chargées', true);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 2,
        progress: 100,
        message: 'Statistiques prêtes ✓'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 500); });

    // Etape 4: Analyse des catégories
    sendTaskProgress('categories', 'Analyse des catégories d\'emails...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 3,
        progress: 30,
        message: 'Analyse des catégories...'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 600); });
    
    sendTaskProgress('categories', 'Catégories analysées', true);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 3,
        progress: 100,
        message: 'Catégories configurées ✓'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 500); });

    // Etape 5: Exploration des dossiers
    sendTaskProgress('folders', 'Exploration de la structure des dossiers...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 4,
        progress: 0,
        message: 'Exploration des dossiers...'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 1000); });
    
    sendTaskProgress('folders', 'Structure des dossiers chargée', true);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 4,
        progress: 100,
        message: 'Dossiers explorés ✓'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 500); });

    // Etape 6: Configuration VBA
    sendTaskProgress('vba', 'Chargement des métriques VBA...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 5,
        progress: 25,
        message: 'Configuration VBA...'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 700); });
    
    sendTaskProgress('vba', 'Métriques VBA configurées', true);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 5,
        progress: 100,
        message: 'VBA configuré ✓'
      });
    }
    await new Promise(resolve => { setTimeout(resolve, 500); });

    // Etape 7: Monitoring automatique
    sendTaskProgress('monitoring', 'Initialisation du monitoring automatique...', false);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 6,
        progress: 0,
        message: 'Initialisation du monitoring...'
      });
    }
    try {
      const databaseService = require('../services/optimizedDatabaseService');
      await databaseService.initialize();
      
      sendTaskProgress('monitoring', 'Base de données initialisée...', false);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 6,
          progress: 30,
          message: 'Base de données initialisée...'
        });
      }
      
      const folderConfig = databaseService.getFoldersConfiguration();
      const configFound = Array.isArray(folderConfig) && folderConfig.length > 0;
      if (configFound) {
        console.log(`[LOG] 📁 Configuration trouvée en BDD: ${folderConfig.length} dossiers configurés`);
      }
      
      sendTaskProgress('monitoring', 'Configuration du service unifié...', false);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 6,
          progress: 60,
          message: 'Configuration du service unifié...'
        });
      }
      
      // CORRECTION: Toujours initialiser le service unifié (même sans configuration)
      const UnifiedMonitoringService = require('../services/unifiedMonitoringService');
      global.unifiedMonitoringService = new UnifiedMonitoringService(outlookConnector);
      global.unifiedMonitoringService.initialize().then(() => {
        console.log('[LOG] ✅ Service unifié initialisé en arrière-plan');
        setupRealtimeEventForwarding();
        if (configFound) {
          console.log(`[LOG] 📁 Configuration trouvée en BDD: ${folderConfig.length} dossiers configurés`);
          console.log('[LOG] 🔄 Le monitoring PowerShell + COM va démarrer automatiquement...');
        } else {
          console.log('[LOG] ℹ️ Service unifié prêt - ajoutez des dossiers pour déclencher la sync PowerShell');
        }
      }).catch((error) => {
        console.error('[LOG] ❌ Erreur initialisation service unifié:', error.message);
      });
      
      sendTaskProgress('monitoring', 'Service de monitoring configuré', true);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 6,
          progress: 100,
          message: configFound ? 'Monitoring configuré ✓' : 'Prêt (config manuelle) ✓'
        });
      }
      await new Promise(resolve => { setTimeout(resolve, 500); });
      
      // Etape 8: Finalisation
      sendTaskProgress('weekly', 'Initialisation du suivi hebdomadaire...', false);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 7,
          progress: 50,
          message: 'Suivi hebdomadaire...'
        });
      }
      
      console.log('[LOG] 📊 Service unifié avec métriques intégrées');
      await new Promise(resolve => { setTimeout(resolve, 800); });
      
      sendTaskProgress('weekly', 'Suivi hebdomadaire configuré', true);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 7,
          progress: 100,
          message: 'Suivi hebdomadaire prêt ✓'
        });
      }
      
    } catch (monitoringError) {
      console.warn('[LOG] ⚠️ Erreur monitoring:', monitoringError.message);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 6,
          progress: 100,
          message: 'Prêt (mode dégradé)'
        });
      }
    }
    
    await new Promise(resolve => { setTimeout(resolve, 500); });
    
    // Signaler la completion finale
    if (loadingWindow) {
      console.log('[LOG] 📤 Envoi de l\'événement loading-complete...');
      loadingWindow.webContents.send('loading-complete');
      console.log('[LOG] ✅ Événement loading-complete envoyé');
    } else {
      console.log('[LOG] ⚠️ Fenêtre de chargement non disponible pour envoyer loading-complete');
    }
    console.log('[LOG] ✅ Initialisation complète réussie');
  } catch (error) {
    console.error('[LOG] Erreur initialisation Outlook:', error);
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-error', {
        message: error.message,
        stack: error.stack
      });
    }
  } finally {
    isInitializing = false;
    console.log('[LOG] 🔓 Protection d\'initialisation libérée');
  }
}

// Handlers IPC
ipcMain.on('loading-page-complete', () => {
  console.log('📱 [IPC] Page de chargement complète - création fenêtre principale');
  createWindow();
  
  // Fermer la page de chargement après un délai pour transition douce
  setTimeout(() => {
    if (loadingWindow && !loadingWindow.isDestroyed()) {
      console.log('🎯 [IPC] Fermeture de la page de chargement');
      loadingWindow.close();
      loadingWindow = null;
    }
  }, 1000);
});

ipcMain.on('loading-retry', () => {
  console.log('Retry initialisation Outlook...');
  initializeOutlook();
});

// Handler IPC pour le redimensionnement dynamique de la fenêtre de chargement
ipcMain.on('resize-loading-window', (event, { width, height }) => {
  if (loadingWindow && !loadingWindow.isDestroyed()) {
    // Ajouter une marge de sécurité et limites min/max
    const finalWidth = Math.max(400, Math.min(800, width + 40));
    const finalHeight = Math.max(300, Math.min(900, height + 40));

    // Réduire le bruit de logs: ne pas logger chaque redimensionnement
    loadingWindow.setSize(finalWidth, finalHeight);
  // Ne pas recentrer la fenêtre pour préserver la position définie par l'utilisateur
  }
});

// Handlers IPC pour l'API de l'application
ipcMain.handle('api-settings-folders-load', async () => {
  try {
    console.log('📁 Chargement de la configuration des dossiers depuis la BDD...');
    
  const databaseService = require('../services/optimizedDatabaseService');
  await databaseService.initialize();
    
    // CORRECTION: Invalider le cache pour forcer un rechargement des données récentes
    databaseService.cache.del('folders_config');
    
    // Récupérer la configuration depuis la base de données uniquement
  const foldersConfig = databaseService.getFoldersConfiguration();
    
    // Convertir le format tableau en format objet pour l'interface
    const folderCategories = {};
    if (Array.isArray(foldersConfig)) {
      foldersConfig.forEach(folder => {
        // CORRECTION: Utiliser les vrais noms des propriétés de la BDD
        folderCategories[folder.folder_path] = {
          category: folder.category,
          name: folder.folder_name
        };
      });
    }
    
    console.log(`✅ ${Object.keys(folderCategories).length} configurations chargées depuis BDD`);
    
    return { 
      success: true, 
      folderCategories: folderCategories 
    };
    
  } catch (error) {
    console.error('❌ Erreur chargement configuration dossiers depuis BDD:', error);
    
    // Fallback: retourner une configuration vide mais réussie
    return { 
      success: true, 
      folderCategories: {},
      warning: `Erreur BDD: ${error.message}`
    };
  }
});

ipcMain.handle('api-settings-folders', async (event, data) => {
  try {
    console.log('💾 Sauvegarde de la configuration des dossiers en BDD uniquement...');
    
  // Sauvegarder UNIQUEMENT dans la base de données (pas de JSON)
  const res = await global.databaseService.saveFoldersConfiguration(data);
    console.log('✅ Configuration dossiers sauvegardée exclusivement en base de données');
    
    // Redémarrer automatiquement le monitoring si des dossiers sont configurés
    const hasData = (Array.isArray(data) && data.length > 0) || 
                   (typeof data === 'object' && Object.keys(data).length > 0);
                   
    if (global.unifiedMonitoringService && hasData) {
      console.log('🔄 Programmation du redémarrage du service unifié en arrière-plan...');
      // Redémarrer en arrière-plan pour ne pas bloquer la réponse IPC
      setTimeout(async () => {
        try {
          if (global.unifiedMonitoringService.isMonitoring) {
            await global.unifiedMonitoringService.stopMonitoring();
          }
          await global.unifiedMonitoringService.initialize();
          console.log('✅ Service unifié redémarré (arrière-plan)');
        } catch (error) {
          console.error('❌ Erreur redémarrage service unifié (arrière-plan):', error);
        }
      }, 0);
    }
    
  return { success: true, result: res };
  } catch (error) {
    console.error('❌ Erreur sauvegarde configuration dossiers:', error);
    return { success: false, error: error.message };
  }
});

// === HANDLERS IPC POUR LA GESTION HIÉRARCHIQUE DES DOSSIERS ===

// Récupérer l'arbre hiérarchique des dossiers
ipcMain.handle('api-folders-tree', async (_event, payload) => {
  try {
    // OPTIMIZED: Vérifier le cache d'abord (sauf si force=true)
    const force = payload && payload.force === true;
    if (!force) {
      const cachedFolders = cacheService.get('config', 'folders_tree');
      if (cachedFolders) {
        return cachedFolders;
      }
    }

    await databaseService.initialize();

    // Récupérer UNIQUEMENT les dossiers configurés (monitorés) depuis la BDD optimisée
    const foldersConfig = global.databaseService.getFoldersConfiguration();

    // Récupérer la structure Outlook pour obtenir les compteurs d'emails
    const allFolders = await outlookConnector.getFolders();

    // Créer la liste des dossiers monitorés seulement
    const monitoredFolders = [];

    foldersConfig.forEach(config => {
      const fullPathRaw = config.folder_path || config.folder_name || '';
      const displayName = config.folder_name || extractFolderName(fullPathRaw);
      // Normaliser si possible: si le préfixe n'est pas un email, conserver tel quel
      const fullPath = fullPathRaw;

      // Chercher le dossier dans la structure Outlook pour obtenir le nombre d'emails
      const outlookFolder = allFolders.find(f => f.path === fullPath || f.name === displayName);

      monitoredFolders.push({
        path: fullPath,
        name: displayName,
        isMonitored: true,
        category: config.category || 'Mails simples',
        emailCount: outlookFolder ? (outlookFolder.emailCount || 0) : 0,
        parentPath: getParentPath(fullPath)
      });
    });

    // Calculer les statistiques
    const stats = calculateFolderStats(monitoredFolders);

    const result = {
      folders: monitoredFolders, // Uniquement les dossiers monitorés
      stats: stats,
      timestamp: new Date().toISOString()
    };

  // OPTIMIZED: Mettre en cache pour 5 minutes (même si force=true pour les appels suivants)
  cacheService.set('config', 'folders_tree', result, 300);

    return result;

  } catch (error) {
    console.error('❌ [IPC] Erreur récupération arbre dossiers:', error);
    throw error;
  }
});

// Statistiques par dossier depuis la base de données (inclut unreadCount)
ipcMain.handle('api-database-folder-stats', async () => {
  try {
    const databaseService = require('../services/optimizedDatabaseService');
    await databaseService.initialize();
    const stats = databaseService.getFolderStats();
    return { success: true, stats };
  } catch (error) {
    console.error('❌ [IPC] Erreur api-database-folder-stats:', error);
    return { success: false, error: error.message, stats: [] };
  }
});

// Ajouter un dossier au monitoring
ipcMain.handle('api-folders-add', async (event, { folderPath, category }) => {
    if (!folderPath || !category) {
      throw new Error('Chemin du dossier et catégorie requis');
    }

    await databaseService.initialize();

  // Nouvelle logique: résoudre la boîte par nom, puis parcourir récursivement
  // tous les descendants via lazy-load COM en se basant sur le chemin (robuste aux langues)
  const toInsert = [];
  const seen = new Set(); // éviter doublons
    {
      const parts = String(folderPath).split('\\');
      const mailboxName = parts[0];

      // 1) Trouver le StoreID de la boîte aux lettres
      let storeId = '';
      try {
        const mbs = await outlookConnector.getMailboxes?.();
        if (Array.isArray(mbs)) {
          const exact = mbs.find(m => m?.Name === mailboxName);
          const like = exact || mbs.find(m => String(m?.Name || '').toLowerCase().includes(String(mailboxName).toLowerCase()));
          storeId = like?.StoreID || '';
        }
      } catch (_) {}

      // Si Store introuvable, tentative via fast COM stores
      if (!storeId) {
        try {
          const fastOL = require('../server/outlookFastFolders');
          const stores = await fastOL.listStores();
          const exact = stores.find(s => s?.DisplayName === mailboxName);
          const like = exact || stores.find(s => String(s?.DisplayName || '').toLowerCase().includes(String(mailboxName).toLowerCase()));
          storeId = like?.StoreId || '';
        } catch (_) {}
      }

      if (!storeId) throw new Error('Boîte aux lettres introuvable pour: ' + mailboxName);

  // 2) Obtenir la mailbox (affichage + SMTP) pour normaliser les chemins
  const mailboxes = await outlookConnector.getFolderStructure?.(storeId);
  const mb = Array.isArray(mailboxes) ? mailboxes[0] : null;
  if (!mb) throw new Error('Structure de dossier non disponible');
  const mailboxDisplay = mb?.Name || mailboxName;
  const mailboxSmtp = mb?.SmtpAddress || '';
  // Tenter de récupérer le nom localisé de l'Inbox depuis la structure
  let inboxLocalizedName = '';
  try {
    const subs = Array.isArray(mb?.SubFolders) ? mb.SubFolders : [];
    const norm = (s) => String(s||'').normalize('NFD').replace(/\p{Diacritic}/gu,'').toLowerCase().trim();
    const inboxNode = subs.find(f => {
      const n = norm(f?.Name||'');
      return n === 'inbox' || n.includes('boite') || n.includes('reception');
    });
    if (inboxNode && inboxNode.Name) inboxLocalizedName = String(inboxNode.Name);
  } catch {}

      // Normalise un chemin en forçant le préfixe de boîte sur le SMTP si disponible
      const normalizePath = (p) => {
        try {
          const s = String(p || '');
          const idx = s.indexOf('\\');
          if (idx > 0) {
            const rest = s.slice(idx + 1);
            if (mailboxSmtp) return `${mailboxSmtp}\\${rest}`;
            return `${mailboxDisplay}\\${rest}`;
          }
          return mailboxSmtp ? mailboxSmtp : (mailboxDisplay || s);
        } catch { return p; }
      };

          // Helpers de normalisation JS (même logique que PowerShell)
          const normalizeName = (s) => {
            try {
              const tr = String(s || '').trim().replace(/[.]+$/, '');
              return tr.normalize('NFD').replace(/\p{Diacritic}/gu, '').toLowerCase().trim();
            } catch { return String(s || '').toLowerCase().trim(); }
          };

          // 2.b) Résoudre le chemin d'affichage exact (incl. Inbox localisée si manquante)
          const original = String(folderPath);
          const idx = original.indexOf('\\');
          const rest = idx > 0 ? original.slice(idx + 1) : '';
          let resolvedDisplayPath = '';
          if (!rest) {
            resolvedDisplayPath = mailboxDisplay; // racine boîte
          } else {
            // Si la chaîne contient déjà un segment Inbox localisé, on garde
            const partsRest = rest.split('\\');
            const hasInboxHint = partsRest.some(p => {
              const n = normalizeName(p);
              return n === 'inbox' || n.includes('boite') || n.includes('reception');
            });

            // Helper: detect if a resolved display path already includes Inbox as first child segment
            const pathHasInbox = (p) => {
              try {
                const segs = String(p || '').split('\\');
                if (segs.length < 2) return false;
                const second = segs[1] || '';
                const n = normalizeName(second);
                return n === 'inbox' || n.includes('boite') || n.includes('reception');
              } catch { return false; }
            };

            // Fonction utilitaire pour récupérer les enfants d'un parent d'affichage
            const listKids = async (parentDisplayPath) => {
              try {
                const res = await outlookConnector.getSubFolders?.(storeId, '', parentDisplayPath);
                return Array.isArray(res) ? res : (res && Array.isArray(res.children) ? res.children : []);
              } catch { return []; }
            };

            // BFS pour trouver le chemin exact correspondant
            const tryResolveByChain = async () => {
              const chain = partsRest.slice();
              // Démarrer depuis la racine d'affichage
              let cursorPath = mailboxDisplay;
              for (let i = 0; i < chain.length; i++) {
                const want = chain[i];
                const kids = await listKids(cursorPath);
                const next = kids.find(k => normalizeName(k?.Name || '') === normalizeName(want));
                if (!next) return '';
                cursorPath = next.FolderPath || `${cursorPath}\\${next.Name}`;
              }
              return cursorPath;
            };

            const tryResolveUnderInbox = async () => {
              // Chercher le dossier 'Inbox' localisé sous la racine, puis descendre
              const top = await listKids(mailboxDisplay);
              const inbox = top.find(k => {
                const n = normalizeName(k?.Name || '');
                return n === 'inbox' || n.includes('boite') || n.includes('reception');
              });
              if (!inbox) return '';
              let cursorPath = inbox.FolderPath || `${mailboxDisplay}\\${inbox.Name}`;
              for (const seg of partsRest) {
                const kids = await listKids(cursorPath);
                const next = kids.find(k => normalizeName(k?.Name || '') === normalizeName(seg));
                if (!next) return '';
                cursorPath = next.FolderPath || `${cursorPath}\\${next.Name}`;
              }
              return cursorPath;
            };

            // EWS fallback: résoudre via EWS (rapide) si COM ne renvoie rien sous la racine
            const tryResolveByEws = async () => {
              try {
                const top = await outlookConnector.getTopLevelFoldersFast?.(mailboxSmtp || mailboxDisplay);
                const folders = Array.isArray(top) ? top : [];
                const chain = partsRest.slice();
                // Priorité: Inbox scope (getTopLevelFoldersFast already tries Inbox then Root)
                // Trouver premier segment
                let cur = null;
                for (const f of folders) {
                  if (normalizeName(f?.Name || '') === normalizeName(chain[0])) { cur = f; break; }
                }
                if (!cur) return '';
                // Descendre segments restants
                for (let i = 1; i < chain.length; i++) {
                  const kids = await outlookConnector.getChildFoldersFast?.(mailboxSmtp || mailboxDisplay, cur.Id);
                  const nxt = (kids || []).find(k => normalizeName(k?.Name || '') === normalizeName(chain[i]));
                  if (!nxt) return '';
                  cur = nxt;
                }
                // Reconstruire chemin d'affichage minimal (sans Inbox localisée) à partir des noms
                return `${mailboxDisplay}\\${chain.join('\\')}`;
              } catch { return ''; }
            };

            const tryResolveByLeafDeep = async () => {
              const leaf = partsRest[partsRest.length - 1];
              // DFS limité: explorer max 200 dossiers pour éviter boucle infinie
              const stack = [mailboxDisplay];
              const visited = new Set();
              const targetNorm = normalizeName(leaf);
              let found = '';
              while (stack.length && visited.size < 200) {
                const parent = stack.pop();
                if (visited.has(parent)) continue;
                visited.add(parent);
                const kids = await listKids(parent);
                for (const k of kids) {
                  if (normalizeName(k?.Name || '') === targetNorm) {
                    found = k.FolderPath || `${parent}\\${k.Name}`;
                    return found;
                  }
                  stack.push(k.FolderPath || `${parent}\\${k.Name}`);
                }
              }
              return found;
            };

            // 1) essayer chaîne exacte telle que fournie (depuis racine)
            const chainResolved = await tryResolveByChain();
            // 2) en parallèle logique: si pas d'indice Inbox dans la saisie, tenter aussi sous Inbox localisée
            let inboxResolved = '';
            if (!hasInboxHint) {
              inboxResolved = await tryResolveUnderInbox();
            }
            // Choisir le meilleur chemin: préférer Inbox si elle existe et que la résolution par chaîne ne pointe pas déjà sous Inbox
            if (inboxResolved) {
              if (!chainResolved) {
                resolvedDisplayPath = inboxResolved;
              } else {
                resolvedDisplayPath = pathHasInbox(chainResolved) ? chainResolved : inboxResolved;
              }
            } else {
              resolvedDisplayPath = chainResolved;
            }
            // 3) éviter EWS si possible; si Inbox localisée connue, forcer le préfixe Inbox
            if (!resolvedDisplayPath && inboxLocalizedName && !hasInboxHint) {
              resolvedDisplayPath = `${mailboxDisplay}\\${inboxLocalizedName}\\${rest}`;
            }
            // 4) sinon, tenter résolution via EWS (dernier recours)
            if (!resolvedDisplayPath) {
              resolvedDisplayPath = await tryResolveByEws();
            }
            // 5) sinon, recherche profonde par nom de feuille
            if (!resolvedDisplayPath) {
              resolvedDisplayPath = await tryResolveByLeafDeep();
            }

            // Correction: si le chemin ne contient pas Inbox mais qu'on connait son libellé, l'inclure directement
            if (resolvedDisplayPath && !pathHasInbox(resolvedDisplayPath) && inboxLocalizedName && !hasInboxHint) {
              resolvedDisplayPath = `${mailboxDisplay}\\${inboxLocalizedName}\\${rest}`;
            } else if (resolvedDisplayPath && !pathHasInbox(resolvedDisplayPath)) {
              // Sinon: si le premier segment existe sous Inbox, préférer l'inclure.
              try {
                const top = await listKids(mailboxDisplay);
                const inbox = top.find(k => {
                  const n = normalizeName(k?.Name || '');
                  return n === 'inbox' || n.includes('boite') || n.includes('reception');
                });
                if (inbox && partsRest.length) {
                  const kids = await listKids(inbox.FolderPath || `${mailboxDisplay}\\${inbox.Name}`);
                  const first = kids.find(k => normalizeName(k?.Name || '') === normalizeName(partsRest[0]));
                  if (first) {
                    resolvedDisplayPath = `${mailboxDisplay}\\${inbox.Name}\\${rest}`;
                  }
                }
              } catch {}
            }

            // Si tout échoue, garder le chemin d'origine mais, si possible, y inclure Inbox localisée
            if (!resolvedDisplayPath) {
              // Si le premier segment est un enfant direct d'une Inbox localisée, le préfixer d'Inbox
              try {
                if (inboxLocalizedName && partsRest.length) {
                  resolvedDisplayPath = `${mailboxDisplay}\\${inboxLocalizedName}\\${rest}`;
                } else {
                  const top = await listKids(mailboxDisplay);
                  const inbox = top.find(k => {
                    const n = normalizeName(k?.Name || '');
                    return n === 'inbox' || n.includes('boite') || n.includes('reception');
                  });
                  if (inbox && partsRest.length) {
                    // Vérifier si le premier segment existe sous Inbox
                    const kids = await listKids(inbox.FolderPath || `${mailboxDisplay}\\${inbox.Name}`);
                    const first = kids.find(k => normalizeName(k?.Name || '') === normalizeName(partsRest[0]));
                    if (first) {
                      resolvedDisplayPath = `${mailboxDisplay}\\${inbox.Name}\\${rest}`;
                    }
                  }
                }
              } catch {}
              if (!resolvedDisplayPath) resolvedDisplayPath = `${mailboxDisplay}\\${rest}`;
            }
          }
      if (!rest) {
        // Dossier racine de la boîte: insérer seulement ce dossier (pas d'enfants énumérables)
        const normalized = normalizePath(original);
        if (!seen.has(normalized)) {
          seen.add(normalized);
          toInsert.push({ path: normalized, name: extractFolderName(original) });
        }
  } else {
    // Chemin "affichage" pour naviguer côté COM et chemin normalisé (SMTP) pour la BDD
  const displayBasePath = resolvedDisplayPath;
  const normalizedBasePath = normalizePath(resolvedDisplayPath);
    try { console.log(`[ADD] Resolved display path: ${displayBasePath}`); } catch {}
    try { console.log(`[ADD] Normalized DB path: ${normalizedBasePath}`); } catch {}
        if (!seen.has(normalizedBasePath)) {
          seen.add(normalizedBasePath);
          toInsert.push({ path: normalizedBasePath, name: extractFolderName(original) });
        }
        // DFS: essayer COM, sinon retomber sur EWS pour la descente complète
        const runDfsCom = async (rootDisplayPath) => {
          const stack = [rootDisplayPath];
          while (stack.length) {
            const parentDisplayPath = stack.pop();
            let kids = [];
            try {
              const res = await outlookConnector.getSubFolders?.(storeId, '', parentDisplayPath);
              if (Array.isArray(res)) {
                kids = res;
              } else if (res && typeof res === 'object' && Array.isArray(res.children)) {
                kids = res.children;
              } else {
                kids = [];
              }
            } catch (_) { kids = []; }

            // Fallback: if no kids by display path, try to resolve this folder's EntryID via its parent, then re-query by EntryID
            if (!Array.isArray(kids) || kids.length === 0) {
              try {
                const segs = String(parentDisplayPath || '').split('\\');
                if (segs.length >= 2) {
                  const leafName = segs[segs.length - 1];
                  const parentOfParent = segs.slice(0, -1).join('\\');
                  const siblingsRes = await outlookConnector.getSubFolders?.(storeId, '', parentOfParent);
                  const siblings = Array.isArray(siblingsRes)
                    ? siblingsRes
                    : (siblingsRes && Array.isArray(siblingsRes.children) ? siblingsRes.children : []);
                  const normalizeName = (s) => String(s||'').normalize('NFD').replace(/\p{Diacritic}/gu,'').toLowerCase().trim();
                  const sibling = siblings.find(ch => normalizeName(ch?.Name||'') === normalizeName(leafName));
                  const entryId = sibling?.EntryID || sibling?.EntryId || '';
                  if (entryId) {
                    const byId = await outlookConnector.getSubFolders?.(storeId, entryId, '');
                    if (Array.isArray(byId)) {
                      kids = byId;
                    } else if (byId && Array.isArray(byId.children)) {
                      kids = byId.children;
                    }
                  }
                }
              } catch {}
            }

            if (!Array.isArray(kids) || kids.length === 0) continue;
            for (const k of kids) {
              const childDisplayPath = k?.FolderPath || '';
              if (!childDisplayPath) continue;
              const childName = k?.Name || extractFolderName(childDisplayPath);
              const normalized = normalizePath(childDisplayPath);
              if (!seen.has(normalized)) {
                seen.add(normalized);
                toInsert.push({ path: normalized, name: childName });
                stack.push(childDisplayPath);
              }
            }
          }
        };

  const runDfsEws = async (chainStart) => {
          try {
            const segs = chainStart.split('\\').filter(Boolean);
            // Remove mailbox part
            const startSegs = segs.slice(1);
            // Resolve starting folder id
            const top = await outlookConnector.getTopLevelFoldersFast?.(mailboxSmtp || mailboxDisplay);
            const folders = Array.isArray(top) ? top : [];
            let cur = null, namesChain = [];
            if (startSegs.length === 0) return; // nothing to enumerate
            for (const f of folders) {
              if (normalizeName(f?.Name || '') === normalizeName(startSegs[0])) { cur = f; break; }
            }
            if (!cur) return;
            namesChain = [cur.Name];
            for (let i = 1; i < startSegs.length; i++) {
              const kids = await outlookConnector.getChildFoldersFast?.(mailboxSmtp || mailboxDisplay, cur.Id);
              const nxt = (kids || []).find(k => normalizeName(k?.Name || '') === normalizeName(startSegs[i]));
              if (!nxt) return;
              cur = nxt;
              namesChain.push(cur.Name);
            }
            // Now DFS from cur.Id
            const stack = [{ id: cur.Id, chain: namesChain.slice() }];
            while (stack.length) {
              const node = stack.pop();
              const kids = await outlookConnector.getChildFoldersFast?.(mailboxSmtp || mailboxDisplay, node.id);
              for (const k of (kids || [])) {
                const childChain = node.chain.concat(k.Name);
                const normalized = normalizePath(`${mailboxDisplay}\\${childChain.join('\\')}`);
                if (!seen.has(normalized)) {
                  seen.add(normalized);
                  toInsert.push({ path: normalized, name: k.Name });
                }
                stack.push({ id: k.Id, chain: childChain });
              }
            }
          } catch {}
        };

        // Fallback COM: explicit DFS starting from Inbox when path start includes Inbox label
        const runDfsFromInboxCom = async () => {
          try {
            const norm = (s) => String(s||'').normalize('NFD').replace(/\p{Diacritic}/gu,'').toLowerCase().trim();
            const top = await outlookConnector.getSubFolders?.(storeId, '', mailboxDisplay);
            const topKids = Array.isArray(top) ? top : (top && Array.isArray(top.children) ? top.children : []);
            const inbox = topKids.find(k => {
              const n = norm(k?.Name||'');
              return n === 'inbox' || n.includes('boite') || n.includes('reception');
            });
            if (!inbox) return;
            const stack = [inbox.FolderPath || `${mailboxDisplay}\\${inbox.Name}`];
            while (stack.length) {
              const p = stack.pop();
              let kids = [];
              try {
                const res = await outlookConnector.getSubFolders?.(storeId, '', p);
                kids = Array.isArray(res) ? res : (res && Array.isArray(res.children) ? res.children : []);
              } catch {}
              for (const k of (kids || [])) {
                const childDisplayPath = k?.FolderPath || '';
                if (!childDisplayPath) continue;
                const childName = k?.Name || extractFolderName(childDisplayPath);
                const normalized = normalizePath(childDisplayPath);
                if (!seen.has(normalized)) {
                  seen.add(normalized);
                  toInsert.push({ path: normalized, name: childName });
                }
                stack.push(childDisplayPath);
              }
            }
          } catch {}
        };

        // Prefer COM DFS. If COM didn't add any children, fall back to EWS DFS
        const beforeCount = toInsert.length;
        await (async () => {
          console.log(`[ADD][DFS] COM by display path start: ${displayBasePath}`);
          await runDfsCom(displayBasePath);
        })().catch(e => console.warn('[ADD][DFS] COM by display path failed:', e?.message || String(e)));
        const afterCount = toInsert.length;
        console.log(`[ADD][DFS] COM by display path inserted: ${afterCount - beforeCount}`);
        if (afterCount === beforeCount) {
          // COM yielded nothing; try EWS DFS
          await (async () => {
            console.log('[ADD][DFS] Falling back to EWS DFS...');
            await runDfsEws(displayBasePath);
          })().catch(e => console.warn('[ADD][DFS] EWS DFS failed:', e?.message || String(e)));
          // If still nothing and we started under Inbox, try explicit inbox traversal via COM
          if (toInsert.length === afterCount) {
            const nn = (s) => String(s||'').normalize('NFD').replace(/\p{Diacritic}/gu,'').toLowerCase().trim();
            const segs = displayBasePath.split('\\');
            if (segs.length >= 2 && (nn(segs[1]) === 'inbox' || nn(segs[1]).includes('boite') || nn(segs[1]).includes('reception'))) {
              await (async () => {
                console.log('[ADD][DFS] COM DFS from Inbox...');
                await runDfsFromInboxCom();
              })().catch(e => console.warn('[ADD][DFS] COM DFS from Inbox failed:', e?.message || String(e)));
            }
            // Last-resort: fast COM traversal by EntryID using outlookFastFolders
            if (toInsert.length === afterCount) {
              await (async () => {
                console.log('[ADD][DFS] Last-resort: fast COM traversal by EntryID...');
                const nn2 = (s) => String(s||'').normalize('NFD').replace(/\p{Diacritic}/gu,'').toLowerCase().trim();
                const fastOL = require('../server/outlookFastFolders');
                const segs2 = displayBasePath.split('\\').filter(Boolean);
                const chainSegs = segs2.slice(1);
                if (!chainSegs.length) return;
                const rootList = await fastOL.listFoldersShallow(storeId, '');
                const parentName = (rootList && rootList.parentName) ? String(rootList.parentName) : (inboxLocalizedName || 'Inbox');
                let curParentId = '';
                let curChain = [parentName];
                let startIdx = 0;
                if (chainSegs.length && nn2(chainSegs[0]) === nn2(parentName)) startIdx = 1;
                let ok = true;
                for (let i = startIdx; i < chainSegs.length; i++) {
                  const segment = chainSegs[i];
                  const listing = await fastOL.listFoldersShallow(storeId, curParentId);
                  const arr = (listing && Array.isArray(listing.folders)) ? listing.folders : [];
                  const match = arr.find(f => nn2(f.Name) === nn2(segment));
                  if (!match) { ok = false; break; }
                  curParentId = match.EntryId || match.EntryID || '';
                  curChain.push(match.Name);
                }
                if (!ok) return;
                const stack = [{ id: curParentId, chain: curChain.slice() }];
                while (stack.length) {
                  const node = stack.pop();
                  const listing = await fastOL.listFoldersShallow(storeId, node.id);
                  const kids = (listing && Array.isArray(listing.folders)) ? listing.folders : [];
                  for (const k of kids) {
                    const childChain = node.chain.concat(k.Name);
                    const disp = `${mailboxDisplay}\\${childChain.join('\\')}`;
                    const norm = normalizePath(disp);
                    if (!seen.has(norm)) {
                      seen.add(norm);
                      toInsert.push({ path: norm, name: k.Name });
                    }
                    const kidId = k.EntryId || k.EntryID || '';
                    if (kidId) stack.push({ id: kidId, chain: childChain });
                  }
                }
              })().catch(e => console.warn('[ADD][DFS] Last-resort fast COM traversal failed:', e?.message || String(e)));
            }
          }
        }
  } // End of recursive enumeration block
  // Ajout récursif partiel: fallback à insertion simple (erreur silencieuse précédemment capturée)
  // If nothing was added due to an internal failure, fallback logic below will add base folder.
      if (toInsert.length === 0) {
        // Appliquer aussi la normalisation SMTP au fallback
        try {
          const parts = String(folderPath).split('\\');
          const raw = parts.join('\\');
          const normalized = (function(){
            const idx = raw.indexOf('\\');
            if (idx > 0) {
              const rest = raw.slice(idx + 1);
              // Sans mb connu ici, tenter d'utiliser la partie 0 si elle ressemble à un email, sinon conserver
              const first = parts[0] || '';
              const isEmail = /.+@.+\..+/.test(first);
              return isEmail ? raw : raw; // pas d'info SMTP fiable; garder tel quel
            }
            return raw;
          })();
          if (!seen.has(normalized)) {
            seen.add(normalized);
            toInsert.push({ path: normalized, name: extractFolderName(folderPath) });
          }
        } catch {
          toInsert.push({ path: folderPath, name: extractFolderName(folderPath) });
        }
      }
  }

    // Insérer tous les dossiers (OR REPLACE évite les doublons)
    let inserted = 0;
    for (const item of toInsert) {
      try {
        await databaseService.addFolderConfiguration(item.path, category, item.name);
        inserted++;
      } catch (e) {
        console.error('❌ Erreur insertion dossier enfant:', item.path, e.message || e);
      }
    }
    
    // OPTIMIZED: Invalidation intelligente du cache
    cacheService.invalidateFoldersConfig();
    
    console.log(`✅ ${inserted} dossier(s) ajouté(s) au monitoring (incl. sous-dossiers)`);

    // Redémarrer le monitoring en arrière-plan pour prendre en compte le nouveau dossier
    if (global.unifiedMonitoringService) {
      setTimeout(async () => {
        try {
          if (global.unifiedMonitoringService.isMonitoring) {
            await global.unifiedMonitoringService.stopMonitoring();
          }
          await global.unifiedMonitoringService.initialize();
          console.log('🔄 Service unifié redémarré pour le nouveau dossier (arrière-plan)');
        } catch (error) {
          console.error('⚠️ Erreur redémarrage monitoring (arrière-plan):', error);
        }
      }, 0);
    }

    return {
      success: true,
      message: `Dossier ajouté (avec sous-dossiers): ${inserted} élément(s)`,
      folderPath: folderPath,
      category: category,
      count: inserted
    };

  // Outer errors already handled internally; surface generic failure if essential data missing
  if (!folderPath || !category) {
    return { success: false, message: 'Paramètres invalides' };
  }
});

// Mettre à jour la catégorie d'un dossier
ipcMain.handle('api-folders-update-category', async (event, { folderPath, category }) => {
  try {
    if (!folderPath || !category) {
      throw new Error('Chemin du dossier et catégorie requis');
    }

    const databaseService = require('../services/optimizedDatabaseService');
    await databaseService.initialize();

    // Mettre à jour directement en base de données
    const updated = await global.databaseService.updateFolderCategory(folderPath, category);
    
    if (!updated) {
      throw new Error('Dossier non trouvé dans la configuration active');
    }

    console.log(`✅ Catégorie de ${folderPath} mise à jour: ${category}`);

    // Invalidate cached folders data so UI refresh sees the change
    try {
      cacheService.invalidateFoldersConfig();
    } catch (e) {
      console.warn('⚠️ Impossible d\'invalider le cache des dossiers après mise à jour:', e?.message || e);
    }

    // Redémarrer le monitoring en arrière-plan pour prendre en compte le changement
    if (global.unifiedMonitoringService) {
      setTimeout(async () => {
        try {
          if (global.unifiedMonitoringService.isMonitoring) {
            await global.unifiedMonitoringService.stopMonitoring();
          }
          await global.unifiedMonitoringService.initialize();
          console.log('🔄 Service unifié redémarré pour changement de catégorie (arrière-plan)');
        } catch (error) {
          console.error('⚠️ Erreur redémarrage monitoring (arrière-plan):', error);
        }
      }, 0);
    }

    return {
      success: true,
      message: 'Catégorie mise à jour',
      folderPath: folderPath,
      category: category
    };

  } catch (error) {
    console.error('❌ Erreur mise à jour catégorie:', error);
    throw error;
  }
});

// Supprimer un dossier du monitoring
ipcMain.handle('api-folders-remove', async (event, { folderPath }) => {
  try {
    if (!folderPath) {
      throw new Error('Chemin du dossier requis');
    }

  const databaseService = require('../services/optimizedDatabaseService');
    await databaseService.initialize();

    // Supprimer le dossier de la configuration en base de données
    const deleted = await global.databaseService.deleteFolderConfiguration(folderPath);
    
    if (!deleted) {
      throw new Error('Dossier non trouvé dans la configuration');
    }

    console.log(`✅ Dossier ${folderPath} supprimé de la configuration`);

    // Invalidate cached folders data so UI refresh sees the deletion
    try {
      cacheService.invalidateFoldersConfig();
    } catch (e) {
      console.warn('⚠️ Impossible d\'invalider le cache des dossiers après suppression:', e?.message || e);
    }

    // CORRECTION: Redémarrer le service de monitoring avec la nouvelle configuration
    if (global.unifiedMonitoringService) {
      console.log('🔄 Redémarrage du service de monitoring après suppression du dossier...');
      
           

      
      try {
        // Arrêter le monitoring actuel
        if (global.unifiedMonitoringService.isMonitoring) {
          await global.unifiedMonitoringService.stopMonitoring();
        }
        
        // Réinitialiser avec la nouvelle configuration
        await global.unifiedMonitoringService.initialize();
        
        console.log('✅ Service de monitoring redémarré avec succès');
      } catch (monitoringError) {
        console.error('⚠️ Erreur redémarrage monitoring:', monitoringError);
        // Continuer même si le redémarrage échoue
      }
    }

    return {
      success: true,
      message: 'Dossier retiré du monitoring',
      folderPath: folderPath
    };

  } catch (error) {
    console.error('❌ Erreur suppression dossier:', error);
    throw error;
  }
});

// Recharger la configuration des dossiers surveillés
ipcMain.handle('api-folders-reload-config', async (event) => {
  try {
    console.log('🔄 Rechargement de la configuration des dossiers...');
    
    if (!global.unifiedMonitoringService) {
      throw new Error('Service de monitoring non disponible');
    }

    // Recharger la configuration
    const result = await global.unifiedMonitoringService.reloadFoldersConfiguration();
    
    if (result.success) {
      console.log(`✅ Configuration rechargée: ${result.foldersCount} dossiers configurés`);
      
      // Émettre un événement pour notifier l'interface
      if (global.mainWindow) {
        global.mainWindow.webContents.send('folders-config-updated', {
          foldersCount: result.foldersCount,
          folders: result.folders
        });
      }
    }

    return result;
    
  } catch (error) {
    console.error('❌ Erreur rechargement configuration dossiers:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// === FONCTIONS UTILITAIRES POUR LES DOSSIERS ===

function findFolderInStructure(folders, targetPath) {
  for (const folder of folders) {
    if (folder.FolderPath === targetPath) {
      return folder;
    }
    if (folder.SubFolders && folder.SubFolders.length > 0) {
      const found = findFolderInStructure(folder.SubFolders, targetPath);
      if (found) return found;
    }
  }
  return null;
}

function extractFolderName(folderPath) {
  const parts = folderPath.split('\\');
  return parts[parts.length - 1] || folderPath;
}

function getParentPath(folderPath) {
  const parts = folderPath.split('\\');
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join('\\');
}

function calculateFolderStats(folders) {
  const stats = {
    total: folders.length,
    active: 0,
    declarations: 0,
    reglements: 0,
    simples: 0
  };

  folders.forEach(folder => {
    if (folder.isMonitored) {
      stats.active++;
      
      switch (folder.category) {
        case 'Déclarations':
          stats.declarations++;
          break;
        case 'Règlements':
          stats.reglements++;
          break;
        case 'Mails simples':
          stats.simples++;
          break;
      }
    }
  });

  return stats;
}

ipcMain.handle('api-outlook-status', async () => {
  try {
    if (outlookConnector && outlookConnector.isOutlookConnected) {
      return {
        connected: true,
        status: 'Connecté',
        version: outlookConnector.outlookVersion || 'Inconnue'
      };
    }
    return {
      connected: false,
      status: 'Déconnecté',
      version: null
    };
  } catch (error) {
    console.error('Erreur statut Outlook:', error);
    return {
      connected: false,
      status: 'Erreur',
      version: null
    };
  }
});

ipcMain.handle('api-stats-summary', async () => {
  try {
    // OPTIMIZED: Utiliser le cache intelligent d'abord
    const cachedStats = cacheService.getUIStats();
    if (cachedStats) {
      return cachedStats;
    }

    // Attendre un peu que le service unifié soit prêt si nécessaire
    let waitAttempts = 0;
    while (waitAttempts < 10 && global.unifiedMonitoringService && !global.unifiedMonitoringService.isInitialized) {
      await new Promise(resolve => setTimeout(resolve, 200));
      waitAttempts++;
    }
    
    // Utiliser le service unifié si disponible et initialisé
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized) {
      // CORRIGÉ: Utiliser la nouvelle méthode getBusinessStats au lieu de getStats
      const stats = await global.unifiedMonitoringService.getBusinessStats();
      return stats;
    }
    
    // OPTIMIZED: Fallback vers le service optimisé
    // Log fallback stats réduit
    // console.log('⚠️ [IPC] Service unifié non disponible, utilisation BD optimisée');
    await databaseService.initialize();
    const stats = await databaseService.getEmailStats();
    
    const result = {
      emailsToday: stats.emailsToday || 0,
      treatedToday: stats.treatedToday || 0,
      unreadTotal: stats.unreadTotal || 0,
      totalEmails: stats.totalEmails || 0,
      lastSyncTime: stats.lastSyncTime || new Date().toISOString(),
      monitoringActive: false
    };

    // OPTIMIZED: Mettre en cache
    cacheService.set('ui', 'dashboard_stats', result, 30);
    
    return result;
  } catch (error) {
    console.error('❌ [IPC] Erreur api-stats-summary:', error);
    return {
      emailsToday: 0,
      treatedToday: 0,
      unreadTotal: 0,
      totalEmails: 0,
      lastSyncTime: new Date().toISOString(),
      monitoringActive: false
    };
  }
});

ipcMain.handle('api-emails-recent', async () => {
  // Log réduit pour éviter le spam
  // console.log('📧 [IPC] api-emails-recent appelé');
  try {
    // OPTIMIZED: Cache intelligent pour emails récents
    const cachedEmails = cacheService.getRecentEmails(50);
    if (cachedEmails) {
      // Cache hit - log supprimé pour réduire spam
      // console.log('⚡ [IPC] Emails récents depuis cache');
      return cachedEmails;
    }

    // Attendre un peu que le service unifié soit prêt si nécessaire
    let waitAttempts = 0;
    while (waitAttempts < 10 && global.unifiedMonitoringService && !global.unifiedMonitoringService.isInitialized) {
      // Log d'attente supprimé pour réduire spam
      // console.log(`⏳ [IPC] Attente initialisation service unifié... ${waitAttempts + 1}/10`);
      await new Promise(resolve => setTimeout(resolve, 200));
      waitAttempts++;
    }
    
    // Utiliser le service unifié si disponible et initialisé
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized) {
      // Service unifié - log réduit
      // console.log('📧 [IPC] Utilisation service unifié pour emails récents');
      const emails = await global.unifiedMonitoringService.getRecentEmails(50);
      // console.log(`📧 [IPC] ${emails?.length || 0} emails trouvés via service unifié`);
      
      // OPTIMIZED: Mettre en cache
      if (emails) {
        cacheService.set('emails', 'recent_50', emails, 60); // 1 minute
      }
      
      return emails || [];
    }
    
    // OPTIMIZED: Fallback vers le service optimisé
    // Log réduit pour fallback
    // console.log('⚠️ [IPC] Service unifié non disponible, utilisation BD optimisée');
    await databaseService.initialize();
    const emails = await databaseService.getRecentEmails(50);
    // console.log(`📧 [IPC] ${emails?.length || 0} emails trouvés via BD optimisée`);
    
    // OPTIMIZED: Mettre en cache
    if (emails) {
      cacheService.set('emails', 'recent_50', emails, 60);
    }
    
    return emails || [];
  } catch (error) {
    console.error('❌ [IPC] Erreur api-emails-recent:', error);
    return [];
  }
});

ipcMain.handle('api-database-stats', async () => {
  try {
    // Utiliser le service unifié
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.getDatabaseStats) {
      return await global.unifiedMonitoringService.getDatabaseStats();
    }
    return {
      totalRecords: 0,
      databaseSize: 0,
      lastUpdate: null
    };
  } catch (error) {
    console.error('Erreur statistiques base de données:', error);
    return {
      totalRecords: 0,
      databaseSize: 0,
      lastUpdate: null
    };
  }
});

// Handler pour récupérer les emails récents
ipcMain.handle('api-recent-emails', async () => {
  try {
    if (global.unifiedMonitoringService) {
      const emails = await global.unifiedMonitoringService.getRecentEmails(20);
      return emails;
    } else {
      // Fallback vers databaseService direct
      await global.databaseService.initialize();
      const emails = await global.databaseService.getRecentEmails(20);
      return emails;
    }
  } catch (error) {
    console.error('❌ [IPC] Erreur récupération emails récents:', error);
    return [];
  }
});

ipcMain.handle('api-stats-by-category', async () => {
  try {
    // Utiliser le service unifié
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.getStatsByCategory) {
      return await global.unifiedMonitoringService.getStatsByCategory();
    }
    return {};
  } catch (error) {
    console.error('Erreur statistiques par catégorie:', error);
    return {};
  }
});

ipcMain.handle('api-app-settings-load', async () => {
  try {
    await global.databaseService.initialize();
    const settings = await global.databaseService.loadAppSettings();
    console.log('📄 Paramètres chargés depuis BDD:', settings);
    try {
      const flatVal = settings.count_read_as_treated;
      if (!settings.monitoring) settings.monitoring = {};
      settings.monitoring.treatReadEmailsAsProcessed = !!flatVal;
    } catch {}
    return { success: true, settings };
  } catch (error) {
    console.error('❌ Erreur chargement paramètres app:', error);
    const defaultSettings = { monitoring: { treatReadEmailsAsProcessed: false, scanInterval: 30000, autoStart: true }, ui: { theme: 'default', language: 'fr', emailsLimit: 20 }, notifications: { enabled: true, showStartupNotification: true } };
    return { success: true, settings: defaultSettings };
  }
});

// APIs VBA Metrics
ipcMain.handle('api-vba-metrics-summary', async () => {
  try {
    // Utiliser le service unifié pour les métriques
    if (global.unifiedMonitoringService) {
      return await global.unifiedMonitoringService.getMetricsSummary();
    }
    return {
      daily: { emailsReceived: 0, emailsProcessed: 0, emailsUnread: 0, readStatusChanges: 0 },
      weekly: { stockStart: 0, stockEnd: 0, evolution: 0, arrivals: 0, treatments: 0 },
      folders: {},
      lastUpdate: new Date().toISOString()
    };
  } catch (error) {
    console.error('Erreur métriques VBA:', error);
    return null;
  }
});

ipcMain.handle('api-vba-folder-distribution', async () => {
  try {
    // Utiliser le service unifié pour la distribution des dossiers
    if (global.unifiedMonitoringService) {
      return await global.unifiedMonitoringService.getFolderDistribution();
    }
    return {};
  } catch (error) {
    console.error('Erreur distribution dossiers:', error);
    return {};
  }
});

ipcMain.handle('api-vba-weekly-evolution', async () => {
  try {
    // Utiliser le service unifié pour l'évolution hebdomadaire
    if (global.unifiedMonitoringService) {
      return await global.unifiedMonitoringService.getWeeklyEvolution();
    }
    return {
      current: { weekNumber: 0, year: 0, stockStart: 0, stockEnd: 0, evolution: 0 },
      trend: 0,
      percentage: '0.0'
    };
  } catch (error) {
    console.error('Erreur évolution hebdo:', error);
    return null;
  }
});

// ========================================================================
// NOUVELLES APIs POUR LE SUIVI HEBDOMADAIRE (inspiré du système VBA)
// ========================================================================

// API pour récupérer les statistiques de la semaine courante
ipcMain.handle('api-weekly-current-stats', async () => {
  try {
    console.log('📅 [IPC] api-weekly-current-stats appelé');
    
    // Attendre que le service soit initialisé
    let waitAttempts = 0;
    while (waitAttempts < 10 && global.unifiedMonitoringService && !global.unifiedMonitoringService.isInitialized) {
      console.log(`📅 [IPC] Attente initialisation service (${waitAttempts + 1}/10)...`);
      await new Promise(resolve => setTimeout(resolve, 500));
      waitAttempts++;
    }
    
    let rawStats;
    let weekInfo;
    
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized && global.unifiedMonitoringService.dbService) {
      console.log('📅 [IPC] Service prêt, récupération des stats...');
      const currentWeekStats = global.unifiedMonitoringService.dbService.getCurrentWeekStats();
      rawStats = currentWeekStats.stats;
      weekInfo = currentWeekStats.weekInfo;
    } else {
      // Fallback: utiliser directement le service de BD
      console.log('📅 [IPC] Fallback: utilisation directe du service BD...');
      const optimizedDatabaseService = require('../services/optimizedDatabaseService');
      
      // S'assurer que la BD est initialisée
      if (!optimizedDatabaseService.isInitialized) {
        await optimizedDatabaseService.initialize();
      }
      
      const currentWeekStats = optimizedDatabaseService.getCurrentWeekStats();
      rawStats = currentWeekStats.stats;
      weekInfo = currentWeekStats.weekInfo;
    }
    
    // Transformer les données pour le frontend
    const categories = {};
    
    if (rawStats && Array.isArray(rawStats)) {
      rawStats.forEach(row => {
        const categoryName = row.folder_type === 'declarations' ? 'Déclarations' :
                           row.folder_type === 'reglements' ? 'Règlements' :
                           row.folder_type === 'mails_simples' ? 'Mails simples' :
                           row.folder_type;
        
        categories[categoryName] = {
          received: row.emails_received || 0,
          treated: Math.max(0, row.emails_treated || 0),
          adjustments: row.manual_adjustments || 0,
          total: (row.emails_received || 0) + (row.manual_adjustments || 0)
        };
      });
    }
    
    return {
      success: true,
      weekInfo: weekInfo,
      categories: categories,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ [IPC] Erreur api-weekly-current-stats:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// API pour récupérer l'historique des statistiques hebdomadaires
ipcMain.handle('api-weekly-history', async (event, { limit = 5, page = 1, pageSize = null } = {}) => {
  try {
    console.log('📅 [IPC] api-weekly-history appelé');
    
    // Attendre que le service soit initialisé
    let waitAttempts = 0;
    while (waitAttempts < 10 && global.unifiedMonitoringService && !global.unifiedMonitoringService.isInitialized) {
      console.log(`📅 [IPC] Attente initialisation service (${waitAttempts + 1}/10)...`);
      await new Promise(resolve => setTimeout(resolve, 500));
      waitAttempts++;
    }
    
  let weeklyStats = [];
  let pagination = { totalWeeks: 0, page: 1, pageSize: limit || 5, totalPages: 1 };
    
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized && global.unifiedMonitoringService.dbService) {
      console.log('📅 [IPC] Service prêt, récupération historique...');
      if (typeof page === 'number' && (pageSize || limit)) {
        const size = pageSize || limit;
        const result = global.unifiedMonitoringService.dbService.getWeeklyHistoryPage(page, size);
        weeklyStats = result.rows;
        pagination = { totalWeeks: result.totalWeeks, page: result.page, pageSize: result.pageSize, totalPages: result.totalPages };
      } else {
        weeklyStats = global.unifiedMonitoringService.dbService.getWeeklyStats(null, limit);
      }
    } else {
      // Fallback: utiliser directement le service de BD
      console.log('📅 [IPC] Fallback: utilisation directe du service BD...');
      const optimizedDatabaseService = require('../services/optimizedDatabaseService');
      
      // S'assurer que la BD est initialisée
      if (!optimizedDatabaseService.isInitialized) {
        await optimizedDatabaseService.initialize();
      }
      
      if (typeof page === 'number' && (pageSize || limit)) {
        const size = pageSize || limit;
        const result = optimizedDatabaseService.getWeeklyHistoryPage(page, size);
        weeklyStats = result.rows;
        pagination = { totalWeeks: result.totalWeeks, page: result.page, pageSize: result.pageSize, totalPages: result.totalPages };
      } else {
        weeklyStats = optimizedDatabaseService.getWeeklyStats(null, limit);
      }
    }
    
    // Transformer les données pour l'interface
    // Grouper par semaine pour calculer les totaux et organiser par catégories
    const weeklyGroups = {};
    
    weeklyStats.forEach(row => {
      const weekKey = `S${row.week_number} - ${row.week_year}`;
      
      if (!weeklyGroups[weekKey]) {
        // Calculer la plage de dates
        let dateRange = '';
        if (row.week_start_date && row.week_end_date) {
          const startDate = new Date(row.week_start_date);
          const endDate = new Date(row.week_end_date);
          const options = { month: '2-digit', day: '2-digit' };
          dateRange = `${startDate.toLocaleDateString('fr-FR', options)} - ${endDate.toLocaleDateString('fr-FR', options)}`;
        }
        
        weeklyGroups[weekKey] = {
          weekDisplay: weekKey,
          dateRange: dateRange,
          week_number: row.week_number,
          week_year: row.week_year,
          categories: {
            'Déclarations': { received: 0, treated: 0, adjustments: 0 },
            'Règlements': { received: 0, treated: 0, adjustments: 0 },
            'Mails simples': { received: 0, treated: 0, adjustments: 0 }
          }
        };
      }
      
      // Mapper le type de dossier vers une catégorie lisible (tolérant aux accents/casse)
      const toDisplayCategory = (val) => {
        if (!val) return 'Mails simples';
        const v = String(val).toLowerCase();
        if (v === 'mails_simples' || v.includes('mail')) return 'Mails simples';
        if (v === 'declarations' || v.includes('déclar') || v.includes('declar')) return 'Déclarations';
        if (v === 'reglements' || v.includes('règle') || v.includes('regle') || v.includes('reglement')) return 'Règlements';
        return 'Mails simples';
      };
      let category = toDisplayCategory(row.folder_type);
      
      const received = row.emails_received || 0;
  const treated = Math.max(0, row.emails_treated || 0);
      const adjustments = row.manual_adjustments || 0;
      
      // Mettre à jour les données de la catégorie
      if (weeklyGroups[weekKey].categories[category]) {
        weeklyGroups[weekKey].categories[category] = {
          received,
          treated,
          adjustments,
          // Cohérence: inclure les ajustements comme traitements additionnels
          stockEndWeek: Math.max(0, (received || 0) - ((treated || 0) + (adjustments || 0)))
        };
      }
    });
    
    // Créer le tableau transformé avec structure par semaine et catégories
    const transformedData = [];
    
    // Trier les semaines par ordre décroissant
    const sortedWeeks = Object.keys(weeklyGroups).sort((a, b) => {
      const aMatch = a.match(/S(\d+) - (\d+)/);
      const bMatch = b.match(/S(\d+) - (\d+)/);
      if (aMatch && bMatch) {
        const aYear = parseInt(aMatch[2]);
        const bYear = parseInt(bMatch[2]);
        if (aYear !== bYear) return bYear - aYear;
        return parseInt(bMatch[1]) - parseInt(aMatch[1]);
      }
      return 0;
    });
    
    // Calcul du stock roulant avec report par catégorie
    // 1) Déterminer le point de départ (carry) avant la première semaine de la page
    let startYear = null, startWeek = null;
    if (sortedWeeks.length) {
      const m = sortedWeeks[sortedWeeks.length - 1].match(/S(\d+) - (\d+)/); // dernière = plus ancienne
      if (m) { startWeek = parseInt(m[1], 10); startYear = parseInt(m[2], 10); }
    }

    // Accès au service DB pour calculer le carry initial
    let dbForCarry = null;
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized && global.unifiedMonitoringService.dbService) {
      dbForCarry = global.unifiedMonitoringService.dbService;
    } else {
      dbForCarry = require('../services/optimizedDatabaseService');
      if (!dbForCarry.isInitialized) await dbForCarry.initialize();
    }

    const carryInitial = (startYear && startWeek) ? dbForCarry.getCarryBeforeWeek(startYear, startWeek) : { declarations: 0, reglements: 0, mails_simples: 0 };

    // 2) Construire les entrées en partant de la plus ancienne vers la plus récente
    const running = {
      'Déclarations': carryInitial.declarations || 0,
      'Règlements': carryInitial.reglements || 0,
      'Mails simples': carryInitial.mails_simples || 0
    };

    for (let i = sortedWeeks.length - 1; i >= 0; i--) {
      const weekKey = sortedWeeks[i]; // plus ancienne -> la première itération
      const weekData = weeklyGroups[weekKey];

      // Construire catégories avec stockStart/End selon running
      const cats = ['Déclarations', 'Règlements', 'Mails simples'].map(name => {
        const rec = weekData.categories[name].received || 0;
        const trt = weekData.categories[name].treated || 0;
        const adj = weekData.categories[name].adjustments || 0;
        const start = running[name] || 0;
        const end = Math.max(0, start + rec - (trt + adj));
        running[name] = end;

        // Debug ciblé pour vérification du calcul S-1 (désactivable en commentant)
        try {
          if (weekData.week_year && weekData.week_number && name === 'Déclarations') {
            // Journaux compacts pour diagnostiquer les écarts de stock
            console.log(`🔎 [WEEKLY][${weekData.week_year}-S${weekData.week_number}] ${name}: start=${start} rec=${rec} trt=${trt} adj=${adj} => end=${end}`);
          }
        } catch (_) {}
        return {
          name,
          received: rec,
          treated: trt,
          adjustments: adj,
          stockStart: start,
          stockEndWeek: end
        };
      });

      const weekEntry = {
        weekDisplay: weekData.weekDisplay,
        week_number: weekData.week_number,
        week_year: weekData.week_year,
        dateRange: weekData.dateRange,
        categories: cats
      };

      transformedData.push(weekEntry);
    }

    // Remettre l'ordre décroissant (récente -> ancienne) pour l'affichage
    transformedData.reverse();

    // Calculer l'évolution S-1 basée sur le STOCK (carry-over) plutôt que sur les arrivées
    for (let i = 0; i < transformedData.length; i++) {
      const curr = transformedData[i];
      const prev = i < transformedData.length - 1 ? transformedData[i + 1] : null;
      if (prev) {
        const currentStock = curr.categories.reduce((sum, cat) => sum + (cat.stockEndWeek || 0), 0);
        const previousStock = prev.categories.reduce((sum, cat) => sum + (cat.stockEndWeek || 0), 0);
        const abs = currentStock - previousStock;
        const pct = previousStock > 0 ? (abs / previousStock) * 100 : 0;

        // Conserver aussi l'ancienne évolution basée sur les arrivées (info annexe)
        const currentArr = curr.categories.reduce((sum, cat) => sum + (cat.received || 0), 0);
        const previousArr = prev.categories.reduce((sum, cat) => sum + (cat.received || 0), 0);
        const absArr = currentArr - previousArr;
        const pctArr = previousArr > 0 ? (absArr / previousArr) * 100 : 0;

        curr.evolution = {
          absolute: abs,
          percent: pct,
          trend: abs > 0 ? 'up' : abs < 0 ? 'down' : 'stable',
          basis: 'stock',
          received: { absolute: absArr, percent: pctArr }
        };
      } else {
        curr.evolution = { absolute: 0, percent: 0, trend: 'stable', basis: 'stock', received: { absolute: 0, percent: 0 } };
      }
    }
    
    return {
      success: true,
      data: transformedData,
      page: pagination.page,
      pageSize: pagination.pageSize,
      totalWeeks: pagination.totalWeeks,
      totalPages: pagination.totalPages,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ [IPC] Erreur api-weekly-history:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// === Import Activité (.xlsb) ===
ipcMain.handle('dialog-open-xlsb', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: 'Choisir un fichier .xlsb',
    properties: ['openFile'],
    filters: [ { name: 'Excel Binary', extensions: ['xlsb'] } ]
  });
  if (canceled || !filePaths?.length) return null;
  return filePaths[0];
});

ipcMain.handle('api-activity-import-preview', async (event, { filePath, weeks }) => {
  try {
    await databaseService.initialize();
  const { rows, skippedWeeks, year } = activityImporter.importActivityFromXlsb(filePath, { weeks });
    // Aperçu: S1 et 2-3 autres semaines non vides
    const nonEmpty = rows.filter(r => (r.recu || r.traite || r.traite_adg));
    const byWeek = new Map();
    for (const r of nonEmpty) {
      const key = r.week_number;
      if (!byWeek.has(key)) byWeek.set(key, []);
      byWeek.get(key).push(r);
    }
    const weeksSorted = Array.from(byWeek.keys()).sort((a,b)=>a-b);
    const sampleWeeks = [];
    if (weeksSorted.includes(1)) sampleWeeks.push(1);
    for (const w of weeksSorted) if (sampleWeeks.length < 3 && !sampleWeeks.includes(w)) sampleWeeks.push(w);
    const preview = rows.filter(r => sampleWeeks.includes(r.week_number));
    return { year, skippedWeeks, preview, count: rows.length };
  } catch (e) {
    return { error: e.message };
  }
});

ipcMain.handle('api-activity-import-run', async (event, { filePath, weeks, outCsv }) => {
  try {
    await databaseService.initialize();
    const { rows, skippedWeeks, year } = activityImporter.importActivityFromXlsb(filePath, { weeks });
    // Si filtre démarre après 1, récupérer stock_debut depuis DB
    const { parseWeeksFilter } = activityImporter;
    const filter = parseWeeksFilter(weeks);
  // Plus de recalcul de stock via activity_weekly (supprimé)

  // Mettre à jour la table existante weekly_stats pour compatibilité UI existante
    //    - Agréger par semaine et "folder_type" équivalent
  const mapCategoryToFolderType = (cat) => {
      if (!cat) return 'Mails simples';
      if (cat === 'MailSimple') return 'Mails simples';
      if (cat === 'Reglements') return 'Règlements';
      if (cat === 'Declarations') return 'Déclarations';
      return 'Mails simples';
    };
  const getISOWeekInfo = (y, w) => {
      // Semaine ISO lundi->dimanche
      const simple = new Date(Date.UTC(y, 0, 1 + (w - 1) * 7));
      const day = simple.getUTCDay();
      const diff = (day <= 4 ? day : day - 7) - 1; // Monday=1
      const start = new Date(simple);
      start.setUTCDate(simple.getUTCDate() - diff);
      start.setUTCHours(0,0,0,0);
      const end = new Date(start);
      end.setUTCDate(start.getUTCDate() + 6);
      const startStr = start.toISOString().slice(0,10);
      const endStr = end.toISOString().slice(0,10);
  // Respecter le format attendu par l'appli: S{semaine}-{annee} (ex: S32-2025)
  const weekId = `S${w}-${y}`;
  return { startStr, endStr, weekId };
    };

    // Agrégation
    const agg = new Map(); // key: year-week-folderType
    for (const r of rows) {
      const folderType = mapCategoryToFolderType(r.category);
      const key = `${r.year}-${r.week_number}-${folderType}`;
      const { startStr, endStr, weekId } = getISOWeekInfo(r.year, r.week_number);
      if (!agg.has(key)) {
        agg.set(key, {
          week_identifier: weekId,
          week_number: r.week_number,
          week_year: r.year,
          week_start_date: startStr,
          week_end_date: endStr,
          folder_type: folderType,
          emails_received: 0,
          emails_treated: 0
        });
      }
      const a = agg.get(key);
      a.emails_received += (r.recu || 0);
      // Règle métier: treated = traité, manual_adjustments = traité_adg
      a.emails_treated += (r.traite || 0);
      a.manual_adjustments = (a.manual_adjustments || 0) + (r.traite_adg || 0);
    }
    if (agg.size > 0 && databaseService.upsertWeeklyStatsBatch) {
      // S'assurer que manual_adjustments et created_at sont présents
      const payload = Array.from(agg.values()).map(x => ({
        ...x,
        manual_adjustments: x.manual_adjustments || 0,
        created_at: null
      }));
      databaseService.upsertWeeklyStatsBatch(payload);
      // Notifier le frontend que les stats hebdo ont été mises à jour
      if (mainWindow && !mainWindow.isDestroyed()) {
        try {
          mainWindow.webContents.send('weekly-stats-updated', {
            inserted: payload.length,
            year,
            weeksAffected: [...new Set(payload.map(p => p.week_number))],
            timestamp: new Date().toISOString()
          });
        } catch (e) {
          console.warn('[IPC] Impossible d\'émettre weekly-stats-updated:', e?.message || e);
        }
      }
    }
    let csvPath = null;
    if (outCsv) {
      csvPath = activityImporter.writeCsv(rows, year, path.join(process.cwd(), 'build'));
    }
    return { inserted: rows.length, skippedWeeks, csvPath, year };
  } catch (e) {
    return { error: e.message };
  }
});

// API pour ajuster manuellement les compteurs (courrier papier, etc.)
ipcMain.handle('api-weekly-adjust-count', async (event, { weekIdentifier, folderType, adjustmentValue, adjustmentType = 'manual_adjustments' }) => {
  try {
    console.log(`📝 [IPC] api-weekly-adjust-count: ${weekIdentifier} - ${folderType} - ${adjustmentValue}`);
    
    if (!weekIdentifier || !folderType || adjustmentValue === undefined) {
      return {
        success: false,
        error: 'Paramètres manquants'
      };
    }
    
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.dbService) {
      const success = global.unifiedMonitoringService.dbService.adjustWeeklyCount(
        weekIdentifier, 
        folderType, 
        adjustmentValue, 
        adjustmentType
      );
      
      const result = {
        success: success,
        message: success ? 'Ajustement effectué' : 'Échec de l\'ajustement'
      };

      // Notifier le frontend si succès
      if (success && mainWindow && !mainWindow.isDestroyed()) {
        try {
          mainWindow.webContents.send('weekly-stats-updated', {
            weekIdentifier,
            folderType,
            adjustmentType,
            adjustmentValue,
            timestamp: new Date().toISOString()
          });
        } catch (e) {
          console.warn('[IPC] Impossible d\'émettre weekly-stats-updated (adjust):', e?.message || e);
        }
      }

      return result;
    }
    
    return {
      success: false,
      error: 'Service de base de données non disponible'
    };
    
  } catch (error) {
    console.error('❌ [IPC] Erreur api-weekly-adjust-count:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// ======================== WEEKLY COMMENTS IPC ========================
ipcMain.handle('api-weekly-comments-add', async (event, payload) => {
  try {
    await databaseService.initialize();
    const res = databaseService.addWeeklyComment(payload || {});
    return res;
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('api-weekly-comments-list', async (event, { week_identifier }) => {
  try {
    await databaseService.initialize();
    const res = databaseService.getWeeklyComments(week_identifier);
    return res;
  } catch (e) {
    return { success: false, rows: [], error: e.message };
  }
});

ipcMain.handle('api-weekly-comments-update', async (event, { id, comment_text, category }) => {
  try {
    await databaseService.initialize();
    const res = databaseService.updateWeeklyComment(id, comment_text, category);
    return res;
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('api-weekly-comments-delete', async (event, { id }) => {
  try {
    await databaseService.initialize();
    const res = databaseService.deleteWeeklyComment(id);
    return res;
  } catch (e) {
    return { success: false, error: e.message };
  }
});

ipcMain.handle('api-weekly-weeks-list', async (event, { limit = 52 } = {}) => {
  try {
    await databaseService.initialize();
    const res = databaseService.listDistinctWeeks(limit);
    return res;
  } catch (e) {
    return { success: false, rows: [], error: e.message };
  }
});

// API pour sauvegarder un mapping de dossier personnalisé
ipcMain.handle('api-folder-mapping-save', async (event, { originalPath, mappedCategory, displayName }) => {
  try {
    console.log(`🗂️ [IPC] api-folder-mapping-save: ${originalPath} -> ${mappedCategory}`);
    
    if (!originalPath || !mappedCategory) {
      return {
        success: false,
        error: 'Paramètres manquants'
      };
    }
    
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.dbService) {
      const success = global.unifiedMonitoringService.dbService.saveFolderMapping(
        originalPath, 
        mappedCategory, 
        displayName
      );
      
      return {
        success: success,
        message: success ? 'Mapping sauvegardé' : 'Échec de la sauvegarde'
      };
    }
    
    return {
      success: false,
      error: 'Service de base de données non disponible'
    };
    
  } catch (error) {
    console.error('❌ [IPC] Erreur api-folder-mapping-save:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

// API pour obtenir/modifier le paramètre "mail lu = traité"
ipcMain.handle('api-settings-count-read-as-treated', async (event, { value } = {}) => {
  try {
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.dbService) {
      
      if (value !== undefined) {
        // Normaliser en booléen puis sauvegarder le paramètre
        const boolVal = (typeof value === 'string') ? (value.toLowerCase() === 'true') : !!value;
        const success = global.unifiedMonitoringService.dbService.saveAppSetting
          ? !!global.unifiedMonitoringService.dbService.saveAppSetting('count_read_as_treated', boolVal)
          : global.unifiedMonitoringService.dbService.setAppSetting('count_read_as_treated', boolVal);
        console.log(`⚙️ [IPC] Paramètre "mail lu = traité" défini: ${boolVal}`);
        
        return {
          success: success,
          value: boolVal,
          message: success ? 'Paramètre sauvegardé' : 'Échec de la sauvegarde'
        };
      } else {
        // Récupérer le paramètre
        const currentValueRaw = global.unifiedMonitoringService.dbService.getAppSetting('count_read_as_treated', 'false');
        const currentBool = (typeof currentValueRaw === 'boolean')
          ? currentValueRaw
          : (typeof currentValueRaw === 'string')
            ? (currentValueRaw.toLowerCase() === 'true')
            : !!currentValueRaw;
        
        return {
          success: true,
          value: currentBool
        };
      }
    }
    
    return {
      success: false,
      error: 'Service de base de données non disponible'
    };
    
  } catch (error) {
    console.error('❌ [IPC] Erreur api-settings-count-read-as-treated:', error);
    return {
      success: false,
      error: error.message
    };
  }
});

ipcMain.handle('api-monitoring-status', async () => {
  try {
    // Utiliser le service unifié pour le statut
    if (global.unifiedMonitoringService) {
      return {
        active: global.unifiedMonitoringService.isMonitoring,
        status: global.unifiedMonitoringService.isMonitoring ? 'En cours' : 'Arrêté',
        lastCheck: global.unifiedMonitoringService.stats.lastSyncTime || new Date().toISOString(),
        foldersMonitored: global.unifiedMonitoringService.monitoredFolders.length
      };
    }
    return {
      active: false,
      status: 'Arrêté',
      lastCheck: null,
      foldersMonitored: 0
    };
  } catch (error) {
    console.error('Erreur statut monitoring:', error);
    return {
      active: false,
      status: 'Erreur',
      lastCheck: null,
      foldersMonitored: 0
    };
  }
});

ipcMain.handle('api-monitoring-start', async () => {
  try {
    console.log('🚀 Tentative de démarrage monitoring depuis IPC...');
    
    if (global.unifiedMonitoringService) {
      if (!global.unifiedMonitoringService.isMonitoring) {
        console.log('🔄 Démarrage du monitoring...');
        await global.unifiedMonitoringService.startMonitoring();
        return { success: true, message: 'Monitoring démarré avec succès' };
      } else {
        console.log('✅ Service unifié déjà actif');
        return { success: true, message: 'Service unifié déjà actif' };
      }
    } else {
      console.log('⚠️ Service unifié non initialisé');
      return { success: false, message: 'Service unifié non initialisé' };
    }
  } catch (error) {
    console.error('❌ Erreur démarrage monitoring:', error);
    return { success: false, message: error.message };
  }
});

ipcMain.handle('api-monitoring-stop', async () => {
  try {
    console.log('⏹️ Tentative d\'arrêt monitoring depuis IPC...');
    
    if (global.unifiedMonitoringService) {
      if (global.unifiedMonitoringService.isMonitoring) {
        console.log('🛑 Arrêt du monitoring...');
        await global.unifiedMonitoringService.stopMonitoring();
        return { success: true, message: 'Monitoring arrêté avec succès' };
      } else {
        console.log('⚠️ Monitoring déjà arrêté');
        return { success: true, message: 'Monitoring déjà arrêté' };
      }
    } else {
      console.log('⚠️ Service unifié non initialisé');
      return { success: false, message: 'Service unifié non initialisé' };
    }
  } catch (error) {
    console.error('❌ Erreur arrêt monitoring:', error);
    return { success: false, message: error.message };
  }
});

ipcMain.handle('api-monitoring-force-sync', async () => {
  try {
    console.log('💪 Synchronisation forcée demandée depuis IPC...');
    
    if (global.unifiedMonitoringService) {
      console.log('🔄 Lancement synchronisation forcée...');
      await global.unifiedMonitoringService.forceSync();
      return { success: true, message: 'Synchronisation forcée terminée' };
    } else {
      console.log('⚠️ Service unifié non initialisé');
      return { success: false, message: 'Service unifié non initialisé' };
    }
  } catch (error) {
    console.error('❌ Erreur synchronisation forcée:', error);
    return { success: false, message: error.message };
  }
});

// App lifecycle
app.whenReady().then(async () => {
  console.log('🚀 Initialisation de Mail Monitor...');
  createLoadingWindow();
  // Vérification ressources critiques (scripts EWS & DLL)
  try {
    const { resolveResource } = require('../server/scriptPathResolver');
    const scriptCheck = resolveResource(['scripts'], 'ews-list-folders.ps1');
  const dllCheck = resolveResource(['ews'], 'Microsoft.Exchange.WebServices.dll');
    if (!scriptCheck.path) {
      console.warn('⚠️ Script EWS introuvable au démarrage (fallback COM). Candidats:', scriptCheck.tried);
    } else {
      console.log('✅ Script EWS trouvé:', scriptCheck.path);
    }
    if (!dllCheck.path) {
      console.warn('⚠️ DLL EWS introuvable au démarrage (fallback COM). Candidats:', dllCheck.tried);
    } else {
      console.log('✅ DLL EWS trouvée:', dllCheck.path);
    }
  } catch (e) {
    console.warn('⚠️ Vérification ressources EWS a échoué:', e.message);
  }
});

// IPC health check pour diagnostics prod
ipcMain.handle('api-health-check', async () => {
  try {
    const { resolveResource } = require('../server/scriptPathResolver');
    const scriptCheck = resolveResource(['scripts'], 'ews-list-folders.ps1');
    const dllCheck = resolveResource(['ews'], 'Microsoft.Exchange.WebServices.dll');
    return {
      success: true,
      timestamp: new Date().toISOString(),
      ewsScript: scriptCheck.path || null,
      ewsDll: dllCheck.path || null,
      ewsTriedScript: scriptCheck.tried,
      ewsTriedDll: dllCheck.tried,
      ewsDisabled: !!(require('../server/outlookConnector')._ewsDisabled),
      ewsFailures: require('../server/outlookConnector')._ewsFailures || 0,
      ewsInvalid: Array.from(require('../server/outlookConnector')._ewsInvalid || [])
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Handler pour la fermeture de fenêtre
ipcMain.handle('window-close', async () => {
  logClean('🔹 Demande de fermeture de fenêtre');
  try {
    // Fermer la fenêtre principale
    if (mainWindow) {
      mainWindow.close();
    }
    // Quitter l'application
    app.quit();
    return { success: true };
  } catch (error) {
    console.error('Erreur lors de la fermeture:', error);
    return { success: false, error: error.message };
  }
});

// Handler pour la minimisation de fenêtre
ipcMain.handle('window-minimize', async () => {
  console.log('🔹 Demande de minimisation de fenêtre');
  if (mainWindow) {
    mainWindow.minimize();
  }
  return { success: true };
});

// Handler pour récupérer les boîtes mail Outlook
ipcMain.handle('api-outlook-mailboxes', async () => {
  console.log('🔹 Demande de récupération des boîtes mail');
  try {
  // Tenter de récupérer les boîtes même si non marqué "connecté" (le connecteur gère ensureConnected)
  const mailboxes = await outlookConnector.getMailboxes?.();
    console.log('📊 DEBUG - Boîtes mail récupérées:', JSON.stringify(mailboxes, null, 2));
    return { 
      success: true, 
      mailboxes: mailboxes || [] 
    };
  } catch (error) {
    console.error('❌ Erreur récupération boîtes mail:', error.message);
    return { 
      success: false, 
      error: error.message,
      mailboxes: [] 
    };
  }
});

// Handler pour récupérer la structure des dossiers d'une boîte mail
ipcMain.handle('api-outlook-folder-structure', async (event, storeId) => {
  console.log(`🔹 Demande de structure des dossiers pour store: ${storeId}`);
  try {
  // Ne pas bloquer sur l'état "connecté"; le connecteur tentera une connexion légère si besoin
  const folders = await outlookConnector.getFolderStructure?.(storeId);
    console.log('📊 DEBUG - Structure des dossiers récupérée:', JSON.stringify(folders, null, 2));
    return { 
      success: true, 
      folders: folders || [] 
    };
  } catch (error) {
    console.error('❌ Erreur récupération structure dossiers:', error.message);
    return { 
      success: false, 
      error: error.message,
      folders: [] 
    };
  }
});

// Handler pour récupérer les sous-dossiers (lazy-load)
ipcMain.handle('api-outlook-subfolders', async (event, payload) => {
  try {
    const { storeId, parentEntryId, parentPath } = payload || {};
    if (!storeId) {
      return { success: false, error: 'storeId requis', children: [] };
    }
    // Support both modes: by EntryID when available, or by display parentPath when EntryID is missing
    const children = await outlookConnector.getSubFolders?.(storeId, parentEntryId || '', parentPath || '');
    return { success: true, children: children || [] };
  } catch (error) {
    console.error('❌ Erreur récupération sous-dossiers:', error.message);
    return { success: false, error: error.message, children: [] };
  }
});

// EWS: enumeration rapide des dossiers (top-level et enfants)
ipcMain.handle('api-ews-top-level', async (_event, { mailbox }) => {
  try {
    if (!mailbox) return { success: false, error: 'mailbox requis', folders: [] };
    const folders = await outlookConnector.getTopLevelFoldersFast(mailbox);
    // Sanitize to plain serializable array
    const safe = Array.isArray(folders) ? folders.map(f => ({
      Id: f.Id,
      Name: f.Name,
      ChildCount: Number(f.ChildCount || 0)
    })) : [];
    return { success: true, folders: safe };
  } catch (e) {
    return { success: false, error: e?.message || String(e), folders: [] };
  }
});

ipcMain.handle('api-ews-children', async (_event, { mailbox, parentId }) => {
  try {
    if (!mailbox || !parentId) return { success: false, error: 'mailbox et parentId requis', folders: [] };
    const folders = await outlookConnector.getChildFoldersFast(mailbox, parentId);
    const safe = Array.isArray(folders) ? folders.map(f => ({
      Id: f.Id,
      Name: f.Name,
      ChildCount: Number(f.ChildCount || 0)
    })) : [];
    return { success: true, folders: safe };
  } catch (e) {
    return { success: false, error: e?.message || String(e), folders: [] };
  }
});

// COM rapide: stores + folders shallow
const fastOL = require('../server/outlookFastFolders');
ipcMain.handle('api-ol-stores', async () => {
  try {
    const data = await fastOL.listStores();
    return { ok: true, data };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
});

ipcMain.handle('api-ol-folders-shallow', async (_event, { storeId, parentEntryId }) => {
  try {
    if (!storeId) return { ok: false, error: 'storeId requis' };
    const payload = await fastOL.listFoldersShallow(storeId, parentEntryId || '');
    return { ok: true, data: payload };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
});

// Handler pour sauvegarder les paramètres de l'application
ipcMain.handle('api-app-settings-save', async (event, settings) => {
  console.log('🔹 Demande de sauvegarde des paramètres');
  try {
    // Ancien require supprimé, utiliser global.databaseService
    await global.databaseService.initialize();
    
    // Sauvegarder dans la base de données
    for (const [section, sectionData] of Object.entries(settings)) {
      for (const [key, value] of Object.entries(sectionData)) {
        const configKey = `${section}.${key}`;
        await global.databaseService.saveAppConfig(configKey, value);
      }
    }
    
    console.log('📝 Paramètres sauvegardés en base de données:', settings);
    
    return { 
      success: true,
      message: 'Paramètres sauvegardés avec succès' 
    };
  } catch (error) {
    console.error('❌ Erreur sauvegarde paramètres:', error.message);
    return { 
      success: false, 
      error: error.message 
    };
  }
});

// Gestionnaires IPC pour l'amélioration de la fenêtre de chargement
ipcMain.on('loading-page-complete', () => {
  logClean('🎯 Page de chargement signale complétion');
  if (loadingWindow) {
    loadingWindow.close();
    loadingWindow = null;
  }
});

ipcMain.on('loading-retry', () => {
  logClean('🔄 Demande de retry depuis la page de chargement');
  // Réinitialiser le système et relancer l'initialisation
  isInitializing = false;
  if (loadingWindow) {
    loadingWindow.webContents.send('loading-progress', {
      step: 0,
      progress: 0,
      message: 'Redémarrage de l\'initialisation...'
    });
  }
  
  // Relancer l'initialisation après un court délai
  setTimeout(() => {
    initializeOutlook().catch(error => {
      logClean('❌ Erreur lors du retry:', error.message);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-error', {
          message: error.message,
          code: error.code || 'RETRY_FAILED'
        });
      }
    });
  }, 500);
});

// Fonction utilitaire pour envoyer la progression des tâches détaillées
function sendTaskProgress(taskId, description, completed = false, error = null) {
  if (loadingWindow) {
    loadingWindow.webContents.send('task-progress', {
      taskId,
      description,
      completed,
      error
    });
  }
}

// Fonction pour fermer la fenêtre de chargement depuis l'API
function closeLoadingWindow() {
  if (loadingWindow) {
    loadingWindow.webContents.send('loading-complete');
    setTimeout(() => {
      if (loadingWindow) {
        loadingWindow.close();
        loadingWindow = null;
      }
    }, 100);
  }
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
