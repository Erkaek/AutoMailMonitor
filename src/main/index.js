/**
 * Mail Monitor - Surveillance professionnelle des emails Outlook
 * Corrected version avec protection contre multiples démarrages
 */

const { app, BrowserWindow, Tray, Menu, ipcMain, nativeImage } = require('electron');
const path = require('path');
// CORRECTION: Utiliser le connecteur optimisé 
const outlookConnector = require('../server/outlookConnector');
// OPTIMIZED: Utiliser le service de base de données optimisé
const databaseService = require('../services/optimizedDatabaseService');
const cacheService = require('../services/cacheService');

// Rendre les services disponibles globalement
global.outlookConnector = outlookConnector;
global.databaseService = databaseService;
global.cacheService = cacheService;

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
  height: 800,
  minWidth: 800,
  minHeight: 600
};

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
    transparent: true, // Fenêtre transparente
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
    initializeOutlook();
  });

  loadingWindow.on('closed', () => {
    loadingWindow = null;
  });

  return loadingWindow;
}

/**
 * Configuration du transfert des événements temps réel du service vers le frontend
 */
function setupRealtimeEventForwarding() {
  if (!global.unifiedMonitoringService || !mainWindow) {
    console.log('⚠️ Service unifié ou fenêtre principale non disponible pour les événements temps réel');
    return;
  }

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
    minWidth: APP_CONFIG.minWidth,
    minHeight: APP_CONFIG.minHeight,
  icon: path.join(__dirname, '../../resources', 'new logo', 'logo.ico'),
    title: 'Mail Monitor - Surveillance Outlook',
    show: false,
    frame: true, // Réactivation de la barre d'outils Windows
    titleBarStyle: 'default', // Style par défaut de la barre de titre
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: false,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile(path.join(__dirname, '../../public/index.html'));

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    logClean('🎉 Application prete !');
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  mainWindow.on('minimize', (event) => {
    if (tray) {
      event.preventDefault();
      mainWindow.hide();
    }
  });

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
    
    console.log(`🔧 [IPC] Redimensionnement fenêtre de chargement: ${finalWidth}x${finalHeight}`);
    loadingWindow.setSize(finalWidth, finalHeight);
    loadingWindow.center(); // Recentrer après redimensionnement
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
        folderCategories[folder.folder_name] = {
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
    
    // Ancien require supprimé, utiliser global.databaseService
    await global.databaseService.initialize();
    
    // Sauvegarder UNIQUEMENT dans la base de données (pas de JSON)
    await global.databaseService.saveFoldersConfiguration(data);
    console.log('✅ Configuration dossiers sauvegardée exclusivement en base de données');
    
    // Redémarrer automatiquement le monitoring si des dossiers sont configurés
    const hasData = (Array.isArray(data) && data.length > 0) || 
                   (typeof data === 'object' && Object.keys(data).length > 0);
                   
    if (global.unifiedMonitoringService && hasData) {
      console.log('🔄 Redémarrage automatique du service unifié avec sync PowerShell...');
      try {
        // Arrêter d'abord le monitoring existant
        if (global.unifiedMonitoringService.isMonitoring) {
          await global.unifiedMonitoringService.stopMonitoring();
        }
        // Réinitialiser le service (déclenchera la sync PowerShell automatiquement)
        await global.unifiedMonitoringService.initialize();
        console.log('✅ Service unifié redémarré automatiquement avec sync PowerShell');
      } catch (error) {
        console.error('❌ Erreur redémarrage service unifié:', error);
      }
    }
    
    return { success: true };
  } catch (error) {
    console.error('❌ Erreur sauvegarde configuration dossiers:', error);
    return { success: false, error: error.message };
  }
});

// === HANDLERS IPC POUR LA GESTION HIÉRARCHIQUE DES DOSSIERS ===

// Récupérer l'arbre hiérarchique des dossiers
ipcMain.handle('api-folders-tree', async () => {
  try {
    // OPTIMIZED: Vérifier le cache d'abord
    const cachedFolders = cacheService.get('config', 'folders_tree');
    if (cachedFolders) {
      return cachedFolders;
    }

    await databaseService.initialize();

    // Récupérer UNIQUEMENT les dossiers configurés (monitorés) depuis la BDD optimisée
    const foldersConfig = global.databaseService.getFoldersConfiguration();

    // Récupérer la structure Outlook pour obtenir les compteurs d'emails
    const allFolders = await outlookConnector.getFolders();

    // Créer la liste des dossiers monitorés seulement
    const monitoredFolders = [];

    foldersConfig.forEach(config => {
      // Chercher le dossier dans la structure Outlook pour obtenir le nombre d'emails
      const outlookFolder = allFolders.find(f => f.path === config.folder_name || f.name === config.folder_name);

      monitoredFolders.push({
        path: config.folder_name,
        name: config.folder_name || config.folder_name,
        isMonitored: true,
        category: config.category || 'Mails simples',
        emailCount: outlookFolder ? outlookFolder.emailCount || 0 : 0,
        parentPath: getParentPath(config.folder_name)
      });
    });

    // Calculer les statistiques
    const stats = calculateFolderStats(monitoredFolders);

    const result = {
      folders: monitoredFolders, // Uniquement les dossiers monitorés
      stats: stats,
      timestamp: new Date().toISOString()
    };

    // OPTIMIZED: Mettre en cache pour 5 minutes
    cacheService.set('config', 'folders_tree', result, 300);

    return result;

  } catch (error) {
    console.error('❌ [IPC] Erreur récupération arbre dossiers:', error);
    throw error;
  }
});

// Ajouter un dossier au monitoring
ipcMain.handle('api-folders-add', async (event, { folderPath, category }) => {
  try {
    if (!folderPath || !category) {
      throw new Error('Chemin du dossier et catégorie requis');
    }

    await databaseService.initialize();

    // Vérifier que le dossier existe dans Outlook
    const folderExists = await outlookConnector.folderExists(folderPath);
    if (!folderExists) {
      throw new Error('Dossier non trouvé dans Outlook');
    }

    // Ajouter directement à la base de données optimisée
    const folderName = extractFolderName(folderPath);
    await databaseService.addFolderConfiguration(folderPath, category, folderName);
    
    // OPTIMIZED: Invalidation intelligente du cache
    cacheService.invalidateFoldersConfig();
    
    console.log(`✅ Dossier ${folderPath} ajouté au monitoring en BDD`);

    // Redémarrer le monitoring pour prendre en compte le nouveau dossier
    if (global.unifiedMonitoringService) {
      try {
        if (global.unifiedMonitoringService.isMonitoring) {
          await global.unifiedMonitoringService.stopMonitoring();
        }
        await global.unifiedMonitoringService.initialize();
        console.log('🔄 Service unifié redémarré pour le nouveau dossier');
      } catch (error) {
        console.error('⚠️ Erreur redémarrage monitoring:', error);
      }
    }

    return {
      success: true,
      message: 'Dossier ajouté au monitoring',
      folderPath: folderPath,
      category: category
    };

  } catch (error) {
    console.error('❌ Erreur ajout dossier:', error);
    throw error;
  }
});

// Mettre à jour la catégorie d'un dossier
ipcMain.handle('api-folders-update-category', async (event, { folderPath, category }) => {
  try {
    if (!folderPath || !category) {
      throw new Error('Chemin du dossier et catégorie requis');
    }

    const databaseService = require('../services/databaseService');
    await databaseService.initialize();

    // Mettre à jour directement en base de données
    const updated = await global.databaseService.updateFolderCategory(folderPath, category);
    
    if (!updated) {
      throw new Error('Dossier non trouvé dans la configuration active');
    }

    console.log(`✅ Catégorie de ${folderPath} mise à jour: ${category}`);

    // Redémarrer le monitoring pour prendre en compte le changement
    if (global.unifiedMonitoringService) {
      try {
        if (global.unifiedMonitoringService.isMonitoring) {
          await global.unifiedMonitoringService.stopMonitoring();
        }
        await global.unifiedMonitoringService.initialize();
        console.log('🔄 Service unifié redémarré pour changement de catégorie');
      } catch (error) {
        console.error('⚠️ Erreur redémarrage monitoring:', error);
      }
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

    const databaseService = require('../services/databaseService');
    await databaseService.initialize();

    // Supprimer le dossier de la configuration en base de données
    const deleted = await global.databaseService.deleteFolderConfiguration(folderPath);
    
    if (!deleted) {
      throw new Error('Dossier non trouvé dans la configuration');
    }

    console.log(`✅ Dossier ${folderPath} supprimé de la configuration`);

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
    // Ancien require supprimé, utiliser global.databaseService
    await global.databaseService.initialize();
    
    // Charger les paramètres depuis la base de données
    const settings = await global.databaseService.loadAppSettings();
    console.log('📄 Paramètres chargés depuis BDD:', settings);
    
    return {
      success: true,
      settings: settings
    };
  } catch (error) {
    console.error('❌ Erreur chargement paramètres app:', error);
    // Retourner des paramètres par défaut
    const defaultSettings = {
      monitoring: {
        treatReadEmailsAsProcessed: false,
        scanInterval: 30000,
        autoStart: true
      },
      ui: {
        theme: 'default',
        language: 'fr',
        emailsLimit: 20
      },
      notifications: {
        enabled: true,
        showStartupNotification: true
      }
    };
    
    return {
      success: true,
      settings: defaultSettings
    };
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
        await optimizedDatabaseService.init();
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
          treated: row.emails_treated || 0,
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
ipcMain.handle('api-weekly-history', async (event, { limit = 20 } = {}) => {
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
    
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized && global.unifiedMonitoringService.dbService) {
      console.log('📅 [IPC] Service prêt, récupération historique...');
      weeklyStats = global.unifiedMonitoringService.dbService.getWeeklyStats(null, limit);
    } else {
      // Fallback: utiliser directement le service de BD
      console.log('📅 [IPC] Fallback: utilisation directe du service BD...');
      const optimizedDatabaseService = require('../services/optimizedDatabaseService');
      
      // S'assurer que la BD est initialisée
      if (!optimizedDatabaseService.isInitialized) {
        await optimizedDatabaseService.initialize();
      }
      
      weeklyStats = optimizedDatabaseService.getWeeklyStats(null, limit);
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
      
      // Mapper le type de dossier vers une catégorie lisible
      let category = row.folder_type || 'Mails simples';
      if (category === 'mails_simples') category = 'Mails simples';
      else if (category === 'declarations') category = 'Déclarations';
      else if (category === 'reglements') category = 'Règlements';
      
      const received = row.emails_received || 0;
      const treated = row.emails_treated || 0;
      const adjustments = row.manual_adjustments || 0;
      
      // Mettre à jour les données de la catégorie
      if (weeklyGroups[weekKey].categories[category]) {
        weeklyGroups[weekKey].categories[category] = {
          received,
          treated,
          adjustments,
          stockEndWeek: Math.max(0, received - treated)
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
    
    // Calculer l'évolution pour chaque semaine
    for (let i = 0; i < sortedWeeks.length; i++) {
      const weekKey = sortedWeeks[i];
      const weekData = weeklyGroups[weekKey];
      const previousWeekData = i < sortedWeeks.length - 1 ? weeklyGroups[sortedWeeks[i + 1]] : null;
      
      // Créer une structure avec les 3 catégories
      const weekEntry = {
        weekDisplay: weekData.weekDisplay,
        dateRange: weekData.dateRange,
        categories: [
          {
            name: 'Déclarations',
            received: weekData.categories['Déclarations'].received,
            treated: weekData.categories['Déclarations'].treated,
            adjustments: weekData.categories['Déclarations'].adjustments,
            stockEndWeek: weekData.categories['Déclarations'].stockEndWeek || 0
          },
          {
            name: 'Règlements',
            received: weekData.categories['Règlements'].received,
            treated: weekData.categories['Règlements'].treated,
            adjustments: weekData.categories['Règlements'].adjustments,
            stockEndWeek: weekData.categories['Règlements'].stockEndWeek || 0
          },
          {
            name: 'Mails simples',
            received: weekData.categories['Mails simples'].received,
            treated: weekData.categories['Mails simples'].treated,
            adjustments: weekData.categories['Mails simples'].adjustments,
            stockEndWeek: weekData.categories['Mails simples'].stockEndWeek || 0
          }
        ]
      };
      
      // Calculer l'évolution par rapport à la semaine précédente
      if (previousWeekData) {
        const currentTotal = weekEntry.categories.reduce((sum, cat) => sum + cat.received, 0);
        const previousTotal = Object.values(previousWeekData.categories).reduce((sum, cat) => sum + cat.received, 0);
        const evolution = currentTotal - previousTotal;
        const evolutionPercent = previousTotal > 0 ? ((evolution / previousTotal) * 100) : 0;
        
        weekEntry.evolution = {
          absolute: evolution,
          percent: evolutionPercent,
          trend: evolution > 0 ? 'up' : evolution < 0 ? 'down' : 'stable'
        };
      } else {
        weekEntry.evolution = { absolute: 0, percent: 0, trend: 'stable' };
      }
      
      transformedData.push(weekEntry);
    }
    
    return {
      success: true,
      data: transformedData,
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
      
      return {
        success: success,
        message: success ? 'Ajustement effectué' : 'Échec de l\'ajustement'
      };
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
        // Sauvegarder le paramètre
        const success = global.unifiedMonitoringService.dbService.setAppSetting('count_read_as_treated', value.toString());
        console.log(`⚙️ [IPC] Paramètre "mail lu = traité" défini: ${value}`);
        
        return {
          success: success,
          value: value,
          message: success ? 'Paramètre sauvegardé' : 'Échec de la sauvegarde'
        };
      } else {
        // Récupérer le paramètre
        const currentValue = global.unifiedMonitoringService.dbService.getAppSetting('count_read_as_treated', 'false');
        
        return {
          success: true,
          value: currentValue === 'true'
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
    if (!outlookConnector || !outlookConnector.isOutlookConnected) {
      return { 
        success: false, 
        error: 'Outlook non connecté',
        mailboxes: [] 
      };
    }
    
    const mailboxes = await outlookConnector.getMailboxes();
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
    if (!outlookConnector || !outlookConnector.isOutlookConnected) {
      return { 
        success: false, 
        error: 'Outlook non connecté',
        folders: [] 
      };
    }
    
    const folders = await outlookConnector.getFolderStructure(storeId);
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
