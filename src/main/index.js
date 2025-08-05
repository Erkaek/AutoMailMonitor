/**
 * Mail Monitor - Surveillance professionnelle des emails Outlook
 * Corrected version avec protection contre multiples démarrages
 */

const { app, BrowserWindow, Tray, Menu, ipcMain, nativeImage } = require('electron');
const path = require('path');
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
    width: 700,
    height: 600,
    frame: false,
    alwaysOnTop: true,
    resizable: false,
    center: true,
    show: false,
    transparent: true, // Fenêtre transparente
    icon: path.join(__dirname, '../../resources/app.ico'),
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

  console.log('✅ Transfert d\'événements temps réel configuré');
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: APP_CONFIG.width,
    height: APP_CONFIG.height,
    minWidth: APP_CONFIG.minWidth,
    minHeight: APP_CONFIG.minHeight,
    icon: path.join(__dirname, '../../resources/app.ico'),
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
  logClean('🚀 Début de l\'initialisation Outlook (protection active)');
  
  try {
    // Etape 1: Verification de l'environnement
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 0,
        progress: 100,
        message: 'Vérification du système...'
      });
    }
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Etape 2: Connexion à Outlook
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 0,
        message: 'Connexion à Outlook...'
      });
    }
    
    // Wait for outlookConnector to be available and try to connect
    let retries = 0;
    const maxRetries = 60;
    let connected = false;
    
    while (retries < maxRetries && !connected) {
      try {
        if (outlookConnector) {
          // Check if already connected
          if (outlookConnector.isOutlookConnected) {
            connected = true;
            break;
          }
          
          // Try to establish connection
          await outlookConnector.establishConnection();
          connected = outlookConnector.isOutlookConnected;
          
          if (connected) {
            break;
          }
        }
      } catch (error) {
        console.log(`[INIT] Tentative ${retries + 1} échouée: ${error.message}`);
      }
      
      await new Promise(resolve => setTimeout(resolve, 500));
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
      throw new Error('Connexion impossible à Outlook après ' + maxRetries + ' tentatives');
    }

    // Finaliser la connexion
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 1,
        progress: 100,
        message: 'Outlook connecté !'
      });
    }
    await new Promise(resolve => setTimeout(resolve, 500));

    // Etape 3-4: Autres étapes...
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 2,
        progress: 100,
        message: 'Configuration chargée !'
      });
    }
    await new Promise(resolve => setTimeout(resolve, 500));

    // Etape 5: Monitoring automatique
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-progress', {
        step: 4,
        progress: 0,
        message: 'Initialisation du monitoring...'
      });
    }
    
    try {
      // Utiliser uniquement la base de données pour la configuration
      const databaseService = require('../services/databaseService');
      await databaseService.initialize();
      const folderConfig = await databaseService.getFoldersConfiguration();
      const configFound = Array.isArray(folderConfig) && folderConfig.length > 0;
      
      if (configFound) {
        console.log(`📁 Configuration trouvée en BDD: ${folderConfig.length} dossiers configurés`);
      }
      
      // CORRECTION: Toujours initialiser le service unifié (même sans configuration)
      const UnifiedMonitoringService = require('../services/unifiedMonitoringService');
      global.unifiedMonitoringService = new UnifiedMonitoringService(outlookConnector);
      
      // Initialiser de manière NON-BLOQUANTE
      console.log('🔧 Initialisation du service unifié en arrière-plan...');
      
      // Faire l'initialisation en arrière-plan sans attendre
      global.unifiedMonitoringService.initialize().then(() => {
        console.log('✅ Service unifié initialisé en arrière-plan');
        // Configurer les listeners d'événements temps réel
        setupRealtimeEventForwarding();
        
        if (configFound) {
          console.log(`📁 Configuration trouvée en BDD: ${folderConfig.length} dossiers configurés`);
          console.log('🔄 Le monitoring PowerShell + COM va démarrer automatiquement...');
          // Le monitoring démarrera automatiquement avec la configuration
        } else {
          console.log('ℹ️ Service unifié prêt - ajoutez des dossiers pour déclencher la sync PowerShell');
        }
      }).catch((error) => {
        console.error('❌ Erreur initialisation service unifié:', error.message);
      });
      
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 4,
          progress: 100,
          message: configFound ? 'Monitoring configuré' : 'Prêt (config manuelle)'
        });
      }

      // Le service unifié remplace à la fois le monitoring et les métriques VBA
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 4,
          progress: 80,
          message: 'Service unifié configuré...'
        });
      }
      
      try {
        // Plus besoin du VBAMetricsService séparé - tout est dans le service unifié
        logClean('📊 Service unifié avec métriques intégrées');
        
        // Petit délai pour que l'interface se mette à jour
        await new Promise(resolve => setTimeout(resolve, 200));
        
        if (loadingWindow) {
          loadingWindow.webContents.send('loading-progress', {
            step: 4,
            progress: 100,
            message: 'Service VBA prêt !'
          });
        }
      } catch (vbaError) {
        console.warn('⚠️ Erreur init métriques VBA:', vbaError.message);
        if (loadingWindow) {
          loadingWindow.webContents.send('loading-progress', {
            step: 4,
            progress: 100,
            message: 'VBA en mode dégradé'
          });
        }
      }
    } catch (monitoringError) {
      console.warn('⚠️ Erreur monitoring:', monitoringError.message);
      if (loadingWindow) {
        loadingWindow.webContents.send('loading-progress', {
          step: 4,
          progress: 100,
          message: 'Prêt (mode dégradé)'
        });
      }
    }

    await new Promise(resolve => setTimeout(resolve, 500));

    // Signaler la completion
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-complete');
    }

  } catch (error) {
    console.error('Erreur initialisation Outlook:', error);
    
    if (loadingWindow) {
      loadingWindow.webContents.send('loading-error', {
        message: error.message,
        stack: error.stack
      });
    }
  } finally {
    // Remettre à zéro la protection pour permettre un retry
    isInitializing = false;
    console.log('🔓 Protection d\'initialisation libérée');
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

// Handlers IPC pour l'API de l'application
ipcMain.handle('api-settings-folders-load', async () => {
  try {
    console.log('📁 Chargement de la configuration des dossiers depuis la BDD...');
    
    const databaseService = require('../services/databaseService');
    await databaseService.initialize();
    
    // Récupérer la configuration depuis la base de données uniquement
    const foldersConfig = await databaseService.getFoldersConfiguration();
    
    // Convertir le format tableau en format objet pour l'interface
    const folderCategories = {};
    if (Array.isArray(foldersConfig)) {
      foldersConfig.forEach(folder => {
        folderCategories[folder.path] = {
          category: folder.category,
          name: folder.name
        };
      });
    }
    
    console.log(`✅ ${Object.keys(folderCategories).length} configurations chargées depuis BDD`);
    console.log('🔍 Configuration finale:', folderCategories);
    
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
    
    const databaseService = require('../services/databaseService');
    await databaseService.initialize();
    
    // Sauvegarder UNIQUEMENT dans la base de données (pas de JSON)
    await databaseService.saveFoldersConfiguration(data);
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
    console.log('📁 [IPC] api-folders-tree appelé');
    
    // OPTIMIZED: Vérifier le cache d'abord
    const cachedFolders = cacheService.getFoldersConfig();
    if (cachedFolders) {
      console.log('⚡ [IPC] Structure dossiers depuis cache');
      return cachedFolders;
    }

    await databaseService.initialize();

    // Récupérer UNIQUEMENT les dossiers configurés (monitorés) depuis la BDD optimisée
    const foldersConfig = await databaseService.getFoldersConfiguration();
    console.log(`📁 [IPC] ${foldersConfig.length} dossiers configurés trouvés en BDD`);

    // Récupérer la structure Outlook pour obtenir les compteurs d'emails
    const allFolders = await outlookConnector.getFolderStructure();
    console.log(`📁 [IPC] Structure Outlook récupérée`);

    // Créer la liste des dossiers monitorés seulement
    const monitoredFolders = [];

    foldersConfig.forEach(config => {
      // Chercher le dossier dans la structure Outlook pour obtenir le nombre d'emails
      const outlookFolder = findFolderInStructure(allFolders, config.path);

      monitoredFolders.push({
        path: config.path,
        name: config.name,
        isMonitored: true,
        category: config.category || 'Mails simples',
        emailCount: outlookFolder ? outlookFolder.Count || 0 : 0,
        parentPath: getParentPath(config.path)
      });
    });

    // Calculer les statistiques
    const stats = calculateFolderStats(monitoredFolders);
    
    console.log(`📁 [IPC] Retour de ${monitoredFolders.length} dossiers monitorés`);
    console.log('📁 [IPC] Données détaillées des dossiers:', JSON.stringify(monitoredFolders, null, 2));
    console.log('📁 [IPC] Stats calculées:', stats);

    const result = {
      folders: monitoredFolders, // Uniquement les dossiers monitorés
      stats: stats,
      timestamp: new Date().toISOString()
    };

    // OPTIMIZED: Mettre en cache pour 5 minutes
    cacheService.set('config', 'folders_tree', result, 300);

    console.log('📁 [IPC] Résultat final envoyé au frontend:', JSON.stringify(result, null, 2));

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
    const updated = await databaseService.updateFolderCategory(folderPath, category);
    
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
    const deleted = await databaseService.deleteFolderConfiguration(folderPath);
    
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
  console.log('📊 [IPC] api-stats-summary appelé');
  try {
    // OPTIMIZED: Utiliser le cache intelligent d'abord
    const cachedStats = cacheService.getUIStats();
    if (cachedStats) {
      console.log('⚡ [IPC] Stats depuis cache (ultra-rapide)');
      return cachedStats;
    }

    // Attendre un peu que le service unifié soit prêt si nécessaire
    let waitAttempts = 0;
    while (waitAttempts < 10 && global.unifiedMonitoringService && !global.unifiedMonitoringService.isInitialized) {
      console.log(`⏳ [IPC] Attente initialisation service unifié... ${waitAttempts + 1}/10`);
      await new Promise(resolve => setTimeout(resolve, 200));
      waitAttempts++;
    }
    
    // Utiliser le service unifié si disponible et initialisé
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized) {
      console.log('📊 [IPC] Utilisation service unifié pour stats');
      const stats = await global.unifiedMonitoringService.getStats();
      
      const result = {
        emailsToday: stats.emailsToday || 0,
        treatedToday: stats.treatedToday || 0,
        unreadTotal: stats.unreadTotal || 0,
        totalEmails: stats.totalEmails || 0,
        lastSyncTime: stats.lastSyncTime || new Date().toISOString(),
        monitoringActive: global.unifiedMonitoringService.isMonitoring
      };
      
      // OPTIMIZED: Mettre en cache pour les prochains appels
      cacheService.set('ui', 'dashboard_stats', result, 30); // 30 secondes
      
      console.log('📊 [IPC] Résultat service unifié:', result);
      return result;
    }
    
    // OPTIMIZED: Fallback vers le service optimisé
    console.log('⚠️ [IPC] Service unifié non disponible, utilisation BD optimisée');
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
  console.log('📧 [IPC] api-emails-recent appelé');
  try {
    // OPTIMIZED: Cache intelligent pour emails récents
    const cachedEmails = cacheService.getRecentEmails(50);
    if (cachedEmails) {
      console.log('⚡ [IPC] Emails récents depuis cache');
      return cachedEmails;
    }

    // Attendre un peu que le service unifié soit prêt si nécessaire
    let waitAttempts = 0;
    while (waitAttempts < 10 && global.unifiedMonitoringService && !global.unifiedMonitoringService.isInitialized) {
      console.log(`⏳ [IPC] Attente initialisation service unifié... ${waitAttempts + 1}/10`);
      await new Promise(resolve => setTimeout(resolve, 200));
      waitAttempts++;
    }
    
    // Utiliser le service unifié si disponible et initialisé
    if (global.unifiedMonitoringService && global.unifiedMonitoringService.isInitialized) {
      console.log('📧 [IPC] Utilisation service unifié pour emails récents');
      const emails = await global.unifiedMonitoringService.getRecentEmails(50);
      console.log(`📧 [IPC] ${emails?.length || 0} emails trouvés via service unifié`);
      
      // OPTIMIZED: Mettre en cache
      if (emails) {
        cacheService.set('emails', 'recent_50', emails, 60); // 1 minute
      }
      
      return emails || [];
    }
    
    // OPTIMIZED: Fallback vers le service optimisé
    console.log('⚠️ [IPC] Service unifié non disponible, utilisation BD optimisée');
    await databaseService.initialize();
    const emails = await databaseService.getRecentEmails(50);
    console.log(`📧 [IPC] ${emails?.length || 0} emails trouvés via BD optimisée`);
    
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
    console.log('📧 [IPC] api-recent-emails appelé');
    
    if (global.unifiedMonitoringService) {
      const emails = await global.unifiedMonitoringService.getRecentEmails(20);
      console.log(`✅ [IPC] ${emails.length} emails récents récupérés`);
      return emails;
    } else {
      // Fallback vers databaseService direct
      const databaseService = require('../services/databaseService');
      await databaseService.initialize();
      const emails = await databaseService.getRecentEmails(20);
      console.log(`✅ [IPC] ${emails.length} emails récents (fallback)`);
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
    const databaseService = require('../services/databaseService');
    await databaseService.initialize();
    
    // Charger les paramètres depuis la base de données
    const settings = await databaseService.loadAppSettings();
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
    const databaseService = require('../services/databaseService');
    await databaseService.initialize();
    
    // Sauvegarder dans la base de données
    for (const [section, sectionData] of Object.entries(settings)) {
      for (const [key, value] of Object.entries(sectionData)) {
        const configKey = `${section}.${key}`;
        await databaseService.saveAppConfig(configKey, value);
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
