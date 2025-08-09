/**
 * Service de monitoring pour Mail Monitor
 * Gère le démarrage et l'arrêt de la surveillance des emails
 */

const path = require('path');
const { EventEmitter } = require('events');

class MonitoringService extends EventEmitter {
  constructor() {
    super();
    this.isRunning = false;
    this.monitoringInterval = null;
    this.checkInterval = 30000; // 30 secondes pour éviter de bloquer Outlook
    this.outlookConnector = null;
    this.databaseService = null;
    this.outlookEventsService = null; // Nouveau service d'écoute COM
    this.monitoredFolders = {};
    this.initialScanComplete = {}; // Suivi des dossiers déjà scannés initialement
    this.lastCheck = null;
    this.isUsingCOMEvents = false; // Flag pour savoir si l'écoute COM est active
    this.fallbackPollingActive = false; // Polling de secours
  }

  /**
   * Initialise le service de monitoring
   */
  async initialize() {
    try {
      console.log('🔧 [MonitoringService] Initialisation...');
      
      // Charger les services nécessaires
      this.outlookConnector = require('../server/outlookConnector');
      this.databaseService = require('./databaseService');
      
      // Initialiser le service d'écoute des événements COM Outlook
      const OutlookEventsService = require('./outlookEventsService');
      this.outlookEventsService = new OutlookEventsService();
      
      // Configurer les listeners d'événements COM
      this.setupCOMEventListeners();
      
      console.log('✅ [MonitoringService] Service initialisé avec succès');
    } catch (error) {
      console.error('❌ [MonitoringService] Erreur initialisation:', error);
      throw error;
    }
  }

  /**
   * Démarre le monitoring des emails
   */
  async startMonitoring(externalConfig = null) {
    try {
      if (this.isRunning) {
        console.log('⚠️ [MonitoringService] Le monitoring est déjà en cours');
        return { success: false, message: 'Monitoring déjà actif' };
      }

      console.log('🚀 [MonitoringService] Démarrage du monitoring...');

      // Initialiser si nécessaire
      if (!this.outlookConnector || !this.databaseService) {
        await this.initialize();
      }

      // Utiliser la configuration externe si fournie, sinon charger depuis les fichiers
      if (externalConfig) {
        this.monitoredFolders = externalConfig;
        console.log(`📁 [MonitoringService] Configuration externe utilisée: ${Object.keys(this.monitoredFolders).length} dossiers`);
      } else {
        await this.loadMonitoringConfiguration();
      }

      // Vérifier qu'il y a des dossiers à surveiller
      if (Object.keys(this.monitoredFolders).length === 0) {
        throw new Error('Aucun dossier configuré pour le monitoring');
      }

      // Démarrer le cycle de surveillance
      this.isRunning = true;
      
      // Effectuer un scan initial des emails existants
      console.log('🔍 [MonitoringService] Démarrage du scan initial...');
      await this.performInitialScan();
      
      // ARCHITECTURE OPTIMISÉE: Démarrer l'écoute COM en temps réel après le scan initial
      console.log('🔔 [MonitoringService] Activation de l\'écoute COM temps réel...');
      await this.startCOMEventListening();
      
      // Si l'écoute COM échoue, utiliser le polling de secours
      if (!this.isUsingCOMEvents) {
        console.log('⚠️ [MonitoringService] Écoute COM indisponible - utilisation du polling de secours');
        this.startFallbackPolling();
      } else {
        console.log('✅ [MonitoringService] Écoute COM temps réel active - polling de secours en standby');
        this.setupFallbackPollingStandby();
      }

      console.log(`✅ [MonitoringService] Monitoring démarré - ${Object.keys(this.monitoredFolders).length} dossiers surveillés`);
      
      this.emit('monitoring-started', {
        folders: Object.keys(this.monitoredFolders).length,
        interval: this.checkInterval
      });

      return { 
        success: true, 
        message: `Monitoring démarré - ${Object.keys(this.monitoredFolders).length} dossiers surveillés` 
      };

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur démarrage monitoring:', error);
      this.isRunning = false;
      throw error;
    }
  }

  /**
   * Arrête le monitoring des emails
   */
  async stopMonitoring() {
    try {
      if (!this.isRunning) {
        console.log('⚠️ [MonitoringService] Le monitoring n\'est pas en cours');
        return { success: false, message: 'Monitoring non actif' };
      }

      console.log('🛑 [MonitoringService] Arrêt du monitoring...');

      // Arrêter le cycle de surveillance
      this.isRunning = false;
      if (this.monitoringInterval) {
        clearInterval(this.monitoringInterval);
        this.monitoringInterval = null;
      }

      console.log('✅ [MonitoringService] Monitoring arrêté avec succès');
      
      this.emit('monitoring-stopped');

      return { success: true, message: 'Monitoring arrêté' };

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur arrêt monitoring:', error);
      throw error;
    }
  }

  /**
   * NOUVEAU: Configure les listeners d'événements COM
   */
  setupCOMEventListeners() {
    if (!this.outlookEventsService) return;

    // Écouter les nouveaux emails
    this.outlookEventsService.on('newEmail', (emailData) => {
      console.log(`📬 [MonitoringService] Nouvel email COM: ${emailData.subject}`);
      this.handleCOMNewEmail(emailData);
    });

    // Écouter les changements d'état des emails
    this.outlookEventsService.on('emailChanged', (emailData) => {
      console.log(`🔄 [MonitoringService] Email modifié COM: ${emailData.subject}`);
      this.handleCOMEmailChanged(emailData);
    });

    // Écouter les événements groupés
    this.outlookEventsService.on('eventsProcessed', (groupedEvents) => {
      console.log(`📊 [MonitoringService] Traitement groupé: ${groupedEvents.totalEvents} événements`);
      this.handleCOMGroupedEvents(groupedEvents);
    });

    // Écouter les problèmes d'écoute
    this.outlookEventsService.on('listening-failed', (error) => {
      console.error('❌ [MonitoringService] Écoute COM échouée, basculement vers polling');
      this.isUsingCOMEvents = false;
      this.startFallbackPolling();
    });

    console.log('✅ [MonitoringService] Listeners événements COM configurés');
  }

  /**
   * NOUVEAU: Démarre l'écoute des événements COM Outlook
   */
  async startCOMEventListening() {
    try {
      if (!this.outlookEventsService) {
        throw new Error('Service d\'écoute COM non initialisé');
      }

      const folderPaths = Object.keys(this.monitoredFolders);
      const result = await this.outlookEventsService.startListening(folderPaths);

      if (result.success) {
        this.isUsingCOMEvents = true;
        console.log('🔔 [MonitoringService] Écoute COM Outlook activée avec succès');
        this.emit('com-listening-started', { folders: folderPaths.length });
      } else {
        throw new Error(result.message);
      }

    } catch (error) {
      console.error('❌ [MonitoringService] Impossible de démarrer l\'écoute COM:', error);
      this.isUsingCOMEvents = false;
    }
  }

  /**
   * NOUVEAU: Arrête l'écoute des événements COM
   */
  async stopCOMEventListening() {
    try {
      if (this.outlookEventsService && this.isUsingCOMEvents) {
        await this.outlookEventsService.stopListening();
        this.isUsingCOMEvents = false;
        console.log('🛑 [MonitoringService] Écoute COM arrêtée');
      }
    } catch (error) {
      console.error('❌ [MonitoringService] Erreur arrêt écoute COM:', error);
    }
  }

  /**
   * NOUVEAU: Démarre le polling de secours en cas d'échec COM
   */
  startFallbackPolling() {
    if (this.fallbackPollingActive) return;

    console.log('🔄 [MonitoringService] Démarrage du polling de secours...');
    this.fallbackPollingActive = true;

    // Polling moins fréquent (toutes les 2 minutes) car c'est un fallback
    this.monitoringInterval = setInterval(() => {
      this.performMonitoringCheck();
    }, 120000); // 2 minutes

    console.log('✅ [MonitoringService] Polling de secours actif (2 min)');
  }

  /**
   * NOUVEAU: Configure le polling de secours en standby
   */
  setupFallbackPollingStandby() {
    // Polling très léger toutes les 5 minutes pour vérifier que l'écoute COM fonctionne
    this.monitoringInterval = setInterval(() => {
      this.checkCOMHealthAndFallback();
    }, 300000); // 5 minutes

    console.log('✅ [MonitoringService] Polling de secours en standby (5 min)');
  }

  /**
   * NOUVEAU: Vérifie la santé de l'écoute COM et bascule si nécessaire
   */
  async checkCOMHealthAndFallback() {
    try {
      if (!this.isUsingCOMEvents) return;

      const stats = this.outlookEventsService.getListeningStats();
      
      if (!stats.isListening) {
        console.warn('⚠️ [MonitoringService] Écoute COM interrompue, basculement vers polling');
        this.isUsingCOMEvents = false;
        this.startFallbackPolling();
      } else {
        console.log('✅ [MonitoringService] Écoute COM opérationnelle');
      }

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur vérification santé COM:', error);
    }
  }

  /**
   * NOUVEAU: Gère les nouveaux emails détectés via COM
   */
  async handleCOMNewEmail(emailData) {
    try {
      // Traiter immédiatement le nouvel email en base
      const result = await this.databaseService.processCOMNewEmail(emailData);
      
      if (result.processed) {
        console.log(`📧 [MonitoringService] Nouvel email traité via COM: ${emailData.subject}`);
        
        // Émettre l'événement pour mise à jour UI temps réel
        this.emit('newEmail', {
          ...emailData,
          category: this.getFolderCategory(emailData.folderPath),
          timestamp: new Date()
        });
      }

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur traitement nouvel email COM:', error);
    }
  }

  /**
   * NOUVEAU: Gère les changements d'état des emails via COM
   */
  async handleCOMEmailChanged(emailData) {
    try {
      // Mettre à jour l'état de l'email en base
      const result = await this.databaseService.processCOMEmailChange(emailData);
      
      if (result.updated) {
        console.log(`🔄 [MonitoringService] État email mis à jour via COM: ${emailData.subject}`);
        
        // Émettre l'événement pour mise à jour UI temps réel
        this.emit('emailUpdated', {
          ...emailData,
          category: this.getFolderCategory(emailData.folderPath),
          timestamp: new Date()
        });
      }

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur traitement changement email COM:', error);
    }
  }

  /**
   * NOUVEAU: Gère les événements groupés COM
   */
  async handleCOMGroupedEvents(groupedEvents) {
    try {
      // Traiter les événements par dossier
      for (const [folderPath, events] of Object.entries(groupedEvents.groupedEvents)) {
        console.log(`📊 [MonitoringService] Traitement ${events.newEmails} nouveaux + ${events.changedEmails} modifiés dans ${folderPath}`);
      }

      // Émettre un événement de synchronisation terminée
      this.emit('syncCompleted', {
        totalEvents: groupedEvents.totalEvents,
        folders: Object.keys(groupedEvents.groupedEvents).length,
        timestamp: groupedEvents.timestamp
      });

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur traitement événements groupés COM:', error);
    }
  }

  /**
   * Récupère la catégorie d'un dossier
   */
  getFolderCategory(folderPath) {
    const folderConfig = this.monitoredFolders[folderPath];
    return folderConfig ? folderConfig.category : 'autres';
  }

  /**
   * Charge la configuration des dossiers à surveiller
   */
  async loadMonitoringConfiguration() {
    try {
      console.log('📋 [MonitoringService] Chargement de la configuration...');

      // Charger la configuration depuis le fichier de settings
      const fs = require('fs');
      const configPaths = [
        path.join(__dirname, '../../data/folders-config.json'),
        path.join(__dirname, '../../data/settings.json')
      ];
      
      let foundConfig = false;
      
      for (const configPath of configPaths) {
        if (fs.existsSync(configPath)) {
          try {
            const configData = fs.readFileSync(configPath, 'utf8');
            const config = JSON.parse(configData);
            
            // Configuration directe : les dossiers sont à la racine du JSON
            this.monitoredFolders = config;
            
            if (Object.keys(this.monitoredFolders).length > 0) {
              foundConfig = true;
              console.log(`📁 [MonitoringService] ${Object.keys(this.monitoredFolders).length} dossiers configurés depuis ${configPath}`);
              break;
            }
          } catch (parseError) {
            console.log(`⚠️ [MonitoringService] Erreur lecture ${configPath}:`, parseError.message);
          }
        }
      }
      
      if (!foundConfig) {
        // Essayer de récupérer depuis la dernière configuration sauvegardée via IPC
        console.log('💡 [MonitoringService] Tentative de récupération de la config depuis le processus principal...');
        
        // Configuration par défaut si rien n'est trouvé
        this.monitoredFolders = {};
        console.log('⚠️ [MonitoringService] Aucune configuration trouvée - utilisez l\'interface pour configurer des dossiers');
      }

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur chargement configuration:', error);
      this.monitoredFolders = {};
    }
  }

  /**
   * Démarre le cycle de surveillance
   */
  startMonitoringCycle() {
    console.log(`🔄 [MonitoringService] Cycle de surveillance démarré (intervalle: ${this.checkInterval}ms)`);

    // Premier check immédiat
    this.performMonitoringCheck();

    // Programmer les checks suivants
    this.monitoringInterval = setInterval(() => {
      if (this.isRunning) {
        this.performMonitoringCheck();
      }
    }, this.checkInterval);
  }

  /**
   * Effectue une vérification de monitoring
   */
  async performMonitoringCheck() {
    try {
      if (!this.isRunning) return;

      console.log('🔍 [MonitoringService] Vérification de monitoring...');

      // Pour chaque dossier surveillé, vérifier les nouveaux emails
      for (const [folderPath, config] of Object.entries(this.monitoredFolders)) {
        await this.checkFolderForNewEmails(folderPath, config);
      }

      this.emit('monitoring-check-completed', {
        timestamp: new Date(),
        foldersChecked: Object.keys(this.monitoredFolders).length
      });

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur lors de la vérification:', error);
      this.emit('monitoring-error', error);
    }
  }

  /**
   * Vérifie un dossier pour de nouveaux emails et changements d'état
   */
  async checkFolderForNewEmails(folderPath, config) {
    try {
      console.log(`📂 [MonitoringService] Vérification du dossier: ${config.name} (${config.category})`);

      // Vérifier si c'est le premier scan de ce dossier
      const isFirstScan = !this.initialScanComplete[folderPath];
      
      if (isFirstScan) {
        console.log(`🔍 [MonitoringService] Premier scan du dossier ${config.name} - Scan complet`);
        await this.performFullFolderScan(folderPath, config);
        this.initialScanComplete[folderPath] = true;
        return;
      }

      // Sinon, monitoring intelligent des changements
      await this.performIncrementalCheck(folderPath, config);

    } catch (error) {
      console.error(`❌ [MonitoringService] Erreur vérification dossier ${config.name}:`, error);
    }
  }

  /**
   * Effectue un scan complet d'un dossier (premier monitoring)
   */
  async performFullFolderScan(folderPath, config) {
    try {
      console.log(`📊 [MonitoringService] Scan complet initial de ${config.name}...`);
      
      // Récupérer TOUS les emails du dossier (sans limite pour le premier scan)
      const folderEmails = await this.outlookConnector.getFolderEmails(folderPath, 10000);
      
      if (!folderEmails || !folderEmails.Emails) {
        console.log(`⚠️ [MonitoringService] Aucun email trouvé dans ${config.name}`);
        return;
      }

      console.log(`📧 [MonitoringService] ${folderEmails.Emails.length} emails trouvés dans ${config.name}`);

      // Traiter et sauvegarder tous les emails avec leurs statuts complets
      const result = await this.databaseService.processFullFolderScan(folderEmails, config.category);
      
      console.log(`✅ [MonitoringService] Scan initial ${config.name}: ${result.processed} nouveaux, ${result.updated} mis à jour, ${result.skipped} existants`);
      
      this.emit('initial-scan-progress', {
        folder: config.name,
        category: config.category,
        processed: result.processed,
        updated: result.updated,
        total: result.total,
        skipped: result.skipped
      });

    } catch (error) {
      console.error(`❌ [MonitoringService] Erreur scan complet ${config.name}:`, error);
      // Ne pas marquer comme complété en cas d'erreur
      delete this.initialScanComplete[folderPath];
      throw error;
    }
  }

  /**
   * Effectue une vérification incrémentale des changements
   */
  async performIncrementalCheck(folderPath, config) {
    try {
      console.log(`🔄 [MonitoringService] Vérification incrémentale de ${config.name}...`);
      
      // Récupérer seulement les emails récents (limite réduite pour l'efficacité)
      const recentEmails = await this.outlookConnector.getFolderEmailsForMonitoring(folderPath, 100);
      
      if (!recentEmails || !recentEmails.Emails) {
        console.log(`⚠️ [MonitoringService] Aucun email récent dans ${config.name}`);
        return;
      }

      // Effectuer la synchronisation incrémentale
      const result = await this.databaseService.syncIncrementalChanges(recentEmails, config.category);
      
      if (result.changes > 0) {
        console.log(`📧 [MonitoringService] ${result.changes} changements détectés dans ${config.name}: ${result.newEmails} nouveaux, ${result.statusChanges} statuts modifiés`);
        
        this.emit('emails-changed', {
          folder: config.name,
          category: config.category,
          newEmails: result.newEmails,
          statusChanges: result.statusChanges,
          total: result.total,
          changes: result.changes
        });
      } else {
        console.log(`✅ [MonitoringService] Aucun changement dans ${config.name} (${result.total} emails vérifiés)`);
      }

    } catch (error) {
      console.error(`❌ [MonitoringService] Erreur vérification incrémentale ${config.name}:`, error);
    }
  }

  /**
   * Effectue un scan initial des emails existants lors de la première activation
   */
  async performInitialScan() {
    try {
      console.log('🔍 [MonitoringService] Scan initial des emails existants...');

      // Réinitialiser le tracking des scans complets
      this.initialScanComplete = {};

      for (const [folderPath, config] of Object.entries(this.monitoredFolders)) {
        console.log(`📁 [MonitoringService] Préparation du scan initial du dossier: ${config.name}`);
        // Le scan complet sera effectué lors du premier appel à checkFolderForNewEmails
      }

      console.log('✅ [MonitoringService] Scan initial préparé - les dossiers seront scannés lors du premier monitoring');
      this.emit('initial-scan-completed');

    } catch (error) {
      console.error('❌ [MonitoringService] Erreur scan initial:', error);
      this.emit('initial-scan-error', error);
    }
  }

  /**
   * Met à jour l'intervalle de monitoring
   */
  setCheckInterval(intervalMs) {
    this.checkInterval = intervalMs;
    
    if (this.isRunning && this.monitoringInterval) {
      // Redémarrer le cycle avec le nouvel intervalle
      clearInterval(this.monitoringInterval);
      this.startMonitoringCycle();
    }
    
    console.log(`⏱️ [MonitoringService] Intervalle mis à jour: ${intervalMs}ms`);
  }

  /**
   * Retourne le statut du monitoring
   */
  getStatus() {
    return {
      isRunning: this.isRunning,
      foldersMonitored: Object.keys(this.monitoredFolders).length,
      checkInterval: this.checkInterval,
      lastCheck: this.lastCheck || null
    };
  }

  /**
   * Nettoie les ressources
   */
  cleanup() {
    if (this.monitoringInterval) {
      clearInterval(this.monitoringInterval);
      this.monitoringInterval = null;
    }
    this.isRunning = false;
    this.initialScanComplete = {}; // Réinitialiser le tracking des scans
    this.lastCheck = null;
    this.removeAllListeners();
    console.log('🧹 [MonitoringService] Ressources nettoyées');
  }
}

// Export d'une instance unique (singleton)
const monitoringService = new MonitoringService();

module.exports = {
  startMonitoring: (config) => monitoringService.startMonitoring(config),
  stopMonitoring: () => monitoringService.stopMonitoring(),
  getStatus: () => monitoringService.getStatus(),
  setCheckInterval: (interval) => monitoringService.setCheckInterval(interval),
  cleanup: () => monitoringService.cleanup(),
  
  // Export de l'instance pour les tests ou usage avancé
  instance: monitoringService
};
