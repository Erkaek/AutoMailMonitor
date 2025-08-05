/**
 * Service de monitoring unifié - Version avec événements COM Outlook
 * 1. Synchronisation complète initiale BDD vs Dossiers
 * 2. Monitoring en temps réel via événements COM Outlook
 */

const EventEmitter = require('events');
const databaseService = require('./databaseService');

class UnifiedMonitoringService extends EventEmitter {
    constructor(outlookConnector = null) {
        super();
        
        this.outlookConnector = outlookConnector;
        this.dbService = databaseService;
        this.isInitialized = false;
        this.isMonitoring = false;
        this.monitoredFolders = [];
        this.outlookEventHandlers = new Map(); // Stockage des handlers d'événements
        
        // Configuration
        this.config = {
            syncBatchSize: 100, // Traiter par lots de 100 emails pour la sync initiale
            enableDetailedLogging: true,
            autoStartMonitoring: true, // Démarrage automatique du monitoring
            skipInitialSync: false, // On garde la sync initiale pour récupérer les emails ajoutés hors ligne
            useComEvents: true // Utiliser les événements COM au lieu du polling
        };
        
        // Statistiques
        this.stats = {
            totalEmailsInFolders: 0,
            totalEmailsInDatabase: 0,
            emailsAdded: 0,
            emailsUpdated: 0,
            lastSyncTime: null,
            eventsReceived: 0 // Compteur d'événements COM reçus
        };
        
        this.log('🚀 Service de monitoring unifié initialisé (mode événements COM)', 'INIT');
    }

    /**
     * Initialisation du service
     */
    async initialize() {
        try {
            this.log('🔧 Initialisation du service de monitoring...', 'INIT');

            // Charger les dossiers configurés pour le monitoring
            await this.loadMonitoredFolders();

            // Vérifier la connexion Outlook
            if (!this.outlookConnector) {
                const OutlookConnector = require('../server/outlookConnector');
                this.outlookConnector = new OutlookConnector();
            }

            // Initialiser la connexion COM avec Outlook
            await this.outlookConnector.initializeComConnection();
            
            this.isInitialized = true;
            this.log('✅ Service de monitoring unifié initialisé', 'SUCCESS');
            
            // Démarrer automatiquement le monitoring si configuré
            if (this.config.autoStartMonitoring) {
                setTimeout(async () => {
                    try {
                        await this.startMonitoring();
                    } catch (error) {
                        this.log(`⚠️ Erreur démarrage automatique: ${error.message}`, 'WARNING');
                    }
                }, 1000);
            }
            
        } catch (error) {
            this.log(`❌ Erreur initialisation: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Chargement des dossiers configurés pour le monitoring
     */
    async loadMonitoredFolders() {
        try {
            this.log('📁 Chargement des dossiers configurés...', 'CONFIG');
            const foldersConfig = await this.dbService.getConfiguration('monitoredFolders');
            
            if (foldersConfig && typeof foldersConfig === 'object' && !Array.isArray(foldersConfig)) {
                this.monitoredFolders = Object.entries(foldersConfig)
                    .filter(([path, config]) => {
                        return path && 
                               path !== 'folderCategories' && 
                               typeof config === 'object' && 
                               config.name && 
                               config.category;
                    })
                    .map(([path, config]) => ({
                        path: path,
                        name: config.name,
                        type: config.category,
                        enabled: true
                    }));
            } else if (Array.isArray(foldersConfig)) {
                this.monitoredFolders = foldersConfig.filter(folder => 
                    folder && 
                    folder.enabled && 
                    folder.path && 
                    folder.path !== 'folderCategories'
                );
            } else {
                this.monitoredFolders = [];
            }
            
            this.log(`📁 ${this.monitoredFolders.length} dossiers configurés pour le monitoring`, 'CONFIG');
            
        } catch (error) {
            this.log(`❌ Erreur chargement dossiers: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Démarrage du monitoring
     */
    async startMonitoring() {
        try {
            if (!this.isInitialized) {
                throw new Error('Service non initialisé');
            }
            
            this.log('🚀 Démarrage du monitoring unifié avec événements COM...', 'START');
            
            if (this.monitoredFolders.length === 0) {
                this.log('⚠️ Aucun dossier configuré pour le monitoring', 'WARNING');
                return;
            }
            
            // Étape 1: Synchronisation initiale pour récupérer les emails ajoutés hors ligne
            if (!this.config.skipInitialSync) {
                this.log('🔄 Synchronisation initiale pour récupérer les changements hors ligne...', 'SYNC');
                await this.performCompleteSync();
            } else {
                this.log('⏭️ Synchronisation initiale ignorée (mode rapide)', 'INFO');
            }
            
            // Étape 2: Démarrage du monitoring en temps réel via événements COM
            await this.startComEventMonitoring();
            
            this.isMonitoring = true;
            this.log('✅ Monitoring démarré avec succès (événements COM)', 'SUCCESS');
            
            // Émettre un événement pour signaler que le monitoring a démarré
            this.emit('monitoring-status', { 
                status: 'active',
                mode: 'com-events',
                folders: this.monitoredFolders.length,
                stats: this.stats
            });
            
        } catch (error) {
            this.log(`❌ Erreur démarrage monitoring: ${error.message}`, 'ERROR');
            this.emit('error', error);
            throw error;
        }
    }

    /**
     * Démarrage du monitoring en temps réel via événements COM
     */
    async startComEventMonitoring() {
        try {
            this.log('🎧 Démarrage du monitoring via événements COM...', 'COM');

            for (const folderConfig of this.monitoredFolders) {
                await this.setupFolderComEvents(folderConfig);
            }

            this.log(`✅ Événements COM configurés pour ${this.monitoredFolders.length} dossiers`, 'COM');
            this.emit('com-monitoring-started', { foldersCount: this.monitoredFolders.length });

        } catch (error) {
            this.log(`❌ Erreur lors du démarrage des événements COM: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Configuration des événements COM pour un dossier spécifique
     */
    async setupFolderComEvents(folderConfig) {
        try {
            this.log(`🎧 Configuration des événements COM pour: ${folderConfig.name}`, 'COM');

            // Récupérer le handler d'événements pour ce dossier
            const eventHandler = await this.outlookConnector.setupFolderEvents(
                folderConfig.path,
                {
                    onNewMail: (mailData) => this.handleNewMail(folderConfig, mailData),
                    onMailChanged: (mailData) => this.handleMailChanged(folderConfig, mailData),
                    onMailDeleted: (mailId) => this.handleMailDeleted(folderConfig, mailId)
                }
            );

            // Stocker le handler pour pouvoir le nettoyer plus tard
            this.outlookEventHandlers.set(folderConfig.path, eventHandler);

            this.log(`✅ Événements COM configurés pour: ${folderConfig.name}`, 'COM');

        } catch (error) {
            this.log(`❌ Erreur configuration événements COM pour ${folderConfig.name}: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Gestionnaire pour les nouveaux emails
     */
    async handleNewMail(folderConfig, mailData) {
        try {
            this.stats.eventsReceived++;
            this.log(`📨 Nouveau mail détecté dans ${folderConfig.name}: ${mailData.subject}`, 'EVENT');

            // Traiter le nouvel email
            const processedEmail = await this.outlookConnector.processEmailData(mailData);
            
            // Sauvegarder en base
            await this.dbService.insertEmail(processedEmail, folderConfig.path);
            
            this.stats.emailsAdded++;
            
            // Émettre l'événement pour l'interface
            this.emit('new-email', {
                folder: folderConfig.name,
                email: processedEmail,
                folderPath: folderConfig.path
            });

            this.log(`✅ Nouveau mail traité: ${processedEmail.subject}`, 'SUCCESS');

        } catch (error) {
            this.log(`❌ Erreur traitement nouveau mail: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Gestionnaire pour les emails modifiés
     */
    async handleMailChanged(folderConfig, mailData) {
        try {
            this.stats.eventsReceived++;
            this.log(`📝 Mail modifié dans ${folderConfig.name}: ${mailData.subject}`, 'EVENT');

            // Traiter l'email modifié
            const processedEmail = await this.outlookConnector.processEmailData(mailData);
            
            // Mettre à jour en base
            await this.dbService.updateEmail(processedEmail, folderConfig.path);
            
            this.stats.emailsUpdated++;
            
            // Émettre l'événement pour l'interface
            this.emit('email-updated', {
                folder: folderConfig.name,
                email: processedEmail,
                folderPath: folderConfig.path
            });

            this.log(`✅ Mail modifié traité: ${processedEmail.subject}`, 'SUCCESS');

        } catch (error) {
            this.log(`❌ Erreur traitement mail modifié: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Gestionnaire pour les emails supprimés
     */
    async handleMailDeleted(folderConfig, mailId) {
        try {
            this.stats.eventsReceived++;
            this.log(`🗑️ Mail supprimé dans ${folderConfig.name}: ${mailId}`, 'EVENT');

            // Supprimer de la base
            await this.dbService.deleteEmail(mailId, folderConfig.path);
            
            // Émettre l'événement pour l'interface
            this.emit('email-deleted', {
                folder: folderConfig.name,
                mailId: mailId,
                folderPath: folderConfig.path
            });

            this.log(`✅ Mail supprimé traité: ${mailId}`, 'SUCCESS');

        } catch (error) {
            this.log(`❌ Erreur traitement mail supprimé: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Arrêt du monitoring et nettoyage des événements COM
     */
    async stopMonitoring() {
        try {
            this.log('🛑 Arrêt du monitoring...', 'STOP');

            // Nettoyer tous les handlers d'événements COM
            for (const [folderPath, handler] of this.outlookEventHandlers) {
                await this.outlookConnector.removeFolderEvents(folderPath, handler);
            }
            this.outlookEventHandlers.clear();

            this.isMonitoring = false;
            this.log('✅ Monitoring arrêté avec succès', 'SUCCESS');
            
            this.emit('monitoring-status', { 
                status: 'stopped',
                mode: 'com-events',
                stats: this.stats
            });

        } catch (error) {
            this.log(`❌ Erreur arrêt monitoring: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Synchronisation complète initiale pour récupérer les emails ajoutés hors ligne
     */
    async performCompleteSync() {
        try {
            this.log('🔄 Début de la synchronisation complète...', 'SYNC');
            const startTime = Date.now();
            
            this.stats.emailsAdded = 0;
            this.stats.emailsUpdated = 0;

            // Traiter chaque dossier configuré
            for (const folder of this.monitoredFolders) {
                await this.syncFolder(folder);
            }

            this.stats.lastSyncTime = new Date();
            const duration = Date.now() - startTime;

            this.log(`✅ Synchronisation complète terminée en ${duration}ms`, 'SYNC');
            this.log(`📊 Résumé: ${this.stats.emailsAdded} ajoutés, ${this.stats.emailsUpdated} mis à jour`, 'STATS');

            this.emit('sync-completed', {
                duration,
                emailsAdded: this.stats.emailsAdded,
                emailsUpdated: this.stats.emailsUpdated,
                folders: this.monitoredFolders.length
            });

        } catch (error) {
            this.log(`❌ Erreur synchronisation complète: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Synchronisation d'un dossier spécifique
     */
    async syncFolder(folder) {
        try {
            this.log(`📁 Synchronisation du dossier: ${folder.name}`, 'SYNC');

            // Récupérer tous les emails du dossier
            const emails = await this.outlookConnector.getFolderEmails(folder.path);
            this.log(`📧 ${emails.length} emails trouvés dans ${folder.name}`, 'INFO');

            // Traiter les emails par batch
            const batchSize = this.config.syncBatchSize;
            for (let i = 0; i < emails.length; i += batchSize) {
                const batch = emails.slice(i, i + batchSize);
                await this.processBatch(batch, folder);
                
                // Émettre un événement de progression
                this.emit('sync-progress', {
                    folder: folder.name,
                    processed: Math.min(i + batchSize, emails.length),
                    total: emails.length,
                    percentage: Math.round((Math.min(i + batchSize, emails.length) / emails.length) * 100)
                });
            }

            this.log(`✅ Dossier ${folder.name} synchronisé`, 'SUCCESS');

        } catch (error) {
            this.log(`❌ Erreur synchronisation dossier ${folder.name}: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Traitement d'un lot d'emails
     */
    async processBatch(emails, folder) {
        try {
            for (const emailData of emails) {
                await this.processEmail(emailData, folder);
            }
        } catch (error) {
            this.log(`❌ Erreur traitement batch: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Traitement d'un email individuel
     */
    async processEmail(emailData, folder) {
        try {
            // Vérifier si l'email existe déjà en base
            const existingEmail = await this.dbService.getEmailById(emailData.id, folder.path);

            if (existingEmail) {
                // Vérifier si une mise à jour est nécessaire
                if (this.needsUpdate(existingEmail, emailData)) {
                    await this.updateEmail(emailData, folder);
                    this.stats.emailsUpdated++;
                }
            } else {
                // Nouvel email, l'ajouter
                await this.addEmail(emailData, folder);
                this.stats.emailsAdded++;
            }

        } catch (error) {
            this.log(`❌ Erreur traitement email ${emailData.id}: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Vérifier si un email nécessite une mise à jour
     */
    needsUpdate(existingEmail, newEmailData) {
        // Comparer les champs importants pour détecter les changements
        return existingEmail.subject !== newEmailData.subject ||
               existingEmail.isRead !== newEmailData.isRead ||
               existingEmail.importance !== newEmailData.importance;
    }

    /**
     * Ajouter un nouvel email
     */
    async addEmail(emailData, folder) {
        try {
            await this.dbService.insertEmail(emailData, folder.path);
            this.emit('email-added', {
                folder: folder.name,
                email: emailData,
                folderPath: folder.path
            });
        } catch (error) {
            this.log(`❌ Erreur ajout email: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Mettre à jour un email existant
     */
    async updateEmail(emailData, folder) {
        try {
            await this.dbService.updateEmail(emailData, folder.path);
            this.emit('email-updated', {
                folder: folder.name,
                email: emailData,
                folderPath: folder.path
            });
        } catch (error) {
            this.log(`❌ Erreur mise à jour email: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Obtenir les statistiques du service
     */
    getStats() {
        return {
            ...this.stats,
            isInitialized: this.isInitialized,
            isMonitoring: this.isMonitoring,
            foldersCount: this.monitoredFolders.length,
            mode: 'com-events'
        };
    }

    /**
     * Obtenir le nombre d'emails en base de données
     */
    async getEmailCountInDatabase() {
        try {
            const count = await this.dbService.getEmailCount();
            return count;
        } catch (error) {
            this.log(`❌ Erreur récupération nombre emails: ${error.message}`, 'ERROR');
            return { total: 0, read: 0, unread: 0 };
        }
    }

    /**
     * Logging avec horodatage
     */
    log(message, level = 'INFO') {
        if (this.config.enableDetailedLogging) {
            const timestamp = new Date().toLocaleTimeString();
            console.log(`[${timestamp}] [${level}] [UnifiedMonitoringService] ${message}`);
        }
    }
}

module.exports = UnifiedMonitoringService;
