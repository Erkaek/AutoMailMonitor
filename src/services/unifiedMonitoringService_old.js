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
            skipInitialSync: false, // MODIFIÉ: On garde la sync initiale pour récupérer les emails ajoutés hors ligne
            useComEvents: true // NOUVEAU: Utiliser les événements COM au lieu du polling
        };
        
        // Statistiques
        this.stats = {
            totalEmailsInFolders: 0,
            totalEmailsInDatabase: 0,
            emailsAdded: 0,
            emailsUpdated: 0,
            lastSyncTime: null,
            eventsReceived: 0 // NOUVEAU: Compteur d'événements COM reçus
        };
        
        this.log('🚀 Service de monitoring unifié initialisé (mode événements COM)', 'INIT');
    }

    async initialize(foldersConfig = []) {
        try {
            this.log('🔧 Initialisation du service de monitoring unifié...', 'INIT');
            
            // Configuration des dossiers - CORRECTION: Filtrer les configurations invalides
            if (typeof foldersConfig === 'object' && !Array.isArray(foldersConfig)) {
                this.monitoredFolders = Object.entries(foldersConfig)
                    .filter(([path, config]) => {
                        // Filtrer les entrées invalides
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
            
            // ÉVITER LA DOUBLE INITIALISATION: la base est déjà initialisée dans l'index.js principal
            this.log('✅ Base de données utilisée (déjà initialisée par l\'application principale)', 'DB');
            
            // Vérifier la connexion Outlook
            if (!this.outlookConnector) {
                throw new Error('Connexion Outlook requise');
            }
            
            // AJOUT: Configurer l'UNIQUE listener PowerShell au niveau du service
            this.setupPowerShellListener();
            
            this.isInitialized = true;
            this.log('✅ Service de monitoring unifié initialisé', 'SUCCESS');
            
            // Démarrer automatiquement le monitoring si configuré
            if (this.config.autoStartMonitoring) {
                this.log('🔄 Démarrage automatique du monitoring...', 'AUTO');
                // CORRECTION: Démarrage immédiat du monitoring, sync en arrière-plan
                setTimeout(async () => {
                    try {
                        await this.startMonitoring();
                        
                        // Puis faire une synchronisation complète en arrière-plan après 5 secondes
                        setTimeout(async () => {
                            this.log('🔄 Lancement de la synchronisation complète en arrière-plan...', 'BACKGROUND');
                            await this.performFullBackgroundSync();
                        }, 5000);
                        
                    } catch (error) {
                        this.log(`⚠️ Erreur démarrage automatique: ${error.message}`, 'WARNING');
                    }
                }, 1000);
            } else {
                this.log('ℹ️ Service prêt - monitoring manuel disponible', 'INFO');
            }
            
        } catch (error) {
            this.log(`❌ Erreur initialisation: ${error.message}`, 'ERROR');
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

    async startMonitoring() {
        try {
            if (!this.isInitialized) {
                throw new Error('Service non initialisé');
            }
            
            this.log('🚀 Démarrage du monitoring unifié avec événements COM...', 'START');
            
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
        this.log('🔄 Début de la synchronisation complète OPTIMISÉE...', 'SYNC');
        
        try {
            // CORRECTION: Synchronisation directe sans timeout pour garantir le traitement
            await this.performSyncInternal();
            this.log('✅ Synchronisation complète terminée avec succès', 'SUCCESS');
            
        } catch (error) {
            this.log(`⚠️ Erreur synchronisation: ${error.message}`, 'WARNING');
            // Continuer même en cas d'erreur
        }
    }

    async performSyncInternal() {
        // Statistiques initiales
        const dbStats = await this.getEmailCountInDatabase();
        this.log(`📊 Emails en base: ${dbStats.total} (lus: ${dbStats.read}, non-lus: ${dbStats.unread})`, 'STATS');
        
        // Vérifier s'il y a des dossiers à synchroniser
        if (this.monitoredFolders.length === 0) {
            this.log(`⚠️ Aucun dossier configuré pour la synchronisation`, 'WARNING');
            return;
        }
        
        // Synchroniser chaque dossier avec gestion d'erreur par dossier
        for (const folder of this.monitoredFolders) {
            try {
                this.log(`📂 Synchronisation du dossier: ${folder.name}`, 'SYNC');
                await this.syncFolder(folder);
            } catch (error) {
                this.log(`❌ Erreur sync dossier ${folder.name}: ${error.message}`, 'ERROR');
                // Continuer avec les autres dossiers
                continue;
            }
        }
        
        // Statistiques finales
        const finalStats = await this.getEmailCountInDatabase();
        this.log(`📊 Synchronisation terminée - Emails en base: ${finalStats.total}`, 'STATS');
        this.log(`📈 Ajoutés: ${this.stats.emailsAdded}, Mis à jour: ${this.stats.emailsUpdated}`, 'STATS');
        
        this.stats.lastSyncTime = new Date();
        this.emit('syncCompleted', this.stats);
    }

    async syncFolder(folder) {
        // VALIDATION: Vérifier que le dossier est valide
        if (!folder || !folder.path || !folder.name || folder.path === 'folderCategories') {
            this.log(`⚠️ Dossier invalide ignoré: ${folder?.name || 'undefined'}`, 'WARNING');
            return;
        }
        
        this.log(`🔍 Récupération de TOUS les emails du dossier: ${folder.name}`, 'FOLDER');
        
        try {
            // Vérifier que le dossier path existe et est valide
            if (!folder.path || folder.path.trim() === '') {
                this.log(`⚠️ Chemin dossier invalide pour ${folder.name}`, 'WARNING');
                return;
            }
            
            // Récupérer TOUS les emails du dossier (sans limite)
            let allEmails = [];
            
            if (typeof this.outlookConnector.getFolderEmails === 'function') {
                this.log(`🔧 Utilisation de getFolderEmails pour récupération RAPIDE (limite: 50)...`, 'FOLDER');
                
                // Réduire la limite pour une sync plus rapide
                const result = await this.outlookConnector.getFolderEmails(folder.path, 50);
                
                if (Array.isArray(result)) {
                    allEmails = result;
                } else if (result && Array.isArray(result.emails)) {
                    allEmails = result.emails;
                } else if (result && result.Emails && Array.isArray(result.Emails)) {
                    allEmails = result.Emails;
                } else {
                    this.log(`⚠️ Format de réponse inattendu pour ${folder.name}`, 'WARNING');
                    this.logDebug('Réponse reçue:', result);
                    allEmails = [];
                }
            } else {
                this.log(`❌ Méthode getFolderEmails non disponible`, 'ERROR');
                return;
            }
            
            this.log(`📧 ${allEmails.length} emails trouvés dans le dossier ${folder.name}`, 'FOLDER');
            this.stats.totalEmailsInFolders += allEmails.length;
            
            // Traiter par petits lots pour un traitement fluide
            const batchSize = Math.min(this.config.syncBatchSize, 10); // Réduire à 10 pour plus de fluidité
            for (let i = 0; i < allEmails.length; i += batchSize) {
                const batch = allEmails.slice(i, i + batchSize);
                this.log(`📦 Traitement rapide du lot ${Math.floor(i/batchSize) + 1}/${Math.ceil(allEmails.length/batchSize)} (${batch.length} emails)`, 'BATCH');
                
                await this.processBatch(batch, folder);
                
                // Pause plus courte pour plus de réactivité
                if (i + batchSize < allEmails.length) {
                    await this.sleep(100); // Pause réduite à 100ms
                }
            }
            
            this.log(`✅ Synchronisation terminée pour ${folder.name}`, 'FOLDER');
            
        } catch (error) {
            this.log(`❌ Erreur sync dossier ${folder.name}: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    async processBatch(emails, folder) {
        this.log(`🔧 Traitement du batch de ${emails.length} emails pour ${folder.name}`, 'BATCH');
        
        for (const email of emails) {
            try {
                await this.processEmail(email, folder);
            } catch (error) {
                this.log(`❌ Erreur traitement email ${email.EntryID}: ${error.message}`, 'ERROR');
            }
        }
        
        this.log(`✅ Batch de ${emails.length} emails traité`, 'BATCH');
    }

    async processEmail(emailData, folder) {
        try {
            this.log(`🔍 Traitement email: ${emailData.Subject?.substring(0, 30) || 'Sans sujet'} (ID: ${emailData.EntryID})`, 'DEBUG');
            
            // Vérifier si l'email existe en base
            const existingEmail = await this.dbService.getEmailByEntryId(emailData.EntryID);
            
            if (existingEmail) {
                this.log(`📧 Email existant trouvé: ${emailData.Subject?.substring(0, 30)}`, 'DEBUG');
                // Mettre à jour si nécessaire
                const needsUpdate = this.needsUpdate(existingEmail, emailData);
                if (needsUpdate) {
                    this.log(`📝 Mise à jour nécessaire pour: ${emailData.Subject?.substring(0, 30)}`, 'UPDATE');
                    await this.updateEmail(emailData, folder);
                    this.stats.emailsUpdated++;
                }
            } else {
                this.log(`📬 Nouvel email à ajouter: ${emailData.Subject?.substring(0, 30)}`, 'NEW');
                // Ajouter le nouvel email
                await this.addEmail(emailData, folder);
                this.stats.emailsAdded++;
                this.log(`✅ Email ajouté avec succès`, 'SUCCESS');
            }
            
        } catch (error) {
            this.log(`❌ Erreur traitement email: ${error.message}`, 'ERROR');
            this.log(`❌ Stack trace: ${error.stack}`, 'ERROR');
        }
    }

    needsUpdate(existingEmail, newEmailData) {
        // Vérifier si les propriétés importantes ont changé
        const currentIsRead = !newEmailData.UnRead;
        const currentSize = newEmailData.Size || 0;
        
        return (
            existingEmail.is_read !== currentIsRead ||
            existingEmail.size !== currentSize ||
            existingEmail.folder_path !== newEmailData.FolderPath
        );
    }

    async addEmail(emailData, folder) {
        const emailRecord = {
            entry_id: emailData.EntryID,
            subject: emailData.Subject || '',
            sender: emailData.SenderEmailAddress || emailData.SenderName || '',
            received_time: new Date(emailData.ReceivedTime),
            is_read: !emailData.UnRead,
            size: emailData.Size || 0,
            folder_path: emailData.FolderPath,
            folder_type: folder.type,
            is_treated: false,
            created_at: new Date()
        };
        
        await this.dbService.insertEmail(emailRecord);
        this.logDebug(`📧 Nouvel email ajouté: ${emailRecord.subject.substring(0, 50)}...`);
        
        // Émettre l'événement de nouvel email
        this.emit('newEmail', {
            type: 'new',
            email: emailRecord,
            folder: folder.name
        });
    }

    async updateEmail(emailData, folder) {
        const isRead = !emailData.UnRead;
        
        await this.dbService.updateEmailStatus(emailData.EntryID, {
            is_read: isRead,
            size: emailData.Size || 0,
            folder_path: emailData.FolderPath
        });
        
        // Marquer comme traité si configuré
        const appSettings = await this.dbService.loadAppSettings();
        if (appSettings.monitoring.treatReadEmailsAsProcessed && isRead) {
            await this.dbService.markEmailAsTreated(emailData.EntryID);
        }
        
        this.logDebug(`📝 Email mis à jour: ${emailData.Subject?.substring(0, 50) || 'Sans sujet'}...`);
        
        // Émettre l'événement de mise à jour d'email
        this.emit('emailUpdated', {
            type: 'updated',
            email: {
                entry_id: emailData.EntryID,
                subject: emailData.Subject || '',
                is_read: isRead,
                folder_path: emailData.FolderPath
            },
            folder: folder.name
        });
    }

    startStreamingMonitoring() {
        this.log('📡 Démarrage du monitoring en streaming...', 'STREAM');
        
        this.monitoringInterval = setInterval(async () => {
            if (this.isMonitoring) {
                this.stats.monitoringCycles++;
                this.log(`🔍 Cycle de monitoring #${this.stats.monitoringCycles}`, 'STREAM');
                
                await this.checkForChanges();
                
                // Émettre l'événement de fin de cycle de monitoring
                this.emit('monitoringCycleComplete', {
                    cycleNumber: this.stats.monitoringCycles,
                    timestamp: new Date(),
                    foldersChecked: this.monitoredFolders.length
                });
            }
        }, this.config.pollingInterval);
        
        this.log(`📡 Monitoring streaming actif (intervalle: ${this.config.pollingInterval/1000}s)`, 'STREAM');
    }

    async checkForChanges() {
        try {
            // Vérifier les changements récents (derniers 50 emails de chaque dossier)
            for (const folder of this.monitoredFolders) {
                await this.checkFolderChanges(folder);
            }
        } catch (error) {
            this.log(`❌ Erreur vérification changements: ${error.message}`, 'ERROR');
        }
    }

    async checkFolderChanges(folder) {
        try {
            // VALIDATION: Vérifier que le dossier est valide
            if (!folder || !folder.path || !folder.name || folder.path === 'folderCategories') {
                this.log(`⚠️ Dossier invalide ignoré pour monitoring: ${folder?.name || 'undefined'}`, 'WARNING');
                return;
            }
            
            // Récupérer les 50 emails les plus récents
            let recentEmails = [];
            
            if (typeof this.outlookConnector.getFolderEmailsForMonitoring === 'function') {
                this.logDebug(`🔧 Vérification des changements récents dans ${folder.name}`);
                const result = await this.outlookConnector.getFolderEmailsForMonitoring(folder.path);
                
                if (result && result.Emails && Array.isArray(result.Emails)) {
                    recentEmails = result.Emails.slice(0, 50); // Les 50 plus récents
                }
            }
            
            // Traiter seulement les nouveaux ou modifiés
            for (const email of recentEmails) {
                const existingEmail = await this.dbService.getEmailByEntryId(email.EntryID);
                
                if (!existingEmail) {
                    // Nouvel email
                    await this.addEmail(email, folder);
                    this.log(`📬 Nouvel email détecté: ${email.Subject?.substring(0, 50) || 'Sans sujet'}`, 'NEW');
                    this.emit('newEmail', email);
                } else if (this.needsUpdate(existingEmail, email)) {
                    // Email modifié
                    await this.updateEmail(email, folder);
                    this.log(`📝 Email modifié: ${email.Subject?.substring(0, 50) || 'Sans sujet'}`, 'UPDATE');
                    this.emit('emailUpdated', email);
                }
            }
            
        } catch (error) {
            this.log(`❌ Erreur vérification dossier ${folder.name}: ${error.message}`, 'ERROR');
        }
    }

    async stopMonitoring() {
        this.log('⏹️ Arrêt du monitoring...', 'STOP');
        
        this.isMonitoring = false;
        
        if (this.monitoringInterval) {
            clearInterval(this.monitoringInterval);
            this.monitoringInterval = null;
        }
        
        this.log('✅ Monitoring arrêté', 'STOP');
        this.emit('monitoringStopped');
    }

    async getMonitoringStatus() {
        const status = {
            isInitialized: this.isInitialized,
            isMonitoring: this.isMonitoring,
            monitoredFoldersCount: this.monitoredFolders.length,
            outlookConnected: this.outlookConnector?.connected || false,
            stats: this.stats
        };
        
        return status;
    }

    async getEmailCountInDatabase() {
        try {
            const result = await this.dbService.getEmailStats();
            return {
                total: result.totalEmails || 0,
                read: (result.totalEmails || 0) - (result.unreadTotal || 0),
                unread: result.unreadTotal || 0,
                treated: result.treatedToday || 0
            };
        } catch (error) {
            this.log(`❌ Erreur récupération stats: ${error.message}`, 'ERROR');
            return { total: 0, read: 0, unread: 0, treated: 0 };
        }
    }

    async forceSync() {
        if (!this.isInitialized) {
            throw new Error('Service non initialisé');
        }
        
        this.log('🔄 Synchronisation forcée demandée...', 'FORCE');
        
        // Réinitialiser les stats
        this.stats.emailsAdded = 0;
        this.stats.emailsUpdated = 0;
        this.stats.totalEmailsInFolders = 0;
        
        await this.performCompleteSync();
    }

    async performFullBackgroundSync() {
        this.log('🔄 Synchronisation complète en arrière-plan...', 'BACKGROUND');
        
        try {
            // Sauvegarder la config actuelle
            const originalSkipSync = this.config.skipInitialSync;
            
            // Forcer la sync
            this.config.skipInitialSync = false;
            
            // Réinitialiser les stats
            this.stats.emailsAdded = 0;
            this.stats.emailsUpdated = 0;
            this.stats.totalEmailsInFolders = 0;
            
            // Faire la sync
            await this.performCompleteSync();
            
            // Restaurer la config
            this.config.skipInitialSync = originalSkipSync;
            
            this.log('✅ Synchronisation en arrière-plan terminée', 'SUCCESS');
            
        } catch (error) {
            this.log(`❌ Erreur sync arrière-plan: ${error.message}`, 'ERROR');
        }
    }

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    log(message, category = 'INFO') {
        const timestamp = new Date().toISOString();
        const prefix = this.getCategoryPrefix(category);
        console.log(`[${timestamp}] UnifiedMonitoring: ${prefix} ${message}`);
    }

    logDebug(message, data = null) {
        if (this.config.enableDetailedLogging) {
            const timestamp = new Date().toISOString();
            if (data) {
                console.log(`[${timestamp}] UnifiedMonitoring: 🔍 ${message}`, data);
            } else {
                console.log(`[${timestamp}] UnifiedMonitoring: 🔍 ${message}`);
            }
        }
    }

    getCategoryPrefix(category) {
        const prefixes = {
            'INIT': '🚀',
            'CONFIG': '⚙️',
            'DB': '🗄️',
            'START': '▶️',
            'AUTO': '🔄',
            'SYNC': '🔄',
            'FOLDER': '📂',
            'BATCH': '📦',
            'NEW': '📬',
            'UPDATE': '📝',
            'STREAM': '📡',
            'FORCE': '💪',
            'STOP': '⏹️',
            'STATS': '📊',
            'DEBUG': '🔍',
            'WARNING': '⚠️',
            'ERROR': '❌',
            'SUCCESS': '✅',
            'INFO': 'ℹ️'
        };
        return prefixes[category] || 'ℹ️';
    }

    async cleanup() {
        this.log('🧹 Nettoyage du service...', 'CLEANUP');
        
        await this.stopMonitoring();
        this.removeAllListeners();
        
        this.log('✅ Service nettoyé', 'SUCCESS');
    }

    // ===== API METHODS POUR COMPATIBILITÉ =====

    async getStats() {
        this.log('📊 UnifiedMonitoring: getStats() appelé', 'DEBUG');
        try {
            const stats = await this.dbService.getEmailStats();
            this.log(`📊 UnifiedMonitoring: Stats reçues: ${JSON.stringify(stats)}`, 'DEBUG');
            
            const result = {
                ...stats,
                ...this.stats,
                lastSyncTime: this.stats.lastSyncTime,
                monitoringActive: this.isMonitoring
            };
            
            this.log(`📊 UnifiedMonitoring: Résultat final: ${JSON.stringify(result)}`, 'DEBUG');
            return result;
        } catch (error) {
            this.log(`❌ Erreur récupération stats: ${error.message}`, 'ERROR');
            return this.stats;
        }
    }

    async getRecentEmails(limit = 50) {
        this.log(`📧 UnifiedMonitoring: getRecentEmails(${limit}) appelé`, 'DEBUG');
        try {
            const emails = await this.dbService.getRecentEmails(limit);
            this.log(`📧 UnifiedMonitoring: ${emails?.length || 0} emails reçus de la DB`, 'DEBUG');
            return emails || [];
        } catch (error) {
            this.log(`❌ Erreur récupération emails récents: ${error.message}`, 'ERROR');
            return [];
        }
    }

    async getDatabaseStats() {
        try {
            return await this.dbService.getDatabaseStats();
        } catch (error) {
            this.log(`❌ Erreur récupération stats BDD: ${error.message}`, 'ERROR');
            return { total: 0, read: 0, unread: 0, withAttachments: 0 };
        }
    }

    async getStatsByCategory() {
        try {
            return await this.dbService.getStatsByCategory();
        } catch (error) {
            this.log(`❌ Erreur récupération stats par catégorie: ${error.message}`, 'ERROR');
            return {};
        }
    }

    // ===== MÉTHODES MÉTRIQUES VBA (REMPLAÇANT VBAMETRICSSERVICE) =====

    async getMetricsSummary() {
        try {
            return await this.dbService.getMetricsSummary();
        } catch (error) {
            this.log(`❌ Erreur récupération métriques: ${error.message}`, 'ERROR');
            return {
                totalEmails: 0,
                readEmails: 0,
                unreadEmails: 0,
                withAttachments: 0,
                averageEmailsPerDay: 0,
                mostActiveFolder: 'Aucun'
            };
        }
    }

    async getFolderDistribution() {
        try {
            return await this.dbService.getFolderDistribution();
        } catch (error) {
            this.log(`❌ Erreur récupération distribution dossiers: ${error.message}`, 'ERROR');
            return [];
        }
    }

    async getWeeklyEvolution() {
        try {
            return await this.dbService.getWeeklyEvolution();
        } catch (error) {
            this.log(`❌ Erreur récupération évolution hebdomadaire: ${error.message}`, 'ERROR');
            return [];
        }
    }
}

module.exports = UnifiedMonitoringService;
