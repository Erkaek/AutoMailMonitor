/**
 * Service de monitoring optimisé avec événements temps réel
 * Évite le polling en utilisant les événements Outlook natifs
 */

const EventEmitter = require('events');
const databaseService = require('./databaseService'); // Import du singleton

class OptimizedMonitoringService extends EventEmitter {
    constructor(outlookConnector = null) {
        super();
        
        this.outlookConnector = outlookConnector; // Utiliser la connexion existante
        this.dbService = databaseService; // Utiliser le singleton
        this.isInitialized = false;
        this.isMonitoring = false;
        this.monitoredFolders = [];
        this.initialSyncCompleted = false;
        
        // Configuration avec debug détaillé
        this.config = {
            enableEventLogging: true,
            performInitialSync: true,
            autoStartMonitoring: true,
            eventDebounceMs: 1000, // Éviter les événements en rafale
            debugMode: true, // Mode debug activé
            detailedLogging: true // Logs détaillés
        };
        
        // Debouncing des événements
        this.eventDebounceTimers = new Map();
        
        // Compteurs pour diagnostics
        this.diagnostics = {
            pollingCycles: 0,
            foldersScanned: 0,
            emailsProcessed: 0,
            errors: 0,
            lastScanTime: null,
            scanDurations: []
        };
        
        this.log('🚀 Service de monitoring optimisé initialisé', 'INIT');
        this.logDebug('📊 Configuration:', this.config);
        this.logDebug('🔗 OutlookConnector disponible:', !!this.outlookConnector);
    }

    async initialize(foldersConfig = []) {
        try {
            this.log('🔧 Initialisation du service de monitoring...', 'INIT');
            this.logDebug('📥 Configuration reçue:', foldersConfig);
            
            // Convertir la configuration des dossiers si nécessaire
            if (typeof foldersConfig === 'object' && !Array.isArray(foldersConfig)) {
                this.logDebug('🔄 Conversion objet vers tableau...');
                // Convertir objet en tableau
                this.monitoredFolders = Object.entries(foldersConfig).map(([path, config]) => ({
                    path: path,
                    name: config.name,
                    type: config.category,
                    enabled: true
                }));
                this.logDebug('✅ Conversion terminée:', this.monitoredFolders);
            } else {
                this.logDebug('📋 Configuration déjà en tableau');
                // Configuration déjà en tableau
                this.monitoredFolders = foldersConfig.filter(folder => folder.enabled);
            }
            
            this.log(`📁 ${this.monitoredFolders.length} dossiers configurés pour le monitoring`, 'CONFIG');
            this.logDebug('📂 Dossiers détaillés:', this.monitoredFolders);
            
            // Initialiser la base de données
            this.log('🗄️ Initialisation de la base de données...', 'DB');
            await this.dbService.initialize();
            this.log('✅ Base de données initialisée', 'DB');
            
            // Vérifier qu'une connexion Outlook est disponible
            if (!this.outlookConnector) {
                this.log('❌ Aucune connexion Outlook fournie', 'ERROR');
                throw new Error('Connexion Outlook requise pour le mode optimisé');
            }
            
            this.logDebug('🔍 Type de connecteur Outlook:', this.outlookConnector.constructor.name);
            this.logDebug('🔍 Méthodes disponibles:', Object.getOwnPropertyNames(Object.getPrototypeOf(this.outlookConnector)));
            
            // Mode optimisé sans événements temps réel
            // Utilise un polling intelligent avec cache
            this.log('⚡ Mode optimisé avec polling intelligent activé', 'MODE');
            
            this.isInitialized = true;
            this.log('✅ Service de monitoring initialisé avec succès', 'INIT');
            
            // Démarrer le monitoring si configuré
            if (this.config.autoStartMonitoring) {
                this.log('🔄 Démarrage automatique du monitoring...', 'AUTO');
                await this.startMonitoring();
            } else {
                this.log('⏸️ Démarrage automatique désactivé', 'CONFIG');
            }
            
        } catch (error) {
            this.log(`❌ Erreur initialisation: ${error.message}`, 'ERROR');
            this.log(`📜 Stack trace: ${error.stack}`, 'ERROR');
            this.diagnostics.errors++;
            throw error;
        }
    }

    setupOutlookEventHandlers() {
        this.log('🎧 Configuration des gestionnaires d\'événements...');
        
        // Nouveaux emails
        this.outlookConnector.on('newMail', async (data) => {
            this.log('📬 Événement: Nouveau mail détecté');
            await this.handleNewMailEvent(data);
        });
        
        // Email ajouté (plus fiable que newMail)
        this.outlookConnector.on('itemAdd', async (data) => {
            this.log(`➕ Événement: Email ajouté (${data.entryId})`);
            await this.handleItemAddEvent(data);
        });
        
        // Email modifié (lecture, déplacement, etc.)
        this.outlookConnector.on('itemChange', async (data) => {
            this.log(`📝 Événement: Email modifié (${data.entryId}, non-lu: ${data.unread})`);
            await this.handleItemChangeEvent(data);
        });
        
        // Email supprimé
        this.outlookConnector.on('itemRemove', async (data) => {
            this.log('🗑️ Événement: Email supprimé');
            await this.handleItemRemoveEvent(data);
        });
        
        // Écouteurs prêts
        this.outlookConnector.on('listenersReady', () => {
            this.log('✅ Écouteurs d\'événements Outlook prêts');
            this.emit('eventListenersReady');
        });
    }

    async waitForEventListeners() {
        return new Promise((resolve) => {
            if (this.outlookConnector.hasEventListeners) {
                resolve();
            } else {
                this.outlookConnector.once('listenersReady', resolve);
            }
        });
    }

    async startMonitoring() {
        try {
            this.log('🚀 Démarrage du monitoring optimisé...', 'INIT');
            this.logDebug('🚀 État avant démarrage:', {
                isInitialized: this.isInitialized,
                isMonitoring: this.isMonitoring,
                foldersCount: this.monitoredFolders.length,
                outlookConnector: !!this.outlookConnector
            });
            
            if (!this.isInitialized) {
                this.log('❌ Service non initialisé', 'ERROR');
                throw new Error('Service non initialisé');
            }
            
            // Synchronisation initiale si nécessaire
            if (this.config.performInitialSync && !this.initialSyncCompleted) {
                this.log('🔄 Lancement de la synchronisation initiale...', 'SYNC');
                await this.performInitialSync();
            } else if (this.initialSyncCompleted) {
                this.log('✅ Synchronisation initiale déjà effectuée', 'SYNC');
            } else {
                this.log('⏭️ Synchronisation initiale désactivée', 'CONFIG');
            }
            
            // Démarrer le polling intelligent
            this.log('📡 Démarrage du polling intelligent...', 'POLLING');
            this.startIntelligentPolling();
            
            this.isMonitoring = true;
            this.log('✅ Monitoring démarré - polling intelligent actif', 'SUCCESS');
            this.logDebug('✅ État après démarrage:', {
                isMonitoring: this.isMonitoring,
                pollingActive: !!this.pollingInterval
            });
            
            this.emit('monitoringStarted');
            
        } catch (error) {
            this.log(`❌ Erreur démarrage monitoring: ${error.message}`, 'ERROR');
            this.log(`📜 Stack trace: ${error.stack}`, 'ERROR');
            this.diagnostics.errors++;
            throw error;
        }
    }

    startIntelligentPolling() {
        if (this.pollingInterval) {
            this.log('🔄 Arrêt du polling précédent...', 'POLLING');
            clearInterval(this.pollingInterval);
        }
        
        this.log('📡 Démarrage du polling intelligent...', 'POLLING');
        this.logDebug('⏱️ Intervalle configuré: 10 secondes');
        
        // Polling plus fréquent que l'ancien système mais optimisé
        this.pollingInterval = setInterval(async () => {
            if (this.isMonitoring) {
                this.diagnostics.pollingCycles++;
                this.diagnostics.lastScanTime = new Date();
                const startTime = Date.now();
                
                this.log(`🔍 Cycle de polling #${this.diagnostics.pollingCycles}`, 'POLLING');
                await this.checkForChanges();
                
                const duration = Date.now() - startTime;
                this.diagnostics.scanDurations.push(duration);
                
                // Garder seulement les 10 dernières durées
                if (this.diagnostics.scanDurations.length > 10) {
                    this.diagnostics.scanDurations.shift();
                }
                
                this.logDebug(`⏱️ Cycle terminé en ${duration}ms`);
                this.logDebug(`📊 Stats: Cycles=${this.diagnostics.pollingCycles}, Dossiers=${this.diagnostics.foldersScanned}, Emails=${this.diagnostics.emailsProcessed}, Erreurs=${this.diagnostics.errors}`);
            } else {
                this.log('⏸️ Polling suspendu (monitoring arrêté)', 'POLLING');
            }
        }, 10000); // 10 secondes au lieu de 30
        
        this.log('📡 Polling intelligent démarré (intervalle: 10s)', 'POLLING');
    }

    async checkForChanges() {
        try {
            this.log(`🔍 Vérification des changements (${this.monitoredFolders.length} dossiers)`, 'SCAN');
            
            for (let i = 0; i < this.monitoredFolders.length; i++) {
                const folder = this.monitoredFolders[i];
                this.log(`📂 [${i+1}/${this.monitoredFolders.length}] Scan dossier: ${folder.name}`, 'SCAN');
                this.logDebug(`📂 Détails dossier:`, folder);
                
                await this.scanFolderForChanges(folder);
                this.diagnostics.foldersScanned++;
            }
            
            this.log('✅ Vérification des changements terminée', 'SCAN');
            
        } catch (error) {
            this.log(`❌ Erreur vérification changements: ${error.message}`, 'ERROR');
            this.log(`📜 Stack trace: ${error.stack}`, 'ERROR');
            this.diagnostics.errors++;
        }
    }

    async scanFolderForChanges(folder) {
        // Implementation optimisée qui ne vérifie que les changements récents
        // au lieu de tous les emails
        this.log(`🔍 Vérification optimisée du dossier: ${folder.name}`, 'FOLDER');
        this.logDebug(`📁 Chemin: ${folder.path}, Type: ${folder.type}`);
        
        try {
            // Utiliser les méthodes disponibles sur le connecteur Outlook existant
            if (typeof this.outlookConnector.getFolderEmailsForMonitoring === 'function') {
                this.logDebug('🔧 Utilisation de getFolderEmailsForMonitoring...');
                const emailsResult = await this.outlookConnector.getFolderEmailsForMonitoring(folder.path);
                
                // Adapter selon le format retourné
                let emails = [];
                if (Array.isArray(emailsResult)) {
                    emails = emailsResult;
                } else if (emailsResult && Array.isArray(emailsResult.emails)) {
                    emails = emailsResult.emails;
                } else if (emailsResult && emailsResult.data && Array.isArray(emailsResult.data)) {
                    emails = emailsResult.data;
                } else {
                    this.logDebug('⚠️ Format de résultat inattendu pour scan:', emailsResult);
                    emails = [];
                }
                
                this.logDebug(`📧 ${emails.length} emails trouvés pour monitoring`);
                
                // Traiter seulement les 10 emails les plus récents pour optimiser
                const recentEmails = emails.slice(0, 10);
                this.logDebug(`📧 Traitement de ${recentEmails.length} emails récents`);
                
                for (const email of recentEmails) {
                    await this.processEmailForDatabase(email, folder);
                    this.diagnostics.emailsProcessed++;
                }
                
            } else if (typeof this.outlookConnector.getFolderEmails === 'function') {
                this.logDebug('🔧 Utilisation de getFolderEmails...');
                const emailsResult = await this.outlookConnector.getFolderEmails(folder.path);
                
                // Même logique d'adaptation
                let emails = [];
                if (Array.isArray(emailsResult)) {
                    emails = emailsResult;
                } else if (emailsResult && Array.isArray(emailsResult.emails)) {
                    emails = emailsResult.emails;
                } else {
                    emails = [];
                }
                
                this.logDebug(`📧 ${emails.length} emails trouvés`);
                
                // Traiter seulement les 10 emails les plus récents pour optimiser
                const recentEmails = emails.slice(0, 10);
                this.logDebug(`📧 Traitement de ${recentEmails.length} emails récents`);
                
                for (const email of recentEmails) {
                    await this.processEmailForDatabase(email, folder);
                    this.diagnostics.emailsProcessed++;
                }
                
            } else {
                this.logDebug('⚠️ Aucune méthode de récupération d\'emails disponible');
                this.log(`⚠️ Connecteur Outlook incompatible pour le scan de ${folder.name}`, 'WARNING');
            }
            
            this.log(`✅ Scan terminé pour: ${folder.name}`, 'FOLDER');
            
        } catch (error) {
            this.log(`❌ Erreur scan dossier ${folder.name}: ${error.message}`, 'ERROR');
            this.logDebug(`📜 Stack trace: ${error.stack}`);
            this.diagnostics.errors++;
        }
    }

    async stopMonitoring() {
        this.log('⏹️ Arrêt du monitoring...', 'STOP');
        this.logDebug('⏹️ État avant arrêt:', {
            isMonitoring: this.isMonitoring,
            pollingInterval: !!this.pollingInterval,
            debounceTimers: this.eventDebounceTimers.size
        });
        
        this.isMonitoring = false;
        
        // Arrêter le polling
        if (this.pollingInterval) {
            this.log('📡 Arrêt du polling intelligent...', 'POLLING');
            clearInterval(this.pollingInterval);
            this.pollingInterval = null;
            this.log('✅ Polling arrêté', 'POLLING');
        } else {
            this.log('ℹ️ Aucun polling actif à arrêter', 'INFO');
        }
        
        // Arrêter les timers de debouncing
        if (this.eventDebounceTimers.size > 0) {
            this.log(`🔄 Nettoyage de ${this.eventDebounceTimers.size} timers de debouncing...`, 'CLEANUP');
            for (const timer of this.eventDebounceTimers.values()) {
                clearTimeout(timer);
            }
            this.eventDebounceTimers.clear();
            this.log('✅ Timers de debouncing nettoyés', 'CLEANUP');
        }
        
        // Afficher les diagnostics
        this.log('📊 Statistiques de la session:', 'INFO');
        this.logDebug('📊 Diagnostics:', this.getDiagnostics());
        
        this.log('✅ Monitoring arrêté', 'SUCCESS');
        this.emit('monitoringStopped');
    }

    async performInitialSync() {
        this.log('🔄 Démarrage de la synchronisation initiale...', 'SYNC');
        
        try {
            let totalEmails = 0;
            let totalProcessed = 0;
            
            for (const folder of this.monitoredFolders) {
                this.log(`📁 Synchronisation du dossier: ${folder.name}`, 'SYNC');
                this.logDebug(`📁 Chemin: ${folder.path}`);
                
                try {
                    let emails = [];
                    
                    // Utiliser les méthodes disponibles
                    if (typeof this.outlookConnector.getFolderEmailsForMonitoring === 'function') {
                        this.logDebug('🔧 Utilisation de getFolderEmailsForMonitoring pour sync...');
                        const emailsResult = await this.outlookConnector.getFolderEmailsForMonitoring(folder.path);
                        
                        // Vérifier le format du résultat
                        this.logDebug('🔍 Format du résultat:', typeof emailsResult);
                        this.logDebug('🔍 Résultat détaillé:', emailsResult);
                        
                        // Adapter selon le format retourné
                        if (Array.isArray(emailsResult)) {
                            emails = emailsResult;
                        } else if (emailsResult && Array.isArray(emailsResult.emails)) {
                            emails = emailsResult.emails;
                        } else if (emailsResult && emailsResult.data && Array.isArray(emailsResult.data)) {
                            emails = emailsResult.data;
                        } else {
                            this.log(`⚠️ Format de résultat inattendu pour ${folder.name}`, 'WARNING');
                            this.logDebug('📊 Résultat reçu:', emailsResult);
                            emails = [];
                        }
                        
                    } else if (typeof this.outlookConnector.getFolderEmails === 'function') {
                        this.logDebug('🔧 Utilisation de getFolderEmails pour sync...');
                        const emailsResult = await this.outlookConnector.getFolderEmails(folder.path);
                        
                        // Même logique d'adaptation
                        if (Array.isArray(emailsResult)) {
                            emails = emailsResult;
                        } else if (emailsResult && Array.isArray(emailsResult.emails)) {
                            emails = emailsResult.emails;
                        } else {
                            emails = [];
                        }
                        
                    } else {
                        this.log(`⚠️ Aucune méthode de récupération disponible pour ${folder.name}`, 'WARNING');
                        continue;
                    }
                    
                    totalEmails += emails.length;
                    this.logDebug(`📧 ${emails.length} emails récupérés de ${folder.name}`);
                    
                    // Traiter les emails récupérés (limiter à 50 pour la sync initiale)
                    const emailsToProcess = emails.slice(0, 50);
                    this.logDebug(`📧 Traitement de ${emailsToProcess.length} emails pour sync initiale`);
                    
                    for (const emailData of emailsToProcess) {
                        await this.processEmailForDatabase(emailData, folder);
                        totalProcessed++;
                    }
                    
                    this.log(`✅ Dossier ${folder.name}: ${emailsToProcess.length}/${emails.length} emails traités`, 'SYNC');
                    
                } catch (error) {
                    this.log(`❌ Erreur sync dossier ${folder.name}: ${error.message}`, 'ERROR');
                    this.logDebug(`📜 Stack trace: ${error.stack}`);
                    this.diagnostics.errors++;
                }
            }
            
            this.initialSyncCompleted = true;
            this.log(`✅ Synchronisation initiale terminée: ${totalProcessed}/${totalEmails} emails traités`, 'SUCCESS');
            this.logDebug(`📊 Résultat sync: ${totalProcessed} traités sur ${totalEmails} trouvés`);
            
            this.emit('initialSyncCompleted', { totalEmails, totalProcessed });
            
        } catch (error) {
            this.log(`❌ Erreur synchronisation initiale: ${error.message}`, 'ERROR');
            this.log(`📜 Stack trace: ${error.stack}`, 'ERROR');
            this.diagnostics.errors++;
            throw error;
        }
    }

    async handleNewMailEvent(data) {
        // Debouncing pour éviter les événements en rafale
        this.debounceEvent('newMail', async () => {
            this.log('📬 Traitement événement nouveau mail');
            
            // Déclencher une vérification ciblée
            await this.checkForNewEmails();
            
            this.emit('newEmailDetected', data);
        });
    }

    async handleItemAddEvent(data) {
        if (!data.entryId) return;
        
        this.debounceEvent(`itemAdd_${data.entryId}`, async () => {
            try {
                this.log(`➕ Traitement ajout email: ${data.entryId}`);
                
                // Récupérer les détails de l'email
                const emailData = await this.outlookConnector.getEmailByEntryId(data.entryId);
                
                if (emailData) {
                    // Vérifier si l'email est dans un dossier monitoré
                    const folder = this.findMonitoredFolder(emailData.FolderPath);
                    
                    if (folder) {
                        await this.processEmailForDatabase(emailData, folder);
                        
                        this.emit('emailAdded', {
                            email: emailData,
                            folder: folder
                        });
                    }
                }
                
            } catch (error) {
                this.log(`❌ Erreur traitement ajout: ${error.message}`);
            }
        });
    }

    async handleItemChangeEvent(data) {
        if (!data.entryId) return;
        
        this.debounceEvent(`itemChange_${data.entryId}`, async () => {
            try {
                this.log(`📝 Traitement modification email: ${data.entryId}`);
                
                // Récupérer les détails de l'email
                const emailData = await this.outlookConnector.getEmailByEntryId(data.entryId);
                
                if (emailData) {
                    // Vérifier si l'email est dans un dossier monitoré
                    const folder = this.findMonitoredFolder(emailData.FolderPath);
                    
                    if (folder) {
                        // Mettre à jour le statut de l'email
                        await this.updateEmailStatus(emailData, folder);
                        
                        this.emit('emailChanged', {
                            email: emailData,
                            folder: folder,
                            unread: data.unread
                        });
                    }
                }
                
            } catch (error) {
                this.log(`❌ Erreur traitement modification: ${error.message}`);
            }
        });
    }

    async handleItemRemoveEvent(data) {
        this.debounceEvent('itemRemove', async () => {
            this.log('🗑️ Traitement suppression email');
            
            // Marquer les emails supprimés dans la BDD
            await this.markDeletedEmailsInDatabase();
            
            this.emit('emailRemoved', data);
        });
    }

    debounceEvent(eventKey, callback) {
        // Annuler le timer précédent s'il existe
        if (this.eventDebounceTimers.has(eventKey)) {
            clearTimeout(this.eventDebounceTimers.get(eventKey));
        }
        
        // Créer un nouveau timer
        const timer = setTimeout(async () => {
            this.eventDebounceTimers.delete(eventKey);
            await callback();
        }, this.config.eventDebounceMs);
        
        this.eventDebounceTimers.set(eventKey, timer);
    }

    findMonitoredFolder(folderPath) {
        return this.monitoredFolders.find(folder => 
            folderPath && folderPath.includes(folder.path.split('\\').slice(-1)[0])
        );
    }

    async processEmailForDatabase(emailData, folderConfig) {
        try {
            // Vérifier si l'email existe déjà
            const existingEmail = await this.dbService.getEmailByEntryId(emailData.EntryID);
            
            if (existingEmail) {
                // Mettre à jour l'email existant
                await this.updateEmailStatus(emailData, folderConfig);
            } else {
                // Ajouter le nouvel email
                await this.addNewEmailToDatabase(emailData, folderConfig);
            }
            
        } catch (error) {
            this.log(`❌ Erreur traitement BDD: ${error.message}`);
        }
    }

    async addNewEmailToDatabase(emailData, folderConfig) {
        const emailRecord = {
            entry_id: emailData.EntryID,
            subject: emailData.Subject || '',
            sender: emailData.SenderEmailAddress || '',
            received_time: new Date(emailData.ReceivedTime),
            is_read: !emailData.UnRead,
            size: emailData.Size || 0,
            folder_path: emailData.FolderPath,
            folder_type: folderConfig.type,
            is_treated: false,
            created_at: new Date()
        };
        
        await this.dbService.insertEmail(emailRecord);
        
        // Logger l'événement
        if (this.config.enableEventLogging) {
            await this.dbService.logEmailEvent(
                emailRecord.entry_id,
                emailRecord.is_read ? 'read' : 'unread',
                new Date()
            );
        }
        
        this.log(`📧 Nouvel email ajouté: ${emailRecord.subject}`);
    }

    async updateEmailStatus(emailData, folderConfig) {
        try {
            const isRead = !emailData.UnRead;
            
            // Mettre à jour le statut
            await this.dbService.updateEmailStatus(emailData.EntryID, {
                is_read: isRead,
                size: emailData.Size || 0,
                folder_path: emailData.FolderPath
            });
            
            // Logger l'événement de changement de statut
            if (this.config.enableEventLogging) {
                await this.dbService.logEmailEvent(
                    emailData.EntryID,
                    isRead ? 'read' : 'unread',
                    new Date()
                );
            }
            
            // Marquer comme traité si configuré
            const appSettings = await this.dbService.loadAppSettings();
            if (appSettings.monitoring.treatReadEmailsAsProcessed && isRead) {
                await this.dbService.markEmailAsTreated(emailData.EntryID);
                
                if (this.config.enableEventLogging) {
                    await this.dbService.logEmailEvent(
                        emailData.EntryID,
                        'treated',
                        new Date()
                    );
                }
            }
            
            this.log(`📝 Email mis à jour: ${emailData.Subject}`);
            
        } catch (error) {
            this.log(`❌ Erreur mise à jour statut: ${error.message}`);
        }
    }

    async checkForNewEmails() {
        // Vérification ciblée des nouveaux emails (méthode de fallback)
        this.log('🔍 Vérification ciblée des nouveaux emails...');
        
        for (const folder of this.monitoredFolders) {
            try {
                // Récupérer seulement les 10 emails les plus récents
                const recentEmails = await this.outlookConnector.performInitialSync(folder.path);
                
                // Ne traiter que les nouveaux
                for (const emailData of recentEmails.emails.slice(0, 10)) {
                    const exists = await this.dbService.getEmailByEntryId(emailData.EntryID);
                    if (!exists) {
                        await this.processEmailForDatabase(emailData, folder);
                    }
                }
                
            } catch (error) {
                this.log(`❌ Erreur vérification dossier ${folder.name}: ${error.message}`);
            }
        }
    }

    async markDeletedEmailsInDatabase() {
        // Marquer les emails qui ne sont plus présents comme supprimés
        try {
            await this.dbService.markMissingEmailsAsDeleted();
            this.log('🗑️ Emails supprimés marqués en base');
        } catch (error) {
            this.log(`❌ Erreur marquage suppressions: ${error.message}`);
        }
    }

    // Méthodes publiques pour l'API
    async getMonitoringStatus() {
        this.log('📊 Récupération du statut de monitoring...', 'INFO');
        
        const status = {
            isInitialized: this.isInitialized,
            isMonitoring: this.isMonitoring,
            initialSyncCompleted: this.initialSyncCompleted,
            outlookConnected: this.outlookConnector?.connected || false,
            eventListenersActive: this.outlookConnector?.hasEventListeners || false,
            monitoredFoldersCount: this.monitoredFolders.length,
            diagnostics: this.getDiagnostics()
        };
        
        this.logDebug('📊 Statut détaillé:', status);
        this.log(`📊 Statut: Init=${status.isInitialized}, Monitor=${status.isMonitoring}, Sync=${status.initialSyncCompleted}, Outlook=${status.outlookConnected}`, 'INFO');
        
        return status;
    }

    async forceSync() {
        if (!this.isInitialized) {
            throw new Error('Service non initialisé');
        }
        
        this.log('🔄 Synchronisation forcée demandée...');
        this.initialSyncCompleted = false;
        await this.performInitialSync();
    }

    async cleanup() {
        this.log('🧹 Arrêt du service de monitoring...', 'CLEANUP');
        this.logDebug('🧹 État avant cleanup:', {
            isInitialized: this.isInitialized,
            isMonitoring: this.isMonitoring,
            pollingInterval: !!this.pollingInterval,
            eventTimers: this.eventDebounceTimers.size
        });
        
        await this.stopMonitoring();
        
        if (this.outlookConnector && typeof this.outlookConnector.cleanup === 'function') {
            this.log('🔌 Nettoyage du connecteur Outlook...', 'CLEANUP');
            await this.outlookConnector.cleanup();
        } else {
            this.log('⚠️ Pas de méthode cleanup sur le connecteur Outlook', 'WARNING');
        }
        
        this.removeAllListeners();
        
        // Afficher les statistiques finales
        this.log('📊 Statistiques finales:', 'SUCCESS');
        this.logDebug('📊 Diagnostics finaux:', this.getDiagnostics());
        
        this.log('✅ Service de monitoring arrêté', 'SUCCESS');
    }

    log(message, category = 'INFO') {
        const timestamp = new Date().toISOString();
        const prefix = this.getCategoryPrefix(category);
        console.log(`[${timestamp}] OptimizedMonitoring: ${prefix} ${message}`);
    }

    logDebug(message, data = null) {
        if (this.config.detailedLogging) {
            const timestamp = new Date().toISOString();
            if (data) {
                console.log(`[${timestamp}] OptimizedMonitoring: 🔍 ${message}`, data);
            } else {
                console.log(`[${timestamp}] OptimizedMonitoring: 🔍 ${message}`);
            }
        }
    }

    getCategoryPrefix(category) {
        const prefixes = {
            'INIT': '🚀',
            'CONFIG': '⚙️',
            'DB': '🗄️',
            'MODE': '⚡',
            'AUTO': '🔄',
            'POLLING': '📡',
            'SCAN': '🔍',
            'FOLDER': '📂',
            'EMAIL': '📧',
            'WARNING': '⚠️',
            'ERROR': '❌',
            'SUCCESS': '✅',
            'INFO': 'ℹ️'
        };
        return prefixes[category] || 'ℹ️';
    }

    // Méthode pour obtenir les diagnostics
    getDiagnostics() {
        const avgScanTime = this.diagnostics.scanDurations.length > 0 
            ? this.diagnostics.scanDurations.reduce((a, b) => a + b, 0) / this.diagnostics.scanDurations.length 
            : 0;
            
        return {
            ...this.diagnostics,
            averageScanTime: Math.round(avgScanTime),
            uptime: this.diagnostics.lastScanTime ? Date.now() - this.diagnostics.lastScanTime.getTime() : 0
        };
    }
}

module.exports = OptimizedMonitoringService;
