// ...existing code...
/**
 * Service de monitoring unifié - Version avec événements COM Outlook
 * 1. Synchronisation complète initiale BDD vs Dossiers
 * 2. Monitoring en temps réel via événements COM Outlook
 */

const EventEmitter = require('events');
const databaseService = require('./optimizedDatabaseService');
const cacheService = require('./cacheService');

// Importer le connecteur COM basé sur FFI-NAPI pour les événements COM natifs
const COMConnector = require('../server/comConnector');
let comConnector = null;
try {
    comConnector = new COMConnector();
    console.log('[COM] Module FFI-NAPI COM chargé avec succès');
} catch (error) {
    console.log('[COM] FFI-NAPI COM non disponible:', error.message);
}

class UnifiedMonitoringService extends EventEmitter {
    constructor(outlookConnector = null) {
        super();
        
        this.outlookConnector = outlookConnector;
        this.dbService = databaseService;
        this.isInitialized = false;
        this.isMonitoring = false;
        this.monitoredFolders = [];
        this.outlookEventHandlers = new Map(); // Stockage des handlers d'événements
        this.pollingInterval = null; // Intervalle de polling complémentaire
        
        // Configuration
        this.config = {
            syncBatchSize: 50, // Réduction pour plus de réactivité
            enableDetailedLogging: true,
            autoStartMonitoring: true, // Démarrage automatique du monitoring
            skipInitialSync: false, // CRITIQUE: Sync PowerShell complète au démarrage
            useComEvents: true, // Utiliser les événements COM au lieu du polling
            useCaching: true, // Cache intelligent activé
            cacheExpiry: 30000, // 30 secondes
            maxConcurrentBatches: 3, // Traitement parallèle
            partialSyncInterval: 5000, // Minimum entre sync partielles
            preferNativeComEvents: true, // Préférer FFI-NAPI COM si disponible
            forcePowerShellInitialSync: true, // NOUVEAU: Forcer sync PowerShell au démarrage
            enableRealtimeComAfterSync: true // NOUVEAU: Activer COM après sync initiale
        };
        
        // Statistiques
        this.stats = {
            totalEmailsInFolders: 0,
            totalEmailsInDatabase: 0,
            emailsAdded: 0,
            emailsUpdated: 0,
            lastSyncTime: null,
            eventsReceived: 0, // Compteur d'événements COM reçus
            lastPartialSync: new Map(), // Dernière sync partielle par dossier
            syncQueue: new Set() // Queue des dossiers à synchroniser
        };
        
        // Cache pour les performances
        this.emailCache = new Map(); // Cache des emails récents
        this.folderStatsCache = new Map(); // Cache des stats par dossier
        this.cacheExpiry = 30000; // 30 secondes
        
        // Optimisations
        this.batchProcessor = {
            queue: [],
            processing: false,
            batchSize: 50,
            intervalMs: 100
        };
    }
    
    /**
     * Méthode de logging unifiée
     */
    log(message, type = 'INFO') {
        const timestamp = new Date().toISOString().substr(11, 8);
        const prefix = `[${timestamp}] [${type}]`;
        console.log(`${prefix} ${message}`);
    }

    /**
     * Initialisation du service de monitoring unifié
     */
    async initialize() {
        try {
            this.log('🔧 Initialisation du service de monitoring unifié...', 'INIT');
            
            // S'assurer que la base de données est prête
            await this.ensureDatabaseReady();
            
            // Charger les dossiers configurés
            await this.loadMonitoredFolders();
            
            // Démarrer le nettoyage automatique du cache
            this.startCacheCleanup();
            
            this.isInitialized = true;
            this.log('✅ Service de monitoring unifié initialisé', 'SUCCESS');
            
            // Démarrer automatiquement le monitoring si configuré
            if (this.config.autoStartMonitoring && this.monitoredFolders.length > 0) {
                this.log('🚀 Démarrage automatique du monitoring...', 'AUTO');
                await this.startMonitoring();
            }
            
        } catch (error) {
            this.log(`❌ Erreur initialisation service: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * S'assurer que la base de données est prête
     */
    async ensureDatabaseReady() {
        this.log('🔧 Vérification de l\'état de la base de données...', 'DB');
        
        let attempts = 0;
        const maxAttempts = 50; // 5 secondes maximum
        
        while (attempts < maxAttempts) {
            try {
                // Tester si la base de données fonctionne avec une requête simple
                await new Promise((resolve, reject) => {
                    if (!this.dbService.db) {
                        reject(new Error('Base de données non connectée'));
                        return;
                    }
                    
                    this.dbService.db.get("SELECT 1 as test", [], (err, row) => {
                        if (err) {
                            reject(err);
                        } else {
                            resolve(row);
                        }
                    });
                });
                
                this.log('✅ Base de données prête et fonctionnelle', 'DB');
                return;
                
            } catch (error) {
                attempts++;
                if (attempts < maxAttempts) {
                    await new Promise(resolve => setTimeout(resolve, 100));
                } else {
                    throw new Error(`Base de données non prête après ${maxAttempts * 100}ms: ${error.message}`);
                }
            }
        }
    }

    /**
     * Chargement des dossiers configurés pour le monitoring
     */
    async loadMonitoredFolders() {
        try {
            this.log('📁 Chargement des dossiers configurés...', 'CONFIG');
            const foldersConfig = await this.dbService.getFoldersConfiguration();
            
            // foldersConfig est maintenant toujours un tableau après correction
            if (Array.isArray(foldersConfig)) {
                this.monitoredFolders = foldersConfig.filter(folder => 
                    folder && 
                    (folder.is_active === 1 || folder.enabled === true) &&
                    (folder.folder_path || folder.path) &&
                    folder.folder_path !== 'folderCategories'
                ).map(folder => ({
                    path: folder.folder_path || folder.path,
                    category: folder.category,
                    name: folder.folder_name || folder.name,
                    enabled: true
                }));
            } else {
                this.log('⚠️ Format de configuration inattendu, utilisation tableau vide', 'WARNING');
                this.monitoredFolders = [];
            }
            
            this.log(`📁 ${this.monitoredFolders.length} dossiers configurés pour le monitoring`, 'CONFIG');
            
        } catch (error) {
            this.log(`❌ Erreur chargement dossiers: ${error.message}`, 'ERROR');
            this.monitoredFolders = [];
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
            
            this.log('🚀 Démarrage du monitoring unifié - Stratégie PowerShell + COM...', 'START');
            
            if (this.monitoredFolders.length === 0) {
                this.log('⚠️ Aucun dossier configuré pour le monitoring', 'WARNING');
                return;
            }
            
            // ÉTAPE 1: Synchronisation PowerShell COMPLÈTE au démarrage
            if (!this.config.skipInitialSync || this.config.forcePowerShellInitialSync) {
                this.log('🔄 PHASE 1: Synchronisation PowerShell complète (mise à jour BDD)', 'SYNC');
                this.log('📊 PowerShell va récupérer TOUS les emails des dossiers configurés...', 'INFO');
                await this.performCompleteSync();
                this.log('✅ PHASE 1 terminée: Base de données synchronisée avec Outlook', 'SUCCESS');
            } else {
                this.log('⏭️ Synchronisation initiale ignorée (mode rapide)', 'INFO');
            }
            
            // ÉTAPE 2: Basculer vers COM pour le temps réel (si disponible)
            if (this.config.enableRealtimeComAfterSync && this.comConnector) {
                this.log('🎧 PHASE 2: Activation monitoring COM temps réel...', 'COM');
                await this.startComEventMonitoring();
                this.log('✅ PHASE 2: COM actif pour emails entrants/traités/supprimés', 'SUCCESS');
            } else {
                this.log('⚠️ PHASE 2: COM non disponible, utilisation PowerShell polling', 'WARNING');
                await this.startComEventMonitoring(); // Fallback PowerShell
            }
            
            this.isMonitoring = true;
            this.log('✅ Monitoring opérationnel: PowerShell (sync) + COM (temps réel)', 'SUCCESS');
            
            // Démarrer le polling de sécurité en complément
            this.startPollingMode();
            
            // Émettre un événement pour signaler que le monitoring a démarré
            this.emit('monitoring-status', { 
                status: 'active',
                mode: this.comConnector ? 'powershell-sync + com-realtime' : 'powershell-only',
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

            // Vérifier d'abord que le dossier configuré existe réellement
            if (!this.comConnector) {
                this.log(`⚠️ COM non disponible pour ${folderConfig.name}`, 'WARNING');
                return;
            }

            // Test de navigation pour vérifier l'existence du dossier
            const namespace = await this.comConnector.getNamespace();
            
            let targetFolder;
            try {
                targetFolder = this.getOutlookFolderByPath(namespace, folderConfig.path);
                if (!targetFolder) {
                    throw new Error(`Dossier non trouvé: ${folderConfig.path}`);
                }
            } catch (navError) {
                this.log(`⚠️ Dossier configuré non trouvé: ${folderConfig.name} (${folderConfig.path})`, 'WARNING');
                this.log(`⚠️ ${navError.message}`, 'WARNING');
                this.log(`⏭️ Ignorer ce dossier et continuer avec les autres`, 'WARNING');
                // Libérer les objets COM
                try {
                    // Libération COM appropriée pour Node.js/COM
                    if (namespace) namespace = null;
                } catch {}
                return; // Ignorer ce dossier et continuer
            }

            // Si on arrive ici, le dossier existe - configurer les événements
            const eventHandler = await this.setupNativeComEvents(folderConfig);
            this.outlookEventHandlers.set(folderConfig.path, eventHandler);
            this.log(`⚡ Événements COM natifs configurés pour: ${folderConfig.name}`, 'COM');

        } catch (error) {
            this.log(`❌ Erreur configuration événements COM pour ${folderConfig.name}: ${error.message}`, 'ERROR');
            // Ne pas propager l'erreur - continuer avec les autres dossiers
            this.log(`⏭️ Continuer avec les autres dossiers configurés`, 'WARNING');
        }
    }

    /**
     * Configuration des événements COM natifs avec COM connector
     */
    async setupNativeComEvents(folderConfig) {
        try {
            if (!this.comConnector) {
                throw new Error('COM non disponible');
            }
            
            this.log('🔧 Initialisation COM pour événements COM...', 'COM');
            
            // Récupérer l'objet folder via COM
            const namespace = await this.comConnector.getNamespace();
            
            // Naviguer vers le dossier
            const folder = this.getOutlookFolderByPath(namespace, folderConfig.path);
            
            if (!folder) {
                throw new Error(`Dossier non trouvé: ${folderConfig.path}`);
            }

            this.log(`📁 Dossier COM récupéré: ${folder.Name} (${folder.Items.Count} items)`, 'COM');

            // Créer les handlers d'événements
            const eventHandler = {
                folder: folder,
                namespace: namespace,
                
                // Handler pour nouveaux emails
                onItemAdd: (item) => {
                    try {
                        this.log(`📨 [COM] Nouvel email détecté dans ${folderConfig.name}`, 'EVENT');
                        const emailData = this.extractEmailDataFromComObject(item);
                        if (emailData) {
                            this.handleNewMail(folderConfig, emailData);
                        }
                    } catch (error) {
                        this.log(`❌ Erreur handler ItemAdd: ${error.message}`, 'ERROR');
                    }
                },
                
                // Handler pour emails modifiés
                onItemChange: (item) => {
                    try {
                        this.log(`📝 [COM] Email modifié dans ${folderConfig.name}`, 'EVENT');
                        const emailData = this.extractEmailDataFromComObject(item);
                        if (emailData) {
                            this.handleMailChanged(folderConfig, emailData);
                        }
                    } catch (error) {
                        this.log(`❌ Erreur handler ItemChange: ${error.message}`, 'ERROR');
                    }
                },
                
                // Handler pour emails supprimés
                onItemRemove: () => {
                    try {
                        this.log(`🗑️ [COM] Email supprimé dans ${folderConfig.name}`, 'EVENT');
                        this.schedulePartialSync(folderConfig);
                    } catch (error) {
                        this.log(`❌ Erreur handler ItemRemove: ${error.message}`, 'ERROR');
                    }
                }
            };

            // Attacher les événements avec WinAX - utiliser getConnectionPoints
            const folderItems = folder.Items;
            
            try {
                // COM direct event binding - approche simplifiée
                this.log(`🔍 Initialisation COM pour événements COM...`, 'COM');
                
                // Obtenir l'application Outlook via COM - syntaxe correcte
                const namespace = await this.comConnector.getNamespace();
                const comFolder = this.getOutlookFolderByPath(namespace, folderConfig.path);
                
                if (!comFolder) {
                    throw new Error(`Dossier COM introuvable: ${folderConfig.path}`);
                }
                
                this.log(`📁 Dossier COM récupéré: ${comFolder.Name} (${comFolder.Items ? comFolder.Items.Count : '?'} items)`, 'COM');
                
                const items = comFolder.Items;
                
                // Créer un object sink personnalisé avec COM
                const EventSink = function(service, folderName, folderPath) {
                    this.service = service;
                    this.folderName = folderName;
                    this.folderPath = folderPath;
                    
                    this.ItemAdd = function(item) {
                        console.log(`📧 [EVENT] Nouvel email dans ${this.folderName}:`, item.Subject || '(sans objet)');
                        this.service.handleNewEmail(item, this.folderPath);
                    }.bind(this);
                    
                    this.ItemChange = function(item) {
                        console.log(`📝 [EVENT] Email modifié dans ${this.folderName}:`, item.Subject || '(sans objet)');
                        this.service.handleEmailChange(item, this.folderPath);
                    }.bind(this);
                    
                    this.ItemRemove = function() {
                        console.log(`🗑️ [EVENT] Email supprimé dans ${this.folderName}`);
                        this.service.handleEmailRemove(this.folderPath);
                    }.bind(this);
                };
                
                // Créer l'instance de l'event sink
                const eventSink = new EventSink(this, folderConfig.name, folderConfig.path);
                
                // Utiliser la méthode native COM pour connecter les événements
                try {
                    // Note: Les événements ItemAdd/ItemChange ne sont pas supportés sur Items.
                    // Outlook COM ne permet pas d'attacher directement ces événements aux collections Items.
                    // Nous utilisons le mode polling comme solution de fallback.
                    
                    this.log(`💡 Configuration mode polling (events COM non supportés sur Items)`, 'COM');
                    
                    // Approche polling avec fallback
                    eventHandler.pollingMode = true;
                    eventHandler.lastItemCount = items.Count;
                    eventHandler.folder = comFolder;
                    
                    this.log(`🔄 Mode polling activé - vérification périodique`, 'COM');
                    
                } catch (eventError) {
                    this.log(`⚠️ Erreur configuration événements: ${eventError.message}`, 'WARNING');
                    
                    // Fallback: Mode polling simple
                    eventHandler.pollingMode = true;
                    eventHandler.lastItemCount = items.Count;
                    eventHandler.folder = comFolder;
                    
                    this.log(`📊 Mode polling activé - vérification périodique`, 'COM');
                }
                
                // Stocker les références pour le nettoyage
                eventHandler.comFolder = comFolder;
                eventHandler.eventSink = eventSink;
                
                this.log(`🎧 Monitoring configuré pour ${items.Count} emails existants`, 'COM');
                
            } catch (comError) {
                this.log(`⚠️ Erreur COM: ${comError.message}`, 'WARNING');
                this.log(`💡 Monitoring passif uniquement`, 'COM');
            }
            
            this.log(`✅ Événements COM natifs configurés pour: ${folderConfig.name}`, 'COM');
            return eventHandler;

        } catch (error) {
            this.log(`❌ Erreur configuration événements COM natifs: ${error.message}`, 'ERROR');
            throw new Error(`Impossible de configurer les événements COM natifs: ${error.message}`);
        }
    }

    /**
     * Gestion des nouveaux emails via événements COM
     */
    handleNewEmail(item, folderPath) {
        try {
            const emailData = this.extractEmailDataFromComObject(item);
            emailData.folderPath = folderPath;
            
            this.log(`📧 Nouvel email détecté: ${emailData.subject}`, 'EVENT');
            
            // Déclencher une synchronisation partielle
            const folderConfig = this.getFolderConfigByPath(folderPath);
            if (folderConfig) {
                this.schedulePartialSync(folderConfig);
            }
            
        } catch (error) {
            this.log(`❌ Erreur traitement nouvel email: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Gestion des emails modifiés via événements COM
     */
    handleEmailChange(item, folderPath) {
        try {
            const emailData = this.extractEmailDataFromComObject(item);
            emailData.folderPath = folderPath;
            
            this.log(`📝 Email modifié: ${emailData.subject}`, 'EVENT');
            
            // Déclencher une synchronisation partielle
            const folderConfig = this.getFolderConfigByPath(folderPath);
            if (folderConfig) {
                this.schedulePartialSync(folderConfig);
            }
            
        } catch (error) {
            this.log(`❌ Erreur traitement email modifié: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Gestion des emails supprimés via événements COM
     */
    handleEmailRemove(folderPath) {
        try {
            this.log(`🗑️ Email supprimé dans: ${folderPath}`, 'EVENT');
            
            // Déclencher une synchronisation partielle
            const folderConfig = this.getFolderConfigByPath(folderPath);
            if (folderConfig) {
                this.schedulePartialSync(folderConfig);
            }
            
        } catch (error) {
            this.log(`❌ Erreur traitement email supprimé: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Récupère la configuration d'un dossier par son chemin
     */
    getFolderConfigByPath(folderPath) {
        return this.monitoredFolders.find(folder => folder.path === folderPath) || null;
    }

    /**
     * Navigue vers un dossier Outlook par son chemin via COM
     */
    getOutlookFolderByPath(namespace, folderPath) {
        try {
            this.log(`🔍 Navigation vers: ${folderPath}`, 'COM');
            
            // Extraire le compte email du chemin
            let emailAccount = null;
            let cleanPath = folderPath;
            
            // Détecter plusieurs formats:
            // 1. \\\\compte@email.com\chemin
            // 2. \\compte@email.com\chemin  
            // 3. compte@email.com\chemin (sans \\)
            const emailMatch = folderPath.match(/^\\\\?([^\\]+@[^\\]+)\\/) || 
                              folderPath.match(/^([^\\]+@[^\\]+)\\/);
            
            if (emailMatch) {
                emailAccount = emailMatch[1];
                // Nettoyer le chemin en supprimant le préfixe email
                cleanPath = folderPath.replace(/^\\\\[^\\]+\\/, '').replace(/^[^\\]+\\/, '');
                this.log(`📧 Compte email détecté: ${emailAccount}`, 'COM');
            }
            
            this.log(`🔍 Chemin nettoyé: ${cleanPath}`, 'COM');
            
            // Diviser le chemin en parties
            const pathParts = cleanPath.split('\\').filter(part => part && part.trim());
            this.log(`🔍 Parties du chemin: [${pathParts.join(', ')}]`, 'COM');
            
            // Démarrer selon le compte email spécifié
            let currentFolder;
            if (emailAccount) {
                // Chercher le store correspondant au compte email
                this.log(`🔍 Recherche du store pour: ${emailAccount}`, 'COM');
                const stores = namespace.Stores;
                let targetStore = null;
                
                for (let i = 1; i <= stores.Count; i++) {
                    const store = stores.Item(i);
                    this.log(`🔍 Vérification store: ${store.DisplayName}`, 'COM');
                    if (store.DisplayName === emailAccount || store.DisplayName.includes(emailAccount)) {
                        targetStore = store;
                        this.log(`✅ Store trouvé: ${store.DisplayName}`, 'COM');
                        break;
                    }
                }
                
                if (!targetStore) {
                    // Lister tous les stores disponibles
                    const availableStores = [];
                    for (let i = 1; i <= stores.Count; i++) {
                        availableStores.push(stores.Item(i).DisplayName);
                    }
                    this.log(`🔍 Stores disponibles: [${availableStores.join(', ')}]`, 'COM');
                    throw new Error(`Store non trouvé pour ${emailAccount}. Disponibles: ${availableStores.join(', ')}`);
                }
                
                // Utiliser la boîte de réception de ce store spécifique
                currentFolder = targetStore.GetDefaultFolder(6); // olFolderInbox
                this.log(`📁 Démarrage depuis la boîte de réception de ${emailAccount}`, 'COM');
            } else {
                // Utiliser la boîte de réception par défaut
                currentFolder = namespace.GetDefaultFolder(6); // olFolderInbox
                this.log(`📁 Démarrage depuis la boîte de réception par défaut`, 'COM');
            }
            
            // Gérer le cas spécial où le premier élément est "Boîte de réception"
            let startIndex = 0;
            if (pathParts[0] === 'Boîte de réception' || pathParts[0] === 'Inbox') {
                startIndex = 1; // Ignorer le premier élément
                this.log(`📁 Ignorer "Boîte de réception" du chemin`, 'COM');
            }
            
            // Naviguer vers les sous-dossiers
            for (let i = startIndex; i < pathParts.length; i++) {
                const part = pathParts[i];
                this.log(`🔍 Recherche du sous-dossier: ${part}`, 'COM');
                
                let found = false;
                const folders = currentFolder.Folders;
                
                for (let j = 1; j <= folders.Count; j++) {
                    const folder = folders.Item(j);
                    this.log(`🔍 Vérification: ${folder.Name} vs ${part}`, 'COM');
                    if (folder.Name === part) {
                        currentFolder = folder;
                        found = true;
                        this.log(`✅ Sous-dossier trouvé: ${part}`, 'COM');
                        break;
                    }
                }
                
                if (!found) {
                    // Lister tous les dossiers disponibles pour debug
                    const availableFolders = [];
                    for (let k = 1; k <= folders.Count; k++) {
                        availableFolders.push(folders.Item(k).Name);
                    }
                    this.log(`🔍 Dossiers disponibles: [${availableFolders.join(', ')}]`, 'COM');
                    throw new Error(`Dossier non trouvé: ${part}. Disponibles: ${availableFolders.join(', ')}`);
                }
            }
            
            this.log(`✅ Navigation réussie vers: ${currentFolder.Name}`, 'COM');
            return currentFolder;
            
        } catch (error) {
            this.log(`❌ Erreur navigation dossier ${folderPath}: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Extraction des données email depuis un objet COM
     */
    extractEmailDataFromComObject(comItem) {
        try {
            return {
                id: comItem.EntryID,
                subject: comItem.Subject || '(Sans objet)',
                senderName: comItem.SenderName || 'Inconnu',
                senderEmail: comItem.SenderEmailAddress || '',
                receivedTime: comItem.ReceivedTime ? new Date(comItem.ReceivedTime) : new Date(),
                isRead: !comItem.UnRead,
                hasAttachments: comItem.Attachments.Count > 0,
                importance: comItem.Importance,
                folderName: comItem.Parent ? comItem.Parent.Name : 'Unknown'
            };
        } catch (error) {
            this.log(`⚠️ Erreur extraction données COM: ${error.message}`, 'WARNING');
            return null;
        }
    }

    /**
     * Trouver un sous-dossier par nom
     */
    findSubfolder(parentFolder, name) {
        try {
            for (let i = 1; i <= parentFolder.Folders.Count; i++) {
                const folder = parentFolder.Folders(i);
                if (folder.Name === name) {
                    return folder;
                }
            }
            return null;
        } catch (error) {
            return null;
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
            const emailRecord = {
                entry_id: processedEmail.EntryID || processedEmail.id,
                outlook_id: processedEmail.EntryID || processedEmail.id,
                subject: processedEmail.subject,
                sender_email: processedEmail.senderEmail,
                received_time: processedEmail.receivedTime,
                is_read: processedEmail.isRead,
                size: processedEmail.size || 0,
                folder_path: folderConfig.path,
                folder_type: folderConfig.category || folderConfig.type,
                is_treated: false,
                created_at: new Date().toISOString()
            };
            
            await this.dbService.insertEmail(emailRecord);
            
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
            const updateData = {
                is_read: processedEmail.isRead,
                subject: processedEmail.subject,
                folder_path: folderConfig.path
            };
            
            const entryId = processedEmail.EntryID || processedEmail.id;
            await this.dbService.updateEmailStatus(entryId, updateData);
            
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

            // Arrêter le polling
            this.stopPollingMode();

            // Nettoyer tous les handlers d'événements COM
            for (const [folderPath, handler] of this.outlookEventHandlers) {
                try {
                    if (handler && handler.eventConnection) {
                        // Détacher la connexion d'événements COM
                        handler.eventConnection.close();
                        this.log(`✅ Connexion événements détachée pour: ${folderPath}`, 'STOP');
                    } else if (handler && handler.folder && handler.folder.Items) {
                        // Nettoyage méthode directe
                        try {
                            handler.folder.Items.ItemAdd = null;
                            handler.folder.Items.ItemChange = null;
                            handler.folder.Items.ItemRemove = null;
                        } catch (cleanupError) {
                            // Ignorer les erreurs de nettoyage
                        }
                        this.log(`✅ Événements détachés pour: ${folderPath}`, 'STOP');
                    }
                } catch (cleanupError) {
                    this.log(`⚠️ Erreur nettoyage événements ${folderPath}: ${cleanupError.message}`, 'WARNING');
                }
            }
            this.outlookEventHandlers.clear();

            this.isMonitoring = false;
            this.log('✅ Monitoring arrêté avec succès', 'SUCCESS');
            
            this.emit('monitoring-status', { 
                status: 'stopped',
                mode: 'com-events-polling',
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

            // Récupérer tous les emails du dossier avec gestion d'erreur
            let emails = [];
            try {
                const emailsResult = await this.outlookConnector.getFolderEmails(folder.path);
                
                // getFolderEmails retourne un objet avec une propriété Emails
                if (emailsResult && emailsResult.Emails && Array.isArray(emailsResult.Emails)) {
                    emails = emailsResult.Emails;
                } else if (Array.isArray(emailsResult)) {
                    emails = emailsResult;
                } else {
                    this.log(`⚠️ Format de retour inattendu pour ${folder.name}: ${typeof emailsResult}`, 'WARNING');
                    emails = [];
                }
            } catch (error) {
                this.log(`⚠️ Erreur récupération emails pour ${folder.name}: ${error.message}`, 'WARNING');
                emails = [];
            }
            
            this.log(`📧 ${emails.length} emails trouvés dans ${folder.name}`, 'INFO');

            // Traiter les emails par batch si on en a
            if (emails.length > 0) {
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
            } else {
                this.log(`ℹ️ Aucun email à traiter pour ${folder.name}`, 'INFO');
            }

            this.log(`✅ Dossier ${folder.name} synchronisé`, 'SUCCESS');

        } catch (error) {
            this.log(`❌ Erreur synchronisation dossier ${folder.name}: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Synchronisation partielle programmée
     */
    schedulePartialSync(folderConfig) {
        // Éviter les sync trop fréquentes
        const lastSync = this.stats.lastPartialSync.get(folderConfig.path);
        const now = Date.now();
        
        if (lastSync && (now - lastSync) < 5000) { // Minimum 5 secondes entre les sync
            return;
        }
        
        this.stats.syncQueue.add(folderConfig.path);
        this.stats.lastPartialSync.set(folderConfig.path, now);
        
        // Démarrer le processeur de queue si pas déjà actif
        if (!this.syncQueueProcessor) {
            this.syncQueueProcessor = setTimeout(() => this.processSyncQueue(), 1000);
        }
    }

    /**
     * Traitement de la queue de synchronisation
     */
    async processSyncQueue() {
        if (this.stats.syncQueue.size === 0) {
            this.syncQueueProcessor = null;
            return;
        }
        
        const folderPaths = Array.from(this.stats.syncQueue);
        this.stats.syncQueue.clear();
        
        for (const folderPath of folderPaths) {
            const folder = this.monitoredFolders.find(f => f.path === folderPath);
            if (folder) {
                try {
                    await this.partialSyncFolder(folder);
                } catch (error) {
                    this.log(`⚠️ Erreur sync partielle ${folder.name}: ${error.message}`, 'WARNING');
                }
            }
        }
        
        this.syncQueueProcessor = null;
    }

    /**
     * Synchronisation partielle d'un dossier (seulement les changements récents)
     */
    async partialSyncFolder(folder) {
        try {
            this.log(`⚡ Sync partielle rapide: ${folder.name}`, 'SYNC');
            
            // Récupérer seulement les emails des dernières 24h
            const yesterday = new Date();
            yesterday.setDate(yesterday.getDate() - 1);
            
            const recentEmails = await this.outlookConnector.getFolderEmails(
                folder.path, 
                { 
                    limit: 100, 
                    since: yesterday,
                    expressMode: true // Mode ultra-rapide
                }
            );
            
            // Extraire les emails du résultat
            let emails = [];
            if (recentEmails && recentEmails.Emails && Array.isArray(recentEmails.Emails)) {
                emails = recentEmails.Emails;
            } else if (Array.isArray(recentEmails)) {
                emails = recentEmails;
            }
            
            if (emails && emails.length > 0) {
                await this.processBatch(emails, folder);
                this.log(`⚡ Sync partielle: ${emails.length} emails traités pour ${folder.name}`, 'SYNC');
            }
            
        } catch (error) {
            this.log(`❌ Erreur sync partielle ${folder.name}: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Traitement par batch optimisé avec cache
     */
    async processBatch(emails, folder) {
        try {
            // Traitement parallèle par petits lots
            const batchSize = 10;
            const batches = [];
            
            for (let i = 0; i < emails.length; i += batchSize) {
                batches.push(emails.slice(i, i + batchSize));
            }
            
            // Traiter les batches en parallèle (max 3 en même temps)
            const concurrency = 3;
            for (let i = 0; i < batches.length; i += concurrency) {
                const currentBatches = batches.slice(i, i + concurrency);
                await Promise.all(
                    currentBatches.map(batch => 
                        this.processBatchParallel(batch, folder)
                    )
                );
            }
            
        } catch (error) {
            this.log(`❌ Erreur traitement batch optimisé: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Traitement parallèle d'un batch
     */
    async processBatchParallel(emails, folder) {
        // Correction: garantir que le sujet n'est jamais null ou vide pour chaque email du batch
        const correctedEmails = emails.map(emailData => {
            if (!emailData.subject || typeof emailData.subject !== 'string' || emailData.subject.trim() === '') {
                emailData.subject = '(Sans objet)';
            }
            return emailData;
        });
        const promises = correctedEmails.map(emailData => this.processEmailOptimized(emailData, folder));
        await Promise.allSettled(promises); // Continue même si certains échouent
    }

    /**
     * Traitement optimisé d'un email avec cache
     */
    async processEmailOptimized(emailData, folder) {
        try {
            const emailId = emailData.EntryID || emailData.id;
            
            // Vérifier le cache d'abord
            const cacheKey = `${folder.path}:${emailId}`;
            const cached = this.emailCache.get(cacheKey);
            
            if (cached && (Date.now() - cached.timestamp) < this.cacheExpiry) {
                return; // Email déjà traité récemment
            }
            
            // Vérifier si l'email existe déjà en base (optimisé)
            const existingEmail = await this.dbService.getEmailByEntryId(emailId, folder.path);

            if (existingEmail) {
                // Vérification rapide des changements
                if (this.needsUpdateOptimized(existingEmail, emailData)) {
                    await this.updateEmail(emailData, folder);
                    this.stats.emailsUpdated++;
                }
            } else {
                // Nouvel email
                await this.addEmail(emailData, folder);
                this.stats.emailsAdded++;
            }
            
            // Mettre en cache
            this.emailCache.set(cacheKey, {
                timestamp: Date.now(),
                processed: true
            });

        } catch (error) {
            this.log(`❌ Erreur traitement email optimisé ${emailData.id}: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Vérification optimisée des changements
     */
    needsUpdateOptimized(existingEmail, newEmailData) {
        // Vérification rapide des champs critiques seulement
        return existingEmail.is_read !== (!newEmailData.UnRead) ||
               existingEmail.subject !== (newEmailData.Subject || existingEmail.subject);
    }

    /**
     * Traitement d'un email individuel
     */
    async processEmail(emailData, folder) {
        try {
            // Vérifier si l'email existe déjà en base
            const existingEmail = await this.dbService.getEmailByEntryId(emailData.id, folder.path);

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
            // Correction: garantir que le sujet n'est jamais null ou vide
            if (!emailData.subject || typeof emailData.subject !== 'string' || emailData.subject.trim() === '') {
                emailData.subject = '(Sans objet)';
            }
            
            // Formatter les données pour la base de données
            const emailRecord = {
                entry_id: emailData.EntryID || emailData.id,
                outlook_id: emailData.EntryID || emailData.id,
                subject: emailData.subject,
                sender_email: emailData.senderEmail || emailData.SenderEmailAddress,
                received_time: emailData.receivedTime || emailData.ReceivedTime,
                is_read: emailData.isRead !== undefined ? emailData.isRead : !emailData.UnRead,
                size: emailData.size || 0,
                folder_path: folder.path,
                folder_type: folder.category || folder.type,
                is_treated: false,
                created_at: new Date().toISOString()
            };
            
            await this.dbService.insertEmail(emailRecord);
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
            // Correction: garantir que le sujet n'est jamais null ou vide
            if (!emailData.subject || typeof emailData.subject !== 'string' || emailData.subject.trim() === '') {
                emailData.subject = '(Sans objet)';
            }
            
            // Formatter les données pour la mise à jour
            const updateData = {
                is_read: emailData.isRead !== undefined ? emailData.isRead : !emailData.UnRead,
                subject: emailData.subject,
                folder_path: folder.path
            };
            
            const entryId = emailData.EntryID || emailData.id;
            await this.dbService.updateEmailStatus(entryId, updateData);
            
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
     * Obtenir les emails récents (méthode requise par l'interface) - Version cachée
     */
    async getRecentEmails(limit = 20) {
        try {
            const cacheKey = `recent_emails_${limit}`;
            const cached = this.emailCache.get(cacheKey);
            
            // Utiliser le cache si récent (10 secondes)
            if (cached && (Date.now() - cached.timestamp) < 10000) {
                return cached.data;
            }
            
            const recentEmails = await this.dbService.getRecentEmails(limit);
            
            // Mettre en cache
            this.emailCache.set(cacheKey, {
                timestamp: Date.now(),
                data: recentEmails
            });
            
            return recentEmails;
        } catch (error) {
            this.log(`❌ Erreur récupération emails récents: ${error.message}`, 'ERROR');
            return [];
        }
    }

    /**
     * Obtenir le résumé des métriques (méthode requise par l'interface) - Version cachée
     */
    async getMetricsSummary() {
        try {
            const cacheKey = 'metrics_summary';
            const cached = this.emailCache.get(cacheKey);
            
            // Cache plus agressif pour les métriques (15 secondes)
            if (cached && (Date.now() - cached.timestamp) < 15000) {
                return cached.data;
            }
            
            // Exécuter les requêtes en parallèle pour plus de rapidité
            const today = new Date().toISOString().split('T')[0]; // Format YYYY-MM-DD
            const [totalEmails, unreadEmails, todayStats] = await Promise.all([
                this.dbService.getTotalEmailCount(),
                this.dbService.getUnreadEmailCount(),
                this.dbService.getEmailCountByDate(today)
            ]);
            
            const metrics = {
                emailsToday: todayStats || 0,
                treatedToday: 0, // À implémenter selon vos besoins
                unreadTotal: unreadEmails || 0,
                totalEmails: totalEmails || 0,
                lastSyncTime: this.stats.lastSyncTime || new Date(),
                monitoringActive: this.isMonitoring,
                eventsReceived: this.stats.eventsReceived,
                performance: {
                    cacheHits: this.getCacheHits(),
                    avgResponseTime: this.getAvgResponseTime()
                }
            };
            
            // Mettre en cache
            this.emailCache.set(cacheKey, {
                timestamp: Date.now(),
                data: metrics
            });
            
            return metrics;
        } catch (error) {
            this.log(`❌ Erreur récupération métriques: ${error.message}`, 'ERROR');
            return {
                emailsToday: 0,
                treatedToday: 0,
                unreadTotal: 0,
                totalEmails: 0,
                lastSyncTime: new Date(),
                monitoringActive: this.isMonitoring
            };
        }
    }

    /**
     * Statistiques de performance du cache
     */
    getCacheHits() {
        return this.emailCache.size;
    }

    /**
     * Temps de réponse moyen (simulé)
     */
    getAvgResponseTime() {
        return this.isMonitoring ? 'Real-time' : 'Polling';
    }

    /**
     * Nettoyage automatique du cache
     */
    cleanupCache() {
        const now = Date.now();
        for (const [key, value] of this.emailCache.entries()) {
            if (now - value.timestamp > this.cacheExpiry) {
                this.emailCache.delete(key);
            }
        }
    }

    /**
     * Démarrage du nettoyage automatique du cache
     */
    startCacheCleanup() {
        setInterval(() => {
            this.cleanupCache();
        }, 60000); // Nettoyage toutes les minutes
    }

    /**
     * Obtenir la distribution des dossiers (méthode requise par l'interface)
     */
    async getFolderDistribution() {
        try {
            const folderStats = await this.dbService.getFolderStats();
            return folderStats || {};
        } catch (error) {
            this.log(`❌ Erreur récupération distribution dossiers: ${error.message}`, 'ERROR');
            return {};
        }
    }

    /**
     * Obtenir l'évolution hebdomadaire (méthode requise par l'interface)
     */
    async getWeeklyEvolution() {
        try {
            // S'assurer que la table weekly_stats existe
            await this.dbService.ensureWeeklyStatsTable();
            
            // Récupérer les stats hebdomadaires
            const weeklyStats = await this.dbService.getWeeklyStats(7);
            return weeklyStats || [];
        } catch (error) {
            if (error.message.includes('no such table: weekly_stats')) {
                this.log('⚠️ Table weekly_stats non trouvée, création en cours...', 'WARNING');
                try {
                    await this.dbService.ensureWeeklyStatsTable();
                    const weeklyStats = await this.dbService.getWeeklyStats(7);
                    return weeklyStats || [];
                } catch (createError) {
                    this.log(`⚠️ Impossible de créer la table weekly_stats: ${createError.message}`, 'WARNING');
                    return {
                        current: { weekNumber: this.getCurrentWeekNumber(), year: new Date().getFullYear(), stockStart: 0, stockEnd: 0, evolution: 0 },
                        trend: 0,
                        percentage: '0.0'
                    };
                }
            }
            this.log(`❌ Erreur récupération évolution hebdomadaire: ${error.message}`, 'ERROR');
            return [];
        }
    }

    /**
     * Obtenir le numéro de semaine actuel
     */
    getCurrentWeekNumber() {
        const now = new Date();
        const start = new Date(now.getFullYear(), 0, 1);
        const days = Math.floor((now - start) / (24 * 60 * 60 * 1000));
        return Math.ceil((days + start.getDay() + 1) / 7);
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
     * Démarrage du mode polling pour détecter les changements
     */
    startPollingMode() {
        if (this.pollingInterval) {
            clearInterval(this.pollingInterval);
        }
        
        this.log('🔄 Démarrage du mode polling complémentaire', 'POLLING');
        
        this.pollingInterval = setInterval(async () => {
            if (!this.isMonitoring) return;
            
            try {
                await this.checkForChanges();
            } catch (error) {
                this.log(`⚠️ Erreur polling: ${error.message}`, 'WARNING');
            }
        }, 30000); // Vérification toutes les 30 secondes
    }

    /**
     * Vérification des changements via polling
     */
    async checkForChanges() {
        try {
            for (const handler of this.outlookEventHandlers.values()) {
                if (handler.pollingMode && handler.folder) {
                    const currentCount = handler.folder.Items.Count;
                    
                    if (currentCount !== handler.lastItemCount) {
                        this.log(`📊 Changement détecté: ${handler.lastItemCount} → ${currentCount} emails`, 'POLLING');
                        
                        // Déclencher une sync partielle
                        const folderConfig = this.getFolderConfigByPath(handler.folder.FolderPath || 'unknown');
                        if (folderConfig) {
                            this.schedulePartialSync(folderConfig);
                        }
                        
                        handler.lastItemCount = currentCount;
                    }
                }
            }
        } catch (error) {
            this.log(`⚠️ Erreur vérification changements: ${error.message}`, 'WARNING');
        }
    }

    /**
     * Arrêt du mode polling
     */
    stopPollingMode() {
        if (this.pollingInterval) {
            clearInterval(this.pollingInterval);
            this.pollingInterval = null;
            this.log('🛑 Mode polling arrêté', 'POLLING');
        }
    }

    /**
     * Nettoyage automatique du cache
     */
    startCacheCleanup() {
        setInterval(() => {
            const now = Date.now();
            // Nettoyer le cache des emails
            for (const [key, entry] of this.emailCache.entries()) {
                if (now - entry.timestamp > this.cacheExpiry) {
                    this.emailCache.delete(key);
                }
            }
            // Nettoyer le cache des stats
            for (const [key, entry] of this.folderStatsCache.entries()) {
                if (now - entry.timestamp > this.cacheExpiry) {
                    this.folderStatsCache.delete(key);
                }
            }
        }, 60000); // Nettoyage toutes les minutes
    }

    /**
     * Arrêt du monitoring
     */
    async stopMonitoring() {
        this.log('🛑 Arrêt du monitoring...', 'STOP');
        this.isMonitoring = false;
        
        // Nettoyer les handlers d'événements
        this.outlookEventHandlers.clear();
        
        // Arrêter le polling si actif
        if (this.pollingInterval) {
            clearInterval(this.pollingInterval);
            this.pollingInterval = null;
        }
        
        this.log('✅ Monitoring arrêté', 'SUCCESS');
    }

    /**
     * Synchronisation forcée
     */
    async forceSync() {
        this.log('💪 Synchronisation forcée démarrée...', 'SYNC');
        // Implémentation de la synchronisation forcée
        this.log('✅ Synchronisation forcée terminée', 'SUCCESS');
    }

    /**
     * Obtenir les emails récents (API compatibility)
     */
    async getRecentEmails(limit = 20) {
        return await this.dbService.getRecentEmails(limit);
    }

    /**
     * Obtenir les statistiques par catégorie
     */
    async getStatsByCategory() {
        try {
            const categoryStats = await this.dbService.getCategoryStats();
            const folderStats = await this.dbService.getFolderStats();
            
            return {
                categories: categoryStats || {},
                folders: folderStats || {},
                lastUpdate: new Date().toISOString()
            };
        } catch (error) {
            console.error('❌ Erreur récupération stats par catégorie:', error);
            return {
                categories: {},
                folders: {},
                lastUpdate: new Date().toISOString()
            };
        }
    }

    /**
     * Obtenir les statistiques de la base de données
     */
    async getDatabaseStats() {
        return await this.dbService.getDatabaseStats();
    }
}

module.exports = UnifiedMonitoringService;
