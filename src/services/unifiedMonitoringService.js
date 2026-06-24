// ...existing code...
/**
 * Service de monitoring unifié OPTIMISÉ - Microsoft Graph API + Better-SQLite3
 * Performance maximale avec API REST native + cache intelligent
 */

const EventEmitter = require('events');
const optimizedDatabaseService = require('./optimizedDatabaseService');
const cacheService = require('./cacheService');
const logService = require('./logService');

// NOUVEAU: Connector optimisé Graph API
// CORRECTION: Utiliser le connecteur principal
const optimizedOutlookConnector = require('../server/outlookConnector');

// NOUVEAU: Service d'écoute événements COM
const OutlookEventsService = require('./outlookEventsService');

class UnifiedMonitoringService extends EventEmitter {
    constructor(outlookConnector = null) {
        super();
        
        // Utiliser le connector optimisé si non fourni
        this.outlookConnector = outlookConnector || optimizedOutlookConnector;
        this.dbService = optimizedDatabaseService;
        this.cacheService = cacheService;
        
    this.isInitialized = false;
    this.isMonitoring = false;
    this._isStarting = false;
    this._isStopping = false;
        this.monitoredFolders = [];
        this.outlookEventHandlers = new Map(); // Stockage des handlers d'événements
        this.pollingInterval = null; // Intervalle de polling complémentaire
        
        // NOUVEAU: Gestion dynamique des dossiers configurés
        this.foldersConfigHash = null;
        this.configCheckInterval = null;
        this.lastConfigCheck = null;
        this._hasLoggedMonitoredFoldersCount = false;
        
        // NOUVEAU: Service d'écoute événements COM moderne
        this.outlookEventsService = new OutlookEventsService();
        this.isUsingCOMEvents = false;
        this.fallbackPollingActive = false;
        
        // Configuration
        this.config = {
            syncBatchSize: 50, // Réduction pour plus de réactivité
            enableDetailedLogging: process.env.ENABLE_DETAILED_LOGGING === '1',
            autoStartMonitoring: true, // Démarrage automatique du monitoring
            skipInitialSync: false, // CRITIQUE: Sync PowerShell complète au démarrage
            useComEvents: true, // Utiliser les événements COM au lieu du polling
            useCaching: true, // Cache intelligent activé
            cacheExpiry: 5000, // 5 secondes seulement
            maxConcurrentBatches: 1, // Traitement parallèle (CPU-friendly)
            partialSyncInterval: 1000, // 1 seconde entre sync partielles
            preferNativeComEvents: true, // Préférer FFI-NAPI COM si disponible
            forcePowerShellInitialSync: true, // NOUVEAU: Forcer sync PowerShell au démarrage
            enableRealtimeComAfterSync: true, // NOUVEAU: Activer COM après sync initiale
            configCheckInterval: 30000, // Vérifier la config des dossiers toutes les 30 secondes
            // Throttling démarrage monitoring
            monitorStartBatchSize: 5, // démarrer par paquets de 5 dossiers
            monitorStartDelayMs: 300 // délai entre démarrages d'items au sein d'un batch
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
    this.cacheCleanupInterval = null;
        
        // Optimisations
        this.batchProcessor = {
            queue: [],
            processing: false,
            batchSize: 50,
            intervalMs: 100
        };

        // Réconciliation anti-écart (déplacements/suppressions manqués)
        this._folderReconcileState = {
            lastRunAtByPath: new Map(),
            inFlightByPath: new Map()
        };
    }

    async reconcileFolderFullScan(folder, reason = '') {
        const folderPath = folder?.path;
        if (!folderPath) return;

        const minIntervalMs = 2 * 60 * 1000; // éviter les scans complets trop fréquents
        const lastAt = this._folderReconcileState.lastRunAtByPath.get(folderPath) || 0;
        const now = Date.now();
        if (now - lastAt < minIntervalMs) {
            return;
        }

        if (this._folderReconcileState.inFlightByPath.get(folderPath)) {
            return;
        }

        const runPromise = (async () => {
            this._folderReconcileState.lastRunAtByPath.set(folderPath, now);
            try {
                this.log(`🧹 Réconciliation complète: ${folder.name}${reason ? ` (${reason})` : ''}`, 'SYNC');
                await this.baselineSyncFolder(folder);
            } catch (e) {
                this.log(`⚠️ Réconciliation échouée pour ${folder.name}: ${e.message}`, 'WARNING');
            }
        })();

        this._folderReconcileState.inFlightByPath.set(folderPath, runPromise);
        try {
            await runPromise;
        } finally {
            this._folderReconcileState.inFlightByPath.delete(folderPath);
        }
    }
    
    /**
     * Méthode de logging unifiée
     */
    log(message, type = 'INFO') {
        const timestamp = new Date().toISOString().substr(11, 8);
        const prefix = `[${timestamp}] [${type}]`;
        console.log(`${prefix} ${message}`);
        
        // Utiliser aussi le nouveau système de logs
        try {
            // Mapper les types vers le nouveau système
            const levelMap = {
                'INIT': 'INFO',
                'INFO': 'INFO',
                'SUCCESS': 'SUCCESS',
                'SYNC': 'INFO',
                'COM': 'INFO',
                'DB': 'INFO',
                'WARNING': 'WARN',
                'WARN': 'WARN',
                'ERROR': 'ERROR',
                'START': 'INFO',
                'STOP': 'INFO',
                'AUTO': 'INFO'
            };
            
            const level = levelMap[type] || 'INFO';
            logService.log(level, type, message);
        } catch (e) {
            // Si le nouveau système n'est pas disponible, ignorer silencieusement
        }
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
            if (this.monitoredFolders.length === 0 && this.dbService && this.dbService.db) {
                try {
                    const rawRows = this.dbService.db.prepare('SELECT folder_path, category, folder_name FROM folder_configurations').all();
                    if (rawRows.length > 0) {
                        this.log(`⚠️ Aucune configuration détectée via cache mais ${rawRows.length} lignes en BDD -> rechargement forcé`, 'WARNING');
                        await this.loadMonitoredFolders({ force: true });
                    }
                } catch (e) {
                    this.log(`⚠️ Vérification directe table configs échouée: ${e.message}`, 'WARNING');
                }
            }
            
            // La surveillance de la configuration est gérée par startConfigWatcher()
            
            // Démarrer le nettoyage automatique du cache
            this.startCacheCleanup();
            
            // NOUVEAU: Configurer les listeners d'événements COM modernes
            this.setupModernCOMEventListeners();
            
            this.isInitialized = true;
            this.log('✅ Service de monitoring unifié initialisé', 'SUCCESS');
            
            // Démarrer la surveillance de la configuration des dossiers
            this.startConfigWatcher();
            
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
                    
                    try {
                        // Better-SQLite3 est synchrone
                        const result = this.dbService.db.prepare("SELECT 1 as test").get();
                        resolve(result);
                    } catch (err) {
                        reject(err);
                    }
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
            if (this.config.enableDetailedLogging) {
                this.log('📁 Chargement des dossiers configurés...', 'CONFIG');
            }
            const foldersConfig = await this.dbService.getFoldersConfiguration();
            
            // Calculer le hash de la nouvelle configuration
            const newConfigHash = this.calculateConfigHash(foldersConfig);
            
            // Tracker les chemins supprimés pour le log
            let cleanedFolderPaths = [];
            let repairedFolderPaths = [];
            
            // foldersConfig est maintenant toujours un tableau après correction
            if (Array.isArray(foldersConfig)) {
                const escapeRegExp = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const normalizeFolderPathForMonitoring = (p) => {
                    let s = String(p || '').replace(/\//g, '\\').replace(/\\+/g, '\\').trim();
                    if (!s) return s;

                    // Corriger les cas où le nom de boîte est collé sans antislash (ex: "FlotteAutoBoîte de réception")
                    // -> "FlotteAuto\\Boîte de réception"
                    const inboxNames = ['Boîte de réception', 'Boite de reception', 'Inbox'];
                    for (const inboxName of inboxNames) {
                        const re = new RegExp(`([^\\\\])(${escapeRegExp(inboxName)})`, 'i');
                        if (re.test(s)) {
                            s = s.replace(re, '$1\\\\$2');
                            s = s.replace(/\\+/g, '\\');
                        }
                    }
                    return s;
                };

                this.monitoredFolders = foldersConfig.filter(folder => {
                    // Validations de base
                    if (!folder || !folder.category || folder.folder_name === 'folderCategories') {
                        return false;
                    }
                    if (!(folder.folder_path || folder.folder_name || folder.path)) {
                        return false;
                    }
                    
                    // IMPORTANT: Filtrer les chemins invalides (trop courts - impossible avec Outlook)
                    // Exemples invalides: "11- Tanguy", "Mails simples" (pas d'antislash, chemin trop court)
                    // Exemples valides: "FlotteAuto\\Boîte de réception\\11- Tanguy", "entryId123"
                    const folderPath = (folder.folder_path || folder.path || folder.folder_name || '').trim();
                    if (!folderPath) return false;
                    
                    // Un dossier valide pour Outlook doit:
                    // 1) Contenir au minimum un antislash (boîte\\dossier) OU
                    // 2) Avoir des IDs de résolution (storeId + entryId)
                    const hasBackslash = folderPath.includes('\\') || folderPath.includes('/');
                    const hasIds = (folder.store_id || folder.storeId) && (folder.entry_id || folder.entryId);
                    
                    if (!hasBackslash && !hasIds) {
                        // Chemin invalide: dossier orpheline qui ne pourra jamais être résolu
                        cleanedFolderPaths.push(folderPath);
                        this.log(`🗑️ [CLEANUP] Dossier orpheline (chemin invalide): "${folderPath}" (pas de boîte-mère, pas d'IDs) - suppression en arrière-plan`, 'INFO');
                        
                        // Nettoyer la BDD en arrière-plan (non-bloquant)
                        try {
                            const toDelete = folder.folder_path || folder.path || folder.folder_name;
                            if (toDelete && typeof this.dbService?.deleteFolderConfiguration === 'function') {
                                // Nettoyage asynchrone sans await
                                Promise.resolve().then(() => 
                                    this.dbService.deleteFolderConfiguration(toDelete)
                                ).catch(e => {
                                    this.log(`⚠️ [CLEANUP] Impossible de supprimer "${toDelete}": ${e?.message || e}`, 'WARNING');
                                });
                            }
                        } catch (_) {
                            // Ignorer les erreurs de nettoyage (ne pas bloquer le monitoring)
                        }
                        
                        return false; // Exclure du monitoring
                    }
                    
                    return true;
                }).map(folder => {
                    // Debug réduit: afficher seulement l'essentiel et éviter undefined
                    if (this.config.enableDetailedLogging) {
                        const dbgFolderPath = folder.folder_path || folder.path || folder.folder_name || '';
                        const dbgFolderName = folder.folder_name || folder.name || '';
                        console.log('🔍 DEBUG MAP - Raw folder:', { folder_path: dbgFolderPath, folder_name: dbgFolderName });
                    }
                    
                    const resolvedPath = normalizeFolderPathForMonitoring(folder.folder_path || folder.path || folder.folder_name || '');
                    const mapped = {
                        path: resolvedPath,
                        category: folder.category,
                        name: folder.folder_name || folder.name || resolvedPath.split('\\').pop() || '',
                        storeId: folder.store_id || folder.storeId || null,
                        entryId: folder.entry_id || folder.entryId || null,
                        storeName: folder.store_name || folder.storeName || null,
                        enabled: true
                    };
                    // Debug réduit
                    if (this.config.enableDetailedLogging) {
                        console.log('🎯 DEBUG MAP - Mapped:', { path: mapped.path, category: mapped.category, name: mapped.name });
                    }
                    return mapped;
                });

                // IMPORTANT (reporting): inclure automatiquement le dossier racine "gestionnaire" (ex: 11- Tanguy)
                // pour compter les arrivées dès qu'un mail est détecté dans ce dossier (même non trié).
                const normalize = (p) => normalizeFolderPathForMonitoring(p).toLowerCase();
                const isManagerSeg = (seg) => /^\d{2,}\s*-\s*.+/.test(String(seg || '').trim());
                const extractManagerRoot = (fullPath) => {
                    const raw = normalizeFolderPathForMonitoring(fullPath);
                    const parts = raw.split('\\').map(s => s.trim()).filter(Boolean);
                    for (let i = 0; i < parts.length; i++) {
                        if (isManagerSeg(parts[i])) {
                            return parts.slice(0, i + 1).join('\\');
                        }
                    }
                    return null;
                };

                const existing = new Set(this.monitoredFolders.map(f => normalize(f.path)));
                const toAdd = [];

                for (const f of this.monitoredFolders) {
                    const root = extractManagerRoot(f.path);
                    if (!root) continue;
                    const key = normalize(root);
                    if (!key || existing.has(key)) continue;
                    existing.add(key);
                    toAdd.push({
                        path: root,
                        category: 'Mails simples',
                        name: root.split('\\').pop() || root,
                        storeId: null,
                        entryId: null,
                        storeName: null,
                        enabled: true,
                        _autoScopeRoot: true
                    });
                }

                if (toAdd.length > 0) {
                    this.monitoredFolders = [...toAdd, ...this.monitoredFolders];

                    // Ne logger l'ajout auto qu'une fois par root (évite le spam toutes les 30s)
                    if (!this._autoScopeRootsLogged) {
                        this._autoScopeRootsLogged = new Set();
                    }
                    const newlyLogged = toAdd.filter(f => {
                        const k = normalize(f.path);
                        if (!k) return false;
                        if (this._autoScopeRootsLogged.has(k)) return false;
                        this._autoScopeRootsLogged.add(k);
                        return true;
                    });
                    if (newlyLogged.length > 0) {
                        this.log(`📁 +${newlyLogged.length} dossier(s) racine gestionnaire ajouté(s) automatiquement pour le scope`, 'CONFIG');
                    }
                }

                // Dédupliquer les dossiers monitorés par chemin normalisé (évite doubles monitorings / logs)
                const pickBetter = (a, b) => {
                    if (!a) return b;
                    if (!b) return a;
                    const aAuto = !!a._autoScopeRoot;
                    const bAuto = !!b._autoScopeRoot;
                    if (aAuto !== bAuto) return aAuto ? b : a;
                    const aHasIds = !!(a.entryId || a.entry_id || a.storeId || a.store_id || a.storeName || a.store_name);
                    const bHasIds = !!(b.entryId || b.entry_id || b.storeId || b.store_id || b.storeName || b.store_name);
                    if (aHasIds !== bHasIds) return aHasIds ? a : b;
                    return a;
                };
                const byPath = new Map();
                for (const f of this.monitoredFolders) {
                    const k = normalize(f?.path);
                    if (!k) continue;
                    byPath.set(k, pickBetter(byPath.get(k), f));
                }
                this.monitoredFolders = Array.from(byPath.values());

                // Résumé des corrections effectuées lors du chargement
                if (repairedFolderPaths.length > 0) {
                    this.log(`🛠️ Chemins réparés automatiquement (${repairedFolderPaths.length}): ${repairedFolderPaths.slice(0, 3).join('; ')}`,'CONFIG');
                }
                if (cleanedFolderPaths.length > 0) {
                    this.log(`🗑️ Chemins invalides supprimés (${cleanedFolderPaths.length}): ${cleanedFolderPaths.slice(0, 3).join('; ')}`,'CONFIG');
                }
            } else {
                this.log('⚠️ Format de configuration inattendu, utilisation tableau vide', 'WARNING');
                this.monitoredFolders = [];
            }
            
            // Détecter les changements de configuration
            const configChanged = this.foldersConfigHash !== newConfigHash;
            this.foldersConfigHash = newConfigHash;

            if (configChanged || !this._hasLoggedMonitoredFoldersCount || this.config.enableDetailedLogging) {
                let msg = `📁 ${this.monitoredFolders.length} dossiers configurés pour le monitoring`;
                if (cleanedFolderPaths.length > 0) {
                    msg += ` (${cleanedFolderPaths.length} orphelins supprimés: ${cleanedFolderPaths.slice(0, 3).map(p => `"${p}"`).join(', ')}${cleanedFolderPaths.length > 3 ? '...' : ''})`;
                }
                this.log(msg, 'CONFIG');
                
                // DEBUG: Afficher les chemins EXACTS pour chaque dossier (aide diagnostic)
                if (this.config.enableDetailedLogging && this.monitoredFolders.length > 0) {
                    this.log('📋 Chemins monitorés (détail):', 'DEBUG');
                    this.monitoredFolders.forEach((f, i) => {
                        const fullPath = f.path || '(vide)';
                        const category = f.category || '(sans catégorie)';
                        const hasIds = f.storeId || f.entryId;
                        const idInfo = hasIds ? ` [store: ${f.storeId}, entry: ${f.entryId}]` : '';
                        console.log(`  ${i + 1}. "${fullPath}" (${category})${idInfo}`);
                    });
                }
                
                this._hasLoggedMonitoredFoldersCount = true;
            }
            
            // Si la configuration a changé et qu'on est en cours de monitoring, redémarrer
            if (configChanged && this.isMonitoring) {
                this.log('🔄 Configuration des dossiers modifiée, redémarrage du monitoring...', 'CONFIG');
                await this.restartMonitoring();
            }
            
            return configChanged;
            
        } catch (error) {
            this.log(`❌ Erreur chargement dossiers: ${error.message}`, 'ERROR');
            this.monitoredFolders = [];
            return false;
        }
    }

    /**
     * Calculer un hash de la configuration des dossiers pour détecter les changements
     */
    calculateConfigHash(foldersConfig) {
        if (!Array.isArray(foldersConfig) || foldersConfig.length === 0) {
            return 'empty';
        }
        
        // Créer une signature basée sur les chemins et catégories
        const signature = foldersConfig
            .filter(folder => folder && (folder.folder_path || folder.folder_name || folder.path))
            .map(folder => `${folder.folder_path || folder.folder_name || folder.path}:${folder.category}`)
            .sort()
            .join('|');
            
        return signature;
    }

    /**
     * Redémarrer le monitoring avec la nouvelle configuration
     */
    async restartMonitoring() {
        try {
            this.log('🔄 Redémarrage du monitoring...', 'RESTART');
            
            // Arrêter le monitoring actuel
            await this.stopMonitoring();
            
            // Attendre un peu pour que l'arrêt soit complet
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // Redémarrer avec la nouvelle configuration
            await this.startMonitoring();
            
            this.log('✅ Monitoring redémarré avec succès', 'RESTART');
            
        } catch (error) {
            this.log(`❌ Erreur redémarrage monitoring: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Vérifier périodiquement si la configuration des dossiers a changé
     */
    // startConfigurationWatcher supprimé (doublon)

    /**
     * Vérifier si la configuration a changé
     */
    // checkConfigurationChanges version 1 supprimée (doublon)

    /**
     * Recharger manuellement la configuration des dossiers
     */
    async reloadFoldersConfiguration() {
        this.log('🔄 Rechargement manuel de la configuration...', 'MANUAL');
        
        try {
            const configChanged = await this.loadMonitoredFolders();
            
            if (configChanged) {
                this.log('✅ Configuration rechargée et monitoring redémarré', 'MANUAL');
                return {
                    success: true,
                    foldersCount: this.monitoredFolders.length,
                    folders: this.monitoredFolders.map(f => f.path)
                };
            } else {
                this.log('ℹ️ Aucun changement de configuration détecté', 'MANUAL');
                return {
                    success: true,
                    foldersCount: this.monitoredFolders.length,
                    noChange: true
                };
            }
            
        } catch (error) {
            this.log(`❌ Erreur rechargement configuration: ${error.message}`, 'ERROR');
            return {
                success: false,
                error: error.message
            };
        }
    }

    /**
     * Démarrer la surveillance de la configuration des dossiers
     */
    startConfigWatcher() {
        if (this.configCheckInterval) {
            clearInterval(this.configCheckInterval);
        }
        
        this.configCheckInterval = setInterval(async () => {
            try {
                await this.checkConfigurationChanges();
            } catch (error) {
                this.log(`⚠️ Erreur vérification configuration: ${error.message}`, 'WARNING');
            }
        }, this.config.configCheckInterval);
        
        this.log(`👁️ Surveillance de la configuration démarrée (intervalle: ${this.config.configCheckInterval}ms)`, 'CONFIG');
    }

    /**
     * Arrêter la surveillance de la configuration
     */
    stopConfigWatcher() {
        if (this.configCheckInterval) {
            clearInterval(this.configCheckInterval);
            this.configCheckInterval = null;
            this.log('🛑 Surveillance de la configuration arrêtée', 'CONFIG');
        }
    }

    /**
     * Vérifier les changements de configuration
     */
    async checkConfigurationChanges() {
        const now = Date.now();
        // Throttle à 10s
        if (this.lastConfigCheck && now - this.lastConfigCheck < 10000) return;
        this.lastConfigCheck = now;
        try {
            const configChanged = await this.loadMonitoredFolders();
            if (configChanged) {
                this.log('🔔 Configuration des dossiers modifiée détectée', 'CONFIG');
                this.emit('configuration-changed', {
                    foldersCount: this.monitoredFolders.length,
                    folders: this.monitoredFolders.map(f => f.path)
                });
            }
        } catch (error) {
            this.log(`❌ Erreur vérification configuration: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Redémarrer le monitoring avec la nouvelle configuration
     */
    async restartMonitoring() {
        try {
            this.log('🔄 Redémarrage du monitoring avec nouvelle configuration...', 'RESTART');
            
            if (this.isMonitoring) {
                await this.stopRealtimeMonitoring();
                await this.sleep(1000); // Attendre un peu
                await this.startRealtimeMonitoring();
            }
            
            this.log('✅ Monitoring redémarré avec succès', 'RESTART');
        } catch (error) {
            this.log(`❌ Erreur redémarrage monitoring: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Démarrage du monitoring
     */
    async startMonitoring() {
        try {
            if (this._isStarting || this.isMonitoring) {
                this.log('ℹ️ startMonitoring ignoré (déjà en cours/actif)', 'INFO');
                return;
            }
            this._isStarting = true;
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
            
            // ÉTAPE 2: Démarrer le monitoring temps réel simple via PowerShell
            this.log('🎧 PHASE 2: Activation monitoring temps réel PowerShell...', 'COM');
            try {
                await this.startRealtimeMonitoring();
                this.log('✅ PHASE 2: Monitoring temps réel PowerShell actif', 'SUCCESS');
            } catch (error) {
                this.log('⚠️ PHASE 2: Monitoring temps réel échoué, utilisation polling', 'WARNING');
                this.startFallbackPolling();
            }
            
            this.isMonitoring = true;
            this.log('✅ Monitoring opérationnel: PowerShell (sync) + COM moderne (temps réel)', 'SUCCESS');
            
            // Démarrer le polling de sécurité en standby
            if (this.isUsingCOMEvents) {
                this.setupFallbackPollingStandby();
            }
            
            // Émettre un événement pour signaler que le monitoring a démarré
            this.emit('monitoring-status', { 
                status: 'active',
                mode: this.isUsingCOMEvents ? 'powershell-sync + com-modern-realtime' : 'powershell-fallback',
                folders: this.monitoredFolders.length,
                stats: this.stats
            });
            
        } catch (error) {
            this.log(`❌ Erreur démarrage monitoring: ${error.message}`, 'ERROR');
            this.emit('error', error);
            throw error;
        } finally {
            this._isStarting = false;
        }
    }

    /**
     * Démarrage du monitoring en temps réel via événements COM modernes
     */
    async startComEventMonitoring() {
        try {
            this.log('🎧 Démarrage du monitoring via événements COM modernes...', 'COM');

            if (!this.outlookEventsService) {
                throw new Error('Service d\'écoute COM non initialisé');
            }

            // Récupérer les dossiers à surveiller (chemin + IDs si dispo)
            const folderTargets = this.monitoredFolders.map(folder => ({
                path: folder.path,
                storeId: folder.storeId || null,
                entryId: folder.entryId || null,
                storeName: folder.storeName || null
            }));
            
            // Démarrer l'écoute COM
            const result = await this.outlookEventsService.startListening(folderTargets);

            if (result.success) {
                this.isUsingCOMEvents = true;
                this.log(`✅ Écoute COM activée pour ${folderTargets.length} dossiers`, 'COM');
                this.emit('com-listening-started', { folders: folderTargets.length });
            } else {
                throw new Error(result.message);
            }

        } catch (error) {
            this.log(`❌ Erreur lors du démarrage des événements COM: ${error.message}`, 'ERROR');
            this.isUsingCOMEvents = false;
            this.emit('com-listening-failed', error.message);
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
            
            // Naviguer vers le dossier avec priorité aux identifiants stables
            const folder = this.resolveOutlookFolder(namespace, folderConfig);
            
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
                const comFolder = this.resolveOutlookFolder(namespace, folderConfig);
                
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
        if (!folderPath) return null;
        const normalize = (p) => String(p || '')
            .replace(/\//g, '\\')
            .replace(/\\+/g, '\\')
            .trim()
            .toLowerCase();

        const target = normalize(folderPath);
        if (!target) return null;

        // 1) Exact match (normalisé)
        let best = null;
        for (const f of (this.monitoredFolders || [])) {
            const fp = normalize(f?.path);
            if (fp && fp === target) return f;
        }

        // 2) Best suffix/prefix match (le plus long gagne)
        const scoreMatch = (fp, t) => {
            if (!fp || !t) return 0;
            if (t.startsWith(fp + '\\') || t === fp) return fp.length;
            if (fp.startsWith(t + '\\') || fp === t) return t.length;
            if (t.endsWith('\\' + fp) || t.endsWith(fp)) return fp.length;
            if (fp.endsWith('\\' + t) || fp.endsWith(t)) return t.length;
            if (t.includes(fp) || fp.includes(t)) return Math.min(fp.length, t.length);
            return 0;
        };

        let bestScore = 0;
        for (const f of (this.monitoredFolders || [])) {
            const fp = normalize(f?.path);
            const s = scoreMatch(fp, target);
            if (s > bestScore) {
                bestScore = s;
                best = f;
            }
        }

        return bestScore > 0 ? best : null;
    }

    /**
     * Résout un dossier Outlook en privilégiant les identifiants stables
     */
    resolveOutlookFolder(namespace, folderConfig) {
        try {
            if (!namespace || !folderConfig) return null;

            const storeId = folderConfig.storeId || folderConfig.store_id || null;
            const entryId = folderConfig.entryId || folderConfig.entry_id || null;

            if (entryId) {
                try {
                    return storeId
                        ? namespace.GetFolderFromID(entryId, storeId)
                        : namespace.GetFolderFromID(entryId);
                } catch (error) {
                    this.log(`⚠️ Résolution par ID impossible pour ${folderConfig.name || folderConfig.path}: ${error.message}`, 'WARNING');
                }
            }

            if (folderConfig.path) {
                return this.getOutlookFolderByPath(namespace, folderConfig.path);
            }

            return null;
        } catch (error) {
            this.log(`❌ Erreur résolution dossier Outlook: ${error.message}`, 'ERROR');
            return null;
        }
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
            const parent = comItem.Parent || null;
            const parentStore = parent?.Store || comItem.Store || null;
            return {
                id: comItem.EntryID,
                internetMessageId: comItem.InternetMessageId || comItem.InternetMessageID || comItem.internetMessageId || '',
                subject: comItem.Subject || '(Sans objet)',
                senderName: comItem.SenderName || 'Inconnu',
                senderEmail: comItem.SenderEmailAddress || '',
                receivedTime: comItem.ReceivedTime ? new Date(comItem.ReceivedTime) : new Date(),
                isRead: !comItem.UnRead,
                hasAttachments: !!(comItem.Attachments && comItem.Attachments.Count > 0),
                importance: comItem.Importance,
                folderName: parent ? parent.Name : 'Unknown',
                folderPath: parent?.FolderPath || '',
                folderStoreId: parentStore?.StoreID || '',
                folderStoreName: parentStore?.DisplayName || ''
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

            // Traiter le nouvel email, avec fallback direct si le connecteur n'enrichit pas les données
            const processedEmail = typeof this.outlookConnector.processEmailData === 'function'
                ? await this.outlookConnector.processEmailData(mailData)
                : mailData;
            
            // Sauvegarder en base
            const emailRecord = {
                outlook_id: processedEmail.EntryID || processedEmail.id,
                subject: processedEmail.subject,
                sender_email: processedEmail.senderEmail || processedEmail.SenderEmailAddress || processedEmail.senderName || processedEmail.SenderName || '',
                received_time: processedEmail.receivedTime || processedEmail.ReceivedTime || new Date().toISOString(),
                is_read: processedEmail.UnRead !== undefined ? !processedEmail.UnRead : (processedEmail.isRead || false),
                size: processedEmail.size || 0,
                folder_name: folderConfig.path,
                folder_type: folderConfig.category || folderConfig.type,
                // Legacy flag kept false on insert; treated_at is the source of truth
                is_treated: false,
                internet_message_id: processedEmail.internetMessageId || processedEmail.internet_message_id || processedEmail.InternetMessageId || '',
                last_modified_time: processedEmail.lastModifiedTime || processedEmail.last_modified_time || processedEmail.LastModificationTime || null,
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

            // Traiter l'email modifié, avec fallback direct si le connecteur n'enrichit pas les données
            const processedEmail = typeof this.outlookConnector.processEmailData === 'function'
                ? await this.outlookConnector.processEmailData(mailData)
                : mailData;
            
            // Mettre à jour en base
            const updateData = {
                is_read: processedEmail.isRead,
                subject: processedEmail.subject,
                folder_name: folderConfig.path,
                internet_message_id: processedEmail.internetMessageId || processedEmail.internet_message_id || processedEmail.InternetMessageId || '',
                last_modified_time: processedEmail.lastModifiedTime || processedEmail.last_modified_time || processedEmail.LastModificationTime || null
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
            if (this._isStopping || !this.isMonitoring) {
                // Ensure timers are cleaned even if already stopped
                if (this.configCheckInterval) { clearInterval(this.configCheckInterval); this.configCheckInterval = null; }
                if (this.pollingInterval) { clearInterval(this.pollingInterval); this.pollingInterval = null; }
                if (this.cacheCleanupInterval) { clearInterval(this.cacheCleanupInterval); this.cacheCleanupInterval = null; }
                this.log('ℹ️ stopMonitoring ignoré (déjà en arrêt ou non actif)', 'INFO');
                return;
            }
            this._isStopping = true;
            this.log('🛑 Arrêt du monitoring...', 'STOP');

            // Arrêter le polling
            this.stopPollingMode();

            // NOUVEAU: Arrêter l'écoute COM moderne
            await this.stopCOMEventListening();

            // NOUVEAU: Arrêter le monitoring temps réel
            await this.stopRealtimeMonitoring();

            // Arrêter le polling de secours
            if (this.pollingInterval) {
                clearInterval(this.pollingInterval);
                this.pollingInterval = null;
                this.fallbackPollingActive = false;
            }
            
            // Arrêter la surveillance de la configuration
            if (this.configCheckInterval) {
                clearInterval(this.configCheckInterval);
                this.configCheckInterval = null;
                this.log('✅ Surveillance de la configuration arrêtée', 'STOP');
            }

            // Nettoyer tous les handlers d'événements COM anciens (si présents)
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

            // Arrêter la surveillance de la configuration
            this.stopConfigWatcher();

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
        } finally {
            this._isStopping = false;
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

            const failedFolders = [];

            // Traiter chaque dossier configuré
            for (const folder of this.monitoredFolders) {
                try {
                    await this.syncFolder(folder);
                } catch (e) {
                    failedFolders.push({ name: folder?.name || folder?.path || 'unknown', error: e?.message || String(e) });
                    this.log(`⚠️ Dossier ignoré (erreur): ${folder?.name || folder?.path} — ${e?.message || e}`, 'WARNING');
                }
            }

            const duration = ((Date.now() - startTime) / 1000).toFixed(1);
            if (failedFolders.length > 0) {
                this.log(`⚠️ Synchronisation terminée avec ${failedFolders.length} dossier(s) en erreur (voir logs)`, 'WARNING');
            }
            this.log(`✅ Synchronisation terminée en ${duration}s (${this.stats.emailsAdded} ajoutés, ${this.stats.emailsUpdated} mis à jour)`, 'SUCCESS');
            this.stats.lastSyncTime = new Date().toISOString();
            
            this.emit('sync-complete', {
                duration,
                emailsAdded: this.stats.emailsAdded,
                emailsUpdated: this.stats.emailsUpdated,
                failedFolders
            });
        } catch (error) {
            this.log(`❌ Erreur synchronisation complète: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Force une resynchronisation complète de tous les dossiers
     * Réinitialise l'état de baseline et relance une synchronisation complète
     */
    async forceFullResync() {
        try {
            this.log('🔄 Forçage resynchronisation complète - Réinitialisation baseline...', 'SYNC');
            
            // Réinitialiser le baseline pour tous les dossiers surveillés
            for (const folder of this.monitoredFolders) {
                try {
                    this.dbService.db.prepare(`
                        UPDATE folder_sync_state 
                        SET baseline_done = 0, 
                            last_modified_cursor = NULL,
                            last_full_scan_at = NULL,
                            updated_at = CURRENT_TIMESTAMP
                        WHERE folder_path = ?
                    `).run(folder.path);
                    this.log(`🔄 Baseline réinitialisé pour: ${folder.name}`, 'SYNC');
                } catch (e) {
                    this.log(`⚠️ Impossible de réinitialiser ${folder.name}: ${e.message}`, 'WARNING');
                }
            }
            
            // Lancer la synchronisation complète
            this.log('🚀 Lancement synchronisation complète forcée...', 'SYNC');
            await this.performCompleteSync();
            
            this.log('✅ Resynchronisation complète forcée terminée avec succès', 'SUCCESS');
            return {
                success: true,
                message: 'Resynchronisation complète effectuée avec succès',
                stats: {
                    emailsAdded: this.stats.emailsAdded,
                    emailsUpdated: this.stats.emailsUpdated
                }
            };
        } catch (error) {
            this.log(`❌ Erreur lors de la resynchronisation forcée: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * Obtenir l'identité de synchronisation d'un dossier
     */
    getFolderSyncIdentity(folder) {
        return {
            folder_path: folder.path,
            store_id: folder.storeId || folder.store_id || null,
            entry_id: folder.entryId || folder.entry_id || null,
            store_name: folder.storeName || folder.store_name || null
        };
    }

    /**
     * Obtenir l'état de synchronisation d'un dossier en toute sécurité
     */
    safeGetFolderSyncState(folder) {
        try {
            return this.dbService.getFolderSyncState(folder.path);
        } catch {
            return null;
        }
    }

    safeUpsertFolderSyncState(folder, patch) {
        try {
            const base = this.getFolderSyncIdentity(folder);
            return this.dbService.upsertFolderSyncState({ ...base, ...(patch || {}) });
        } catch (e) {
            this.log(`⚠️ Impossible d'upsert folder_sync_state pour ${folder?.name || folder?.path}: ${e.message}`, 'WARNING');
            return null;
        }
    }

    async baselineSyncFolder(folder) {
        // Baseline = scan complet (tous les items) paginé via ModifiedBefore.
        const pageSize = Math.max(50, Number(this.config.syncBatchSize) || 50);
        const psPageSize = Math.max(100, Math.min(1000, pageSize * 10));

        // Cutoff de réconciliation: tout email du dossier non revu depuis ce moment sera considéré absent
        // IMPORTANT: la BDD stocke last_seen_at en SQLite CURRENT_TIMESTAMP (UTC) => "YYYY-MM-DD HH:MM:SS".
        // Comparer avec un ISO ("...T...Z") peut provoquer des comparaisons lexicographiques fausses et
        // soft-delete des emails fraîchement insérés pendant le baseline.
        const baselineStartedAtIso = new Date().toISOString();
        const baselineStartedAtSql = baselineStartedAtIso.slice(0, 19).replace('T', ' ');

        let before = null;
        let maxCursor = null;
        let totalFetched = 0;
        let page = 0;

        this.log(`🧱 Baseline scan: ${folder.name}`, 'SYNC');

        while (true) {
            page++;
            const res = await this.outlookConnector.getFolderEmails(folder.path, {
                limit: psPageSize,
                allItems: true,
                useLastModificationTime: true,
                modifiedBefore: before,
                storeId: folder.storeId || folder.store_id,
                storeName: folder.storeName || folder.store_name,
                folderEntryId: folder.entryId || folder.entry_id
            });

            if (!res || res.success === false) {
                throw new Error(res?.error || 'Baseline scan PowerShell échoué');
            }

            const emails = Array.isArray(res.emails) ? res.emails : [];
            if (emails.length === 0) {
                break;
            }

            // Cursor = dernière date de modification max observée (ISO UTC)
            const batchMax = res.maxLastModificationTime || emails[0]?.LastModificationTime || '';
            if (batchMax) {
                const batchMaxMs = Date.parse(batchMax);
                const curMs = maxCursor ? Date.parse(maxCursor) : NaN;
                if (!maxCursor || (Number.isFinite(batchMaxMs) && (!Number.isFinite(curMs) || batchMaxMs > curMs))) {
                    maxCursor = batchMax;
                }
            }

            // Traiter en batch “app-level”
            for (let i = 0; i < emails.length; i += this.config.syncBatchSize) {
                const chunk = emails.slice(i, i + this.config.syncBatchSize);
                await this.processBatch(chunk, folder);
            }

            totalFetched += emails.length;

            // Préparer la page suivante
            const minLmt = res.minLastModificationTime || emails[emails.length - 1]?.LastModificationTime || '';
            if (!res.hasMore || !minLmt) {
                break;
            }

            const minMs = Date.parse(minLmt);
            if (!Number.isFinite(minMs)) {
                break;
            }
            // -1s pour éviter de boucler sur la même frontière
            before = new Date(Math.max(0, minMs - 1000)).toISOString();

            // Sécurité: ne pas boucler indéfiniment
            if (page > 10000) {
                throw new Error('Baseline scan interrompu (trop de pages)');
            }
        }

        this.safeUpsertFolderSyncState(folder, {
            last_modified_cursor: maxCursor,
            last_full_scan_at: new Date().toISOString(),
            baseline_done: 1
        });

        // Réconciliation: marquer supprimés/déplacés les emails qui étaient en BDD pour ce dossier
        // mais qui n'ont pas été revus pendant ce baseline complet.
        try {
            const changes = this.dbService.softDeleteMissingEmailsByFolderSince(folder.path, baselineStartedAtSql);
            if (changes > 0) {
                this.log(`🧹 Réconciliation: ${changes} email(s) marqués absents dans ${folder.name}`, 'SYNC');
                try { this.emailCache?.clear?.(); } catch (_) {}
                try { this.cacheService?.invalidateStats?.(); } catch (_) {}
            }
        } catch (e) {
            this.log(`⚠️ Réconciliation baseline ignorée (${folder.name}): ${e.message}`, 'WARNING');
        }

        this.log(`✅ Baseline terminé: ${folder.name} (${totalFetched} items traités)`, 'SUCCESS');
    }

    async incrementalSyncFolder(folder) {
        const state = this.safeGetFolderSyncState(folder);
        const cursorStr = state?.last_modified_cursor || null;
        const cursor = cursorStr ? new Date(cursorStr) : null;

        const res = await this.outlookConnector.getFolderEmails(folder.path, {
            limit: 1000,
            allItems: true,
            useLastModificationTime: true,
            modifiedSince: cursor && !Number.isNaN(cursor.getTime()) ? cursor : null,
            storeId: folder.storeId || folder.store_id,
            storeName: folder.storeName || folder.store_name,
            folderEntryId: folder.entryId || folder.entry_id
        });

        if (!res || res.success === false) {
            throw new Error(res?.error || 'Sync incrémentale PowerShell échouée');
        }

        const emails = Array.isArray(res.emails) ? res.emails : [];
        if (emails.length > 0) {
            for (let i = 0; i < emails.length; i += this.config.syncBatchSize) {
                const chunk = emails.slice(i, i + this.config.syncBatchSize);
                await this.processBatch(chunk, folder);
            }
        }

        // Détection d'écart offline: si Outlook indique un total plus petit que la BDD,
        // on a probablement raté des suppressions/déplacements pendant que l'app était arrêtée.
        try {
            const totalInFolder = Number.isFinite(res.totalInFolder) ? Number(res.totalInFolder) : null;
            if (Number.isFinite(totalInFolder)) {
                const dbCount = this.dbService.getActiveEmailCountByFolder(folder.path);
                // Tolérance: 1 (évite de rescan pour des micro-diff temporaires)
                if (Number.isFinite(dbCount) && dbCount > (totalInFolder + 1)) {
                    this.reconcileFolderFullScan(folder, 'écart BDD vs Outlook');
                }
            }
        } catch (_) {}

        const nextCursor = res.maxLastModificationTime || emails[0]?.LastModificationTime || null;
        if (nextCursor) {
            this.safeUpsertFolderSyncState(folder, { last_modified_cursor: nextCursor });
        }

        return { count: emails.length, cursor: nextCursor };
    }

    /**
     * Synchronisation d'un dossier spécifique
     */
    async syncFolder(folder) {
        try {
            // Afficher le chemin COMPLET pour plus de clarté (aide diagnostic)
            const displayPath = folder.path || folder.name || '(inconnu)';
            const displayMsg = displayPath.includes('\\') ? displayPath : `${folder.name} [path: ${folder.path}]`;
            this.log(`📁 Synchronisation du dossier: ${displayMsg}`, 'SYNC');
            
            if (this.config.enableDetailedLogging) {
                console.log(`🔍 DEBUG folder.path: "${folder.path}"`);
                console.log(`🔍 DEBUG folder.name: "${folder.name}"`);
                console.log(`🔍 DEBUG folder.storeId: "${folder.storeId}"`);
                console.log(`🔍 DEBUG folder.entryId: "${folder.entryId}"`);
            }

            const state = this.safeGetFolderSyncState(folder);
            const baselineDone = state && Number(state.baseline_done) === 1;

            if (!baselineDone) {
                await this.baselineSyncFolder(folder);
            } else {
                const inc = await this.incrementalSyncFolder(folder);
                if (inc && inc.count > 0) {
                    this.log(`📧 ${inc.count} changement(s) appliqué(s) dans ${folder.name}`, 'INFO');
                }
            }

            this.log(`✅ Dossier ${folder.name} synchronisé`, 'SUCCESS');

        } catch (error) {
            const dbgPath = folder.path || '(path inconnu)';
            const dbgStore = folder.storeId || folder.store_id || '(store inconnu)';
            const dbgEntry = folder.entryId || folder.entry_id || '(entryId inconnu)';
            this.log(`❌ Erreur synchronisation dossier ${folder.name}: ${error.message} — path="${dbgPath}", storeId="${dbgStore}", entryId="${dbgEntry}"`, 'ERROR');
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
                    expressMode: true, // Mode ultra-rapide
                    storeId: folder.storeId || folder.store_id,
                    storeName: folder.storeName || folder.store_name,
                    folderEntryId: folder.entryId || folder.entry_id
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
            const concurrency = Math.max(1, Number(this.config.maxConcurrentBatches) || 1);
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
        // Normaliser le sujet pour chaque email (les scripts COM/PS fournissent souvent "Subject")
        const correctedEmails = emails.map(emailData => {
            // Prélever depuis plusieurs sources possibles
            const rawSubject = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
            const normalized = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
            emailData.subject = normalized;
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
            
            // Vérifier si l'email existe déjà en base (recherche par outlook_id uniquement pour gérer les déplacements entre dossiers)
            const existingEmail = await this.dbService.getEmailByEntryId(emailId);

            if (existingEmail) {
                // Détecter un déplacement de dossier
                const movedFolder = existingEmail.folder_name !== folder.path;
                const needsUpdate = this.needsUpdateOptimized(existingEmail, emailData);

                if (movedFolder || needsUpdate) {
                    // Normaliser les champs et utiliser saveEmail pour mettre à jour dossier/statut/sujet
                    const rawSubject = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
                    const subject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
                    const saveRecord = {
                        outlook_id: emailId,
                        subject,
                        sender_email: emailData.senderEmail || emailData.SenderEmailAddress || emailData.senderName || emailData.SenderName || '',
                        received_time: emailData.receivedTime || emailData.ReceivedTime || existingEmail.received_time,
                        folder_name: folder.path,
                        category: folder.category || folder.type,
                        is_read: emailData.UnRead !== undefined ? !emailData.UnRead : (emailData.isRead ?? existingEmail.is_read ?? false),
                        is_treated: existingEmail.is_treated || 0,
                        internet_message_id: emailData.InternetMessageId || emailData.internet_message_id || emailData.internetMessageId || '',
                        last_modified_time: emailData.LastModificationTime || emailData.last_modified_time || emailData.lastModifiedTime || ''
                    };
                    await this.dbService.saveEmail(saveRecord);
                    this.stats.emailsUpdated++;
                } else {
                    // IMPORTANT: même sans changement, marquer l'email comme "vu" pendant les scans
                    // pour permettre la réconciliation (emails absents) au démarrage.
                    // Et si "lu = traité" est actif, appliquer la règle même sans transition détectée.
                    try {
                        const isReadNow = emailData.UnRead !== undefined ? !emailData.UnRead : (emailData.isRead ?? existingEmail.is_read ?? false);
                        if (isReadNow && !existingEmail?.treated_at) {
                            const when = emailData.LastModificationTime || emailData.last_modified_time || emailData.lastModifiedTime || null;
                            this.dbService.markReadAsTreatedIfNeeded?.(emailId, when);
                        }
                    } catch (_) {}
                    try { this.dbService.touchEmailSeen(emailId); } catch (_) {}
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
            // Normaliser le sujet (compat: Subject/subject/ConversationTopic)
            const rawSubject = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
            emailData.subject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
            
            // Formatter les données pour la base de données
            const emailRecord = {
                outlook_id: emailData.EntryID || emailData.id,
                subject: emailData.subject,
                sender_email: emailData.senderEmail || emailData.SenderEmailAddress || emailData.senderName || emailData.SenderName || '',
                received_time: emailData.receivedTime || emailData.ReceivedTime,
                is_read: emailData.UnRead !== undefined ? !emailData.UnRead : (emailData.isRead || false),
                folder_name: folder.path,
                category: folder.category || folder.type,
                is_treated: false,
                internet_message_id: emailData.InternetMessageId || emailData.internet_message_id || emailData.internetMessageId || '',
                last_modified_time: emailData.LastModificationTime || emailData.last_modified_time || emailData.lastModifiedTime || ''
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
            // Normaliser le sujet (compat: Subject/subject/ConversationTopic)
            const rawSubject = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
            emailData.subject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
            
            // Formatter les données pour la mise à jour
            const updateData = {
                is_read: emailData.UnRead !== undefined ? !emailData.UnRead : (emailData.isRead || false),
                subject: emailData.subject,
                folder_name: folder.path
            };
            
            const entryId = emailData.EntryID || emailData.id;
            await this.dbService.updateEmailStatus(entryId, updateData.is_read, updateData.folder_name);
            
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
     * NOUVEAU: Obtenir les statistiques business réelles depuis la BDD
     */
    async getBusinessStats() {
        try {
            // Utiliser le cache si récent
            const cacheKey = 'business_stats';
            const cached = this.emailCache.get(cacheKey);
            
            if (cached && (Date.now() - cached.timestamp) < 15000) { // 15 secondes de cache
                return cached.data;
            }
            
            // Calculer les vraies stats depuis la BDD
            const dbStats = await this.dbService.getEmailStats();
            
            const businessStats = {
                emailsToday: dbStats.emailsToday || 0,
                treatedToday: dbStats.treatedToday || 0,
                unreadTotal: dbStats.unreadTotal || 0,
                totalEmails: dbStats.totalEmails || 0,
                lastSyncTime: this.stats.lastSyncTime || new Date(),
                monitoringActive: this.isMonitoring
            };
            
            // Mettre en cache
            this.emailCache.set(cacheKey, {
                timestamp: Date.now(),
                data: businessStats
            });
            
            return businessStats;
            
        } catch (error) {
            this.log(`❌ Erreur récupération stats business: ${error.message}`, 'ERROR');
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
        for (const [key, entry] of this.folderStatsCache.entries()) {
            if (now - entry.timestamp > this.cacheExpiry) {
                this.folderStatsCache.delete(key);
            }
        }
    }

    /**
     * Démarrage du nettoyage automatique du cache
     */
    startCacheCleanup() {
        if (this.cacheCleanupInterval) {
            clearInterval(this.cacheCleanupInterval);
        }
        this.cacheCleanupInterval = setInterval(() => {
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
        }, 5000); // Vérification toutes les 5 secondes
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

    // startCacheCleanup et stopMonitoring en double supprimés

    /**
     * Synchronisation forcée
     */
    async forceSync() {
        this.log('💪 Synchronisation forcée démarrée...', 'SYNC');
        // Implémentation de la synchronisation forcée
        this.log('✅ Synchronisation forcée terminée', 'SUCCESS');
    }

    /**
     * Invalider le cache d'emails - méthode pour synchronisation en temps réel
     */
    invalidateEmailCache() {
        try {
            console.log('📊 [CACHE-UNIFIED] Invalidation START');
            
            // Vider complètement le cache d'emails
            const emailCacheSize = this.emailCache.size;
            this.emailCache.clear();
            console.log(`  ✓ emailCache vidé (${emailCacheSize} entrées supprimées)`);
            
            // Invalider aussi le cache des stats de dossiers
            const folderCacheSize = this.folderStatsCache.size;
            this.folderStatsCache.clear();
            console.log(`  ✓ folderStatsCache vidé (${folderCacheSize} entrées supprimées)`);
            
            // Invalider le cache du service de base de données
            if (this.dbService && this.dbService.cache) {
                // Invalider spécifiquement les clés d'emails récents
                const keys = this.dbService.cache.keys();
                const emailKeys = keys.filter(key => 
                    key.startsWith('recent_emails_') || 
                    key.startsWith('stats_') ||
                    key.startsWith('folder_')
                );
                
                emailKeys.forEach(key => this.dbService.cache.del(key));
                
                console.log(`  ✓ DB cache invalidé: ${emailKeys.length} clés emails/stats supprimées`);
            }

            // Invalider aussi le cache partagé (cacheService) utilisé par les handlers IPC
            try {
                if (this.cacheService) {
                    if (typeof this.cacheService.invalidateStats === 'function') {
                        this.cacheService.invalidateStats();
                        console.log(`  ✓ cacheService.invalidateStats() appelée`);
                    } else {
                        this.cacheService?.del?.('ui', 'dashboard_stats');
                        console.log(`  ✓ cacheService: dashboard_stats supprimée`);
                    }

                    if (typeof this.cacheService.invalidateEmailsCache === 'function') {
                        this.cacheService.invalidateEmailsCache();
                        console.log(`  ✓ cacheService.invalidateEmailsCache() appelée`);
                    } else {
                        this.cacheService?.del?.('emails', 'recent_20');
                        this.cacheService?.del?.('emails', 'recent_50');
                        console.log(`  ✓ cacheService: recent_20 et recent_50 supprimées`);
                    }

                    if (typeof this.cacheService.invalidateFoldersTree === 'function') {
                        this.cacheService.invalidateFoldersTree();
                        console.log(`  ✓ cacheService.invalidateFoldersTree() appelée`);
                    } else {
                        this.cacheService?.del?.('config', 'folders_tree');
                        console.log(`  ✓ cacheService: folders_tree supprimée`);
                    }
                }
            } catch (e) {
                console.error('❌ [CACHE-UNIFIED] Erreur lors invalidation cacheService:', e.message);
            }

            console.log('📊 [CACHE-UNIFIED] Invalidation DONE');
            
        } catch (error) {
            console.error('❌ [CACHE] Erreur invalidation cache emails:', error);
        }
    }

    // getRecentEmails (compat) supprimé: utiliser la version avec cache plus haut

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

    /**
     * NOUVEAU: Configure les listeners d'événements COM modernes
     */
    setupModernCOMEventListeners() {
        if (!this.outlookEventsService) return;

        this.log('🔧 Configuration des listeners COM modernes...', 'COM');

        // Écouter les nouveaux emails
        this.outlookEventsService.on('newEmail', (emailData) => {
            this.log(`📬 Nouvel email COM: ${emailData.subject}`, 'COM');
            this.handleCOMNewEmail(emailData);
        });

        // Écouter les changements d'état des emails (COM Events)
        this.outlookEventsService.on('emailChanged', (emailData) => {
            this.log(`🔄 Email modifié COM: ${emailData.subject}`, 'COM');
            this.handleCOMEmailChanged(emailData);
        });

        // Écouter les changements d'état des emails (Polling Intelligent)
        this.outlookEventsService.on('email-changed', (eventData) => {
            this.log(`🔄 Email modifié POLLING: ${eventData.Subject || 'N/A'} - ${eventData.ChangeType}`, 'POLLING');
            this.handlePollingEmailChanged(eventData);
        });

        // Écouter les événements groupés
        this.outlookEventsService.on('eventsProcessed', (groupedEvents) => {
            this.log(`📊 Traitement groupé: ${groupedEvents.totalEvents} événements`, 'COM');
            this.handleCOMGroupedEvents(groupedEvents);
        });

        // Écouter les problèmes d'écoute
        this.outlookEventsService.on('listening-failed', (error) => {
            this.log('❌ Écoute COM échouée, basculement vers polling', 'ERROR');
            this.isUsingCOMEvents = false;
            this.startFallbackPolling();
        });

        this.log('✅ Listeners événements COM modernes configurés', 'COM');
    }

    /**
     * NOUVEAU: Gère les nouveaux emails détectés via COM
     */
    async handleCOMNewEmail(emailData) {
        try {
            // Traiter immédiatement le nouvel email en base
            const result = await this.dbService.processCOMNewEmail(emailData);
            
            if (result.processed) {
                this.log(`📧 Nouvel email traité via COM: ${emailData.subject}`, 'COM');
                
                // Émettre l'événement pour mise à jour UI temps réel
                this.emit('realtime-new-email', {
                    ...emailData,
                    category: this.getFolderCategory(emailData.folderPath),
                    timestamp: new Date()
                });

                // Incrémenter les stats
                this.stats.emailsAdded++;
                this.stats.eventsReceived++;
            }

        } catch (error) {
            this.log(`❌ Erreur traitement nouvel email COM: ${error.message}`, 'ERROR');
        }
    }

    /**
     * NOUVEAU: Gère les changements d'état des emails via COM
     */
    async handleCOMEmailChanged(emailData) {
        try {
            // Mettre à jour l'état de l'email en base
            const result = await this.dbService.processCOMEmailChange(emailData);
            
            if (result.updated) {
                this.log(`🔄 État email mis à jour via COM: ${emailData.subject}`, 'COM');
                
                // Émettre l'événement pour mise à jour UI temps réel
                this.emit('realtime-email-update', {
                    ...emailData,
                    category: this.getFolderCategory(emailData.folderPath),
                    timestamp: new Date()
                });

                // Incrémenter les stats
                this.stats.emailsUpdated++;
                this.stats.eventsReceived++;
            }

        } catch (error) {
            this.log(`❌ Erreur traitement changement email COM: ${error.message}`, 'ERROR');
        }
    }

    /**
     * NOUVEAU: Gère les changements d'état des emails via Polling Intelligent
     */
    async handlePollingEmailChanged(eventData) {
        try {
            this.log(`🔄 [POLLING] Traitement changement détecté: ${eventData.Subject} - ${eventData.ChangeType}`, 'POLLING');
            
            // Convertir les données du polling en format compatible base de données
            const emailUpdateData = {
                messageId: eventData.EntryID,
                folderPath: eventData.FolderPath,
                subject: eventData.Subject,
                isRead: !eventData.UnRead, // UnRead est inversé
                lastModificationTime: eventData.LastModificationTime,
                changeType: eventData.ChangeType,
                changes: eventData.Changes || [],
                internetMessageId: eventData.InternetMessageId || eventData.internetMessageId || eventData.internet_message_id || '',
                storeId: eventData.StoreId || eventData.StoreID || '',
                storeName: eventData.StoreName || ''
            };

            // Mettre à jour l'état de l'email en base
            const result = await this.dbService.processPollingEmailChange(emailUpdateData);
            
            if (result && result.updated) {
                this.log(`✅ [POLLING] État email mis à jour en BDD: ${eventData.Subject}`, 'SUCCESS');
                
                // Émettre l'événement pour mise à jour UI temps réel
                this.emit('realtime-email-update', {
                    ...emailUpdateData,
                    category: this.getFolderCategory(eventData.FolderPath),
                    timestamp: new Date()
                });

                // Incrémenter les stats
                this.stats.emailsUpdated++;
                this.stats.eventsReceived++;
            } else {
                this.log(`⚠️ [POLLING] Changement non traité: ${eventData.Subject}`, 'WARNING');
            }

        } catch (error) {
            this.log(`❌ [POLLING] Erreur traitement changement email: ${error.message}`, 'ERROR');
            console.error('Stack trace:', error.stack);
        }
    }

    /**
     * NOUVEAU: Gère les événements groupés COM
     */
    async handleCOMGroupedEvents(groupedEvents) {
        try {
            // Traiter les événements par dossier
            for (const [folderPath, events] of Object.entries(groupedEvents.groupedEvents)) {
                this.log(`📊 Traitement ${events.newEmails} nouveaux + ${events.changedEmails} modifiés dans ${folderPath}`, 'COM');
            }

            // Émettre un événement de synchronisation terminée
            this.emit('syncCompleted', {
                totalEvents: groupedEvents.totalEvents,
                folders: Object.keys(groupedEvents.groupedEvents).length,
                timestamp: groupedEvents.timestamp,
                source: 'COM'
            });

        } catch (error) {
            this.log(`❌ Erreur traitement événements groupés COM: ${error.message}`, 'ERROR');
        }
    }

    /**
     * NOUVEAU: Démarre le polling de secours en cas d'échec COM
     */
    startFallbackPolling() {
        if (this.fallbackPollingActive) return;

        this.log('🔄 Démarrage du polling de secours...', 'FALLBACK');
        this.fallbackPollingActive = true;

        // Polling moins fréquent (toutes les 2 minutes) car c'est un fallback
        if (this.pollingInterval) {
            clearInterval(this.pollingInterval);
            this.pollingInterval = null;
        }
        this.pollingInterval = setInterval(() => {
            this.performLightMonitoringCheck();
        }, 120000); // 2 minutes

        this.log('✅ Polling de secours actif (2 min)', 'FALLBACK');
    }

    /**
     * NOUVEAU: Configure le polling de secours en standby
     */
    setupFallbackPollingStandby() {
        // Polling très léger toutes les 5 minutes pour vérifier que l'écoute COM fonctionne
        if (this.pollingInterval) {
            clearInterval(this.pollingInterval);
            this.pollingInterval = null;
        }
        this.pollingInterval = setInterval(() => {
            this.checkCOMHealthAndFallback();
        }, 300000); // 5 minutes

        this.log('✅ Polling de secours en standby (5 min)', 'STANDBY');
    }

    /**
     * NOUVEAU: Vérifie la santé de l'écoute COM et bascule si nécessaire
     */
    async checkCOMHealthAndFallback() {
        try {
            if (!this.isUsingCOMEvents) return;

            const stats = this.outlookEventsService.getListeningStats();
            
            if (!stats.isListening) {
                this.log('⚠️ Écoute COM interrompue, basculement vers polling', 'WARNING');
                this.isUsingCOMEvents = false;
                this.startFallbackPolling();
            } else {
                this.log('✅ Écoute COM opérationnelle', 'COM');
            }

        } catch (error) {
            this.log(`❌ Erreur vérification santé COM: ${error.message}`, 'ERROR');
        }
    }

    /**
     * NOUVEAU: Démarre le monitoring en temps réel avec PowerShell
     */
    async startRealtimeMonitoring() {
        try {
            this.log('🚀 Démarrage du monitoring en temps réel complet...', 'REALTIME');
            
            if (!this.outlookConnector) {
                throw new Error('Outlook connector non disponible');
            }

            // Vérifier que nous avons des dossiers à surveiller
            if (!this.monitoredFolders || this.monitoredFolders.length === 0) {
                this.log('⚠️ Aucun dossier configuré pour le monitoring', 'WARNING');
                return;
            }

            // ===== CONFIGURER TOUS LES GESTIONNAIRES D'ÉVÉNEMENTS =====
            // Éviter les doublons
            this.outlookConnector.removeAllListeners('newEmailDetected');
            this.outlookConnector.removeAllListeners('emailStatusChanged');
            this.outlookConnector.removeAllListeners('emailSubjectChanged');
            this.outlookConnector.removeAllListeners('emailModified');
            this.outlookConnector.removeAllListeners('emailDeleted');
            this.outlookConnector.removeAllListeners('folderCountChanged');
            
            // 1. Nouveaux emails détectés
            this.outlookConnector.on('newEmailDetected', (emailData) => {
                this.handleRealtimeNewEmail(emailData);
            });

            // 2. Changements de statut lu/non lu
            this.outlookConnector.on('emailStatusChanged', async (data) => {
                try {
                    this.log(`📝 Statut changé: ${data.subject} -> ${data.isRead ? 'Lu' : 'Non lu'}`, 'STATUS');
                    
                    // Mettre à jour la base de données
                    await this.dbService.updateEmailStatus(data.entryId, data.isRead);

                    // IMPORTANT: un passage en lu/non-lu ne change pas le count total du dossier,
                    // donc on ne reçoit pas forcément folderCountChanged. On force une sync partielle
                    // (debounced) pour garantir la prise en compte (y compris si EntryID change).
                    const folderPath = data.folderPath;
                    const folderConfig = folderPath ? this.getFolderConfigByPath(folderPath) : null;
                    if (folderConfig) {
                        this.schedulePartialSync(folderConfig);
                    }
                    
                    // Émettre événement pour mise à jour de l'interface
                    this.emit('emailStatusUpdated', {
                        entryId: data.entryId,
                        isRead: data.isRead,
                        subject: data.subject,
                        folderPath: data.folderPath
                    });

                    // Invalider les caches internes (sinon l'UI peut rester figée ~15s)
                    try { this.emailCache?.clear?.(); } catch (_) {}
                    
                    // Invalider le cache des statistiques
                    if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        this.cacheService.invalidateStats();
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }
                    
                } catch (error) {
                    this.log(`❌ Erreur mise à jour statut: ${error.message}`, 'ERROR');
                }
            });

            // 3. Modifications du sujet d'email
            this.outlookConnector.on('emailSubjectChanged', async (data) => {
                try {
                    this.log(`📝 Sujet modifié: "${data.oldSubject}" -> "${data.newSubject}"`, 'MODIFY');
                    
                    // Mettre à jour la base de données
                    await this.dbService.updateEmailField(data.entryId, 'subject', data.newSubject);
                    
                    // Émettre événement
                    this.emit('emailSubjectUpdated', data);
                    
                } catch (error) {
                    this.log(`❌ Erreur mise à jour sujet: ${error.message}`, 'ERROR');
                }
            });

            // 4. Modifications générales d'emails
            this.outlookConnector.on('emailModified', async (data) => {
                try {
                    this.log(`🔄 Email modifié: ${data.subject}`, 'MODIFY');
                    
                    // Marquer comme modifié dans la base
                    await this.dbService.updateEmailField(data.entryId, 'last_modified', new Date());
                    
                    // Émettre événement
                    this.emit('emailModified', data);
                    
                } catch (error) {
                    this.log(`❌ Erreur traitement modification: ${error.message}`, 'ERROR');
                }
            });

            // 5. Emails supprimés
            this.outlookConnector.on('emailDeleted', async (data) => {
                try {
                    this.log(`🗑️ Email supprimé: ${data.subject}`, 'DELETE');
                    
                    // Marquer comme supprimé ou supprimer de la base
                    await this.dbService.markEmailAsDeleted(data.entryId);
                    
                    // Émettre événement
                    this.emit('emailDeleted', data);

                    // Invalider les caches internes
                    try { this.emailCache?.clear?.(); } catch (_) {}
                    
                    // Invalider le cache des statistiques
                    if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        this.cacheService.invalidateStats();
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }
                    
                } catch (error) {
                    this.log(`❌ Erreur traitement suppression: ${error.message}`, 'ERROR');
                }
            });

            // 6. Changements de nombre d'emails dans un dossier
            this.outlookConnector.on('folderCountChanged', async (data) => {
                try {
                    this.log(`📊 Nombre d'emails changé dans ${data.folderPath}: ${data.oldCount} -> ${data.newCount}`, 'COUNT');
                    
                    // Émettre événement pour mise à jour de l'interface
                    this.emit('folderCountUpdated', data);

                    // Invalider les caches internes (stats business, emails récents)
                    try { this.emailCache?.clear?.(); } catch (_) {}
                    
                    // Invalider le cache des statistiques
                    if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        if (this.cacheService && typeof this.cacheService.invalidateStats === 'function') {
                        this.cacheService.invalidateStats();
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }
                    } else if (this.dbService && this.dbService.cache) {
                        // Fallback: invalider le cache de la base de données
                        this.dbService.cache.flushAll();
                    }

                    // Déclencher une synchronisation partielle ciblée pour capter les emails déplacés
                    const folderConfig = this.getFolderConfigByPath(data.folderPath);
                    if (folderConfig) {
                        this.schedulePartialSync(folderConfig);

                        // Si le nombre baisse, c'est typiquement une suppression ou un déplacement hors du dossier.
                        // Une sync partielle ne peut pas "deviner" ce qui manque => lancer une réconciliation complète (throttled).
                        if (Number.isFinite(data.oldCount) && Number.isFinite(data.newCount) && data.newCount < data.oldCount) {
                            this.reconcileFolderFullScan(folderConfig, 'baisse compteur');
                        }
                    }
                    
                } catch (error) {
                    this.log(`❌ Erreur traitement changement de nombre: ${error.message}`, 'ERROR');
                }
            });

            // ===== DÉMARRER LE MONITORING AVEC THROTTLING =====
            const batchSize = this.config.monitorStartBatchSize || 5;
            const delayMs = this.config.monitorStartDelayMs || 300;
            const folders = this.monitoredFolders.slice();
            for (let i = 0; i < folders.length; i += batchSize) {
                const batch = folders.slice(i, i + batchSize);
                // Démarrer les items d'un batch en série avec petit délai pour éviter les pics CPU/IO
                for (const folderConfig of batch) {
                    try {
                        await this.outlookConnector.startFolderMonitoring(folderConfig.path, {
                            storeId: folderConfig.storeId || folderConfig.store_id,
                            storeName: folderConfig.storeName || folderConfig.store_name,
                            folderEntryId: folderConfig.entryId || folderConfig.entry_id
                        });
                        this.log(`✅ Monitoring activé: ${folderConfig.name}`, 'REALTIME');
                        if (delayMs > 0) { await this.sleep(delayMs); }
                    } catch (error) {
                        this.log(`❌ Erreur monitoring dossier ${folderConfig.name}: ${error.message}`, 'ERROR');
                    }
                }
                // Petite pause entre paquets lorsque grand nombre de dossiers
                if (folders.length > batchSize && delayMs > 0) {
                    await this.sleep(Math.min(1000, delayMs * 2));
                }
            }

            this.log('🎯 Monitoring temps réel complet démarré avec succès', 'REALTIME');
            this.log('📋 Événements surveillés: Nouveaux emails, Statuts lu/non lu, Modifications, Suppressions', 'INFO');
            
        } catch (error) {
            this.log(`❌ Erreur démarrage monitoring temps réel: ${error.message}`, 'ERROR');
            throw error;
        }
    }

    /**
     * NOUVEAU: Gère les nouveaux emails détectés en temps réel
     */
    async handleRealtimeNewEmail(emailData) {
        try {
            // Normaliser le sujet pour la persistance
            const rawSubjectRT = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
            emailData.subject = rawSubjectRT.trim() !== '' ? rawSubjectRT : '(Sans objet)';
            this.log(`📧 Nouvel email détecté: ${emailData.subject}`, 'REALTIME');
            
            // Trouver la configuration du dossier
            const folderConfig = this.monitoredFolders.find(f => f.path === emailData.folderPath);
            if (!folderConfig) {
                this.log(`⚠️ Dossier non configuré: ${emailData.folderPath}`, 'WARNING');
                return;
            }

            // Mapper et enrichir les données pour le schéma BDD optimisé
            const enrichedEmailData = {
                outlook_id: emailData.entryId || emailData.EntryID || emailData.id || '',
                subject: emailData.subject,
                sender_email: emailData.senderEmail || emailData.SenderEmailAddress || emailData.senderName || emailData.SenderName || '',
                received_time: emailData.receivedTime || emailData.ReceivedTime || new Date().toISOString(),
                folder_name: folderConfig.path,
                category: folderConfig.category || 'Mails simples',
                is_read: emailData.UnRead !== undefined ? !emailData.UnRead : (emailData.isRead || false),
                is_treated: false
            };

            // Sauvegarder en base de données via le service optimisé
            await this.dbService.saveEmail(enrichedEmailData);
            
            this.log(`💾 Email sauvegardé en temps réel: ${emailData.subject}`, 'REALTIME');

            // Émettre l'événement de mise à jour
            this.emit('realtimeEmailAdded', {
                email: enrichedEmailData,
                folder: folderConfig.name,
                category: folderConfig.category
            });

        } catch (error) {
            this.log(`❌ Erreur traitement email temps réel: ${error.message}`, 'ERROR');
        }
    }

    /**
     * NOUVEAU: Arrête le monitoring en temps réel
     */
    async stopRealtimeMonitoring() {
        try {
            this.log('🛑 Arrêt du monitoring en temps réel...', 'REALTIME');
            
            if (this.outlookConnector) {
                // Arrêter le monitoring pour chaque dossier
                for (const folderConfig of this.monitoredFolders) {
                    try {
                        await this.outlookConnector.stopFolderMonitoring(folderConfig.path);
                        this.log(`⏹️ Monitoring arrêté pour: ${folderConfig.name}`, 'REALTIME');
                    } catch (error) {
                        this.log(`⚠️ Erreur arrêt monitoring ${folderConfig.name}: ${error.message}`, 'WARNING');
                    }
                }

                // Supprimer les écouteurs d'événements
                this.outlookConnector.removeAllListeners('newEmailDetected');
            }

            this.log('✅ Monitoring en temps réel arrêté', 'REALTIME');
            
        } catch (error) {
            this.log(`❌ Erreur arrêt monitoring temps réel: ${error.message}`, 'ERROR');
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
                this.log('🛑 Écoute COM arrêtée', 'COM');
            }
        } catch (error) {
            this.log(`❌ Erreur arrêt écoute COM: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Récupère la catégorie d'un dossier
     */
    getFolderCategory(folderPath) {
        const folderConfig = this.monitoredFolders.find(f => f.path === folderPath);
        return folderConfig ? folderConfig.category : 'autres';
    }

    /**
     * NOUVEAU: Vérification légère de monitoring (pour fallback)
     */
    async performLightMonitoringCheck() {
        try {
            this.log('🔍 Vérification légère de monitoring (fallback)...', 'FALLBACK');
            
            // Sync rapide seulement si nécessaire
            const stats = await this.performQuickSync();
            
            if (stats.totalProcessed > 0) {
                this.log(`📊 Fallback: ${stats.totalProcessed} emails traités`, 'FALLBACK');
                this.emit('syncCompleted', {
                    ...stats,
                    source: 'FALLBACK'
                });
            }
            
        } catch (error) {
            this.log(`❌ Erreur vérification fallback: ${error.message}`, 'ERROR');
        }
    }

    /**
     * Utilitaire pour attendre
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

module.exports = UnifiedMonitoringService;
