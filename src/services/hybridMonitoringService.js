/**
 * Service de monitoring hybride - Version complète avec système VBA
 * Combine le service unifié avec le monitoring avancé style VBA
 */

const EventEmitter = require('events');
const fs = require('fs');
const path = require('path');
const databaseService = require('./databaseService');

class HybridMonitoringService extends EventEmitter {
    constructor(outlookConnector) {
        super();
        
        this.outlookConnector = outlookConnector;
        this.config = this.loadConfig();
        
        // Services de monitoring
        this.unifiedService = null;
        this.advancedMonitor = null;
        
        this.isInitialized = false;
        this.isActive = false;
        
        console.log('🔄 [HYBRID] Service de monitoring hybride créé');
    }

    loadConfig() {
        try {
            const settingsPath = path.join(__dirname, '../../config/app-settings.json');
            if (fs.existsSync(settingsPath)) {
                const settingsData = fs.readFileSync(settingsPath, 'utf8');
                const settings = JSON.parse(settingsData);
                return {
                    useAdvancedMonitoring: true,  // Utiliser le système style VBA
                    enableUnifiedFallback: true,  // Garder l'ancien système en backup
                    enableEventLogging: true,
                    syncOnStartup: true,
                    ...settings.monitoring
                };
            }
        } catch (error) {
            console.warn('⚠️ Erreur chargement configuration:', error.message);
        }
        
        return {
            useAdvancedMonitoring: true,
            enableUnifiedFallback: false,  // Désactiver le backup par défaut pour éviter la concurrence
            enableEventLogging: true,
            syncOnStartup: true,
            treatReadEmailsAsProcessed: false,
            scanInterval: 30000,
            autoStart: true
        };
    }

    /**
     * Configurer la transmission des événements
     */
    setupEventForwarding() {
        if (this.advancedMonitor) {
            // Événements de l'AdvancedMonitor (style VBA)
            this.advancedMonitor.on('emailArrived', (data) => {
                console.log(`📡 [HYBRID] Email arrivé: ${data.subject.substring(0, 50)}...`);
                this.emit('emailArrived', data);
            });
            
            this.advancedMonitor.on('emailRemoved', (data) => {
                console.log(`📡 [HYBRID] Email supprimé: ${data.subject.substring(0, 50)}...`);
                this.emit('emailRemoved', data);
            });
            
            this.advancedMonitor.on('monitoringStarted', () => {
                console.log('📡 [HYBRID] Monitoring avancé démarré');
                this.emit('monitoringStarted');
            });
            
            this.advancedMonitor.on('monitoringStopped', () => {
                console.log('📡 [HYBRID] Monitoring avancé arrêté');
                this.emit('monitoringStopped');
            });
        }
        
        if (this.unifiedService) {
            // Événements PowerShell du UnifiedService pour l'interface de chargement
            this.unifiedService.on('analysis-progress', (progressData) => {
                console.log(`📡 [HYBRID] Propagation événement analysis-progress: ${progressData.data.message}`);
                this.emit('analysis-progress', progressData);
            });
        }
    }

    /**
     * Initialiser le service hybride
     */
    async initialize(foldersConfig = {}) {
        try {
            console.log('🚀 [HYBRID] Initialisation du service hybride...');
            
            // Sauvegarder la configuration
            this.foldersConfig = foldersConfig;
            
            if (this.config.useAdvancedMonitoring) {
                // Initialiser le monitoring avancé (style VBA)
                console.log('🔧 [HYBRID] Initialisation du monitoring avancé (style VBA)...');
                const AdvancedEmailMonitor = require('./advancedEmailMonitor');
                this.advancedMonitor = new AdvancedEmailMonitor(this.outlookConnector, databaseService);
                await this.advancedMonitor.initialize(foldersConfig);
                
                if (this.config.enableUnifiedFallback) {
                    // Initialiser aussi le service unifié en backup
                    console.log('🔧 [HYBRID] Initialisation du service unifié (backup)...');
                    const UnifiedMonitoringService = require('./unifiedMonitoringService');
                    this.unifiedService = new UnifiedMonitoringService(this.outlookConnector);
                    await this.unifiedService.initialize(foldersConfig);
                }
            } else {
                // Utiliser uniquement le service unifié
                console.log('🔧 [HYBRID] Initialisation du service unifié uniquement...');
                const UnifiedMonitoringService = require('./unifiedMonitoringService');
                this.unifiedService = new UnifiedMonitoringService(this.outlookConnector);
                await this.unifiedService.initialize(foldersConfig);
            }
            
            // Configurer la transmission des événements
            this.setupEventForwarding();
            
            this.isInitialized = true;
            console.log('✅ [HYBRID] Service hybride initialisé avec succès');
            
            // Démarrer automatiquement le monitoring
            if (this.config.syncOnStartup) {
                await this.startMonitoring();
            }
            
        } catch (error) {
            console.error('❌ [HYBRID] Erreur initialisation:', error);
            throw error;
        }
    }

    /**
     * Démarrer le monitoring hybride
     */
    async startMonitoring() {
        try {
            if (!this.isInitialized) {
                throw new Error('Service non initialisé');
            }
            
            console.log('▶️ [HYBRID] Démarrage du monitoring hybride...');
            
            if (this.config.useAdvancedMonitoring && this.advancedMonitor) {
                // Démarrer le monitoring avancé (style VBA)
                await this.advancedMonitor.startActiveMonitoring();
                
                // Si configuré, démarrer aussi le service unifié en arrière-plan
                if (this.config.enableUnifiedFallback && this.unifiedService) {
                    try {
                        await this.unifiedService.startMonitoring();
                    } catch (error) {
                        console.warn('⚠️ [HYBRID] Erreur démarrage service unifié (backup):', error.message);
                    }
                }
            } else if (this.unifiedService) {
                // Utiliser uniquement le service unifié
                await this.unifiedService.startMonitoring();
            }
            
            this.isActive = true;
            console.log('✅ [HYBRID] Monitoring hybride démarré');
            
        } catch (error) {
            console.error('❌ [HYBRID] Erreur démarrage monitoring:', error);
            throw error;
        }
    }

    /**
     * Arrêter le monitoring
     */
    async stopMonitoring() {
        try {
            console.log('⏹️ [HYBRID] Arrêt du monitoring hybride...');
            
            if (this.advancedMonitor) {
                await this.advancedMonitor.stopMonitoring();
            }
            
            if (this.unifiedService && this.unifiedService.isMonitoring) {
                await this.unifiedService.stopMonitoring();
            }
            
            this.isActive = false;
            console.log('✅ [HYBRID] Monitoring hybride arrêté');
            
        } catch (error) {
            console.error('❌ [HYBRID] Erreur arrêt monitoring:', error);
        }
    }

    /**
     * Obtenir les statistiques combinées
     */
    async getStats() {
        try {
            const stats = {
                hybrid: {
                    isActive: this.isActive,
                    useAdvancedMonitoring: this.config.useAdvancedMonitoring,
                    enableUnifiedFallback: this.config.enableUnifiedFallback
                }
            };
            
            if (this.advancedMonitor) {
                stats.advanced = this.advancedMonitor.getMonitoringStats();
            }
            
            if (this.unifiedService) {
                stats.unified = await this.unifiedService.getMonitoringStatus();
            }
            
            // Statistiques de base de données
            try {
                stats.database = await databaseService.getEmailStats();
                stats.weeklyStats = await databaseService.getWeeklyStats(5);
            } catch (error) {
                console.warn('⚠️ [HYBRID] Erreur récupération stats DB:', error.message);
                stats.database = { total: 0, read: 0, unread: 0 };
                stats.weeklyStats = [];
            }
            
            return stats;
            
        } catch (error) {
            console.error('❌ [HYBRID] Erreur récupération stats:', error);
            return {
                hybrid: { isActive: false },
                error: error.message
            };
        }
    }

    /**
     * Obtenir les emails récents
     */
    async getRecentEmails(limit = 50) {
        try {
            return await databaseService.getRecentEmails(limit);
        } catch (error) {
            console.error('❌ [HYBRID] Erreur récupération emails récents:', error);
            return [];
        }
    }

    /**
     * Obtenir les statistiques de base de données
     */
    async getDatabaseStats() {
        try {
            const stats = await databaseService.getEmailStats();
            const weeklyStats = await databaseService.getWeeklyStats(10);
            
            return {
                ...stats,
                weeklyStatsAvailable: weeklyStats.length,
                lastWeekStats: weeklyStats[0] || null
            };
        } catch (error) {
            console.error('❌ [HYBRID] Erreur stats base de données:', error);
            return {
                totalEmails: 0,
                unreadCount: 0,
                totalEvents: 0,
                weeklyStatsAvailable: 0
            };
        }
    }

    /**
     * Obtenir les statistiques par catégorie
     */
    async getStatsByCategory() {
        try {
            const categoryStats = await databaseService.getCategoryStats();
            
            // Convertir le format pour correspondre aux 3 catégories attendues
            const categories = {};
            
            categoryStats.forEach(stat => {
                // Normaliser les noms de catégories selon notre configuration
                let categoryName = stat.category;
                
                // Mapping des anciennes catégories vers les nouvelles
                if (categoryName === 'declaration' || categoryName === 'declarations') {
                    categoryName = 'Déclarations';
                } else if (categoryName === 'reglement' || categoryName === 'reglements') {
                    categoryName = 'Règlements';
                } else if (categoryName === 'mails_simple' || categoryName === 'mails_simples') {
                    categoryName = 'Mails simples';
                } else {
                    // Ignorer les autres catégories qui ne font pas partie des 3 principales
                    return;
                }
                
                categories[categoryName] = {
                    emailsReceived: stat.count,
                    unreadCount: stat.unread
                };
            });
            
            return { categories };
        } catch (error) {
            console.error('❌ [HYBRID] Erreur stats par catégorie:', error);
            return { categories: {} };
        }
    }

    /**
     * Générer un rapport hebdomadaire (style VBA)
     */
    async generateWeeklyReport(weekIdentifier = null) {
        try {
            if (this.advancedMonitor) {
                return await this.advancedMonitor.generateWeeklyReport(weekIdentifier);
            } else {
                // Rapport basique depuis la base de données
                const weeklyStats = await databaseService.getWeeklyStats(1);
                return {
                    weekIdentifier: weekIdentifier || 'Semaine courante',
                    summary: weeklyStats[0] || { arrivals: 0, treatments: 0 },
                    generatedAt: new Date()
                };
            }
        } catch (error) {
            console.error('❌ [HYBRID] Erreur génération rapport:', error);
            throw error;
        }
    }

    /**
     * Forcer une synchronisation complète
     */
    async forceSync() {
        try {
            console.log('🔄 [HYBRID] Synchronisation forcée...');
            
            if (this.advancedMonitor) {
                // Re-synchroniser le monitoring avancé
                await this.advancedMonitor.performInitialSynchronization();
            }
            
            if (this.unifiedService) {
                // Synchroniser aussi le service unifié
                await this.unifiedService.forceSync();
            }
            
            console.log('✅ [HYBRID] Synchronisation forcée terminée');
            
        } catch (error) {
            console.error('❌ [HYBRID] Erreur synchronisation forcée:', error);
            throw error;
        }
    }

    /**
     * Méthode de compatibilité pour getMonitoringStatus
     */
    async getMonitoringStatus() {
        return await this.getStats();
    }

    /**
     * Nettoyer les ressources
     */
    async cleanup() {
        await this.stopMonitoring();
        
        if (this.advancedMonitor) {
            await this.advancedMonitor.cleanup();
        }
        
        if (this.unifiedService) {
            await this.unifiedService.cleanup();
        }
        
        console.log('🧹 [HYBRID] Nettoyage terminé');
    }

    saveConfig() {
        try {
            const settingsPath = path.join(__dirname, '../../config/app-settings.json');
            fs.writeFileSync(settingsPath, JSON.stringify(this.config, null, 2));
            console.log('💾 Configuration sauvegardée');
        } catch (error) {
            console.error('❌ Erreur sauvegarde configuration:', error);
        }
    }
}

module.exports = HybridMonitoringService;
