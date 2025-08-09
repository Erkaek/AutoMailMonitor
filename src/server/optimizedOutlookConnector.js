/**
 * OUTLOOK CONNECTOR OPTIMISÉ - MICROSOFT GRAPH API
 * Performance maximale avec API REST + Better-SQLite3 + Cache
 */

const { EventEmitter } = require('events');
const GraphOutlookConnector = require('./graphOutlookConnector');

class OptimizedOutlookConnector extends EventEmitter {
    constructor() {
        super();
        
        // Graph API Connector (haute performance)
        this.graphConnector = new GraphOutlookConnector();
        this.isConnected = false;
        this.connectionState = 'disconnected';
        
        // Configuration optimisée
        this.config = {
            realtimePollingInterval: 15000, // 15s - optimal pour Graph API
            batchSize: 100, // emails par batch
            enableDetailedLogs: false, // Performance
            maxRetries: 3
        };
        
        // Cache et statistiques
        this.folders = new Map();
        this.stats = {
            totalEmails: 0,
            apiCalls: 0,
            lastSyncTime: null,
            avgResponseTime: 0
        };
        
        // Monitoring interval
        this.monitoringInterval = null;
    }

    /**
     * PERFORMANCE: Connexion Graph API rapide
     */
    async connect() {
        if (this.isConnected) {
            console.log('✅ Déjà connecté à Graph API');
            return true;
        }

        try {
            this.connectionState = 'connecting';
            console.log('🚀 Connexion à Microsoft Graph API...');
            
            const startTime = Date.now();
            
            // Initialiser Graph API
            await this.graphConnector.initialize();
            
            // Test de connexion
            const connected = await this.graphConnector.testConnection();
            
            if (connected) {
                this.isConnected = true;
                this.connectionState = 'connected';
                
                const connectionTime = Date.now() - startTime;
                console.log(`✅ Connecté à Graph API en ${connectionTime}ms`);
                
                this.emit('connected');
                return true;
            }
            
            throw new Error('Test de connexion échoué');
            
        } catch (error) {
            this.connectionState = 'error';
            console.error('❌ Erreur connexion Graph API:', error);
            this.emit('error', error);
            return false;
        }
    }

    /**
     * HAUTE PERFORMANCE: Récupération dossiers optimisée
     */
    async getFolders() {
        if (!this.isConnected) {
            await this.connect();
        }

        try {
            const startTime = Date.now();
            
            const folders = await this.graphConnector.getFolders();
            
            // Mise en cache
            this.folders.clear();
            folders.forEach(folder => {
                this.folders.set(folder.path, folder);
            });
            
            const responseTime = Date.now() - startTime;
            this.updateStats('getFolders', responseTime);
            
            console.log(`⚡ ${folders.length} dossiers récupérés en ${responseTime}ms`);
            return folders;
            
        } catch (error) {
            console.error('❌ Erreur récupération dossiers:', error);
            throw error;
        }
    }

    /**
     * BATCH PROCESSING: Récupération emails par lots
     */
    async getEmailsFromFolder(folderPath, limit = 100) {
        if (!this.isConnected) {
            await this.connect();
        }

        try {
            const startTime = Date.now();
            
            const emails = await this.graphConnector.getEmailsBatch(folderPath, limit);
            
            const responseTime = Date.now() - startTime;
            this.updateStats('getEmails', responseTime);
            
            console.log(`⚡ ${emails.length} emails de ${folderPath} en ${responseTime}ms`);
            
            return emails;
            
        } catch (error) {
            console.error(`❌ Erreur emails ${folderPath}:`, error);
            return [];
        }
    }

    /**
     * MONITORING TEMPS RÉEL optimisé
     */
    async startRealtimeMonitoring(foldersConfig = []) {
        if (!this.isConnected) {
            await this.connect();
        }

        console.log('🔄 Démarrage monitoring temps réel optimisé...');
        
        // Monitoring avec Graph API
        await this.graphConnector.startRealtimeMonitoring(foldersConfig);
        
        // Écouter les nouveaux emails
        this.graphConnector.on('emails-received', (data) => {
            this.emit('emails-received', data);
            this.stats.totalEmails += data.emails.length;
        });
        
        // Monitoring des statistiques
        this.monitoringInterval = setInterval(() => {
            this.stats.lastSyncTime = new Date().toISOString();
            this.emit('stats-updated', this.getStats());
        }, 30000); // 30s
        
        this.emit('monitoring-started');
        console.log('✅ Monitoring temps réel actif');
    }

    /**
     * SYNC COMPLET performant
     */
    async performFullSync(foldersConfig = []) {
        if (!this.isConnected) {
            await this.connect();
        }

        console.log('🚀 Synchronisation complète optimisée...');
        const startTime = Date.now();
        
        const results = {
            totalEmails: 0,
            foldersProcessed: 0,
            errors: []
        };

        try {
            // Traitement par batch pour performance
            const batchPromises = foldersConfig.map(async (folderConfig) => {
                try {
                    const emails = await this.getEmailsFromFolder(
                        folderConfig.folder_path, 
                        this.config.batchSize
                    );
                    
                    // Ajouter métadonnées
                    const enrichedEmails = emails.map(email => ({
                        ...email,
                        category: folderConfig.category,
                        folder_path: folderConfig.folder_path
                    }));
                    
                    results.totalEmails += emails.length;
                    results.foldersProcessed++;
                    
                    // Émettre par batch pour traitement
                    this.emit('emails-batch', {
                        folder: folderConfig.folder_path,
                        category: folderConfig.category,
                        emails: enrichedEmails
                    });
                    
                    return enrichedEmails;
                    
                } catch (error) {
                    console.error(`❌ Erreur sync ${folderConfig.folder_path}:`, error);
                    results.errors.push({
                        folder: folderConfig.folder_path,
                        error: error.message
                    });
                    return [];
                }
            });

            await Promise.all(batchPromises);
            
            const syncTime = Date.now() - startTime;
            console.log(`✅ Sync complète: ${results.totalEmails} emails en ${syncTime}ms`);
            
            results.syncTime = syncTime;
            this.emit('sync-completed', results);
            
            return results;
            
        } catch (error) {
            console.error('❌ Erreur sync complète:', error);
            throw error;
        }
    }

    /**
     * Statistiques de performance
     */
    getStats() {
        return {
            ...this.stats,
            isConnected: this.isConnected,
            connectionState: this.connectionState,
            foldersCount: this.folders.size,
            uptime: this.isConnected ? Date.now() - (this.stats.startTime || Date.now()) : 0
        };
    }

    /**
     * Mise à jour des métriques de performance
     */
    updateStats(operation, responseTime) {
        this.stats.apiCalls++;
        this.stats.avgResponseTime = (this.stats.avgResponseTime + responseTime) / 2;
        this.stats.lastSyncTime = new Date().toISOString();
    }

    /**
     * Test de performance
     */
    async runPerformanceTest() {
        console.log('🧪 Test de performance Graph API...');
        
        const tests = [
            { name: 'Connection', fn: () => this.connect() },
            { name: 'Get Folders', fn: () => this.getFolders() },
            { name: 'Get Inbox Emails', fn: () => this.getEmailsFromFolder('Inbox', 50) }
        ];

        const results = [];
        
        for (const test of tests) {
            const startTime = Date.now();
            try {
                await test.fn();
                const time = Date.now() - startTime;
                results.push({ test: test.name, time, status: 'success' });
                console.log(`✅ ${test.name}: ${time}ms`);
            } catch (error) {
                const time = Date.now() - startTime;
                results.push({ test: test.name, time, status: 'error', error: error.message });
                console.log(`❌ ${test.name}: ${time}ms (erreur)`);
            }
        }
        
        return results;
    }

    /**
     * Arrêt propre
     */
    async disconnect() {
        console.log('🔌 Déconnexion Graph API...');
        
        if (this.monitoringInterval) {
            clearInterval(this.monitoringInterval);
        }
        
        this.graphConnector.stopMonitoring();
        this.isConnected = false;
        this.connectionState = 'disconnected';
        
        this.emit('disconnected');
        console.log('✅ Déconnexion propre effectuée');
    }
}

module.exports = OptimizedOutlookConnector;
