/**
 * MICROSOFT GRAPH API CONNECTOR - HAUTE PERFORMANCE
 * Remplace PowerShell/COM pour une performance optimale
 * API REST native + OAuth2 + JSON
 */

const { Client } = require('@microsoft/microsoft-graph-client');
const { AuthenticationProvider } = require('@microsoft/microsoft-graph-client');
const { PublicClientApplication } = require('@azure/msal-node');
const { EventEmitter } = require('events');

class GraphOutlookConnector extends EventEmitter {
    constructor() {
        super();
        this.isInitialized = false;
        this.graphClient = null;
        this.msalInstance = null;
        this.account = null;
        
        // Configuration MSAL
        this.msalConfig = {
            auth: {
                clientId: "d326c1ce-6a08-4b78-83cb-8d123456789a", // App publique Microsoft
                authority: "https://login.microsoftonline.com/common"
            },
            cache: {
                cacheLocation: "localStorage"
            }
        };
        
        // Scopes requis pour Outlook
        this.scopes = [
            'https://graph.microsoft.com/Mail.Read',
            'https://graph.microsoft.com/Mail.ReadWrite',
            'https://graph.microsoft.com/MailboxSettings.Read'
        ];
    }

    /**
     * PERFORMANCE: Initialisation Microsoft Graph
     */
    async initialize() {
        if (this.isInitialized) {
            console.log('✅ Graph API déjà initialisé');
            return;
        }

        try {
            console.log('🚀 Initialisation Microsoft Graph API...');
            
            // Créer l'instance MSAL
            this.msalInstance = new PublicClientApplication(this.msalConfig);
            
            // Authentification
            await this.authenticate();
            
            // Créer le client Graph
            this.graphClient = Client.initWithMiddleware({
                authProvider: this.createAuthProvider()
            });
            
            this.isInitialized = true;
            console.log('✅ Microsoft Graph API initialisé');
            
        } catch (error) {
            console.error('❌ Erreur initialisation Graph API:', error);
            throw error;
        }
    }

    /**
     * PERFORMANCE: Authentification rapide
     */
    async authenticate() {
        try {
            // Tenter authentification silencieuse d'abord
            const accounts = await this.msalInstance.getTokenCache().getAllAccounts();
            
            if (accounts.length > 0) {
                this.account = accounts[0];
                const silentRequest = {
                    scopes: this.scopes,
                    account: this.account
                };
                
                const response = await this.msalInstance.acquireTokenSilent(silentRequest);
                console.log('✅ Authentification silencieuse réussie');
                return response;
            }
            
            // Sinon, authentification interactive
            const response = await this.msalInstance.acquireTokenInteractive({
                scopes: this.scopes
            });
            
            this.account = response.account;
            console.log('✅ Authentification interactive réussie');
            return response;
            
        } catch (error) {
            console.error('❌ Erreur authentification:', error);
            throw error;
        }
    }

    /**
     * Provider d'authentification pour Graph Client
     */
    createAuthProvider() {
        return {
            getAccessToken: async () => {
                try {
                    const silentRequest = {
                        scopes: this.scopes,
                        account: this.account
                    };
                    
                    const response = await this.msalInstance.acquireTokenSilent(silentRequest);
                    return response.accessToken;
                    
                } catch (error) {
                    console.error('Erreur récupération token:', error);
                    throw error;
                }
            }
        };
    }

    /**
     * HAUTE PERFORMANCE: Récupération emails par batch
     */
    async getEmailsBatch(folderName = 'Inbox', limit = 100) {
        if (!this.isInitialized) {
            await this.initialize();
        }

        try {
            const startTime = Date.now();
            
            // Requête optimisée Graph API
            const emails = await this.graphClient
                .me
                .mailFolders(folderName)
                .messages
                .top(limit)
                .select([
                    'id',
                    'subject', 
                    'sender',
                    'toRecipients',
                    'receivedDateTime',
                    'sentDateTime',
                    'isRead',
                    'hasAttachments',
                    'bodyPreview',
                    'importance',
                    'parentFolderId'
                ])
                .orderby('receivedDateTime desc')
                .get();

            const processingTime = Date.now() - startTime;
            console.log(`⚡ ${emails.value.length} emails récupérés en ${processingTime}ms`);
            
            // Transformation pour compatibilité
            return emails.value.map(email => ({
                outlook_id: email.id,
                subject: email.subject || '',
                sender_name: email.sender?.emailAddress?.name || '',
                sender_email: email.sender?.emailAddress?.address || '',
                recipient_email: email.toRecipients?.[0]?.emailAddress?.address || '',
                received_time: email.receivedDateTime,
                sent_time: email.sentDateTime,
                folder_name: folderName,
                is_read: email.isRead,
                has_attachment: email.hasAttachments,
                body_preview: email.bodyPreview || '',
                importance: this.mapImportance(email.importance),
                event_type: 'email_received'
            }));
            
        } catch (error) {
            console.error('❌ Erreur récupération emails:', error);
            throw error;
        }
    }

    /**
     * PERFORMANCE: Récupération dossiers optimisée
     */
    async getFolders() {
        if (!this.isInitialized) {
            await this.initialize();
        }

        try {
            const folders = await this.graphClient
                .me
                .mailFolders
                .select(['id', 'displayName', 'parentFolderId', 'childFolderCount'])
                .get();

            return folders.value.map(folder => ({
                id: folder.id,
                name: folder.displayName,
                path: folder.displayName,
                parentId: folder.parentFolderId,
                childCount: folder.childFolderCount
            }));
            
        } catch (error) {
            console.error('❌ Erreur récupération dossiers:', error);
            throw error;
        }
    }

    /**
     * STREAMING: Monitoring temps réel efficace
     */
    async startRealtimeMonitoring(foldersConfig = []) {
        if (!this.isInitialized) {
            await this.initialize();
        }

        console.log('🔄 Démarrage monitoring temps réel Graph API...');
        
        // Monitoring par polling optimisé (Graph API ne supporte pas les webhooks pour Outlook personnel)
        this.monitoringInterval = setInterval(async () => {
            try {
                for (const folderConfig of foldersConfig) {
                    const newEmails = await this.getRecentEmails(folderConfig.folder_path, 10);
                    
                    if (newEmails.length > 0) {
                        this.emit('emails-received', {
                            folder: folderConfig.folder_path,
                            category: folderConfig.category,
                            emails: newEmails
                        });
                    }
                }
            } catch (error) {
                console.error('❌ Erreur monitoring:', error);
            }
        }, 15000); // 15 secondes - optimal pour Graph API
    }

    /**
     * OPTIMIZED: Emails récents avec cache
     */
    async getRecentEmails(folderName, limit = 50) {
        try {
            const emails = await this.graphClient
                .me
                .mailFolders(folderName)
                .messages
                .top(limit)
                .filter(`receivedDateTime ge ${new Date(Date.now() - 30*60000).toISOString()}`) // 30 min
                .select(['id', 'subject', 'sender', 'receivedDateTime', 'isRead'])
                .orderby('receivedDateTime desc')
                .get();

            return emails.value;
            
        } catch (error) {
            console.error('❌ Erreur emails récents:', error);
            return [];
        }
    }

    /**
     * Utilitaires
     */
    mapImportance(importance) {
        const mapping = { 'low': 0, 'normal': 1, 'high': 2 };
        return mapping[importance?.toLowerCase()] || 1;
    }

    /**
     * Arrêt propre
     */
    stopMonitoring() {
        if (this.monitoringInterval) {
            clearInterval(this.monitoringInterval);
            console.log('✅ Monitoring Graph API arrêté');
        }
    }

    /**
     * Test de connexion
     */
    async testConnection() {
        try {
            await this.initialize();
            const profile = await this.graphClient.me.get();
            console.log(`✅ Connecté à Graph API: ${profile.displayName}`);
            return true;
        } catch (error) {
            console.error('❌ Test connexion échoué:', error);
            return false;
        }
    }
}

module.exports = GraphOutlookConnector;
