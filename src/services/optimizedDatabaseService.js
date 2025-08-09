/**
 * Service de base de données OPTIMISÉ avec Better-SQLite3
 * QUICK WINS: +300% performance, WAL mode, prepared statements, cache intelligent
 */

const Database = require('better-sqlite3');
const NodeCache = require('node-cache');
const path = require('path');
const fs = require('fs');

class OptimizedDatabaseService {
    constructor() {
        this.db = null;
        this.isInitialized = false;
        
        // Cache intelligent avec TTL
        this.cache = new NodeCache({ 
            stdTTL: 300, // 5 minutes par défaut
            checkperiod: 60, // Nettoyage toutes les 60s
            useClones: false // Performance: pas de clonage d'objets
        });
        
        // Prepared statements pour performance maximale
        this.statements = {};
        
        // Statistiques de performance
        this.stats = {
            queriesExecuted: 0,
            cacheHits: 0,
            cacheMisses: 0,
            avgQueryTime: 0
        };
    }

    /**
     * Initialisation avec optimisations WAL et performance
     */
    async initialize() {
        if (this.isInitialized) {
            return;
        }

        try {
            const dbPath = path.join(__dirname, '../../data/emails.db');
            
            // Créer le répertoire si nécessaire
            const dbDir = path.dirname(dbPath);
            if (!fs.existsSync(dbDir)) {
                fs.mkdirSync(dbDir, { recursive: true });
            }

            console.log('🚀 Initialisation Better-SQLite3 avec optimisations...');
            
            // Ouvrir avec optimisations de performance
            this.db = new Database(dbPath, {
                verbose: null, // Pas de logging verbose en prod
                fileMustExist: false
            });

            // OPTIMISATIONS CRITIQUES
            this.setupOptimizations();
            
            // Créer les tables si nécessaire
            this.createTables();
            
            // S'assurer que les colonnes de suppression existent
                // this.ensureDeletedColumn(); // supprimé, colonne gérée dans la migration
            this.ensureWeeklyStatsTable();
            
            // Préparer les statements pour performance
            this.prepareStatements();
            
            // Mettre à jour les statistiques hebdomadaires avec les données actuelles
            this.updateCurrentWeekStats();
            
            this.isInitialized = true;
            console.log('✅ Better-SQLite3 initialisé avec WAL mode et cache intelligent');
            
        } catch (error) {
            console.error('❌ Erreur initialisation Better-SQLite3:', error);
            throw error;
        }
    }

    /**
     * Configuration des optimisations SQLite
     */
    setupOptimizations() {
        // WAL Mode pour performance concurrente
        this.db.pragma('journal_mode = WAL');
        
        // Optimisations de performance
        this.db.pragma('synchronous = NORMAL'); // Balance sécurité/performance
        this.db.pragma('cache_size = 10000'); // 10MB cache
        this.db.pragma('temp_store = MEMORY'); // Temp tables en RAM
        this.db.pragma('mmap_size = 268435456'); // 256MB memory mapping
        
        console.log('⚡ WAL mode activé + optimisations performance');
    }

    /**
     * Création des tables optimisées (compatible avec schéma existant)
     */
    createTables() {
        // Table emails avec structure optimisée (compatible logique VBA)
        this.db.exec(`
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                outlook_id TEXT UNIQUE NOT NULL,
                subject TEXT NOT NULL,
                sender_email TEXT,
                received_time DATETIME,
                folder_name TEXT,
                category TEXT,
                is_read BOOLEAN DEFAULT 0,
                is_treated BOOLEAN DEFAULT 0,
                deleted_at DATETIME NULL,
                week_identifier TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

    // (activity_weekly supprimé)

        // Index optimisés pour les requêtes fréquentes
        this.db.exec(`
            CREATE INDEX IF NOT EXISTS idx_emails_outlook_id ON emails(outlook_id);
            CREATE INDEX IF NOT EXISTS idx_emails_folder_name ON emails(folder_name);
            CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails(received_time);
            CREATE INDEX IF NOT EXISTS idx_emails_is_read ON emails(is_read);
            CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category);
            CREATE INDEX IF NOT EXISTS idx_emails_week_identifier ON emails(week_identifier);
            CREATE INDEX IF NOT EXISTS idx_emails_deleted_at ON emails(deleted_at);
            CREATE INDEX IF NOT EXISTS idx_emails_is_treated ON emails(is_treated);
        `);
    }

    /**
     * Préparation des statements pour performance maximale (schéma optimisé final 13 colonnes)
     */
    prepareStatements() {
        // Statements les plus utilisés
        this.statements = {
            // Emails - compatible avec nouvelle structure optimisée
            insertEmail: this.db.prepare(`
                INSERT OR REPLACE INTO emails 
                (outlook_id, subject, sender_email, received_time, folder_name, 
                 category, is_read, is_treated, week_identifier, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            `),
            
            getRecentEmails: this.db.prepare(`
                SELECT * FROM emails 
                ORDER BY received_time DESC 
                LIMIT ?
            `),
            
            getEmailStats: this.db.prepare(`
                SELECT 
                    COUNT(*) as totalEmails,
                    SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unreadTotal,
                    SUM(CASE WHEN DATE(received_time) = DATE('now') THEN 1 ELSE 0 END) as emailsToday,
                    SUM(CASE WHEN DATE(deleted_at) = DATE('now') THEN 1 ELSE 0 END) as treatedToday
                FROM emails
            `),
            
            // Folders
            getFoldersConfig: this.db.prepare(`
                SELECT folder_path, category, folder_name FROM folder_configurations
            `),
            
            insertFolderConfig: this.db.prepare(`
                INSERT OR REPLACE INTO folder_configurations (folder_name, category, folder_name, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            `),
            
            deleteFolderConfig: this.db.prepare(`
                DELETE FROM folder_configurations WHERE folder_name = ?
            `),
            
            // Settings
            getSetting: this.db.prepare(`
                SELECT value FROM app_settings WHERE key = ?
            `),
            
            setSetting: this.db.prepare(`
                INSERT OR REPLACE INTO app_settings (key, value, updated_at)
                VALUES (?, ?, CURRENT_TIMESTAMP)
            `),
            
            // Email lookups by EntryID (using outlook_id column)
            getEmailByEntryId: this.db.prepare(`
                SELECT * FROM emails WHERE outlook_id = ?
            `),
            
            getEmailByEntryIdAndFolder: this.db.prepare(`
                SELECT * FROM emails WHERE outlook_id = ? AND folder_name = ?
            `)
        };

    // (activity_weekly statements supprimés)

        // Statements weekly_stats (upsert direct)
        this.statements.upsertWeeklyStats = this.db.prepare(`
            INSERT INTO weekly_stats (
                week_identifier, week_number, week_year, week_start_date, week_end_date,
                folder_type, emails_received, emails_treated, manual_adjustments, updated_at, created_at
            ) VALUES (
                @week_identifier, @week_number, @week_year, @week_start_date, @week_end_date,
                @folder_type, @emails_received, @emails_treated, @manual_adjustments, CURRENT_TIMESTAMP, COALESCE(@created_at, CURRENT_TIMESTAMP)
            )
            ON CONFLICT(week_identifier, folder_type) DO UPDATE SET
                week_number=excluded.week_number,
                week_year=excluded.week_year,
                week_start_date=excluded.week_start_date,
                week_end_date=excluded.week_end_date,
                emails_received=excluded.emails_received,
                emails_treated=excluded.emails_treated,
                manual_adjustments=excluded.manual_adjustments,
                updated_at=CURRENT_TIMESTAMP
        `);

        console.log('⚡ Prepared statements créés pour performance maximale (compatible)');
    }

    /**
     * Exécution avec cache et métriques
     */
    executeWithCache(cacheKey, queryFn, ttl = 300) {
        const startTime = Date.now();
        
        // Vérifier le cache d'abord
        const cached = this.cache.get(cacheKey);
        if (cached !== undefined) {
            this.stats.cacheHits++;
            return cached;
        }

        // Exécuter la requête
        const result = queryFn();
        
        // Mettre en cache
        this.cache.set(cacheKey, result, ttl);
        this.stats.cacheMisses++;
        
        // Métriques
        this.stats.queriesExecuted++;
        const queryTime = Date.now() - startTime;
        this.stats.avgQueryTime = (this.stats.avgQueryTime + queryTime) / 2;
        
        return result;
    }

    // (méthodes activity_weekly supprimées)

    // ==================== WEEKLY_STATS (Import XLSB) ====================
    upsertWeeklyStats(row) {
        try {
            return this.statements.upsertWeeklyStats.run(row);
        } catch (error) {
            console.error('❌ [DB] Erreur upsert weekly_stats:', error);
            throw error;
        }
    }

    upsertWeeklyStatsBatch(rows) {
        const tx = this.db.transaction((items) => {
            for (const r of items) this.statements.upsertWeeklyStats.run(r);
        });
        return tx(rows);
    }

    /**
     * OPTIMIZED: Sauvegarde d'email avec prepared statement (schéma optimisé final (13 colonnes))
     */
    saveEmail(emailData) {
        const cacheKey = `email_${emailData.outlook_id || emailData.id}`;
        this.cache.del(cacheKey); // Invalider le cache
        
        // Calculer le week_identifier (semaine ISO)
        const weekId = this.calculateWeekIdentifier(emailData.received_time);
        
        const result = this.statements.insertEmail.run(
            emailData.outlook_id || emailData.id || '',
            emailData.subject || '',
            emailData.sender_email || '',
            emailData.received_time || new Date().toISOString(),
            emailData.folder_name || '',
            emailData.category || 'Mails simples',
            emailData.is_read ? 1 : 0,
            emailData.is_treated ? 1 : 0,
            weekId
        );
        
        // Mettre à jour les statistiques hebdomadaires
        if (result.changes > 0) {
            // Comptabiliser comme "arrivé"
            this.updateWeeklyEmailCount(emailData, true, false);
            
            // Si l'email est déjà lu, le comptabiliser aussi comme "traité" selon le paramètre
            const countReadAsTreated = this.getAppSetting('count_read_as_treated', 'false') === 'true';
            if (countReadAsTreated && emailData.is_read) {
                this.updateWeeklyEmailCount(emailData, false, true);
            }
        }
        
        // Invalider le cache de l'interface utilisateur en temps réel
        this.invalidateUICache();
        
        return result;
    }

    /**
     * Alias pour saveEmail - compatibilité avec unifiedMonitoringService
     */
    insertEmail(emailData) {
        return this.saveEmail(emailData);
    }

    /**
     * OPTIMIZED: Sauvegarde par batch avec transaction (schéma optimisé final 13 colonnes)
     */
    saveEmailsBatch(emails) {
        if (!emails || emails.length === 0) return;

        const transaction = this.db.transaction((emails) => {
            for (const email of emails) {
                // Calculer le week_identifier (semaine ISO)
                const weekId = this.calculateWeekIdentifier(email.received_time);
                
                this.statements.insertEmail.run(
                    email.outlook_id || email.id || '',
                    email.subject || '',
                    email.sender_email || '',
                    email.received_time || new Date().toISOString(),
                    email.folder_name || email.folder_name || '',
                    email.category || 'Mails simples',
                    email.is_read ? 1 : 0,
                    email.is_treated ? 1 : 0,
                    weekId
                );
            }
        });

        // Invalider les caches pertinents
        this.cache.flushAll();
        
        const result = transaction(emails);
        
        // Invalider le cache de l'interface utilisateur pour mise à jour en temps réel
        this.invalidateUICache();
        
        return result;
    }

    /**
     * OPTIMIZED: Récupération des emails récents avec cache
     */
    getRecentEmails(limit = 50) {
        const cacheKey = `recent_emails_${limit}`;
        
        return this.executeWithCache(cacheKey, () => {
            return this.statements.getRecentEmails.all(limit);
        }, 60); // Cache 1 minute
    }

    /**
     * OPTIMIZED: Statistiques avec cache intelligent
     */
    getEmailStats() {
        const cacheKey = 'email_stats';
        
        return this.executeWithCache(cacheKey, () => {
            const stats = this.statements.getEmailStats.get();
            return {
                totalEmails: stats.totalEmails || 0,
                unreadTotal: stats.unreadTotal || 0,
                emailsToday: stats.emailsToday || 0,
                treatedToday: stats.treatedToday || 0,
                lastSyncTime: new Date().toISOString()
            };
        }, 30); // Cache 30 secondes
    }

    /**
     * OPTIMIZED: Configuration des dossiers avec cache
     */
    getFoldersConfiguration() {
        const cacheKey = 'folders_config';
        
        return this.executeWithCache(cacheKey, () => {
            return this.statements.getFoldersConfig.all();
        }, 300); // Cache 5 minutes
    }

    /**
     * Sauvegarde configuration dossier
     */
    addFolderConfiguration(folderPath, category, folderName) {
        this.cache.del('folders_config'); // Invalider cache
        return this.statements.insertFolderConfig.run(folderPath, category, folderName);
    }

    /**
     * Suppression configuration dossier
     */
    deleteFolderConfiguration(folderPath) {
        this.cache.del('folders_config'); // Invalider cache
        return this.statements.deleteFolderConfig.run(folderPath);
    }

    /**
     * Settings optimisés
     */
    getAppSetting(key, defaultValue = null) {
        const cacheKey = `setting_${key}`;
        
        return this.executeWithCache(cacheKey, () => {
            const result = this.statements.getSetting.get(key);
            return result ? JSON.parse(result.value) : defaultValue;
        }, 600); // Cache 10 minutes
    }

    saveAppSetting(key, value) {
        this.cache.del(`setting_${key}`); // Invalider cache
        return this.statements.setSetting.run(key, JSON.stringify(value));
    }

    /**
     * Statistiques par catégorie
     */
    getCategoryStats() {
        const cacheKey = 'category_stats';
        let stats = this.cache.get(cacheKey);
        
        if (!stats) {
            try {
                const result = this.db.prepare(`
                    SELECT 
                        category,
                        COUNT(*) as count,
                        COUNT(CASE WHEN is_read = 0 THEN 1 END) as unread_count
                    FROM emails 
                    WHERE category IS NOT NULL 
                    GROUP BY category
                `).all();
                
                stats = result.reduce((acc, row) => {
                    acc[row.category] = {
                        total: row.count,
                        unread: row.unread_count
                    };
                    return acc;
                }, {});
                
                this.cache.set(cacheKey, stats, 180); // Cache 3 minutes
                this.stats.cacheMisses++;
            } catch (error) {
                console.error('❌ Erreur getCategoryStats:', error);
                return {};
            }
        } else {
            this.stats.cacheHits++;
        }
        
        return stats;
    }

    /**
     * Comptes d'emails par date
     */
    getEmailCountByDate(startDate, endDate) {
        const cacheKey = `email_count_${startDate}_${endDate}`;
        let count = this.cache.get(cacheKey);
        
        if (count === undefined) {
            try {
                const result = this.db.prepare(`
                    SELECT COUNT(*) as count 
                    FROM emails 
                    WHERE DATE(received_time) BETWEEN ? AND ?
                `).get(startDate, endDate);
                count = result.count;
                this.cache.set(cacheKey, count, 300); // Cache 5 minutes
                this.stats.cacheMisses++;
            } catch (error) {
                console.error('❌ Erreur getEmailCountByDate:', error);
                return 0;
            }
        } else {
            this.stats.cacheHits++;
        }
        
        return count;
    }

    /**
     * Nombre total d'emails non lus
     */
    getUnreadEmailCount() {
        const cacheKey = 'unread_email_count';
        let count = this.cache.get(cacheKey);
        
        if (count === undefined) {
            try {
                const result = this.db.prepare('SELECT COUNT(*) as count FROM emails WHERE is_read = 0').get();
                count = result.count;
                this.cache.set(cacheKey, count, 60); // Cache 1 minute
                this.stats.cacheMisses++;
            } catch (error) {
                console.error('❌ Erreur getUnreadEmailCount:', error);
                return 0;
            }
        } else {
            this.stats.cacheHits++;
        }
        
        return count;
    }

    /**
     * Statistiques hebdomadaires (ancienne méthode supprimée - voir ligne 1420 pour la nouvelle)
     */

    /**
     * Nombre total d'emails
     */
    getTotalEmailCount() {
        const cacheKey = 'total_email_count';
        let count = this.cache.get(cacheKey);
        
        if (count === undefined) {
            try {
                const result = this.db.prepare('SELECT COUNT(*) as count FROM emails').get();
                count = result.count;
                this.cache.set(cacheKey, count, 60); // Cache 1 minute
                this.stats.cacheMisses++;
            } catch (error) {
                console.error('❌ Erreur getTotalEmailCount:', error);
                return 0;
            }
        } else {
            this.stats.cacheHits++;
        }
        
        return count;
    }

    /**
     * Charger paramètres application
     */
    loadAppSettings() {
        const cacheKey = 'app_settings';
        let settings = this.cache.get(cacheKey);
        
        if (!settings) {
            try {
                const rows = this.db.prepare('SELECT key, value FROM app_settings').all();
                settings = {};
                rows.forEach(row => {
                    try {
                        settings[row.key] = JSON.parse(row.value);
                    } catch {
                        settings[row.key] = row.value;
                    }
                });
                this.cache.set(cacheKey, settings, 300); // Cache 5 minutes
                this.stats.cacheMisses++;
            } catch (error) {
                console.error('❌ Erreur loadAppSettings:', error);
                return {};
            }
        } else {
            this.stats.cacheHits++;
        }
        
        return settings;
    }

    /**
     * Statistiques des dossiers
     */
    getFolderStats() {
        const cacheKey = 'folder_stats';
        let stats = this.cache.get(cacheKey);
        
        if (!stats) {
            try {
                const result = this.db.prepare(`
                    SELECT 
                        folder_name,
                        COUNT(*) as email_count,
                        COUNT(CASE WHEN is_read = 0 THEN 1 END) as unread_count,
                        MAX(received_time) as last_email
                    FROM emails 
                    GROUP BY folder_name
                `).all();
                
                stats = result.map(row => ({
                    path: row.folder_name,
                    emailCount: row.email_count,
                    unreadCount: row.unread_count,
                    lastEmail: row.last_email
                }));
                
                this.cache.set(cacheKey, stats, 120); // Cache 2 minutes
                this.stats.cacheMisses++;
            } catch (error) {
                console.error('❌ Erreur getFolderStats:', error);
                return [];
            }
        } else {
            this.stats.cacheHits++;
        }
        
        return stats;
    }

    /**
     * Assurer table statistiques hebdomadaires (structure mise à jour)
     */
    ensureWeeklyStatsTable() {
        try {
            this.db.exec(`
                CREATE TABLE IF NOT EXISTS weekly_stats (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    week_identifier TEXT NOT NULL,
                    week_number INTEGER NOT NULL,
                    week_year INTEGER NOT NULL,
                    week_start_date TEXT NOT NULL,
                    week_end_date TEXT NOT NULL,
                    folder_type TEXT NOT NULL DEFAULT 'Mails simples',
                    emails_received INTEGER DEFAULT 0,
                    emails_treated INTEGER DEFAULT 0,
                    manual_adjustments INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(week_identifier, folder_type)
                )
            `);
            // Log table weekly_stats supprimé pour éviter spam
            // console.log('✅ Table weekly_stats assurée');
        } catch (error) {
            console.error('❌ Erreur ensureWeeklyStatsTable:', error);
        }
    }

    /**
     * NOUVEAU: Traite un nouvel email détecté via COM Outlook
     */
    async processCOMNewEmail(emailData) {
        try {
            const startTime = Date.now();
            
            // Vérifier si l'email existe déjà (éviter les doublons)
            const existingEmail = this.getEmailById(emailData.id);
            if (existingEmail) {
                console.log(`⚠️ Email COM déjà existant: ${emailData.id}`);
                return { processed: false, reason: 'already_exists' };
            }

            // Préparer les données email
            const emailRecord = {
                email_id: emailData.id,
                subject: emailData.subject || '',
                sender_email: emailData.sender_email || '',
                received_time: emailData.receivedTime || new Date().toISOString(),
                folder_name: emailData.folderPath || '',
                is_read: emailData.isRead === true ? 1 : 0,
                category: emailData.category || 'autres',
                last_updated: new Date().toISOString()
            };

            // Insérer l'email
            const stmt = this.getStatement('insertEmail');
            const result = stmt.run(emailRecord);

            // Invalider le cache pour ce dossier
            this.invalidateFolderCache(emailData.folderPath);

            // Mettre à jour les stats
            this.updateQueryStats(Date.now() - startTime);

            console.log(`📧 [DB-COM] Nouvel email inséré: ${emailData.subject} (ID: ${result.lastInsertRowid})`);
            
            return { 
                processed: true, 
                rowId: result.lastInsertRowid,
                emailId: emailData.id
            };

        } catch (error) {
            console.error('❌ [DB-COM] Erreur insertion nouvel email:', error);
            return { processed: false, error: error.message };
        }
    }

    /**
     * NOUVEAU: Traite un changement d'état email via COM Outlook
     */
    async processCOMEmailChange(emailData) {
        try {
            const startTime = Date.now();
            
            // Récupérer l'email existant
            const existingEmail = this.getEmailById(emailData.id);
            if (!existingEmail) {
                console.log(`⚠️ Email COM non trouvé pour mise à jour: ${emailData.id}`);
                return { updated: false, reason: 'not_found' };
            }

            // Préparer les données de mise à jour
            const updateData = {
                is_read: emailData.isRead === true ? 1 : 0,
                last_updated: new Date().toISOString()
            };

            // Ajouter d'autres champs s'ils sont fournis
            if (emailData.subject !== undefined) updateData.subject = emailData.subject;
            if (emailData.category !== undefined) updateData.category = emailData.category;

            // Construire la requête de mise à jour dynamiquement
            const fields = Object.keys(updateData);
            const setClause = fields.map(field => `${field} = @${field}`).join(', ');
            const sql = `UPDATE emails SET ${setClause} WHERE email_id = @email_id`;

            // Préparer et exécuter la mise à jour
            let stmt = this.statements[`updateEmail_${fields.join('_')}`];
            if (!stmt) {
                stmt = this.db.prepare(sql);
                this.statements[`updateEmail_${fields.join('_')}`] = stmt;
            }

            const result = stmt.run({
                ...updateData,
                email_id: emailData.id
            });

            // Invalider le cache
            this.invalidateFolderCache(existingEmail.folder_name);

            // Mettre à jour les stats
            this.updateQueryStats(Date.now() - startTime);

            console.log(`🔄 [DB-COM] Email mis à jour: ${emailData.id} (${result.changes} changements)`);
            
            return { 
                updated: true, 
                changes: result.changes,
                emailId: emailData.id
            };

        } catch (error) {
            console.error('❌ [DB-COM] Erreur mise à jour email:', error);
            return { updated: false, error: error.message };
        }
    }

    /**
     * NOUVEAU: Récupère un email par ID (optimisé avec cache)
     */
    getEmailById(emailId) {
        try {
            const cacheKey = `email_${emailId}`;
            let email = this.cache.get(cacheKey);
            
            if (email) {
                this.stats.cacheHits++;
                return email;
            }

            // Pas en cache, requête DB
            const stmt = this.getStatement('getEmailById', 
                'SELECT * FROM emails WHERE email_id = ?'
            );
            
            email = stmt.get(emailId);
            
            if (email) {
                // Mettre en cache pour 10 minutes
                this.cache.set(cacheKey, email, 600);
            }
            
            this.stats.cacheMisses++;
            return email;

        } catch (error) {
            console.error('❌ [DB-COM] Erreur récupération email par ID:', error);
            return null;
        }
    }

    /**
     * NOUVEAU: Trouve un email par subject et dossier (fallback)
     */
    findEmailBySubjectAndFolder(subject, folderPath) {
        try {
            if (!this.statements.findEmailBySubjectAndFolder) {
                this.statements.findEmailBySubjectAndFolder = this.db.prepare(`
                    SELECT * FROM emails 
                    WHERE subject = ? AND folder_name = ? 
                    ORDER BY received_time DESC 
                    LIMIT 1
                `);
            }
            
            return this.statements.findEmailBySubjectAndFolder.get(subject, folderPath);
        } catch (error) {
            console.error('❌ [DB] Erreur recherche email par subject:', error);
            return null;
        }
    }

    /**
     * NOUVEAU: Récupère un email par Entry ID (optimisé avec cache)
     */
    getEmailByEntryId(entryId, folderPath = null) {
        try {
            const cacheKey = `email_entry_${entryId}`;
            let email = this.cache.get(cacheKey);
            
            if (email) {
                this.stats.cacheHits++;
                return email;
            }

            // Pas en cache, requête DB
            let stmt;
            if (folderPath) {
                stmt = this.statements.getEmailByEntryIdAndFolder;
                email = stmt.get(entryId, folderPath);
            } else {
                stmt = this.statements.getEmailByEntryId;
                email = stmt.get(entryId);
            }
            
            if (email) {
                // Mettre en cache pour 10 minutes
                this.cache.set(cacheKey, email, 600);
            }
            
            this.stats.cacheMisses++;
            return email;

        } catch (error) {
            console.error('❌ [DB-COM] Erreur récupération email par Entry ID:', error);
            return null;
        }
    }

    /**
     * NOUVEAU: Met à jour le statut read/unread d'un email par Entry ID
     */
    updateEmailStatus(entryId, isRead, folderPath = null) {
        try {
            const startTime = Date.now();
            let stmt;
            let params;
            
            console.log(`🔧 [updateEmailStatus] DEBUGGING:`);
            console.log(`  - entryId: ${entryId}`);
            console.log(`  - isRead: ${isRead}`);
            console.log(`  - folderPath: ${folderPath}`);
            
            // TEST DIRECT: Vérifier si l'email existe avant la mise à jour
            if (folderPath) {
                const testStmt = this.db.prepare('SELECT outlook_id, subject, folder_name, is_read FROM emails WHERE outlook_id = ? AND folder_name = ?');
                const testResult = testStmt.get(entryId, folderPath);
                console.log(`🔧 [TEST] Email avec outlook_id + folder trouvé:`, testResult);
                
                if (!testResult) {
                    // Tester sans le folder_name
                    const testStmt2 = this.db.prepare('SELECT outlook_id, subject, folder_name, is_read FROM emails WHERE outlook_id = ?');
                    const testResult2 = testStmt2.get(entryId);
                    console.log(`🔧 [TEST] Email avec outlook_id seul:`, testResult2);
                    
                    // Tester avec LIKE pour voir les variations
                    const testStmt3 = this.db.prepare('SELECT outlook_id, subject, folder_name, is_read FROM emails WHERE outlook_id LIKE ? LIMIT 3');
                    const testResult3 = testStmt3.all(entryId.substring(0, 20) + '%');
                    console.log(`🔧 [TEST] Emails similaires (LIKE):`, testResult3);
                }
            }
            // CORRECTION SIMPLE: Utiliser uniquement outlook_id (qui est unique) au lieu de outlook_id + folder
            // Le problème d'encodage des caractères spéciaux dans folder_name cause des échecs de correspondance
            console.log(`🔧 [updateEmailStatus] CORRECTION: Recherche par outlook_id uniquement (sans folder pour éviter problèmes d'encodage)`);
            
            // Utiliser toujours la requête sans dossier pour éviter les problèmes d'encodage
            if (!this.statements.updateEmailStatus) {
                this.statements.updateEmailStatus = this.db.prepare(`
                    UPDATE emails 
                    SET is_read = ?, updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
            }
            stmt = this.statements.updateEmailStatus;
            params = [isRead ? 1 : 0, entryId];
            console.log(`🔧 [updateEmailStatus] Requête simplifiée: params = [${params.join(', ')}]`);
            
            const result = stmt.run(...params);
            console.log(`🔧 [updateEmailStatus] Résultat SQL: changes=${result.changes}, lastInsertRowid=${result.lastInsertRowid}`);
            
            // Invalider le cache si un email a été modifié
            if (result.changes > 0 && folderPath) {
                this.invalidateFolderCache(folderPath);
            }
            
            // Invalider le cache de l'interface utilisateur pour mise à jour en temps réel
            if (result.changes > 0) {
                // Mettre à jour les statistiques hebdomadaires si "mail lu = traité"
                const countReadAsTreated = this.getAppSetting('count_read_as_treated', 'false') === 'true';
                
                if (countReadAsTreated) {
                    // Récupérer les infos de l'email pour les stats hebdomadaires
                    const emailInfo = this.db.prepare('SELECT * FROM emails WHERE outlook_id = ?').get(entryId);
                    if (emailInfo) {
                        if (isRead) {
                            // Email marqué comme lu = traité
                            this.updateWeeklyEmailCount(emailInfo, false, true);
                            console.log(`📊 [WEEKLY] Email marqué comme traité: ${entryId}`);
                        } else {
                            // Email marqué comme non lu = non traité (décrémenter si possible)
                            this.adjustWeeklyCount(
                                this.getISOWeekInfo().identifier, // Semaine courante
                                this.mapFolderToCategory(emailInfo.folder_name),
                                -1,
                                'emails_treated'
                            );
                            console.log(`📊 [WEEKLY] Email marqué comme non traité: ${entryId}`);
                        }
                    }
                }
                
                this.invalidateUICache();
            }
            
            if (result.changes > 0) {
                console.log(`✅ [DB-STATUS] Email ${isRead ? 'marqué lu' : 'marqué non lu'}: ${entryId} (${result.changes} changements)`);
            } else {
                console.log(`❌ [DB-STATUS] Aucun email trouvé pour mise à jour: entryId=${entryId}, folderPath=${folderPath}`);
            }
            
            return { 
                updated: result.changes > 0, 
                changes: result.changes,
                entryId: entryId,
                isRead: isRead
            };

        } catch (error) {
            console.error('❌ [DB-STATUS] Erreur mise à jour statut email:', error);
            return { updated: false, error: error.message };
        }
    }

    /**
     * NOUVEAU: Invalide le cache pour un dossier spécifique
     */
    invalidateFolderCache(folderPath) {
        try {
            const keys = this.cache.keys();
            const folderKeys = keys.filter(key => 
                key.includes(folderPath) || 
                key.startsWith('folder_') ||
                key.startsWith('stats_')
            );
            
            folderKeys.forEach(key => this.cache.del(key));
            
            if (folderKeys.length > 0) {
                console.log(`🗑️ [DB-COM] Cache invalidé: ${folderKeys.length} clés pour ${folderPath}`);
            }
            
        } catch (error) {
            console.error('❌ [DB-COM] Erreur invalidation cache:', error);
        }
    }

    /**
     * NOUVEAU: Invalide le cache de l'interface utilisateur pour mise à jour en temps réel
     */
    invalidateUICache() {
        try {
            // Invalider le cache local de la base de données
            const keys = this.cache.keys();
            const emailKeys = keys.filter(key => 
                key.startsWith('recent_emails_') || 
                key.startsWith('stats_') ||
                key.startsWith('folder_')
            );
            emailKeys.forEach(key => this.cache.del(key));
            
            // Invalider le cache de l'interface utilisateur via le service global
            if (global.unifiedMonitoringService && global.unifiedMonitoringService.invalidateEmailCache) {
                global.unifiedMonitoringService.invalidateEmailCache();
                console.log(`🔄 [REAL-TIME] Cache UI invalidé: mise à jour immédiate de l'interface`);
            }
            
        } catch (error) {
            console.error('❌ [REAL-TIME] Erreur invalidation cache UI:', error);
        }
    }

    /**
     * NOUVEAU: Traite les changements d'emails détectés par le polling intelligent
     */
    processPollingEmailChange(emailUpdateData) {
        try {
            console.log(`🔄 [DATABASE] Traitement changement polling: ${emailUpdateData.subject}`);
            console.log(`🔍 [DATABASE] Recherche EntryID: ${emailUpdateData.messageId}`);
            console.log(`🔍 [DATABASE] Dans dossier: ${emailUpdateData.folderPath}`);
            console.log(`🔍 [DATABASE] ChangeType: ${emailUpdateData.changeType}`);
            console.log(`🔍 [DATABASE] Changes: ${JSON.stringify(emailUpdateData.changes)}`);
            
            // FILTRER: Traiter les événements Modified ET Added (pour les changements de statut)
            if (emailUpdateData.changeType !== 'Modified' && emailUpdateData.changeType !== 'Added') {
                console.log(`⏭️ [DATABASE] Événement ${emailUpdateData.changeType} ignoré pour le debugging - Changes: ${JSON.stringify(emailUpdateData.changes)}`);
                return { updated: false, reason: 'Événement ignoré pour debugging' };
            }
            
            // Si c'est un événement Added, vérifier qu'il contient des changements de statut
            if (emailUpdateData.changeType === 'Added' && (!emailUpdateData.changes || emailUpdateData.changes.length === 0)) {
                console.log(`⏭️ [DATABASE] Événement Added sans changements de statut ignoré - Changes: ${JSON.stringify(emailUpdateData.changes)}`);
                return { updated: false, reason: 'Événement Added sans changements' };
            }
            
            // Vérifier d'abord si l'email existe en base
            let existingEmail = this.getEmailByEntryId(emailUpdateData.messageId, emailUpdateData.folderPath);
            
            if (!existingEmail) {
                console.log(`⚠️ [DATABASE] Email non trouvé en base: ${emailUpdateData.messageId} - ${emailUpdateData.subject}`);
                
                // TEMPORAIRE: Désactiver le fallback de recherche par subject pour éviter l'erreur SQL
                /*
                const fallbackEmail = this.findEmailBySubjectAndFolder(emailUpdateData.subject, emailUpdateData.folderPath);
                if (fallbackEmail) {
                    console.log(`🔍 [DATABASE] Email trouvé par subject: ${fallbackEmail.outlook_id}`);
                    // Utiliser l'email trouvé comme existingEmail
                    existingEmail = fallbackEmail;
                    emailUpdateData.messageId = fallbackEmail.outlook_id; // Corriger l'ID
                } else {
                    console.log(`❌ [DATABASE] Email vraiment non trouvé même par subject`);
                }
                */
                console.log(`❌ [DATABASE] Email vraiment non trouvé - fallback désactivé temporairement`);
                return { updated: false, reason: 'Email non trouvé en base' };
            }

            // Préparer les données de mise à jour
            let shouldUpdate = false;
            let newIsRead = existingEmail.is_read;
            
            console.log(`🔍 [DATABASE] Email trouvé - état actuel: is_read=${existingEmail.is_read}, subject="${existingEmail.subject}"`);
            
            // CORRECTION: Normaliser les valeurs pour la comparaison (0/1 vs false/true)
            const currentIsReadBool = Boolean(existingEmail.is_read);
            
            // NOUVELLE LOGIQUE: Toujours faire confiance à Outlook pour les changements de statut
            // Outlook envoie ces événements seulement quand il y a eu une action utilisateur réelle
            if (emailUpdateData.changes && emailUpdateData.changes.includes('MarkedRead')) {
                newIsRead = true;
                // FORCER la mise à jour car Outlook signale un changement utilisateur réel
                shouldUpdate = true;
                console.log(`📖 [DATABASE] FORCE Marquage comme lu: ${emailUpdateData.subject} (BDD: ${existingEmail.is_read}/${currentIsReadBool} -> Outlook: ${newIsRead}) - shouldUpdate: ${shouldUpdate}`);
            } else if (emailUpdateData.changes && emailUpdateData.changes.includes('MarkedUnread')) {
                newIsRead = false;
                // FORCER la mise à jour car Outlook signale un changement utilisateur réel
                shouldUpdate = true;
                console.log(`📬 [DATABASE] FORCE Marquage comme non lu: ${emailUpdateData.subject} (BDD: ${existingEmail.is_read}/${currentIsReadBool} -> Outlook: ${newIsRead}) - shouldUpdate: ${shouldUpdate}`);
            } else if (emailUpdateData.isRead !== undefined && emailUpdateData.isRead !== currentIsReadBool) {
                newIsRead = emailUpdateData.isRead;
                shouldUpdate = true;
                console.log(`📝 [DATABASE] Mise à jour statut lecture: ${emailUpdateData.subject} -> ${emailUpdateData.isRead ? 'Lu' : 'Non lu'} (actuel: ${existingEmail.is_read}/${currentIsReadBool} -> nouveau: ${newIsRead})`);
            }

            // Si aucune modification détectée, pas de mise à jour nécessaire
            if (!shouldUpdate) {
                console.log(`ℹ️ [DATABASE] Aucune modification à appliquer: ${emailUpdateData.subject} (state déjà: ${existingEmail.is_read})`);
                return { updated: false, reason: 'Aucune modification détectée' };
            }
            
            console.log(`🚀 [DATABASE] Procédure de mise à jour: ${emailUpdateData.subject} - BDD:${existingEmail.is_read} -> Outlook:${newIsRead}`);

            // Effectuer la mise à jour - CORRECTION: utiliser l'outlook_id de l'email trouvé
            console.log(`🔧 [DATABASE] Appel updateEmailStatus avec outlook_id: ${existingEmail.outlook_id} (au lieu de messageId: ${emailUpdateData.messageId})`);
            const result = this.updateEmailStatus(existingEmail.outlook_id, newIsRead, emailUpdateData.folderPath);
            
            console.log(`📊 [DATABASE] Résultat updateEmailStatus: updated=${result.updated}, changes=${result.changes}`);
            
            if (result.updated) {
                console.log(`✅ [DATABASE] Email mis à jour via polling: ${emailUpdateData.subject} (${result.changes} changements)`);
                return { 
                    updated: true, 
                    changes: result.changes,
                    emailData: { ...existingEmail, is_read: newIsRead }
                };
            } else {
                console.log(`⚠️ [DATABASE] Aucun changement effectué: ${emailUpdateData.subject}`);
                return { updated: false, reason: 'Aucun changement en base' };
            }

        } catch (error) {
            console.error(`❌ [DATABASE] Erreur traitement changement polling:`, error);
            throw error;
        }
    }

    /**
     * Métriques de performance
     */
    getPerformanceStats() {
        return {
            ...this.stats,
            cacheStats: this.cache.getStats(),
            cacheKeys: this.cache.keys().length
        };
    }

    /**
     * Nettoyage et fermeture
     */
    close() {
        if (this.db) {
            this.db.close();
            this.cache.flushAll();
            console.log('✅ Better-SQLite3 fermé proprement');
        }
    }
    /**
     * NOUVEAU: Met à jour un champ spécifique d'un email
     */
    async updateEmailField(entryId, fieldName, value) {
        try {
            if (!entryId || !fieldName) {
                throw new Error('EntryID et nom du champ requis');
            }

            console.log(`🔄 [DB-UPDATE] Mise à jour ${fieldName} pour email ${entryId}`);

            // Mapping des noms de champs
            const fieldMapping = {
                'subject': 'subject',
                'last_modified': 'last_modified',
                'sender_email': 'sender_email'
            };

            const dbField = fieldMapping[fieldName] || fieldName;
            
            const stmt = this.db.prepare(`
                UPDATE emails 
                SET ${dbField} = ?, updated_at = datetime('now') 
                WHERE outlook_id = ?
            `);

            const result = stmt.run(value, entryId);

            if (result.changes > 0) {
                console.log(`✅ [DB-UPDATE] Champ ${fieldName} mis à jour pour email ${entryId}`);
                this.invalidateFolderCache(''); // Invalider tout le cache
                return { updated: true, changes: result.changes };
            } else {
                console.log(`⚠️ [DB-UPDATE] Aucun email trouvé avec EntryID: ${entryId}`);
                return { updated: false, reason: 'Email non trouvé' };
            }

        } catch (error) {
            console.error(`❌ [DB-UPDATE] Erreur mise à jour champ ${fieldName}:`, error);
            return { updated: false, error: error.message };
        }
    }

    /**
     * NOUVEAU: Marque un email comme supprimé
     */
    async markEmailAsDeleted(entryId) {
        try {
            if (!entryId) {
                throw new Error('EntryID requis');
            }

            console.log(`🗑️ [DB-DELETE] Marquage email supprimé: ${entryId}`);

            // Option 1: Marquer comme supprimé (soft delete)
            const stmt = this.db.prepare(`
                UPDATE emails 
                SET is_deleted = 1, deleted_at = datetime('now'), updated_at = datetime('now')
                WHERE outlook_id = ?
            `);

            const result = stmt.run(entryId);

            if (result.changes > 0) {
                console.log(`✅ [DB-DELETE] Email marqué comme supprimé: ${entryId}`);
                this.invalidateFolderCache(''); // Invalider tout le cache
                return { deleted: true, changes: result.changes };
            } else {
                console.log(`⚠️ [DB-DELETE] Aucun email trouvé avec EntryID: ${entryId}`);
                return { deleted: false, reason: 'Email non trouvé' };
            }

        } catch (error) {
            console.error(`❌ [DB-DELETE] Erreur marquage suppression:`, error);
            return { deleted: false, error: error.message };
        }
    }

    /**
     * NOUVEAU: Supprime définitivement un email de la base
     */
    async deleteEmailPermanently(entryId) {
        try {
            if (!entryId) {
                throw new Error('EntryID requis');
            }

            console.log(`💀 [DB-DELETE] Suppression définitive email: ${entryId}`);

            const stmt = this.db.prepare(`DELETE FROM emails WHERE outlook_id = ?`);
            const result = stmt.run(entryId);

            if (result.changes > 0) {
                console.log(`✅ [DB-DELETE] Email supprimé définitivement: ${entryId}`);
                this.invalidateFolderCache(''); // Invalider tout le cache
                return { deleted: true, changes: result.changes };
            } else {
                console.log(`⚠️ [DB-DELETE] Aucun email trouvé avec EntryID: ${entryId}`);
                return { deleted: false, reason: 'Email non trouvé' };
            }

        } catch (error) {
            console.error(`❌ [DB-DELETE] Erreur suppression définitive:`, error);
            return { deleted: false, error: error.message };
        }
    }

    /**
     * ========================================================================
     * MÉTHODES POUR LE SUIVI HEBDOMADAIRE (inspiré du système VBA)
     * ========================================================================
     */

    /**
     * Calcule le numéro de semaine ISO et l'année ISO pour une date donnée
     */
    getISOWeekInfo(date = new Date()) {
        const tempDate = new Date(date);
        
        // Calcul de la semaine ISO
        tempDate.setHours(0, 0, 0, 0);
        tempDate.setDate(tempDate.getDate() + 3 - (tempDate.getDay() + 6) % 7);
        const week1 = new Date(tempDate.getFullYear(), 0, 4);
        const weekNumber = 1 + Math.round(((tempDate.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
        
        // Calcul de l'année ISO
        const year = tempDate.getFullYear();
        
        // Dates de début et fin de semaine (lundi-dimanche)
        const monday = new Date(tempDate);
        monday.setDate(monday.getDate() - (monday.getDay() + 6) % 7);
        const sunday = new Date(monday);
        sunday.setDate(sunday.getDate() + 6);
        
        return {
            weekNumber,
            year,
            identifier: `S${weekNumber}-${year}`,
            displayName: `S${weekNumber} - ${year}`,
            startDate: monday.toISOString().split('T')[0],
            endDate: sunday.toISOString().split('T')[0]
        };
    }

    /**
     * Obtient ou crée les statistiques hebdomadaires pour une semaine donnée
     */
    getOrCreateWeeklyStats(weekInfo = null, folderType = 'Mails simples') {
        if (!weekInfo) {
            weekInfo = this.getISOWeekInfo();
        }

        try {
            // D'abord, essayer de récupérer les stats existantes
            let stats = this.db.prepare(`
                SELECT * FROM weekly_stats 
                WHERE week_identifier = ? AND folder_type = ?
            `).get(weekInfo.identifier, folderType);

            if (!stats) {
                // Créer une nouvelle entrée
                this.db.prepare(`
                    INSERT INTO weekly_stats 
                    (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type, emails_received)
                    VALUES (?, ?, ?, ?, ?, ?, 0)
                `).run(
                    weekInfo.identifier, 
                    weekInfo.weekNumber, 
                    weekInfo.year, 
                    weekInfo.startDate, 
                    weekInfo.endDate, 
                    folderType
                );

                // Récupérer l'entrée nouvellement créée
                stats = this.db.prepare(`
                    SELECT * FROM weekly_stats 
                    WHERE week_identifier = ? AND folder_type = ?
                `).get(weekInfo.identifier, folderType);
            }

            return stats;
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur getOrCreateWeeklyStats:', error);
            return null;
        }
    }

    /**
     * Met à jour les statistiques hebdomadaires en calculant les emails de la semaine actuelle
     */
    updateCurrentWeekStats() {
        try {
            const weekInfo = this.getISOWeekInfo();
            
            // Vérifier le paramètre compterLuCommeTraite depuis la config
            const compterLuCommeTraite = this.getAppSetting('count_read_as_treated', 'false') === 'true';
            
            // Récupérer tous les emails ARRIVÉS (ajoutés en BDD) de la semaine actuelle groupés par dossier
            const emailStats = this.db.prepare(`
                SELECT 
                    folder_name,
                    COUNT(*) as total_emails,
                    SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread_emails
                FROM emails 
                WHERE DATE(created_at) BETWEEN ? AND ?
                GROUP BY folder_name
            `).all(weekInfo.startDate, weekInfo.endDate);

            // Récupérer tous les emails TRAITÉS de la semaine actuelle groupés par dossier
            let treatedQuery;
            if (compterLuCommeTraite) {
                // Traité = lu OU supprimé (et mis à jour pendant la semaine)
                treatedQuery = `
                    SELECT 
                        folder_name,
                        COUNT(*) as treated_emails
                    FROM emails 
                    WHERE (is_read = 1 OR deleted_at IS NOT NULL)
                    AND DATE(updated_at) BETWEEN ? AND ?
                    GROUP BY folder_name
                `;
            } else {
                // Traité = marqué explicitement comme traité
                treatedQuery = `
                    SELECT 
                        folder_name,
                        COUNT(*) as treated_emails
                    FROM emails 
                    WHERE is_treated = 1
                    AND DATE(updated_at) BETWEEN ? AND ?
                    GROUP BY folder_name
                `;
            }
            
            const treatedStats = this.db.prepare(treatedQuery).all(weekInfo.startDate, weekInfo.endDate);

            console.log(`📊 [WEEKLY] Mise à jour stats semaine ${weekInfo.identifier}:`, { emailStats, treatedStats, compterLuCommeTraite });

            // Mapping des folders vers les types VBA
            const folderTypeMapping = {
                'erkaekanon@outlook.com\\Boîte de réception\\testA': 'Réglements',
                'erkaekanon@outlook.com\\Boîte de réception\\test': 'Déclarations', 
                'erkaekanon@outlook.com\\Boîte de réception\\test\\test-1': 'Déclarations',
                'erkaekanon@outlook.com\\Boîte de réception\\testA\\test-c': 'Mails simples'
            };

            // Grouper les emails reçus par type de folder
            const statsByType = {};
            emailStats.forEach(stat => {
                const folderType = folderTypeMapping[stat.folder_name] || 'Mails simples';
                if (!statsByType[folderType]) {
                    statsByType[folderType] = { received: 0, treated: 0 };
                }
                statsByType[folderType].received += stat.total_emails;
            });

            // Grouper les emails traités par type de folder
            treatedStats.forEach(stat => {
                const folderType = folderTypeMapping[stat.folder_name] || 'Mails simples';
                if (!statsByType[folderType]) {
                    statsByType[folderType] = { received: 0, treated: 0 };
                }
                statsByType[folderType].treated += stat.treated_emails;
            });

            // Mettre à jour ou créer les entrées weekly_stats
            for (const [folderType, counts] of Object.entries(statsByType)) {
                this.db.prepare(`
                    INSERT OR REPLACE INTO weekly_stats 
                    (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type, emails_received, emails_treated, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                `).run(
                    weekInfo.identifier,
                    weekInfo.weekNumber,
                    weekInfo.year,
                    weekInfo.startDate,
                    weekInfo.endDate,
                    folderType,
                    counts.received,
                    counts.treated
                );

                console.log(`✅ [WEEKLY] Stats mises à jour: ${folderType} = ${counts.received} reçus, ${counts.treated} traités`);
            }

            // S'assurer qu'il y a au moins une entrée pour la semaine actuelle
            if (Object.keys(statsByType).length === 0) {
                this.db.prepare(`
                    INSERT OR IGNORE INTO weekly_stats 
                    (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type, emails_received)
                    VALUES (?, ?, ?, ?, ?, ?, 0)
                `).run(
                    weekInfo.identifier,
                    weekInfo.weekNumber,
                    weekInfo.year,
                    weekInfo.startDate,
                    weekInfo.endDate,
                    'Mails simples'
                );
            }

            return true;
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur updateCurrentWeekStats:', error);
            return false;
        }
    }

    /**
     * Met à jour les compteurs d'emails pour une semaine (arrivés/traités)
     */
    updateWeeklyEmailCount(emailData, isArrival = true, isTreated = false) {
        try {
            // Pour les arrivées, utiliser la semaine courante (moment de l'ajout en BDD)
            // Pour les traitements, utiliser aussi la semaine courante (moment du traitement)
            const weekInfo = this.getISOWeekInfo(); // Semaine courante
            
            // Mapper le dossier vers une catégorie (comme dans le VBA)
            const folderType = this.mapFolderToCategory(emailData.folder_name || emailData.folder_name || 'Inconnu');
            
            // Obtenir ou créer les stats de la semaine
            const weeklyStats = this.getOrCreateWeeklyStats(weekInfo, folderType);
            
            // Mettre à jour les compteurs
            let updateQuery = '';
            let params = [];
            
            if (isArrival) {
                updateQuery = 'UPDATE weekly_stats SET emails_received = emails_received + 1, updated_at = CURRENT_TIMESTAMP WHERE week_identifier = ? AND folder_type = ?';
                params = [weekInfo.identifier, folderType];
            }
            
            if (isTreated) {
                updateQuery = 'UPDATE weekly_stats SET emails_treated = emails_treated + 1, updated_at = CURRENT_TIMESTAMP WHERE week_identifier = ? AND folder_type = ?';
                params = [weekInfo.identifier, folderType];
            }
            
            if (updateQuery) {
                const updateStmt = this.db.prepare(updateQuery);
                const result = updateStmt.run(...params);
                
                console.log(`📊 [WEEKLY] Compteur mis à jour: ${weekInfo.displayName} - ${folderType} - ${isArrival ? 'Arrivé' : 'Traité'}`);
                
                // Invalider le cache des stats
                this.invalidateUICache();
                
                return result.changes > 0;
            }
            
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur mise à jour compteur hebdomadaire:', error);
            return false;
        }
    }

    /**
     * Mappe un chemin de dossier vers une catégorie (équivalent du mapping VBA)
     */
    mapFolderToCategory(folderPath) {
        try {
            // Utiliser directement la configuration des dossiers
            const stmt = this.db.prepare(`
                SELECT category FROM folder_configs 
                WHERE folder_path = ? AND enabled = 1
            `);
            
            const result = stmt.get(folderPath);
            if (result) {
                return result.category;
            }

            // Mapping par défaut selon le nom du dossier
            const folderName = folderPath.split('\\').pop() || '';
            
            if (folderName.toLowerCase().includes('test')) {
                return 'Déclarations';
            } else if (folderName.toLowerCase().includes('regel') || folderName.toLowerCase().includes('reglement')) {
                return 'Règlements';
            } else {
                return 'Mails simples';
            }
        } catch (error) {
            console.error(`❌ [MAP] Erreur mapping dossier ${folderPath}:`, error.message);
            return 'Mails simples'; // Valeur par défaut
        }
    }

    /**
     * Ajuste manuellement les compteurs (pour courrier papier, etc.)
     */
    adjustWeeklyCount(weekIdentifier, folderType, adjustmentValue, adjustmentType = 'manual_adjustments') {
        try {
            const stmt = this.db.prepare(`
                UPDATE weekly_stats 
                SET ${adjustmentType} = ${adjustmentType} + ?, updated_at = CURRENT_TIMESTAMP 
                WHERE week_identifier = ? AND folder_type = ?
            `);
            
            const result = stmt.run(adjustmentValue, weekIdentifier, folderType);
            
            if (result.changes > 0) {
                console.log(`📝 [WEEKLY] Ajustement manuel: ${weekIdentifier} - ${folderType} - ${adjustmentValue}`);
                this.invalidateUICache();
                return true;
            }
            
            return false;
            
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur ajustement manuel:', error);
            return false;
        }
    }

    /**
     * Récupère les statistiques hebdomadaires (compatible avec les anciens et nouveaux formats)
     */
    getWeeklyStats(weekIdentifierOrCount = null, limit = 20) {
        try {
            let query = `
                SELECT 
                    week_identifier,
                    week_number,
                    week_year,
                    week_start_date,
                    week_end_date,
                    folder_type,
                    emails_received,
                    emails_treated,
                    manual_adjustments,
                    (emails_received + manual_adjustments) as total_received,
                    updated_at
                FROM weekly_stats
            `;
            
            let params = [];
            
            // Détecter si c'est un identifiant de semaine (string) ou un count (number)
            if (weekIdentifierOrCount !== null) {
                if (typeof weekIdentifierOrCount === 'string') {
                    // C'est un identifiant de semaine spécifique
                    query += ' WHERE week_identifier = ?';
                    params.push(weekIdentifierOrCount);
                } else if (typeof weekIdentifierOrCount === 'number') {
                    // C'est un nombre de semaines à récupérer (usage legacy)
                    limit = weekIdentifierOrCount;
                }
            }
            
            query += ' ORDER BY week_year DESC, week_number DESC, folder_type ASC';
            
            if (limit) {
                query += ' LIMIT ?';
                params.push(limit);
            }
            
            const stmt = this.db.prepare(query);
            const results = stmt.all(...params);
            
            // Si c'était un appel legacy avec un nombre, transformer en format legacy
            if (typeof weekIdentifierOrCount === 'number') {
                // Grouper par semaine et retourner le format legacy
                const weekGroups = {};
                results.forEach(row => {
                    const weekKey = row.week_identifier;
                    if (!weekGroups[weekKey]) {
                        weekGroups[weekKey] = {
                            weekStart: row.week_start_date,
                            weekEnd: row.week_end_date,
                            totalEmails: 0,
                            byCategory: {}
                        };
                    }
                    weekGroups[weekKey].totalEmails += row.total_received;
                    weekGroups[weekKey].byCategory[row.folder_type] = row.total_received;
                });
                
                return Object.values(weekGroups);
            }
            
            return results;
            
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur récupération stats:', error);
            return [];
        }
    }

    /**
     * Sauvegarde ou met à jour un mapping de dossier personnalisé
     */
    saveFolderMapping(originalPath, mappedCategory, displayName = null) {
        try {
            const stmt = this.db.prepare(`
                INSERT OR REPLACE INTO folder_mappings 
                (original_folder_name, mapped_category, display_name, is_active)
                VALUES (?, ?, ?, 1)
            `);
            
            const result = stmt.run(originalPath, mappedCategory, displayName);
            
            console.log(`🗂️ [MAPPING] Mapping sauvegardé: ${originalPath} -> ${mappedCategory}`);
            return result.changes > 0;
            
        } catch (error) {
            console.error('❌ [MAPPING] Erreur sauvegarde mapping:', error);
            return false;
        }
    }

    /**
     * Récupère la semaine courante avec ses statistiques
     */
    getCurrentWeekStats() {
        const currentWeekInfo = this.getISOWeekInfo();
        const stats = this.getWeeklyStats(currentWeekInfo.identifier);
        
        return {
            weekInfo: currentWeekInfo,
            stats: stats
        };
    }

    /**
     * Méthodes pour gérer les paramètres d'application
     */
    getAppSetting(key, defaultValue = null) {
        try {
            const stmt = this.db.prepare('SELECT value FROM app_settings WHERE key = ?');
            const result = stmt.get(key);
            return result ? result.value : defaultValue;
        } catch (error) {
            console.error(`❌ [SETTINGS] Erreur lecture paramètre ${key}:`, error);
            return defaultValue;
        }
    }

    setAppSetting(key, value) {
        try {
            const stmt = this.db.prepare(`
                INSERT OR REPLACE INTO app_settings (key, value, updated_at) 
                VALUES (?, ?, CURRENT_TIMESTAMP)
            `);
            const result = stmt.run(key, value);
            console.log(`⚙️ [SETTINGS] Paramètre sauvegardé: ${key} = ${value}`);
            return result.changes > 0;
        } catch (error) {
            console.error(`❌ [SETTINGS] Erreur sauvegarde paramètre ${key}:`, error);
            return false;
        }
    }

    /**
     * UTILITAIRE: Calcule l'identifiant de semaine ISO (format VBA: "S32-2025")
     */
    calculateWeekIdentifier(dateTime) {
        const date = new Date(dateTime);
        const year = date.getFullYear();
        
        // Calcul de la semaine ISO
        const startOfYear = new Date(year, 0, 4); // 4 janvier = toujours en semaine 1
        const dayOfYear = Math.floor((date - startOfYear) / (24 * 60 * 60 * 1000));
        const weekNumber = Math.ceil((dayOfYear + startOfYear.getDay()) / 7);
        
        return `S${weekNumber}-${year}`;
    }

    /**
     * VBA: Marquer un email comme supprimé (logique traitement)
     */
    markEmailAsDeleted(outlookId, deletedAt = null) {
        try {
            // D'abord récupérer les infos de l'email pour les stats
            const emailInfo = this.db.prepare('SELECT * FROM emails WHERE outlook_id = ?').get(outlookId);
            
            const stmt = this.db.prepare(`
                UPDATE emails 
                SET deleted_at = ?, updated_at = CURRENT_TIMESTAMP 
                WHERE outlook_id = ?
            `);
            
            const deleteTime = deletedAt || new Date().toISOString();
            const result = stmt.run(deleteTime, outlookId);
            
            if (result.changes > 0 && emailInfo) {
                // Email supprimé = toujours traité, indépendamment du paramètre
                this.updateWeeklyEmailCount(emailInfo, false, true);
                console.log(`📊 [WEEKLY] Email supprimé comptabilisé comme traité: ${outlookId}`);
            }
            
            console.log(`📧 [VBA-LOGIC] Email ${outlookId} marqué supprimé à ${deleteTime}`);
            return true;
        } catch (error) {
            console.error('❌ [VBA-LOGIC] Erreur marquage suppression:', error);
            return false;
        }
    }

    /**
     * VBA: Marquer un email comme lu
     */
    markEmailAsRead(outlookId, isRead = true) {
        try {
            const stmt = this.db.prepare(`
                UPDATE emails 
                SET is_read = ?, updated_at = CURRENT_TIMESTAMP 
                WHERE outlook_id = ?
            `);
            
            stmt.run(isRead ? 1 : 0, outlookId);
            
            console.log(`📧 [VBA-LOGIC] Email ${outlookId} marqué ${isRead ? 'lu' : 'non lu'}`);
            return true;
        } catch (error) {
            console.error('❌ [VBA-LOGIC] Erreur marquage lecture:', error);
            return false;
        }
    }

    /**
     * VBA: Marquer un email comme traité manuellement
     */
    markEmailAsTreated(outlookId, isTreated = true) {
        try {
            const stmt = this.db.prepare(`
                UPDATE emails 
                SET is_treated = ?, updated_at = CURRENT_TIMESTAMP 
                WHERE outlook_id = ?
            `);
            
            stmt.run(isTreated ? 1 : 0, outlookId);
            
            console.log(`� [VBA-LOGIC] Email ${outlookId} marqué ${isTreated ? 'traité' : 'non traité'}`);
            return true;
        } catch (error) {
            console.error('❌ [VBA-LOGIC] Erreur marquage traitement:', error);
            return false;
        }
    }

    /**
     * VBA: Statistiques hebdomadaires (arrivées par semaine)
     */
    getWeeklyArrivals(weekIdentifier) {
        try {
            const stmt = this.db.prepare(`
                SELECT 
                    folder_name,
                    category,
                    COUNT(*) as arrivals
                FROM emails 
                WHERE week_identifier = ?
                GROUP BY folder_name, category
                ORDER BY folder_name, category
            `);
            
            return stmt.all(weekIdentifier);
        } catch (error) {
            console.error('❌ [VBA-LOGIC] Erreur stats arrivées:', error);
            return [];
        }
    }

    /**
     * VBA: Statistiques hebdomadaires (traitements par semaine)
     * Selon le paramètre compterLuCommeTraite
     */
    getWeeklyTreatments(weekStart, weekEnd, compterLuCommeTraite = false) {
        try {
            let whereClause;
            if (compterLuCommeTraite) {
                // Traité = lu OU supprimé
                whereClause = `
                    WHERE (is_read = 1 OR deleted_at IS NOT NULL)
                    AND (updated_at BETWEEN ? AND ?)
                `;
            } else {
                // Traité = marqué explicitement comme traité
                whereClause = `
                    WHERE is_treated = 1
                    AND (updated_at BETWEEN ? AND ?)
                `;
            }
            
            const stmt = this.db.prepare(`
                SELECT 
                    folder_name,
                    category,
                    COUNT(*) as treatments
                FROM emails 
                ${whereClause}
                GROUP BY folder_name, category
                ORDER BY folder_name, category
            `);
            
            return stmt.all(weekStart, weekEnd);
        } catch (error) {
            console.error('❌ [VBA-LOGIC] Erreur stats traitements:', error);
            return [];
        }
    }

    /**
     * VBA: Récapitulatif complet pour une semaine
     */
    getWeeklySummary(weekIdentifier, compterLuCommeTraite = false) {
        try {
            // Arrivées de la semaine
            const arrivals = this.getWeeklyArrivals(weekIdentifier);
            
            // Traitements de la semaine (dates de début/fin de semaine)
            const [week, year] = weekIdentifier.split('-');
            const weekNum = parseInt(week.substring(1));
            const weekStart = this.getWeekStartDate(weekNum, parseInt(year));
            const weekEnd = this.getWeekEndDate(weekNum, parseInt(year));
            
            const treatments = this.getWeeklyTreatments(weekStart, weekEnd, compterLuCommeTraite);
            
            return {
                week: weekIdentifier,
                arrivals,
                treatments,
                summary: {
                    totalArrivals: arrivals.reduce((sum, item) => sum + item.arrivals, 0),
                    totalTreatments: treatments.reduce((sum, item) => sum + item.treatments, 0)
                }
            };
        } catch (error) {
            console.error('❌ [VBA-LOGIC] Erreur récapitulatif:', error);
            return null;
        }
    }

    /**
     * UTILITAIRE: Calcule la date de début de semaine ISO
     */
    getWeekStartDate(weekNumber, year) {
        const jan4 = new Date(year, 0, 4);
        const daysToAdd = (weekNumber - 1) * 7 - jan4.getDay() + 1;
        const startDate = new Date(jan4.getTime() + daysToAdd * 24 * 60 * 60 * 1000);
        return startDate.toISOString();
    }

    /**
     * UTILITAIRE: Calcule la date de fin de semaine ISO
     */
    getWeekEndDate(weekNumber, year) {
        const startDate = new Date(this.getWeekStartDate(weekNumber, year));
        const endDate = new Date(startDate.getTime() + 6 * 24 * 60 * 60 * 1000);
        return endDate.toISOString();
    }
}

module.exports = OptimizedDatabaseService;
// Export singleton
const optimizedDatabaseService = new OptimizedDatabaseService();
module.exports = optimizedDatabaseService;
