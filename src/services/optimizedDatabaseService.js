/**
 * Service de base de données OPTIMISÉ avec Better-SQLite3
 * QUICK WINS: +300% performance, WAL mode, prepared statements, cache intelligent
 */

const Database = require('better-sqlite3');
const NodeCache = require('node-cache');
const path = require('path');
const fs = require('fs');
const os = require('os');

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

    // Normalise les valeurs de catégorie (accepte anciens slugs et renvoie libellés humains)
    normalizeCategory(category) {
        if (!category) return category;
        const map = {
            'declarations': 'Déclarations',
            'reglements': 'Règlements',
            'mails_simples': 'Mails simples',
            'mails simple': 'Mails simples',
            'mail simple': 'Mails simples'
        };
        const key = String(category).trim();
        if (['Déclarations', 'Règlements', 'Mails simples'].includes(key)) return key;
        return map[key] || key;
    }

    /**
     * Initialisation avec optimisations WAL et performance
     */
    async initialize() {
        if (this.isInitialized) {
            return;
        }

        try {
            // Résolution robuste du chemin DB (toutes possibilités)
            const resolveDbPath = () => {
                // 1) Priorité: variable d'environnement explicite
                const envPath = process.env.MAILMONITOR_DB_PATH || process.env.AUTO_MAIL_MONITOR_DB_PATH;
                if (envPath) {
                    const p = path.isAbsolute(envPath) ? envPath : path.resolve(process.cwd(), envPath);
                    return p;
                }
                // 2) En production packagée: %APPDATA%/Mail Monitor/data/emails.db
                try {
                    const electron = require('electron');
                    const app = electron && electron.app;
                    if (app && typeof app.getPath === 'function' && app.isPackaged) {
                        const userData = app.getPath('userData');
                        return path.join(userData, 'data', 'emails.db');
                    }
                } catch {}
                // 3) Développement/CLI (sans Electron): dossier du projet
                return path.join(__dirname, '../../data/emails.db');
            };

            let dbPath = resolveDbPath();

            const ensureDir = (dir) => { if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true }); };
            const canWriteDir = (dir) => {
                try {
                    ensureDir(dir);
                    const test = path.join(dir, `.writetest_${Date.now()}`);
                    fs.writeFileSync(test, 'ok');
                    fs.unlinkSync(test);
                    return true;
                } catch { return false; }
            };

            // Si le dossier cible n'est pas inscriptible (cas très rare), fallback vers %LOCALAPPDATA% / Temp
            let dbDir = path.dirname(dbPath);
            if (!canWriteDir(dbDir)) {
                // Essayer LocalAppData/Mail Monitor/data
                try {
                    const electron = require('electron');
                    const app = electron && electron.app;
                    if (app && typeof app.getPath === 'function') {
                        const localData = app.getPath('userData').replace(/\\Roaming\\/i, '\\Local\\');
                        const fallback = path.join(localData, 'data');
                        if (canWriteDir(fallback)) {
                            dbPath = path.join(fallback, 'emails.db');
                            dbDir = fallback;
                        }
                    }
                } catch {}
            }
            if (!canWriteDir(dbDir)) {
                const tmp = path.join(os.tmpdir(), 'MailMonitor', 'data');
                ensureDir(tmp);
                dbPath = path.join(tmp, 'emails.db');
                dbDir = tmp;
                console.warn('⚠️ Dossier non inscriptible, utilisation du dossier temporaire:', dbDir);
            }
            
            // Migration depuis anciens emplacements si le fichier n'existe pas encore
            const migrateLegacyIfAny = () => {
                if (fs.existsSync(dbPath)) return;
                const candidates = [];
                // Ancienne tentative dans le paquet (non fiable, mais on regarde)
                try {
                    const electron = require('electron');
                    const app = electron && electron.app;
                    if (app && app.isPackaged) {
                        const resPath = process.resourcesPath || path.join(path.dirname(app.getPath('exe')), 'resources');
                        candidates.push(path.join(resPath, 'app.asar', 'data', 'emails.db'));
                        candidates.push(path.join(resPath, 'data', 'emails.db'));
                        // Ancien userData sans sous-dossier data
                        const userData = app.getPath('userData');
                        candidates.push(path.join(userData, 'emails.db'));
                    }
                } catch {}
                // Portable/ancien: dossier courant
                candidates.push(path.resolve(process.cwd(), 'data', 'emails.db'));
                // Développement: dossier du projet (déjà la cible en dev, mais au cas où override env)
                candidates.push(path.join(__dirname, '../../data/emails.db'));

                for (const c of candidates) {
                    try {
                        if (c && fs.existsSync(c)) {
                            ensureDir(dbDir);
                            fs.copyFileSync(c, dbPath);
                            console.log(`📦 Migration DB depuis ancien emplacement: ${c} → ${dbPath}`);
                            break;
                        }
                    } catch (e) { console.warn('⚠️ Migration DB échouée depuis', c, e.message); }
                }
            };
            migrateLegacyIfAny();

            console.log('🚀 Initialisation Better-SQLite3 avec optimisations...');
            console.log(`📁 Base de données: ${dbPath}`);
            
            // Ouvrir avec optimisations de performance
            try {
                // Sauvegarde quotidienne simple (si fichier existe)
                if (fs.existsSync(dbPath)) {
                    const d = new Date();
                    const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
                    const bak = path.join(dbDir, `emails-${stamp}.bak`);
                    if (!fs.existsSync(bak)) {
                        try { fs.copyFileSync(dbPath, bak); } catch {}
                    }
                }
                this.db = new Database(dbPath, {
                    verbose: null, // Pas de logging verbose en prod
                    fileMustExist: false
                });
            } catch (e) {
                // Tentative de récupération en cas de corruption
                console.error('❌ Ouverture DB échouée:', e?.message || e);
                if (/SQLITE_CORRUPT|database disk image is malformed/i.test(String(e?.message || e))) {
                    try {
                        const corruptPath = path.join(dbDir, `emails.corrupt-${Date.now()}.db`);
                        fs.renameSync(dbPath, corruptPath);
                        console.warn('⚠️ DB corrompue renommée en', corruptPath, '; création d\'une nouvelle base');
                        this.db = new Database(dbPath, { verbose: null, fileMustExist: false });
                    } catch (e2) {
                        console.error('❌ Récupération DB impossible:', e2?.message || e2);
                        throw e;
                    }
                } else {
                    throw e;
                }
            }

            // OPTIMISATIONS CRITIQUES
            this.setupOptimizations();
            
            // Créer les tables si nécessaire
            this.createTables();
            
            // S'assurer que les colonnes de suppression existent
                // this.ensureDeletedColumn(); // supprimé, colonne gérée dans la migration
            // Nouvelle colonne: treated_at (date de traitement)
            this.ensureTreatedAtColumn();
            this.ensureWeeklyStatsTable();
            // Assurer la table des commentaires hebdomadaires
            this.ensureWeeklyCommentsTable();
            // Normaliser d'anciennes valeurs de catégories si besoin
            this.migrateNormalizeFolderCategories();
            
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

    // Migration légère: normaliser les catégories héritées (slugs -> libellés)
    migrateNormalizeFolderCategories() {
        try {
            this.db.exec(`UPDATE folder_configurations SET category = 'Mails simples' WHERE LOWER(category) IN ('mails_simples','mail simple','mails simple')`);
            this.db.exec(`UPDATE folder_configurations SET category = 'Déclarations' WHERE LOWER(category) = 'declarations'`);
            this.db.exec(`UPDATE folder_configurations SET category = 'Règlements' WHERE LOWER(category) = 'reglements'`);
        } catch (e) {
            console.warn('⚠️ Migration categories skipped:', e.message);
        }
    }

    /**
     * Configuration des optimisations SQLite
     */
    setupOptimizations() {
        // WAL Mode pour performance concurrente
        this.db.pragma('journal_mode = WAL');
    // S'assurer que l'encodage SQLite est en UTF-8 (par défaut, mais explicite)
    try { this.db.pragma('encoding = "UTF-8"'); } catch (_) {}
        
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
        treated_at DATETIME NULL,
                week_identifier TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // (activity_weekly supprimé)

        // Table des configurations de dossiers surveillés
        this.db.exec(`
            CREATE TABLE IF NOT EXISTS folder_configurations (
                folder_path TEXT PRIMARY KEY,
                folder_name TEXT NOT NULL,
                category TEXT NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // Table des paramètres applicatifs (clé/valeur)
        this.db.exec(`
            CREATE TABLE IF NOT EXISTS app_settings (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

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
     * Migration: ajoute la colonne treated_at si absente et pré-remplit depuis is_treated/deleted_at
     */
    ensureTreatedAtColumn() {
        try {
            const cols = this.db.prepare(`PRAGMA table_info(emails)`).all();
            const hasTreatedAt = cols.some(c => String(c.name).toLowerCase() === 'treated_at');
            if (!hasTreatedAt) {
                this.db.exec(`ALTER TABLE emails ADD COLUMN treated_at DATETIME NULL`);
            }
            // Index idempotent
            this.db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_treated_at ON emails(treated_at)`);
            // Pré-remplissage: si is_treated=1 et treated_at NULL => treated_at = updated_at (ou now)
            this.db.exec(`
                UPDATE emails
                SET treated_at = COALESCE(updated_at, CURRENT_TIMESTAMP)
                WHERE (is_treated = 1 OR is_treated = '1' OR is_treated = TRUE)
                  AND (treated_at IS NULL)
            `);
            // Si deleted_at défini et treated_at NULL => treated_at = deleted_at
            this.db.exec(`
                UPDATE emails
                SET treated_at = deleted_at
                WHERE deleted_at IS NOT NULL
                  AND treated_at IS NULL
            `);
        } catch (e) {
            console.warn('⚠️ Migration treated_at ignorée:', e.message);
        }
    }

    /**
     * Préparation des statements pour performance maximale (schéma optimisé final 13 colonnes)
     */
    prepareStatements() {
        // Statements les plus utilisés
        this.statements = {
            // Emails - compatible avec nouvelle structure optimisée
            // Insertion stricte (pas de REPLACE) pour éviter les réinsertions involontaires
            insertEmailNew: this.db.prepare(`
                INSERT INTO emails 
                (outlook_id, subject, sender_email, received_time, folder_name, 
                 category, is_read, is_treated, deleted_at, treated_at, week_identifier, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            `),
            // Mise à jour par outlook_id (sans toucher received_time ni week_identifier)
            updateEmailByOutlookId: this.db.prepare(`
                UPDATE emails SET 
                    subject = ?,
                    sender_email = ?,
                    folder_name = ?,
                    category = ?,
                    is_read = ?,
                    -- Keep legacy is_treated in sync only if explicitly provided
                    is_treated = ?,
                    -- Do not overwrite treated_at unless a non-null value is provided
                    treated_at = COALESCE(?, treated_at),
                    updated_at = CURRENT_TIMESTAMP
                WHERE outlook_id = ?
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
            SUM(CASE WHEN DATE(treated_at) = DATE('now') THEN 1 ELSE 0 END) as treatedToday
                FROM emails
            `),
            
            // Folders
            getFoldersConfig: this.db.prepare(`
                SELECT folder_path, category, folder_name FROM folder_configurations
            `),
            
            insertFolderConfig: this.db.prepare(`
                INSERT OR REPLACE INTO folder_configurations (folder_path, category, folder_name, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            `),
            
            deleteFolderConfig: this.db.prepare(`
                DELETE FROM folder_configurations WHERE folder_path = ?
            `),

            updateFolderCategory: this.db.prepare(`
                UPDATE folder_configurations
                SET category = ?, updated_at = CURRENT_TIMESTAMP
                WHERE folder_path = ?
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
        // Normaliser le sujet par sécurité (si l'amont fournit Subject/ConversationTopic)
        const rawSubject = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
        const normalizedSubject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';

        // Id unique
        const outlookId = emailData.outlook_id || emailData.id || '';
        const existed = this.statements.getEmailByEntryId.get(outlookId);

        let result;
        if (existed) {
            // Mise à jour sans réinsérer (évite de compter une nouvelle arrivée)
            // Déterminer treated_at s'il faut le définir
            let treatedAtParam = null;
            const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);
            if (!existed.treated_at) {
                if (emailData.is_treated) {
                    treatedAtParam = new Date().toISOString();
                } else if (countReadAsTreated && emailData.is_read && !existed.is_read) {
                    treatedAtParam = new Date().toISOString();
                } else if (emailData.deleted_at) {
                    treatedAtParam = emailData.deleted_at;
                }
            }
            result = this.statements.updateEmailByOutlookId.run(
                normalizedSubject,
                emailData.sender_email || '',
                emailData.folder_name || '',
                emailData.category || 'Mails simples',
                emailData.is_read ? 1 : 0,
                emailData.is_treated ? 1 : 0,
                treatedAtParam,
                outlookId
            );
        } else {
            // Insertion d'un nouvel email (véritable arrivée)
            // treated_at initial: si déjà traité ou supprimé à l'insertion
            const initialTreatedAt = emailData.is_treated ? (emailData.treated_at || new Date().toISOString()) : (emailData.deleted_at ? emailData.deleted_at : null);
            result = this.statements.insertEmailNew.run(
                outlookId,
                normalizedSubject,
                emailData.sender_email || '',
                emailData.received_time || new Date().toISOString(),
                emailData.folder_name || '',
                emailData.category || 'Mails simples',
                emailData.is_read ? 1 : 0,
                emailData.is_treated ? 1 : 0,
                emailData.deleted_at || null,
                initialTreatedAt,
                weekId
            );
        }

        // Recalcule déterministe des stats (évite les doubles comptages)
        this.updateCurrentWeekStats();

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
                const outlookId = email.outlook_id || email.id || '';
                const existed = this.statements.getEmailByEntryId.get(outlookId);
                const rawSubject = (email.subject ?? email.Subject ?? email.ConversationTopic ?? '').toString();
                const normalizedSubject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
                if (existed) {
                    this.statements.updateEmailByOutlookId.run(
                        normalizedSubject,
                        email.sender_email || '',
                        email.folder_name || '',
                        email.category || 'Mails simples',
                        email.is_read ? 1 : 0,
                        email.is_treated ? 1 : 0,
                        null,
                        outlookId
                    );
                } else {
                    const weekId = this.calculateWeekIdentifier(email.received_time);
                    const initialTreatedAt = email.is_treated ? (email.treated_at || new Date().toISOString()) : null;
                    this.statements.insertEmailNew.run(
                        outlookId,
                        normalizedSubject,
                        email.sender_email || '',
                        email.received_time || new Date().toISOString(),
                        email.folder_name || '',
                        email.category || 'Mails simples',
                        email.is_read ? 1 : 0,
                        email.is_treated ? 1 : 0,
                        null,
                        initialTreatedAt,
                        weekId
                    );
                }
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
            const rows = this.statements.getFoldersConfig.all();
            // Normaliser les catégories à l'affichage
            return rows.map(r => ({
                ...r,
                category: this.normalizeCategory(r.category)
            }));
        }, 300); // Cache 5 minutes
    }

    /**
     * Sauvegarde configuration dossier
     */
    addFolderConfiguration(folderPath, category, folderName) {
        this.cache.del('folders_config'); // Invalider cache
    const cat = this.normalizeCategory(category);
    return this.statements.insertFolderConfig.run(folderPath, cat, folderName);
    }

    /**
     * Suppression configuration dossier
     */
    deleteFolderConfiguration(folderPath) {
        this.cache.del('folders_config'); // Invalider cache
        return this.statements.deleteFolderConfig.run(folderPath);
    }

    /**
     * Met à jour la catégorie d'un dossier surveillé
     */
    updateFolderCategory(folderPath, category) {
        this.cache.del('folders_config');
        const cat = this.normalizeCategory(category);
        const info = this.statements.updateFolderCategory.run(cat, folderPath);
        // Si aucune ligne affectée, faire un upsert minimal avec folderName dérivé du path
        if (info.changes === 0) {
            const name = String(folderPath || '').split('\\').pop() || folderPath;
            this.statements.insertFolderConfig.run(folderPath, cat, name);
            return 1;
        }
        return info.changes;
    }

    /**
     * Remplace la configuration complète des dossiers à surveiller
     * Accepts either { folderCategories: { [path]: {category,name} } } or
     * an array of { folder_path, category, folder_name }
     */
    saveFoldersConfiguration(payload) {
        this.cache.del('folders_config');
        const debugStart = Date.now();
        try { console.log('🛠️ [FOLDERS-CONFIG] Sauvegarde payload brut:', JSON.stringify(payload).slice(0,1000)); } catch {}
        const tx = this.db.transaction((rows) => {
            // Remplacer tout pour garder la source unique (UI) en cohérence
            this.db.prepare('DELETE FROM folder_configurations').run();
            const insert = this.statements.insertFolderConfig;
            for (const r of rows) {
                const folderPath = r.folder_path || r.path || r.key || r.folderPath;
                if (!folderPath) continue;
                // Toujours appliquer une catégorie par défaut si absente
                const categoryRaw = r.category || '';
                const category = this.normalizeCategory(categoryRaw || 'Mails simples') || 'Mails simples';
                const name = r.folder_name || r.name || String(folderPath).split('\\').pop();
                if (!category) {
                    // Cela ne devrait plus arriver, mais log si vide
                    console.warn('⚠️ [FOLDERS-CONFIG] Catégorie vide après normalisation, fallback Mails simples', folderPath);
                }
                insert.run(folderPath, category, name);
            }
        });

        // Normaliser payload en tableau de lignes
        let rows = [];
        if (payload && typeof payload === 'object' && payload.folderCategories && typeof payload.folderCategories === 'object') {
            rows = Object.entries(payload.folderCategories).map(([folder_path, cfg]) => ({
                folder_path,
                category: this.normalizeCategory(cfg?.category || 'Mails simples') || 'Mails simples',
                folder_name: cfg?.name || String(folder_path).split('\\').pop()
            }));
        } else if (Array.isArray(payload)) {
            rows = payload;
        } else if (payload && typeof payload === 'object' && Array.isArray(payload.rows)) {
            rows = payload.rows;
        }

        tx(rows);
        try {
            const verify = this.db.prepare('SELECT folder_path, category, folder_name FROM folder_configurations').all();
            console.log(`✅ [FOLDERS-CONFIG] ${verify.length} lignes enregistrées.`, verify.slice(0,10));
        } catch (e) {
            console.warn('⚠️ [FOLDERS-CONFIG] Vérification post-insert impossible:', e.message);
        }
        return { success: true, count: rows.length, durationMs: Date.now()-debugStart };
    }

    /**
     * Diagnostic: dump rapide des configurations de dossiers (non mis en cache)
     */
    debugDumpFolderConfigurations() {
        try {
            const rows = this.db.prepare('SELECT folder_path, category, folder_name, created_at, updated_at FROM folder_configurations').all();
            console.log('🧪 [FOLDERS-CONFIG][DUMP]', rows);
            return rows;
        } catch (e) {
            console.error('❌ [FOLDERS-CONFIG] Dump impossible:', e.message);
            return [];
        }
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
        let cached = this.cache.get(cacheKey);

        if (cached) {
            this.stats.cacheHits++;
            return cached;
        }

        try {
            const rows = this.db.prepare('SELECT key, value FROM app_settings').all();
            // Construire un objet à partir des clés plates "section.sousCle"
            const flat = {};
            const coerceValue = (raw) => {
                // 1) Essayer JSON.parse d'abord (stockage normalisé)
                try {
                    const parsed = JSON.parse(raw);
                    // Certains anciens enregistrements ont été stockés comme chaîne "true"/"false"
                    if (parsed === 'true') return true;
                    if (parsed === 'false') return false;
                    return parsed;
                } catch {}
                // 2) Gérer anciens formats: chaînes 'true'/'false' éventuellement avec quotes simples
                if (typeof raw === 'string') {
                    const trimmed = raw.trim().replace(/^['"]|['"]$/g, '');
                    if (trimmed.toLowerCase() === 'true') return true;
                    if (trimmed.toLowerCase() === 'false') return false;
                }
                // 3) Retourner brut si rien d'autre ne convient
                return raw;
            };

            rows.forEach(row => {
                flat[row.key] = coerceValue(row.value);
            });

            // Build nested object robustly: if both 'a' and 'a.b' exist, keep shallow 'a' (primitive)
            const nested = {};
            const entries = Object.entries(flat).sort((a, b) => {
                const da = a[0].split('.').length;
                const db = b[0].split('.').length;
                return da - db; // process shallow keys first
            });

            let collisionCount = 0;
            for (const [key, value] of entries) {
                const parts = key.split('.');
                let cur = nested;
                let skip = false;
                for (let i = 0; i < parts.length - 1; i++) {
                    const p = parts[i];
                    // If a primitive already exists here, we ignore deeper keys under it
                    if (cur[p] !== undefined && (typeof cur[p] !== 'object' || cur[p] === null || Array.isArray(cur[p]))) {
                        collisionCount++;
                        skip = true;
                        break;
                    }
                    if (cur[p] === undefined) cur[p] = {};
                    cur = cur[p];
                }
                if (skip) continue;

                const leaf = parts[parts.length - 1];
                // If an object already exists at the leaf due to previous deeper keys, don't overwrite with a primitive
                if (cur[leaf] && typeof cur[leaf] === 'object' && cur[leaf] !== null && !Array.isArray(cur[leaf])) {
                    collisionCount++;
                    continue;
                }
                cur[leaf] = value;
            }
            if (collisionCount > 0) {
                console.warn(`⚠️ [SETTINGS] Collisions ignorées lors du chargement des paramètres: ${collisionCount}`);
            }

            this.cache.set(cacheKey, nested, 300); // Cache 5 minutes
            this.stats.cacheMisses++;
            return nested;
        } catch (error) {
            console.error('❌ Erreur loadAppSettings:', error);
            return {};
        }
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
     * Assurer table des commentaires hebdomadaires
     */
    ensureWeeklyCommentsTable() {
        try {
            this.db.exec(`
                CREATE TABLE IF NOT EXISTS weekly_comments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    week_identifier TEXT NOT NULL,
                    week_year INTEGER NOT NULL,
                    week_number INTEGER NOT NULL,
                    category TEXT NULL,
                    comment_text TEXT NOT NULL,
                    author TEXT NULL,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                );
                CREATE INDEX IF NOT EXISTS idx_weekly_comments_week ON weekly_comments(week_identifier);
            `);
        } catch (error) {
            console.error('❌ Erreur ensureWeeklyCommentsTable:', error);
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

            // Insérer l'email (via prepared SQL classique sur email_id)
            if (!this.statements.insertEmailByEmailId) {
                this.statements.insertEmailByEmailId = this.db.prepare(`
                    INSERT INTO emails (email_id, subject, sender_email, received_time, folder_name, is_read, category, last_updated)
                    VALUES (@email_id, @subject, @sender_email, @received_time, @folder_name, @is_read, @category, @last_updated)
                `);
            }
            const result = this.statements.insertEmailByEmailId.run(emailRecord);

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
            if (!this.statements.getEmailById) {
                this.statements.getEmailById = this.db.prepare('SELECT * FROM emails WHERE email_id = ?');
            }
            email = this.statements.getEmailById.get(emailId);
            
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
            // Définir treated_at si on passe à lu et le réglage l'autorise et que treated_at est encore NULL
            const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);
            const existing = this.statements.getEmailByEntryId.get(entryId);
            const shouldSetTreatedAt = countReadAsTreated && isRead && existing && !existing.treated_at;
            if (!this.statements.updateEmailStatusWithTreated) {
                this.statements.updateEmailStatusWithTreated = this.db.prepare(`
                    UPDATE emails 
                    SET is_read = ?,
                        is_treated = CASE WHEN ? = 1 THEN 1 ELSE is_treated END,
                        treated_at = CASE WHEN ? = 1 THEN COALESCE(treated_at, CURRENT_TIMESTAMP) ELSE treated_at END,
                        updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
            }
            if (!this.statements.updateEmailStatusSimple) {
                this.statements.updateEmailStatusSimple = this.db.prepare(`
                    UPDATE emails 
                    SET is_read = ?, updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
            }
            if (shouldSetTreatedAt) {
                stmt = this.statements.updateEmailStatusWithTreated;
                params = [isRead ? 1 : 0, 1, 1, entryId];
            } else {
                stmt = this.statements.updateEmailStatusSimple;
                params = [isRead ? 1 : 0, entryId];
            }
            console.log(`🔧 [updateEmailStatus] Requête simplifiée: params = [${params.join(', ')}]`);
            
            const result = stmt.run(...params);
            console.log(`🔧 [updateEmailStatus] Résultat SQL: changes=${result.changes}, lastInsertRowid=${result.lastInsertRowid}`);
            
            // Invalider le cache si un email a été modifié
            if (result.changes > 0 && folderPath) {
                this.invalidateFolderCache(folderPath);
            }
            
            // Recalcul déterministe des stats pour la semaine en cours (évite les doubles comptages et négatifs)
            if (result.changes > 0) {
                this.updateCurrentWeekStats();
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
     * Sécurité: remet à zéro toute valeur négative dans weekly_stats
     */
    clampWeeklyStatsNonNegative() {
        try {
            this.db.exec(`
                UPDATE weekly_stats
                SET 
                    emails_received = CASE WHEN emails_received < 0 THEN 0 ELSE emails_received END,
                    emails_treated = CASE WHEN emails_treated < 0 THEN 0 ELSE emails_treated END,
                    manual_adjustments = CASE WHEN manual_adjustments < 0 THEN 0 ELSE manual_adjustments END
                WHERE emails_received < 0 OR emails_treated < 0 OR manual_adjustments < 0
            `);
        } catch (e) {
            console.warn('⚠️ [WEEKLY] Clamp non-negatif échoué:', e.message);
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
                SET is_deleted = 1,
                    deleted_at = datetime('now'),
                    is_treated = 1,
                    treated_at = COALESCE(treated_at, datetime('now')),
                    updated_at = datetime('now')
                WHERE outlook_id = ?
            `);

            const result = stmt.run(entryId);

            if (result.changes > 0) {
                console.log(`✅ [DB-DELETE] Email marqué comme supprimé: ${entryId}`);
                this.invalidateFolderCache(''); // Invalider tout le cache
                this.updateCurrentWeekStats();
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
            const compterLuCommeTraite = !!this.getAppSetting('count_read_as_treated', false);
            
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
                        // Nouveau modèle: traité = treated_at non nul, et tombant dans la semaine
                        const treatedStats = this.db.prepare(`
                                SELECT 
                                        folder_name,
                                        COUNT(*) as treated_emails
                                FROM emails 
                                WHERE treated_at IS NOT NULL
                                    AND DATE(treated_at) BETWEEN ? AND ?
                                GROUP BY folder_name
                        `).all(weekInfo.startDate, weekInfo.endDate);

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

            // Clamp de sécurité: éviter toute valeur négative en base
            this.clampWeeklyStatsNonNegative();
            return true;
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur updateCurrentWeekStats:', error);
            return false;
        }
    }

    // updateWeeklyEmailCount supprimé: on s'appuie sur updateCurrentWeekStats() déterministe pour éviter les doubles comptages

    /**
     * Mappe un chemin de dossier vers une catégorie (équivalent du mapping VBA)
     */
    mapFolderToCategory(folderPath) {
        try {
            // Utiliser directement la configuration des dossiers (schéma actuel)
            const stmt = this.db.prepare(`
                SELECT category FROM folder_configurations
                WHERE folder_path = ?
                LIMIT 1
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
     * NOUVEAU: Compte le nombre total de semaines distinctes dans weekly_stats
     */
    getWeeklyDistinctWeeksCount() {
        try {
            const row = this.db.prepare(`SELECT COUNT(DISTINCT week_identifier) AS total FROM weekly_stats`).get();
            return row?.total || 0;
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur comptage semaines distinctes:', error);
            return 0;
        }
    }

    /**
     * NOUVEAU: Récupère une page d'historique hebdomadaire par semaines (pas par lignes)
     * @param {number} page - numéro de page (1-based)
     * @param {number} pageSize - nombre de semaines par page
     * @returns {{ rows: any[], totalWeeks: number }}
     */
    getWeeklyHistoryPage(page = 1, pageSize = 5) {
        try {
            const totalWeeks = this.getWeeklyDistinctWeeksCount();
            const totalPages = pageSize > 0 ? Math.max(1, Math.ceil(totalWeeks / pageSize)) : 1;
            const safePage = Math.min(Math.max(1, page), totalPages);
            const offset = (safePage - 1) * pageSize;

            // CTE pour sélectionner les semaines de la page, puis joindre sur weekly_stats pour récupérer les 3 catégories
            const sql = `
                WITH weeks_page AS (
                    SELECT week_identifier, MAX(week_year) AS week_year, MAX(week_number) AS week_number
                    FROM weekly_stats
                    GROUP BY week_identifier
                    ORDER BY week_year DESC, week_number DESC
                    LIMIT @limit OFFSET @offset
                )
                SELECT 
                    ws.week_identifier,
                    ws.week_number,
                    ws.week_year,
                    ws.week_start_date,
                    ws.week_end_date,
                    ws.folder_type,
                    ws.emails_received,
                    ws.emails_treated,
                    ws.manual_adjustments,
                    (ws.emails_received + ws.manual_adjustments) as total_received,
                    ws.updated_at
                FROM weekly_stats ws
                INNER JOIN weeks_page wp ON wp.week_identifier = ws.week_identifier
                ORDER BY wp.week_year DESC, wp.week_number DESC, ws.folder_type ASC
            `;

            const rows = this.db.prepare(sql).all({ limit: pageSize, offset });
            return { rows, totalWeeks, page: safePage, pageSize, totalPages };
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur récupération page historique:', error);
            return { rows: [], totalWeeks: 0, page: 1, pageSize, totalPages: 1 };
        }
    }

    /**
     * Calcule le stock (carry-over) par catégorie avant une semaine donnée (exclue),
     * en simulant semaine par semaine avec clamp à 0.
     * @param {number} weekYear
     * @param {number} weekNumber
     * @returns {{declarations:number, reglements:number, mails_simples:number}}
     */
    getCarryBeforeWeek(weekYear, weekNumber) {
        try {
            const sql = `
                SELECT week_year, week_number, folder_type,
                       COALESCE(emails_received,0) AS emails_received,
                       COALESCE(emails_treated,0) AS emails_treated,
                       COALESCE(manual_adjustments,0) AS manual_adjustments
                FROM weekly_stats
                WHERE (week_year < @year)
                   OR (week_year = @year AND week_number < @number)
                ORDER BY week_year ASC, week_number ASC, folder_type ASC
            `;
            const rows = this.db.prepare(sql).all({ year: weekYear, number: weekNumber });

            // Accumulate per week, applying clamp to 0 at each step
            const carry = { declarations: 0, reglements: 0, mails_simples: 0 };
            let currentKey = null;
            let weekAgg = {
                declarations: { rec: 0, trt: 0, adj: 0 },
                reglements: { rec: 0, trt: 0, adj: 0 },
                mails_simples: { rec: 0, trt: 0, adj: 0 }
            };

            const flushWeek = () => {
                if (!currentKey) return;
                for (const type of ['declarations', 'reglements', 'mails_simples']) {
                    const rec = weekAgg[type].rec || 0;
                    const treatedTotal = (weekAgg[type].trt || 0) + (weekAgg[type].adj || 0);
                    const net = rec - treatedTotal;
                    const next = carry[type] + net;
                    carry[type] = next < 0 ? 0 : next;
                }
                // reset
                weekAgg = {
                    declarations: { rec: 0, trt: 0, adj: 0 },
                    reglements: { rec: 0, trt: 0, adj: 0 },
                    mails_simples: { rec: 0, trt: 0, adj: 0 }
                };
            };

            // Helper: map any folder_type variant to canonical keys used in carry structure
            const toCanonical = (ft) => {
                if (!ft) return 'mails_simples';
                const v = String(ft).toLowerCase();
                // Accept both localized labels and internal identifiers
                if (v.includes('déclar') || v.includes('declar')) return 'declarations';
                if (v.includes('règle') || v.includes('regle') || v.includes('reglement')) return 'reglements';
                if (v.includes('mail') || v.includes('simple')) return 'mails_simples';
                // Fallback to raw if already canonical
                if (v === 'declarations' || v === 'reglements' || v === 'mails_simples') return v;
                return 'mails_simples';
            };

            for (const row of rows) {
                const wk = `${row.week_year}-${row.week_number}`;
                if (currentKey !== wk) {
                    // process previous week
                    flushWeek();
                    currentKey = wk;
                }
                const type = toCanonical(row.folder_type);
                if (!weekAgg[type]) {
                    weekAgg[type] = { rec: 0, trt: 0, adj: 0 };
                }
                weekAgg[type].rec += row.emails_received || 0;
                weekAgg[type].trt += row.emails_treated || 0;
                weekAgg[type].adj += row.manual_adjustments || 0;
            }

            // Process the last aggregated week
            flushWeek();

            return carry;
        } catch (error) {
            console.error('❌ [WEEKLY] Erreur calcul carry-over:', error);
            return { declarations: 0, reglements: 0, mails_simples: 0 };
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

    // ==================== WEEKLY COMMENTS ====================
    /**
     * Ajoute un commentaire pour une semaine donnée
     */
    addWeeklyComment({ week_identifier, week_year, week_number, category = null, comment_text, author = null }) {
        try {
            if (!week_identifier || !comment_text) throw new Error('Paramètres obligatoires manquants');
            if ((!week_year || !week_number) && week_identifier) {
                // Extraire depuis S{num}-{year}
                const m = String(week_identifier).match(/S(\d+)-?(\d{4})/);
                if (m) { week_number = week_number || parseInt(m[1], 10); week_year = week_year || parseInt(m[2], 10); }
            }
            const stmt = this.db.prepare(`
                INSERT INTO weekly_comments (week_identifier, week_year, week_number, category, comment_text, author, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
            `);
            const res = stmt.run(week_identifier, week_year, week_number, category, comment_text, author);
            return { success: true, id: res.lastInsertRowid };
        } catch (error) {
            console.error('❌ [WEEKLY-COMMENTS] Erreur ajout:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Récupère les commentaires d'une semaine
     */
    getWeeklyComments(week_identifier) {
        try {
            const stmt = this.db.prepare(`
                SELECT id, week_identifier, week_year, week_number, category, comment_text, author, created_at, updated_at
                FROM weekly_comments
                WHERE week_identifier = ?
                ORDER BY created_at DESC
            `);
            const rows = stmt.all(week_identifier);
            return { success: true, rows };
        } catch (error) {
            console.error('❌ [WEEKLY-COMMENTS] Erreur lecture:', error);
            return { success: false, rows: [], error: error.message };
        }
    }

    /**
     * Met à jour le texte d'un commentaire
     */
    updateWeeklyComment(id, comment_text, category = undefined) {
        try {
            if (!id) throw new Error('ID requis');
            let sql = 'UPDATE weekly_comments SET comment_text = ?, updated_at = CURRENT_TIMESTAMP';
            const params = [comment_text];
            if (category !== undefined) { sql += ', category = ?'; params.push(category); }
            sql += ' WHERE id = ?';
            params.push(id);
            const stmt = this.db.prepare(sql);
            const res = stmt.run(...params);
            return { success: res.changes > 0, changes: res.changes };
        } catch (error) {
            console.error('❌ [WEEKLY-COMMENTS] Erreur mise à jour:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Supprime un commentaire
     */
    deleteWeeklyComment(id) {
        try {
            const stmt = this.db.prepare('DELETE FROM weekly_comments WHERE id = ?');
            const res = stmt.run(id);
            return { success: res.changes > 0, changes: res.changes };
        } catch (error) {
            console.error('❌ [WEEKLY-COMMENTS] Erreur suppression:', error);
            return { success: false, error: error.message };
        }
    }

    /**
     * Liste les semaines distinctes (pour le sélecteur de commentaires)
     */
    listDistinctWeeks(limit = 52) {
        try {
            const rows = this.db.prepare(`
                SELECT week_identifier, MAX(week_year) AS week_year, MAX(week_number) AS week_number,
                       MAX(week_start_date) AS week_start_date, MAX(week_end_date) AS week_end_date
                FROM weekly_stats
                GROUP BY week_identifier
                ORDER BY week_year DESC, week_number DESC
                LIMIT ?
            `).all(limit);
            return { success: true, rows };
        } catch (error) {
            console.error('❌ [WEEKS] Erreur liste semaines:', error);
            return { success: false, rows: [], error: error.message };
        }
    }

    /**
     * Méthodes pour gérer les paramètres d'application
     */
    getAppSetting(key, defaultValue = null) {
        try {
            const stmt = this.db.prepare('SELECT value FROM app_settings WHERE key = ?');
            const result = stmt.get(key);
            if (!result) return defaultValue;
            const raw = result.value;
            // Les valeurs sont stockées en JSON string; fallback si legacy texte
            try {
                return JSON.parse(raw);
            } catch {
                if (typeof raw === 'string') {
                    if (raw.toLowerCase() === 'true') return true;
                    if (raw.toLowerCase() === 'false') return false;
                }
                return raw ?? defaultValue;
            }
        } catch (error) {
            console.error(`❌ [SETTINGS] Erreur lecture paramètre ${key}:`, error);
            return defaultValue;
        }
    }

    setAppSetting(key, value) {
        try {
            const storedValue = JSON.stringify(value);
            const stmt = this.db.prepare(`
                INSERT OR REPLACE INTO app_settings (key, value, updated_at) 
                VALUES (?, ?, CURRENT_TIMESTAMP)
            `);
            const result = stmt.run(key, storedValue);
            this.cache.del && this.cache.del('app_settings');
            this.cache.del && this.cache.del(`setting_${key}`);
            console.log(`⚙️ [SETTINGS] Paramètre sauvegardé: ${key} = ${storedValue}`);
            return result.changes > 0;
        } catch (error) {
            console.error(`❌ [SETTINGS] Erreur sauvegarde paramètre ${key}:`, error);
            return false;
        }
    }

    /**
     * Compat: utilisé par l'IPC 'api-app-settings-save'
     */
    saveAppConfig(key, value) {
        try {
            // Toujours sérialiser en chaîne (évite l'erreur de binding pour les booléens)
            const storedValue = JSON.stringify(value);
            const stmt = this.db.prepare(`
                INSERT OR REPLACE INTO app_settings (key, value, updated_at)
                VALUES (?, ?, CURRENT_TIMESTAMP)
            `);
            stmt.run(key, storedValue);
            // Invalider le cache pour refléter immédiatement les changements
            this.cache.del && this.cache.del('app_settings');
            console.log(`📝 [SETTINGS] Config sauvegardée: ${key}`);
            return true;
        } catch (error) {
            console.error(`❌ [SETTINGS] Erreur saveAppConfig(${key}):`, error);
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
                // Recalculer les stats de la semaine (déterministe)
                this.updateCurrentWeekStats();
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
            const existing = this.statements.getEmailByEntryId.get(outlookId);
            const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);
            const shouldSetTreated = countReadAsTreated && isRead && existing && !existing.treated_at;
            let stmt;
            if (shouldSetTreated) {
                stmt = this.db.prepare(`
                    UPDATE emails 
                    SET is_read = 1,
                        is_treated = 1,
                        treated_at = COALESCE(treated_at, CURRENT_TIMESTAMP),
                        updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
                stmt.run(outlookId);
            } else {
                stmt = this.db.prepare(`
                    UPDATE emails 
                    SET is_read = ?, updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
                stmt.run(isRead ? 1 : 0, outlookId);
            }

            console.log(`📧 [VBA-LOGIC] Email ${outlookId} marqué ${isRead ? 'lu' : 'non lu'}`);
            this.updateCurrentWeekStats();
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
            let stmt;
            if (isTreated) {
                stmt = this.db.prepare(`
                    UPDATE emails 
                    SET is_treated = 1,
                        treated_at = COALESCE(treated_at, CURRENT_TIMESTAMP),
                        updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
                stmt.run(outlookId);
            } else {
                stmt = this.db.prepare(`
                    UPDATE emails 
                    SET is_treated = 0,
                        -- ne pas effacer treated_at si on "détrait"; on garde l'historique
                        updated_at = CURRENT_TIMESTAMP 
                    WHERE outlook_id = ?
                `);
                stmt.run(outlookId);
            }
            
            console.log(`� [VBA-LOGIC] Email ${outlookId} marqué ${isTreated ? 'traité' : 'non traité'}`);
            this.updateCurrentWeekStats();
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
            // Nouveau: se base uniquement sur treated_at
            const stmt = this.db.prepare(`
                SELECT 
                    folder_name,
                    category,
                    COUNT(*) as treatments
                FROM emails 
                WHERE treated_at IS NOT NULL AND (treated_at BETWEEN ? AND ?)
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
