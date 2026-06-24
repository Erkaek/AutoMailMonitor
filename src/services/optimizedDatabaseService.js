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
            this.ensureFolderIdColumns();
            
            // S'assurer que les colonnes de suppression existent
                // this.ensureDeletedColumn(); // supprimé, colonne gérée dans la migration
            // Nouvelle colonne: treated_at (date de traitement)
            this.ensureTreatedAtColumn();
            // Tracking: internet_message_id + last_modified_time + first/last seen
            this.ensureEmailTrackingColumns();
            // Curseur par dossier
            this.ensureFolderSyncStateTable();
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
                internet_message_id TEXT NULL,
                subject TEXT NOT NULL,
                sender_email TEXT,
                received_time DATETIME,
                last_modified_time DATETIME NULL,
                folder_name TEXT,
                category TEXT,
                is_read BOOLEAN DEFAULT 0,
                is_treated BOOLEAN DEFAULT 0,
                deleted_at DATETIME NULL,
        treated_at DATETIME NULL,
                week_identifier TEXT,
                first_seen_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                last_seen_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        `);

        // (activity_weekly supprimé)

        // Table des configurations de dossiers surveillés (chemin + IDs)
        this.db.exec(`
            CREATE TABLE IF NOT EXISTS folder_configurations (
                folder_path TEXT PRIMARY KEY,
                folder_name TEXT NOT NULL,
                category TEXT NOT NULL,
                store_id TEXT NULL,
                entry_id TEXT NULL,
                store_name TEXT NULL,
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

        // Table d'état de synchronisation par dossier (curseur persistant)
        this.db.exec(`
            CREATE TABLE IF NOT EXISTS folder_sync_state (
                folder_path TEXT PRIMARY KEY,
                store_id TEXT NULL,
                entry_id TEXT NULL,
                store_name TEXT NULL,
                last_modified_cursor DATETIME NULL,
                last_full_scan_at DATETIME NULL,
                baseline_done INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
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
            CREATE INDEX IF NOT EXISTS idx_folder_sync_store_entry ON folder_sync_state(store_id, entry_id);
            CREATE INDEX IF NOT EXISTS idx_folder_sync_baseline_done ON folder_sync_state(baseline_done);
        `);
    }

    /**
     * Migration: ajoute store_id/entry_id/store_name aux folder_configurations si absents
     */
    ensureFolderIdColumns() {
        try {
            const cols = this.db.prepare(`PRAGMA table_info(folder_configurations)`).all();
            const have = (name) => cols.some(c => String(c.name).toLowerCase() === name);
            if (!have('store_id')) {
                this.db.exec(`ALTER TABLE folder_configurations ADD COLUMN store_id TEXT NULL`);
            }
            if (!have('entry_id')) {
                this.db.exec(`ALTER TABLE folder_configurations ADD COLUMN entry_id TEXT NULL`);
            }
            if (!have('store_name')) {
                this.db.exec(`ALTER TABLE folder_configurations ADD COLUMN store_name TEXT NULL`);
            }
            this.db.exec(`CREATE INDEX IF NOT EXISTS idx_folder_cfg_store_entry ON folder_configurations(store_id, entry_id)`);
        } catch (e) {
            console.warn('⚠️ Migration folder_configurations (store/entry) ignorée:', e.message);
        }
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
     * Migration: ajoute colonnes de tracking email si absentes
     */
    ensureEmailTrackingColumns() {
        try {
            const cols = this.db.prepare(`PRAGMA table_info(emails)`).all();
            const have = (name) => cols.some(c => String(c.name).toLowerCase() === name);
            if (!have('internet_message_id')) {
                this.db.exec(`ALTER TABLE emails ADD COLUMN internet_message_id TEXT NULL`);
            }
            if (!have('last_modified_time')) {
                this.db.exec(`ALTER TABLE emails ADD COLUMN last_modified_time DATETIME NULL`);
            }
            if (!have('first_seen_at')) {
                this.db.exec(`ALTER TABLE emails ADD COLUMN first_seen_at DATETIME NULL`);
            }
            if (!have('last_seen_at')) {
                this.db.exec(`ALTER TABLE emails ADD COLUMN last_seen_at DATETIME NULL`);
            }
            // Index idempotents
            try { this.db.exec(`CREATE UNIQUE INDEX IF NOT EXISTS idx_emails_internet_message_id ON emails(internet_message_id)`); } catch {}
            try { this.db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_last_modified_time ON emails(last_modified_time)`); } catch {}
            try { this.db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_last_seen_at ON emails(last_seen_at)`); } catch {}

            // Backfill léger
            this.db.exec(`UPDATE emails SET first_seen_at = COALESCE(first_seen_at, created_at, CURRENT_TIMESTAMP) WHERE first_seen_at IS NULL`);
            this.db.exec(`UPDATE emails SET last_seen_at = COALESCE(last_seen_at, updated_at, CURRENT_TIMESTAMP) WHERE last_seen_at IS NULL`);
        } catch (e) {
            console.warn('⚠️ Migration emails tracking ignorée:', e.message);
        }
    }

    /**
     * Migration: s'assurer que la table folder_sync_state existe
     */
    ensureFolderSyncStateTable() {
        try {
            this.db.exec(`
                CREATE TABLE IF NOT EXISTS folder_sync_state (
                    folder_path TEXT PRIMARY KEY,
                    store_id TEXT NULL,
                    entry_id TEXT NULL,
                    store_name TEXT NULL,
                    last_modified_cursor DATETIME NULL,
                    last_full_scan_at DATETIME NULL,
                    baseline_done INTEGER DEFAULT 0,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            `);
            this.db.exec(`CREATE INDEX IF NOT EXISTS idx_folder_sync_store_entry ON folder_sync_state(store_id, entry_id)`);
            this.db.exec(`CREATE INDEX IF NOT EXISTS idx_folder_sync_baseline_done ON folder_sync_state(baseline_done)`);
        } catch (e) {
            console.warn('⚠️ Migration folder_sync_state ignorée:', e.message);
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
                 category, is_read, is_treated, deleted_at, treated_at, week_identifier,
                 internet_message_id, last_modified_time, first_seen_at, last_seen_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
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
                    internet_message_id = COALESCE(?, internet_message_id),
                    last_modified_time = COALESCE(?, last_modified_time),
                    last_seen_at = CURRENT_TIMESTAMP,
                    -- Allow clearing treated_at when a mail becomes unread (if user expects treated to follow read)
                    treated_at = CASE WHEN ? = 1 THEN NULL ELSE COALESCE(?, treated_at) END,
                    updated_at = CURRENT_TIMESTAMP
                WHERE outlook_id = ?
            `),
            
            getRecentEmails: this.db.prepare(`
                SELECT * FROM emails 
                WHERE deleted_at IS NULL
                ORDER BY received_time DESC 
                LIMIT ?
            `),
            
        getEmailStats: this.db.prepare(`
                SELECT 
                    COUNT(*) as totalEmails,
                    SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unreadTotal,
                    SUM(CASE WHEN DATE(first_seen_at) = DATE('now') THEN 1 ELSE 0 END) as emailsToday,
            SUM(CASE WHEN DATE(treated_at) = DATE('now') THEN 1 ELSE 0 END) as treatedToday
                FROM emails
                WHERE deleted_at IS NULL
            `),

            // Réconciliation (dossier = folder_name == folder.path)
            getActiveEmailCountByFolder: this.db.prepare(`
                SELECT COUNT(*) as count
                FROM emails
                WHERE deleted_at IS NULL AND folder_name = ?
            `),
            softDeleteMissingEmailsByFolderSince: this.db.prepare(`
                UPDATE emails
                SET deleted_at = COALESCE(deleted_at, CURRENT_TIMESTAMP),
                    treated_at = COALESCE(treated_at, CURRENT_TIMESTAMP),
                    is_treated = 1,
                    updated_at = CURRENT_TIMESTAMP
                WHERE deleted_at IS NULL
                  AND folder_name = ?
                  AND (last_seen_at IS NULL OR last_seen_at < ?)
            `),

            // Rétro-migration: appliquer "lu = traité" aux emails déjà en base
            markExistingReadAsTreated: this.db.prepare(`
                UPDATE emails
                SET is_treated = 1,
                    treated_at = COALESCE(treated_at, last_modified_time, received_time, first_seen_at, CURRENT_TIMESTAMP),
                    updated_at = CURRENT_TIMESTAMP
                WHERE deleted_at IS NULL
                  AND treated_at IS NULL
                  AND (is_read = 1 OR is_read = '1' OR is_read = TRUE)
            `),

            // Rétro-migration inverse: si on désactive "lu = traité", remettre "non traité" les mails lus
            // (hors soft-deleted) pour coller à la règle.
            clearExistingReadTreated: this.db.prepare(`
                UPDATE emails
                SET is_treated = 0,
                    treated_at = NULL,
                    updated_at = CURRENT_TIMESTAMP
                WHERE deleted_at IS NULL
                  AND treated_at IS NOT NULL
                  AND (is_read = 1 OR is_read = '1' OR is_read = TRUE)
            `),
            
            // Folders
            getFoldersConfig: this.db.prepare(`
                SELECT folder_path, category, folder_name, store_id, entry_id, store_name FROM folder_configurations
            `),
            
            insertFolderConfig: this.db.prepare(`
                INSERT OR REPLACE INTO folder_configurations (folder_path, category, folder_name, store_id, entry_id, store_name, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
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

            getEmailByInternetMessageId: this.db.prepare(`
                SELECT * FROM emails WHERE internet_message_id = ?
            `),

            updateEmailOutlookIdByInternetMessageId: this.db.prepare(`
                UPDATE emails
                SET outlook_id = ?,
                    folder_name = ?,
                    category = ?,
                    is_read = ?,
                    last_seen_at = CURRENT_TIMESTAMP,
                    updated_at = CURRENT_TIMESTAMP
                WHERE internet_message_id = ?
            `),
            
            getEmailByEntryIdAndFolder: this.db.prepare(`
                SELECT * FROM emails WHERE outlook_id = ? AND folder_name = ?
            `)
            ,
            touchEmailSeen: this.db.prepare(`
                UPDATE emails
                SET last_seen_at = CURRENT_TIMESTAMP,
                    updated_at = CURRENT_TIMESTAMP
                WHERE outlook_id = ?
            `)
            ,
            markReadAsTreatedIfNull: this.db.prepare(`
                UPDATE emails
                SET is_treated = 1,
                    treated_at = COALESCE(treated_at, ?),
                    updated_at = CURRENT_TIMESTAMP
                WHERE outlook_id = ?
                  AND deleted_at IS NULL
                  AND treated_at IS NULL
                  AND is_read = 1
            `)
        };

        // Folder sync state
        this.statements.getFolderSyncState = this.db.prepare(`
            SELECT * FROM folder_sync_state WHERE folder_path = ?
        `);
        this.statements.upsertFolderSyncState = this.db.prepare(`
            INSERT INTO folder_sync_state (
                folder_path, store_id, entry_id, store_name,
                last_modified_cursor, last_full_scan_at, baseline_done,
                updated_at
            ) VALUES (
                @folder_path, @store_id, @entry_id, @store_name,
                @last_modified_cursor, @last_full_scan_at, COALESCE(@baseline_done, 0),
                CURRENT_TIMESTAMP
            )
            ON CONFLICT(folder_path) DO UPDATE SET
                store_id=excluded.store_id,
                entry_id=excluded.entry_id,
                store_name=excluded.store_name,
                last_modified_cursor=COALESCE(excluded.last_modified_cursor, folder_sync_state.last_modified_cursor),
                last_full_scan_at=COALESCE(excluded.last_full_scan_at, folder_sync_state.last_full_scan_at),
                baseline_done=CASE WHEN excluded.baseline_done > folder_sync_state.baseline_done THEN excluded.baseline_done ELSE folder_sync_state.baseline_done END,
                updated_at=CURRENT_TIMESTAMP
        `);

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
     * Applique rétroactivement la règle "mail lu = traité" sur les emails existants.
     * Important: met aussi à jour les agrégats (weekly_stats) pour refléter ce changement.
     * @returns {{ success: boolean, updatedEmails?: number, updatedWeeklyRows?: number, error?: string }}
     */
    applyReadAsTreatedRetroactively() {
        try {
            const tx = this.db.transaction(() => {
                const res = this.statements.markExistingReadAsTreated.run();
                const weeklyRes = this.rebuildWeeklyStatsFromEmails();
                return { updatedEmails: res?.changes || 0, updatedWeeklyRows: weeklyRes?.updatedRows || 0 };
            });

            const { updatedEmails, updatedWeeklyRows } = tx();

            // Invalider caches et forcer les refresh UI
            this.invalidateFolderCache('');
            this.invalidateUICache();

            return { success: true, updatedEmails, updatedWeeklyRows };
        } catch (e) {
            console.error('❌ [DB] Erreur applyReadAsTreatedRetroactively:', e);
            return { success: false, error: e?.message || String(e) };
        }
    }

    /**
     * Applique rétroactivement la règle inverse quand on désactive "mail lu = traité":
     * les emails non supprimés (deleted_at NULL) qui sont à la fois lus et traités repassent non traités.
     * Important: met aussi à jour les agrégats (weekly_stats).
     * @returns {{ success: boolean, updatedEmails?: number, updatedWeeklyRows?: number, error?: string }}
     */
    unapplyReadAsTreatedRetroactively() {
        try {
            const tx = this.db.transaction(() => {
                const res = this.statements.clearExistingReadTreated.run();
                const weeklyRes = this.rebuildWeeklyStatsFromEmails();
                return { updatedEmails: res?.changes || 0, updatedWeeklyRows: weeklyRes?.updatedRows || 0 };
            });

            const { updatedEmails, updatedWeeklyRows } = tx();

            this.invalidateFolderCache('');
            this.invalidateUICache();

            return { success: true, updatedEmails, updatedWeeklyRows };
        } catch (e) {
            console.error('❌ [DB] Erreur unapplyReadAsTreatedRetroactively:', e);
            return { success: false, error: e?.message || String(e) };
        }
    }

    /**
     * Recalcule complètement la table weekly_stats à partir de la table emails.
     * Conserve les champs manual_adjustments existants.
     * @returns {{ updatedRows: number }}
     */
    rebuildWeeklyStatsFromEmails() {
        const manual = new Map();
        const manualMeta = new Map();
        try {
            const existingManual = this.db.prepare(`
                SELECT week_identifier, folder_type,
                       week_number, week_year, week_start_date, week_end_date,
                       COALESCE(manual_adjustments, 0) AS manual_adjustments
                FROM weekly_stats
            `).all();

            for (const row of existingManual) {
                manual.set(`${row.week_identifier}||${row.folder_type}`, Number(row.manual_adjustments) || 0);
                manualMeta.set(`${row.week_identifier}||${row.folder_type}`, {
                    week_number: Number(row.week_number) || 0,
                    week_year: Number(row.week_year) || 0,
                    week_start_date: row.week_start_date,
                    week_end_date: row.week_end_date
                });
            }
        } catch (e) {
            console.warn('⚠️ [WEEKLY] Impossible de charger manual_adjustments:', e?.message || e);
        }

        // Cache local des catégories (éviter une requête SQL par email)
        const folderCategory = new Map();
        try {
            const rows = this.db.prepare(`SELECT folder_path, category FROM folder_configurations`).all();
            for (const r of rows) {
                folderCategory.set(r.folder_path, this.normalizeCategory(r.category));
            }
        } catch (e) {
            // fallback: mapping par défaut plus bas
        }

        const defaultCategoryFromPath = (folderPath) => {
            const folderName = String(folderPath || '').split('\\').pop() || '';
            const name = folderName.toLowerCase();
            if (name.includes('test')) return 'Déclarations';
            if (name.includes('regel') || name.includes('reglement')) return 'Règlements';
            return 'Mails simples';
        };

        const resolveCategory = (folderPath) => {
            return folderCategory.get(folderPath) || defaultCategoryFromPath(folderPath);
        };

        const byKey = new Map();
        const ensureEntry = (weekInfo, folderType) => {
            const key = `${weekInfo.identifier}||${folderType}`;
            if (!byKey.has(key)) {
                byKey.set(key, {
                    week_identifier: weekInfo.identifier,
                    week_number: weekInfo.weekNumber,
                    week_year: weekInfo.year,
                    week_start_date: weekInfo.startDate,
                    week_end_date: weekInfo.endDate,
                    folder_type: folderType,
                    emails_received: 0,
                    emails_treated: 0,
                    manual_adjustments: manual.get(key) || 0
                });
            }
            return byKey.get(key);
        };

        const parseDateOnly = (value) => {
            if (!value) return null;
            const s = String(value);
            const d = s.length >= 10 ? s.slice(0, 10) : null;
            if (!d || !/^\d{4}-\d{2}-\d{2}$/.test(d)) return null;
            return new Date(`${d}T00:00:00Z`);
        };

        // 1) Reçus: basé sur first_seen_at (même logique que updateCurrentWeekStats)
        const receivedRows = this.db.prepare(`
            SELECT folder_name, first_seen_at
            FROM emails
            WHERE first_seen_at IS NOT NULL
        `).iterate();

        for (const row of receivedRows) {
            const dt = parseDateOnly(row.first_seen_at);
            if (!dt) continue;
            const weekInfo = this.getISOWeekInfo(dt);
            const folderType = resolveCategory(row.folder_name) || 'Mails simples';
            const entry = ensureEntry(weekInfo, folderType);
            entry.emails_received += 1;
        }

        // 2) Traités: basé sur treated_at
        const treatedRows = this.db.prepare(`
            SELECT folder_name, treated_at
            FROM emails
            WHERE treated_at IS NOT NULL
        `).iterate();

        for (const row of treatedRows) {
            const dt = parseDateOnly(row.treated_at);
            if (!dt) continue;
            const weekInfo = this.getISOWeekInfo(dt);
            const folderType = resolveCategory(row.folder_name) || 'Mails simples';
            const entry = ensureEntry(weekInfo, folderType);
            entry.emails_treated += 1;
        }

        // Inclure les semaines existantes qui n'auraient plus d'emails (mais ont des ajustements)
        for (const [k, adj] of manual.entries()) {
            if (!byKey.has(k)) {
                const [week_identifier, folder_type] = k.split('||');
                const meta = manualMeta.get(k) || {};
                // Fallback: reconstruire week_number/year depuis l'identifiant Sxx-YYYY si meta manquante
                const m = /^S(\d+)-(\d{4})$/.exec(week_identifier);
                const weekNumber = meta.week_number || (m ? Number(m[1]) : 0);
                const year = meta.week_year || (m ? Number(m[2]) : 0);

                // Si on n'a pas start/end, approx sur la base de l'année ISO
                const approxDate = new Date(Date.UTC(year || 1970, 0, 4));
                const weekInfo = this.getISOWeekInfo(approxDate);
                byKey.set(k, {
                    week_identifier,
                    week_number: weekNumber || weekInfo.weekNumber,
                    week_year: year || weekInfo.year,
                    week_start_date: meta.week_start_date || weekInfo.startDate,
                    week_end_date: meta.week_end_date || weekInfo.endDate,
                    folder_type,
                    emails_received: 0,
                    emails_treated: 0,
                    manual_adjustments: adj
                });
            }
        }

        const upsert = this.db.prepare(`
            INSERT OR REPLACE INTO weekly_stats
            (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type,
             emails_received, emails_treated, manual_adjustments, updated_at)
            VALUES (@week_identifier, @week_number, @week_year, @week_start_date, @week_end_date, @folder_type,
                    @emails_received, @emails_treated, @manual_adjustments, CURRENT_TIMESTAMP)
        `);

        let updatedRows = 0;
        for (const entry of byKey.values()) {
            upsert.run(entry);
            updatedRows += 1;
        }

        // Clamp de sécurité: éviter valeurs négatives
        this.clampWeeklyStatsNonNegative();

        return { updatedRows };
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

    getActiveEmailCountByFolder(folderPath) {
        try {
            if (!folderPath) return 0;
            const row = this.statements?.getActiveEmailCountByFolder?.get(folderPath);
            return Number(row?.count || 0);
        } catch (e) {
            console.warn('⚠️ [DB] getActiveEmailCountByFolder failed:', e.message);
            return 0;
        }
    }

    softDeleteMissingEmailsByFolderSince(folderPath, cutoffIso) {
        try {
            if (!folderPath || !cutoffIso) return 0;
            const res = this.statements?.softDeleteMissingEmailsByFolderSince?.run(folderPath, cutoffIso);
            return Number(res?.changes || 0);
        } catch (e) {
            console.warn('⚠️ [DB] softDeleteMissingEmailsByFolderSince failed:', e.message);
            return 0;
        }
    }

    touchEmailSeen(outlookId) {
        try {
            if (!outlookId) return 0;
            const res = this.statements?.touchEmailSeen?.run(outlookId);
            return Number(res?.changes || 0);
        } catch (e) {
            console.warn('⚠️ [DB] touchEmailSeen failed:', e.message);
            return 0;
        }
    }

    // Applique la règle "lu = traité" sans repasser par saveEmail (utile quand aucun changement n'est détecté)
    markReadAsTreatedIfNeeded(outlookId, treatedAtIso = null) {
        try {
            if (!outlookId) return 0;
            const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);
            if (!countReadAsTreated) return 0;

            const existing = this.statements?.getEmailByEntryId?.get(outlookId);
            if (!existing) return 0;
            if (existing.deleted_at) return 0;
            if (existing.treated_at) return 0;
            const isRead = existing.is_read === 1 || existing.is_read === true || String(existing.is_read) === '1';
            if (!isRead) return 0;

            const toIsoIfPossible = (raw) => {
                if (!raw) return null;
                const ms = Date.parse(String(raw));
                if (!Number.isFinite(ms)) return null;
                return new Date(ms).toISOString();
            };
            const when = treatedAtIso
                || toIsoIfPossible(existing.last_modified_time)
                || new Date().toISOString();

            const res = this.statements?.markReadAsTreatedIfNull?.run(when, outlookId);
            const changes = Number(res?.changes || 0);
            if (changes > 0) {
                // Invalider les agrégats pour que les compteurs se mettent à jour sans attendre.
                try {
                    for (const k of ['email_stats', 'category_stats', 'folder_stats', 'unread_email_count', 'total_email_count']) {
                        this.cache.del(k);
                    }
                    // Caches récents (quelques limites possibles)
                    for (const k of this.cache.keys()) {
                        if (k && typeof k === 'string' && k.startsWith('recent_emails_')) this.cache.del(k);
                    }
                } catch (_) {}
            }
            return changes;
        } catch (e) {
            console.warn('⚠️ [DB] markReadAsTreatedIfNeeded failed:', e.message);
            return 0;
        }
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

        // IMPORTANT (reporting): la semaine d'"arrivée" est basée sur la détection
        // dans le scope monitoré (et donc sur l'insertion), pas sur received_time.
        const weekId = this.calculateWeekIdentifier(new Date());
        // Normaliser le sujet par sécurité (si l'amont fournit Subject/ConversationTopic)
        const rawSubject = (emailData.subject ?? emailData.Subject ?? emailData.ConversationTopic ?? '').toString();
        const normalizedSubject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';

        // Id unique
        const outlookId = emailData.outlook_id || emailData.id || '';
        const internetMessageId = (emailData.internet_message_id || emailData.internetMessageId || emailData.InternetMessageId || '').toString().trim() || null;
        const lastModifiedTime = (emailData.last_modified_time || emailData.lastModifiedTime || emailData.LastModificationTime || '').toString().trim() || null;

        // Dedupe/re-key si EntryID change (mails déplacés entre dossiers, surtout boîtes partagées)
        let existed = this.statements.getEmailByEntryId.get(outlookId);
        if (!existed && internetMessageId) {
            try {
                const byImid = this.statements.getEmailByInternetMessageId.get(internetMessageId);
                if (byImid && byImid.outlook_id && byImid.outlook_id !== outlookId) {
                    // Ne pas écraser un éventuel mail déjà présent avec le même outlook_id
                    const collision = this.statements.getEmailByEntryId.get(outlookId);
                    if (!collision) {
                        this.statements.updateEmailOutlookIdByInternetMessageId.run(
                            outlookId,
                            emailData.folder_name || byImid.folder_name || '',
                            emailData.category || byImid.category || 'Mails simples',
                            emailData.is_read ? 1 : 0,
                            internetMessageId
                        );
                        existed = this.statements.getEmailByEntryId.get(outlookId);
                    }
                }
            } catch {}
        }

        const toIsoIfPossible = (raw) => {
            if (!raw) return null;
            const ms = Date.parse(String(raw));
            if (!Number.isFinite(ms)) return null;
            return new Date(ms).toISOString();
        };

        let result;
        if (existed) {
            // Mise à jour sans réinsérer (évite de compter une nouvelle arrivée)
            // Déterminer treated_at s'il faut le définir
            let treatedAtParam = null;
            const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);

            // Traité si déplacé hors du scope monitoré (dossier gestionnaire 11-… + descendants)
            const prevFolderPath = existed.folder_name || '';
            const nextFolderPath = emailData.folder_name || '';
            const prevInScope = this.isInMonitoredScope(prevFolderPath);
            const nextInScope = this.isInMonitoredScope(nextFolderPath);

            const nowIso = new Date().toISOString();
            if (!existed.treated_at) {
                if (emailData.is_treated) {
                    treatedAtParam = nowIso;
                } else if (countReadAsTreated && emailData.is_read) {
                    // Important au démarrage: si l'email arrive déjà en "lu", on le marque traité.
                    // On tente d'utiliser LastModificationTime (proche du moment où il est passé lu), sinon now.
                    treatedAtParam = toIsoIfPossible(lastModifiedTime) || nowIso;
                } else if (emailData.deleted_at) {
                    treatedAtParam = emailData.deleted_at;
                } else if (prevInScope && !nextInScope) {
                    treatedAtParam = nowIso;
                }
            }

            // Demande utilisateur: si un mail repasse NON LU, il doit repasser NON TRAITÉ
            // (quand le paramètre "lu = traité" est actif). On n'applique pas ce "dé-traitement"
            // aux mails supprimés ou sortis du scope monitoré.
            const shouldClearTreated = !!(
                countReadAsTreated &&
                !emailData.deleted_at &&
                nextInScope &&
                (emailData.is_read === false || emailData.is_read === 0) &&
                existed.treated_at
            );
            const clearTreatedFlag = shouldClearTreated ? 1 : 0;

            const effectiveIsTreated = shouldClearTreated
                ? false
                : !!(emailData.is_treated || emailData.deleted_at || (prevInScope && !nextInScope) || (countReadAsTreated && emailData.is_read) || treatedAtParam);

            if (shouldClearTreated) {
                treatedAtParam = null;
            }
            result = this.statements.updateEmailByOutlookId.run(
                normalizedSubject,
                emailData.sender_email || '',
                emailData.folder_name || '',
                emailData.category || 'Mails simples',
                emailData.is_read ? 1 : 0,
                effectiveIsTreated ? 1 : 0,
                internetMessageId,
                lastModifiedTime,
                clearTreatedFlag,
                treatedAtParam,
                outlookId
            );
        } else {
            // Insertion d'un nouvel email (véritable arrivée)
            const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);
            const nowIso = new Date().toISOString();
            const shouldTreatOnInsert = !!(emailData.is_treated || emailData.deleted_at || (countReadAsTreated && emailData.is_read));
            // treated_at initial: si déjà traité/supprimé OU si "lu = traité" et l'email arrive déjà en lu
            const initialTreatedAt = shouldTreatOnInsert
                ? (
                    emailData.deleted_at
                        ? emailData.deleted_at
                        : (emailData.treated_at || (countReadAsTreated && emailData.is_read ? (toIsoIfPossible(lastModifiedTime) || nowIso) : nowIso))
                )
                : null;
            result = this.statements.insertEmailNew.run(
                outlookId,
                normalizedSubject,
                emailData.sender_email || '',
                emailData.received_time || new Date().toISOString(),
                emailData.folder_name || '',
                emailData.category || 'Mails simples',
                emailData.is_read ? 1 : 0,
                shouldTreatOnInsert ? 1 : 0,
                emailData.deleted_at || null,
                initialTreatedAt,
                weekId,
                internetMessageId,
                lastModifiedTime
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

        const toIsoIfPossible = (raw) => {
            if (!raw) return null;
            const ms = Date.parse(String(raw));
            if (!Number.isFinite(ms)) return null;
            return new Date(ms).toISOString();
        };

        const countReadAsTreated = !!this.getAppSetting('count_read_as_treated', false);

        const transaction = this.db.transaction((emails) => {
            for (const email of emails) {
                const outlookId = email.outlook_id || email.id || '';
                const internetMessageId = (email.internet_message_id || email.internetMessageId || email.InternetMessageId || '').toString().trim() || null;
                const lastModifiedTime = (email.last_modified_time || email.lastModifiedTime || email.LastModificationTime || '').toString().trim() || null;

                let existed = this.statements.getEmailByEntryId.get(outlookId);
                if (!existed && internetMessageId) {
                    try {
                        const byImid = this.statements.getEmailByInternetMessageId.get(internetMessageId);
                        if (byImid && byImid.outlook_id && byImid.outlook_id !== outlookId) {
                            const collision = this.statements.getEmailByEntryId.get(outlookId);
                            if (!collision) {
                                this.statements.updateEmailOutlookIdByInternetMessageId.run(
                                    outlookId,
                                    email.folder_name || byImid.folder_name || '',
                                    email.category || byImid.category || 'Mails simples',
                                    email.is_read ? 1 : 0,
                                    internetMessageId
                                );
                                existed = this.statements.getEmailByEntryId.get(outlookId);
                            }
                        }
                    } catch {}
                }
                const rawSubject = (email.subject ?? email.Subject ?? email.ConversationTopic ?? '').toString();
                const normalizedSubject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
                if (existed) {
                    const nowIso = new Date().toISOString();
                    // Si l'email arrive déjà lu (au démarrage) et que "lu = traité" est actif,
                    // on ne peut pas toujours détecter la transition => on marque treated_at si absent.
                    let treatedAtParam = null;
                    if (!existed.treated_at) {
                        if (email.deleted_at) {
                            treatedAtParam = email.deleted_at;
                        } else if (email.is_treated) {
                            treatedAtParam = nowIso;
                        } else if (countReadAsTreated && email.is_read) {
                            treatedAtParam = toIsoIfPossible(lastModifiedTime) || nowIso;
                        }
                    }

                    const nextInScope = this.isInMonitoredScope(email.folder_name || '');
                    const shouldClearTreated = !!(
                        countReadAsTreated &&
                        !email.deleted_at &&
                        nextInScope &&
                        (email.is_read === false || email.is_read === 0) &&
                        existed.treated_at
                    );
                    const clearTreatedFlag = shouldClearTreated ? 1 : 0;
                    const effectiveIsTreated = shouldClearTreated
                        ? false
                        : !!(email.is_treated || email.deleted_at || (countReadAsTreated && email.is_read) || treatedAtParam);
                    if (shouldClearTreated) {
                        treatedAtParam = null;
                    }
                    this.statements.updateEmailByOutlookId.run(
                        normalizedSubject,
                        email.sender_email || '',
                        email.folder_name || '',
                        email.category || 'Mails simples',
                        email.is_read ? 1 : 0,
                        effectiveIsTreated ? 1 : 0,
                        internetMessageId,
                        lastModifiedTime,
                        clearTreatedFlag,
                        treatedAtParam,
                        outlookId
                    );
                } else {
                    // IMPORTANT: semaine d'arrivée = détection (insertion), pas received_time
                    const weekId = this.calculateWeekIdentifier(new Date());
                    const nowIso = new Date().toISOString();
                    const shouldTreatOnInsert = !!(email.is_treated || email.deleted_at || (countReadAsTreated && email.is_read));
                    const initialTreatedAt = shouldTreatOnInsert
                        ? (
                            email.deleted_at
                                ? email.deleted_at
                                : (email.treated_at || (countReadAsTreated && email.is_read ? (toIsoIfPossible(lastModifiedTime) || nowIso) : nowIso))
                        )
                        : null;
                    this.statements.insertEmailNew.run(
                        outlookId,
                        normalizedSubject,
                        email.sender_email || '',
                        email.received_time || new Date().toISOString(),
                        email.folder_name || '',
                        email.category || 'Mails simples',
                        email.is_read ? 1 : 0,
                        shouldTreatOnInsert ? 1 : 0,
                        null,
                        initialTreatedAt,
                        weekId,
                        internetMessageId,
                        lastModifiedTime
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

    // ==================== FOLDER SYNC STATE ====================
    getFolderSyncState(folderPath) {
        try {
            if (!folderPath) return null;
            return this.statements.getFolderSyncState.get(folderPath);
        } catch {
            return null;
        }
    }

    upsertFolderSyncState(row) {
        try {
            if (!row || !row.folder_path) return null;
            return this.statements.upsertFolderSyncState.run(row);
        } catch (e) {
            console.warn('⚠️ [DB] upsertFolderSyncState failed:', e.message);
            return null;
        }
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

    // ==================== MONITORED SCOPE (dossier gestionnaire 11-… + descendants) ====================
    _normalizeFolderPathForScope(p) {
        let s = String(p || '')
            .replace(/\//g, '\\')
            .replace(/\\+/g, '\\')
            .trim();

        // Corriger les cas où le nom de boîte est collé sans antislash (ex: "FlotteAutoBoîte de réception")
        // => "FlotteAuto\\Boîte de réception". On le fait ici pour que les comparisons de scope ne cassent pas.
        const inboxNames = ['Boîte de réception', 'Boite de reception', 'Inbox'];
        for (const inboxName of inboxNames) {
            const re = new RegExp(`([^\\\\])(${inboxName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'i');
            if (re.test(s)) {
                s = s.replace(re, '$1\\$2');
                s = s.replace(/\\+/g, '\\');
            }
        }

        return s.toLowerCase();
    }

    _extractManagerRootFromPath(folderPath) {
        const raw = String(folderPath || '').replace(/\//g, '\\');
        const parts = raw.split('\\').map(s => s.trim()).filter(Boolean);
        const isManagerSeg = (seg) => /^\d{2,}\s*-\s*.+/.test(seg);
        for (let i = 0; i < parts.length; i++) {
            if (isManagerSeg(parts[i])) {
                return parts.slice(0, i + 1).join('\\');
            }
        }
        return null;
    }

    getMonitoredScopeRoots() {
        const cacheKey = 'monitored_scope_roots_v1';
        const cached = this.cache.get(cacheKey);
        if (Array.isArray(cached)) return cached;

        let rows = [];
        try {
            rows = this.db.prepare(`SELECT folder_path FROM folder_configurations`).all();
        } catch {
            rows = [];
        }

        const roots = new Map();
        const addRoot = (p) => {
            const k = this._normalizeFolderPathForScope(p);
            if (!k) return;
            if (!roots.has(k)) roots.set(k, String(p || '').replace(/\//g, '\\'));
        };

        for (const r of rows) {
            const p = r?.folder_path;
            if (!p) continue;
            addRoot(p);
            const managerRoot = this._extractManagerRootFromPath(p);
            if (managerRoot) addRoot(managerRoot);
        }

        const result = Array.from(roots.values());
        this.cache.set(cacheKey, result, 60);
        return result;
    }

    isInMonitoredScope(folderPath) {
        const p = this._normalizeFolderPathForScope(folderPath);
        if (!p) return false;

        const roots = this.getMonitoredScopeRoots();
        if (!roots || roots.length === 0) return false;

        const pathSegs = p.split('\\').filter(Boolean);
        for (const root of roots) {
            const rNorm = this._normalizeFolderPathForScope(root);
            if (!rNorm) continue;

            // Match direct (prefix) pour couvrir le dossier racine + descendants
            if (p === rNorm) return true;
            if (p.startsWith(rNorm + '\\')) return true;

            // Fallback: match sur séquence de segments (utile si préfixe mailbox varie)
            const rSegs = rNorm.split('\\').filter(Boolean);
            if (rSegs.length === 0) continue;
            for (let i = 0; i <= pathSegs.length - rSegs.length; i++) {
                let ok = true;
                for (let j = 0; j < rSegs.length; j++) {
                    if (pathSegs[i + j] !== rSegs[j]) { ok = false; break; }
                }
                if (ok) return true;
            }
        }

        return false;
    }

    /**
     * Sauvegarde configuration dossier
     */
    addFolderConfiguration(folderPath, category, folderName, storeId = null, entryId = null, storeName = null) {
        this.cache.del('folders_config'); // Invalider cache
        const cat = this.normalizeCategory(category);
        return this.statements.insertFolderConfig.run(folderPath, cat, folderName, storeId || null, entryId || null, storeName || null);
    }

    /**
     * Insertion par lot des configurations de dossiers
     * rows: Array<{ folder_path|path, category, folder_name|name }>
     * Retourne { inserted, unique }
     */
    addFolderConfigurationsBatch(rows = []) {
        try {
            // Normaliser et dédupliquer par folder_path
            const map = new Map();
            for (const r of rows) {
                if (!r) continue;
                const folder_path = r.folder_path || r.path || r.key || r.folderPath;
                if (!folder_path) continue;
                const category = this.normalizeCategory(r.category || 'Mails simples') || 'Mails simples';
                const folder_name = r.folder_name || r.name || String(folder_path).split('\\').pop();
                const store_id = r.store_id || r.storeId;
                const entry_id = r.entry_id || r.entryId;
                const store_name = r.store_name || r.storeName;
                map.set(folder_path, { folder_path, category, folder_name, store_id, entry_id, store_name });
            }

            const uniqueRows = Array.from(map.values());
            if (uniqueRows.length === 0) {
                return { inserted: 0, unique: 0 };
            }

            const insert = this.statements.insertFolderConfig;
            const tx = this.db.transaction((items) => {
                for (const it of items) {
                    insert.run(it.folder_path, it.category, it.folder_name, it.store_id || null, it.entry_id || null, it.store_name || null);
                }
            });
            tx(uniqueRows);
            // Invalider le cache une seule fois
            this.cache.del('folders_config');
            return { inserted: uniqueRows.length, unique: uniqueRows.length };
        } catch (e) {
            console.error('❌ [DB] Erreur insertion par lot folder_configurations:', e.message);
            return { inserted: 0, unique: 0, error: e.message };
        }
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
            this.statements.insertFolderConfig.run(folderPath, cat, name, null, null, null);
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
                const storeId = r.store_id || r.storeId || null;
                const entryId = r.entry_id || r.entryId || null;
                const storeName = r.store_name || r.storeName || null;
                if (!category) {
                    // Cela ne devrait plus arriver, mais log si vide
                    console.warn('⚠️ [FOLDERS-CONFIG] Catégorie vide après normalisation, fallback Mails simples', folderPath);
                }
                insert.run(folderPath, category, name, storeId, entryId, storeName);
            }
        });

        // Normaliser payload en tableau de lignes
        let rows = [];
        if (payload && typeof payload === 'object' && payload.folderCategories && typeof payload.folderCategories === 'object') {
            rows = Object.entries(payload.folderCategories).map(([folder_path, cfg]) => ({
                folder_path,
                category: this.normalizeCategory(cfg?.category || 'Mails simples') || 'Mails simples',
                folder_name: cfg?.name || String(folder_path).split('\\').pop(),
                store_id: cfg?.store_id || cfg?.storeId || null,
                entry_id: cfg?.entry_id || cfg?.entryId || null,
                store_name: cfg?.store_name || cfg?.storeName || null
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
            const rows = this.db.prepare('SELECT folder_path, category, folder_name, store_id, entry_id, store_name, created_at, updated_at FROM folder_configurations').all();
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
                    WHERE deleted_at IS NULL AND category IS NOT NULL 
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
                                        WHERE deleted_at IS NULL
                                            AND DATE(received_time) BETWEEN ? AND ?
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
                const result = this.db.prepare('SELECT COUNT(*) as count FROM emails WHERE deleted_at IS NULL AND is_read = 0').get();
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
                const result = this.db.prepare('SELECT COUNT(*) as count FROM emails WHERE deleted_at IS NULL').get();
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
    getFolderStats(opts = {}) {
        const cacheKey = 'folder_stats';
        const force = !!opts.force;

        let stats = force ? null : this.cache.get(cacheKey);

        if (!stats) {
            try {
                const result = this.db.prepare(`
                    SELECT 
                        folder_name,
                        COUNT(*) as email_count,
                        COUNT(CASE WHEN is_read = 0 THEN 1 END) as unread_count,
                        MAX(received_time) as last_email
                    FROM emails
                    WHERE deleted_at IS NULL
                    GROUP BY folder_name
                `).all();

                stats = result.map(row => ({
                    path: row.folder_name,
                    emailCount: row.email_count,
                    unreadCount: row.unread_count,
                    lastEmail: row.last_email
                }));

                // Garder un cache par défaut (pour limiter les requêtes), mais permettre au renderer de bypass.
                // TTL réduit pour limiter les "compteurs figés" en cas d'appel non-forcé.
                this.cache.set(cacheKey, stats, 10);
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
            const emailRecord = {
                outlook_id: emailData.id,
                subject: emailData.subject || '(Sans objet)',
                sender_email: emailData.sender_email || '',
                received_time: emailData.receivedTime || new Date().toISOString(),
                folder_name: emailData.folderPath || '',
                is_read: emailData.isRead === true ? 1 : 0,
                is_treated: false,
                category: emailData.category || 'autres',
                internet_message_id: emailData.internet_message_id || emailData.internetMessageId || emailData.InternetMessageId || null,
                last_modified_time: emailData.lastModifiedTime || emailData.last_modified_time || emailData.LastModificationTime || null,
                deleted_at: null
            };

            const result = this.saveEmail(emailRecord);

            // Invalider le cache pour ce dossier
            this.invalidateFolderCache(emailData.folderPath);

            // Mettre à jour les stats
            this.updateQueryStats(Date.now() - startTime);

            console.log(`📧 [DB-COM] Nouvel email traité: ${emailData.subject} (outlook_id: ${emailRecord.outlook_id})`);
            
            return { 
                processed: true,
                rowId: result?.lastInsertRowid || null,
                changes: result?.changes || 0,
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
            const internetMessageId = (emailData.internet_message_id || emailData.internetMessageId || emailData.InternetMessageId || '').toString().trim() || null;
            let existingEmail = this.getEmailByEntryId(emailData.id);
            if (!existingEmail && internetMessageId) {
                existingEmail = this.getEmailByInternetMessageId(internetMessageId);
            }
            if (!existingEmail) {
                console.log(`⚠️ Email COM non trouvé pour mise à jour: ${emailData.id}`);
                return { updated: false, reason: 'not_found' };
            }
            const emailRecord = {
                outlook_id: emailData.id,
                subject: emailData.subject || existingEmail?.subject || '(Sans objet)',
                sender_email: emailData.sender_email || existingEmail?.sender_email || '',
                received_time: emailData.receivedTime || existingEmail?.received_time || new Date().toISOString(),
                folder_name: emailData.folderPath || existingEmail?.folder_name || '',
                is_read: emailData.isRead === true ? 1 : 0,
                is_treated: existingEmail?.is_treated ?? false,
                category: emailData.category || existingEmail?.category || 'autres',
                internet_message_id: internetMessageId || existingEmail?.internet_message_id || null,
                last_modified_time: emailData.lastModifiedTime || emailData.last_modified_time || emailData.LastModificationTime || null,
                deleted_at: existingEmail?.deleted_at || null
            };

            const result = this.saveEmail(emailRecord);

            // Invalider le cache
            this.invalidateFolderCache(emailData.folderPath || existingEmail?.folder_name || '');

            // Mettre à jour les stats
            this.updateQueryStats(Date.now() - startTime);

            console.log(`🔄 [DB-COM] Email mis à jour: ${emailData.id} (${result?.changes || 0} changements)`);
            
            return { 
                updated: true,
                changes: result?.changes || 0,
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

            // IMPORTANT: invalider aussi les agrégats globaux utilisés par l'UI.
            // Sinon les compteurs restent "figés" même si les données emails changent.
            for (const k of ['email_stats', 'category_stats', 'folder_stats', 'unread_email_count', 'total_email_count']) {
                try { this.cache.del(k); } catch (_) {}
            }
            // Invalider les caches dérivés (ex: email_count_YYYY-MM-DD_YYYY-MM-DD)
            try {
                for (const k of keys) {
                    if (k && typeof k === 'string' && k.startsWith('email_count_')) {
                        this.cache.del(k);
                    }
                }
            } catch (_) {}
            
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

            // IMPORTANT: invalider aussi les agrégats globaux.
            for (const k of ['email_stats', 'category_stats', 'folder_stats', 'unread_email_count', 'total_email_count']) {
                try { this.cache.del(k); } catch (_) {}
            }

            // Invalider les caches dérivés (ex: email_count_...)
            try {
                for (const k of keys) {
                    if (k && typeof k === 'string' && k.startsWith('email_count_')) {
                        this.cache.del(k);
                    }
                }
            } catch (_) {}
            
            // Invalider le cache de l'interface utilisateur via le service global
            if (global.unifiedMonitoringService && global.unifiedMonitoringService.invalidateEmailCache) {
                global.unifiedMonitoringService.invalidateEmailCache();
                console.log(`🔄 [REAL-TIME] Cache UI invalidé: mise à jour immédiate de l'interface`);
            }

            // IMPORTANT: invalider aussi le cache IPC/UI (cacheService) utilisé par les handlers du main.
            // Sinon le renderer peut continuer à recevoir des valeurs "figées" (ex: dashboard_stats, recent_XX).
            try {
                const cs = global.cacheService;
                if (cs) {
                    console.log('📊 [CACHE-INVALIDATE] Invalidation cacheService START');
                    if (typeof cs.invalidateStats === 'function') {
                        cs.invalidateStats();
                        console.log('  ✓ invalidateStats() appelée');
                    } else {
                        cs?.del?.('ui', 'dashboard_stats');
                        console.log('  ✓ dashboard_stats supprimée du cache ui');
                    }

                    if (typeof cs.invalidateEmailsCache === 'function') {
                        cs.invalidateEmailsCache();
                        console.log('  ✓ invalidateEmailsCache() appelée');
                    } else {
                        cs?.del?.('emails', 'recent_20');
                        cs?.del?.('emails', 'recent_50');
                        console.log('  ✓ recent_20 et recent_50 supprimées du cache emails');
                    }

                    if (typeof cs.invalidateFoldersTree === 'function') {
                        cs.invalidateFoldersTree();
                        console.log('  ✓ invalidateFoldersTree() appelée');
                    } else {
                        cs?.del?.('config', 'folders_tree');
                        console.log('  ✓ folders_tree supprimée du cache config');
                    }
                    console.log('📊 [CACHE-INVALIDATE] Invalidation cacheService DONE');
                } else {
                    console.warn('⚠️ [CACHE-INVALIDATE] cacheService est null/undefined');
                }
            } catch (e) {
                console.error('❌ [CACHE-INVALIDATE] Erreur lors invalidation cacheService:', e.message);
            }

            // Émettre un événement IPC pour forcer le renderer à rafraîchir immédiatement le dashboard
            try {
                if (global.mainWindow && global.mainWindow.webContents) {
                    global.mainWindow.webContents.send('stats-cache-invalidated', {
                        timestamp: new Date().toISOString(),
                        reason: 'DB cache invalidation'
                    });
                    console.log('📣 [IPC] Événement stats-cache-invalidated émis au renderer');
                }
            } catch (e) {
                console.warn('⚠️ [REAL-TIME] Impossible d\'émettre stats-cache-invalidated:', e?.message || e);
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
            console.log(`🔍 [DATABASE] Recherche InternetMessageId: ${emailUpdateData.internetMessageId || emailUpdateData.internet_message_id || ''}`);
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
            
            const internetMessageId = (emailUpdateData.internetMessageId || emailUpdateData.internet_message_id || '').toString().trim() || null;

            // Vérifier d'abord si l'email existe en base
            let existingEmail = this.getEmailByEntryId(emailUpdateData.messageId, emailUpdateData.folderPath);
            let matchedBy = existingEmail ? 'outlook_id' : null;
            if (!existingEmail && internetMessageId) {
                existingEmail = this.getEmailByInternetMessageId(internetMessageId);
                matchedBy = existingEmail ? 'internet_message_id' : null;
            }
            
            if (!existingEmail) {
                console.log(`⚠️ [DATABASE] Email non trouvé en base: ${emailUpdateData.messageId} - ${emailUpdateData.subject}`);
                return { updated: false, reason: 'Email non trouvé en base' };
            }

            const rawSubject = (emailUpdateData.subject ?? existingEmail.subject ?? '').toString();
            const normalizedSubject = rawSubject.trim() !== '' ? rawSubject : '(Sans objet)';
            let newIsRead = existingEmail.is_read;
            
            console.log(`🔍 [DATABASE] Email trouvé - état actuel: is_read=${existingEmail.is_read}, subject="${existingEmail.subject}"`);
            
            // CORRECTION: Normaliser les valeurs pour la comparaison (0/1 vs false/true)
            const currentIsReadBool = Boolean(existingEmail.is_read);
            
            // NOUVELLE LOGIQUE: Toujours faire confiance à Outlook pour les changements de statut
            // Outlook envoie ces événements seulement quand il y a eu une action utilisateur réelle
            if (emailUpdateData.changes && emailUpdateData.changes.includes('MarkedRead')) {
                newIsRead = true;
                console.log(`📖 [DATABASE] FORCE Marquage comme lu: ${emailUpdateData.subject} (BDD: ${existingEmail.is_read}/${currentIsReadBool} -> Outlook: ${newIsRead})`);
            } else if (emailUpdateData.changes && emailUpdateData.changes.includes('MarkedUnread')) {
                newIsRead = false;
                console.log(`📬 [DATABASE] FORCE Marquage comme non lu: ${emailUpdateData.subject} (BDD: ${existingEmail.is_read}/${currentIsReadBool} -> Outlook: ${newIsRead})`);
            } else if (emailUpdateData.isRead !== undefined && emailUpdateData.isRead !== currentIsReadBool) {
                newIsRead = emailUpdateData.isRead;
                console.log(`📝 [DATABASE] Mise à jour statut lecture: ${emailUpdateData.subject} -> ${emailUpdateData.isRead ? 'Lu' : 'Non lu'} (actuel: ${existingEmail.is_read}/${currentIsReadBool} -> nouveau: ${newIsRead})`);
            }

            const hasMeaningfulChange = newIsRead !== currentIsReadBool ||
                normalizedSubject !== (existingEmail.subject || '(Sans objet)') ||
                (emailUpdateData.folderPath && emailUpdateData.folderPath !== existingEmail.folder_name);

            if (!hasMeaningfulChange) {
                console.log(`ℹ️ [DATABASE] Aucune modification à appliquer: ${emailUpdateData.subject} (state déjà identique)`);
                return { updated: false, reason: 'Aucune modification détectée', matchedBy };
            }

            const emailRecord = {
                outlook_id: emailUpdateData.messageId,
                subject: normalizedSubject,
                sender_email: existingEmail.sender_email || '',
                received_time: existingEmail.received_time || new Date().toISOString(),
                folder_name: emailUpdateData.folderPath || existingEmail.folder_name || '',
                is_read: newIsRead ? 1 : 0,
                is_treated: existingEmail.is_treated,
                category: existingEmail.category || 'autres',
                internet_message_id: internetMessageId || existingEmail.internet_message_id || null,
                last_modified_time: emailUpdateData.lastModificationTime || existingEmail.last_modified_time || null,
                deleted_at: existingEmail.deleted_at || null
            };

            console.log(`🚀 [DATABASE] Procédure de mise à jour: ${emailUpdateData.subject} - BDD:${existingEmail.is_read} -> Outlook:${newIsRead}`);
            const result = this.saveEmail(emailRecord);

            console.log(`📊 [DATABASE] Résultat saveEmail: changes=${result?.changes || 0}`);
            this.invalidateFolderCache(emailUpdateData.folderPath || existingEmail.folder_name || '');
            console.log(`✅ [DATABASE] Email mis à jour via polling: ${emailUpdateData.subject} (${result?.changes || 0} changements)`);
            return {
                updated: true,
                changes: result?.changes || 0,
                emailData: { ...existingEmail, is_read: newIsRead },
                matchedBy
            };

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
            
            // Récupérer tous les emails ARRIVÉS de la semaine actuelle groupés par dossier
            // IMPORTANT (reporting): arrivés = première détection dans le scope monitoré => first_seen_at.
            const emailStats = this.db.prepare(`
                SELECT 
                    folder_name,
                    COUNT(*) as total_emails,
                    SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread_emails
                FROM emails 
                WHERE first_seen_at IS NOT NULL
                  AND DATE(first_seen_at) BETWEEN ? AND ?
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

            // Grouper les emails reçus par type de folder
            const statsByType = {};
            emailStats.forEach(stat => {
                const folderType = this.mapFolderToCategory(stat.folder_name) || 'Mails simples';
                if (!statsByType[folderType]) {
                    statsByType[folderType] = { received: 0, treated: 0 };
                }
                statsByType[folderType].received += stat.total_emails;
            });

            // Grouper les emails traités par type de folder
            treatedStats.forEach(stat => {
                const folderType = this.mapFolderToCategory(stat.folder_name) || 'Mails simples';
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
     * Ajustements de lancement (stock initial) par catégorie.
     * Stockés dans app_settings sous la clé "startup_adjustments".
     * @returns {{declarations:number, reglements:number, mails_simples:number}}
     */
    getStartupAdjustments() {
        const defaults = { declarations: 0, reglements: 0, mails_simples: 0 };
        try {
            const raw = this.getAppSetting('startup_adjustments', defaults);
            const obj = (raw && typeof raw === 'object') ? raw : defaults;
            const toInt = (v) => {
                const n = Number.parseInt(v, 10);
                return Number.isFinite(n) ? n : 0;
            };
            return {
                declarations: toInt(obj.declarations),
                reglements: toInt(obj.reglements),
                mails_simples: toInt(obj.mails_simples)
            };
        } catch (e) {
            return defaults;
        }
    }

    saveStartupAdjustments(values) {
        const toInt = (v) => {
            const n = Number.parseInt(v, 10);
            return Number.isFinite(n) ? n : 0;
        };
        const payload = {
            declarations: toInt(values?.declarations),
            reglements: toInt(values?.reglements),
            mails_simples: toInt(values?.mails_simples)
        };
        return this.saveAppSetting('startup_adjustments', payload);
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

            // Stock initial (ajustements de lancement) + accumulation semaine par semaine
            const startup = this.getStartupAdjustments();
            const carry = {
                declarations: startup.declarations || 0,
                reglements: startup.reglements || 0,
                mails_simples: startup.mails_simples || 0
            };
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
    hasAppSetting(key) {
        try {
            const stmt = this.db.prepare('SELECT 1 AS ok FROM app_settings WHERE key = ? LIMIT 1');
            const row = stmt.get(key);
            return !!row;
        } catch (error) {
            console.error(`❌ [SETTINGS] Erreur existence paramètre ${key}:`, error);
            return false;
        }
    }

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

            // Définition "traité": supprimé => traité (si dans scope monitoré), sans écraser un treated_at existant
            if (result.changes > 0 && emailInfo && !emailInfo.treated_at) {
                const wasInScope = this.isInMonitoredScope(emailInfo.folder_name || '');
                if (wasInScope) {
                    try {
                        this.db.prepare(`
                            UPDATE emails
                            SET treated_at = COALESCE(treated_at, ?),
                                is_treated = 1,
                                updated_at = CURRENT_TIMESTAMP
                            WHERE outlook_id = ?
                        `).run(deleteTime, outlookId);
                    } catch {}
                }
            }
            
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
            // Arrivées = first_seen_at dans la semaine (détection dans scope)
            const weekRow = this.db.prepare(`
                SELECT week_start_date, week_end_date
                FROM weekly_stats
                WHERE week_identifier = ?
                LIMIT 1
            `).get(weekIdentifier);

            let startDate;
            let endDate;
            if (weekRow && weekRow.week_start_date && weekRow.week_end_date) {
                startDate = weekRow.week_start_date;
                endDate = weekRow.week_end_date;
            } else {
                const [week, year] = String(weekIdentifier || '').split('-');
                const weekNum = parseInt(String(week || '').substring(1), 10);
                startDate = this.getWeekStartDate(weekNum, parseInt(year, 10)).split('T')[0];
                endDate = this.getWeekEndDate(weekNum, parseInt(year, 10)).split('T')[0];
            }

            const stmt = this.db.prepare(`
                SELECT 
                    folder_name,
                    category,
                    COUNT(*) as arrivals
                FROM emails 
                WHERE first_seen_at IS NOT NULL
                  AND DATE(first_seen_at) BETWEEN ? AND ?
                GROUP BY folder_name, category
                ORDER BY folder_name, category
            `);

            return stmt.all(startDate, endDate);
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

    // =====================================================================
    // BDD: lecture brute (UI Debug) - READ ONLY
    // =====================================================================

    listDatabaseTables() {
        try {
            if (!this.db) throw new Error('DB non initialisée');
            const rows = this.db.prepare(`
                SELECT name
                FROM sqlite_master
                WHERE type = 'table'
                  AND name NOT LIKE 'sqlite_%'
                ORDER BY name
            `).all();
            return rows.map(r => r.name).filter(Boolean);
        } catch (error) {
            console.error('❌ [DB-VIEW] Erreur listDatabaseTables:', error);
            return [];
        }
    }

    getDatabaseTablePreview(tableName, limit = 200, offset = 0) {
        try {
            if (!this.db) throw new Error('DB non initialisée');

            const table = String(tableName || '').trim();
            if (!table) throw new Error('Nom de table requis');

            const tables = new Set(this.listDatabaseTables());
            if (!tables.has(table)) throw new Error('Table inconnue');

            const safeLimit = Math.max(1, Math.min(1000, Number(limit) || 200));
            const safeOffset = Math.max(0, Math.min(500000, Number(offset) || 0));

            const quoted = `"${table.replace(/"/g, '""')}"`;
            const cols = this.db.prepare(`PRAGMA table_info(${quoted})`).all();
            const columns = (cols || []).map(c => c.name).filter(Boolean);

            const rows = this.db.prepare(`SELECT * FROM ${quoted} LIMIT ? OFFSET ?`).all(safeLimit, safeOffset);
            return { table, columns, rows, limit: safeLimit, offset: safeOffset };
        } catch (error) {
            console.error('❌ [DB-VIEW] Erreur getDatabaseTablePreview:', error);
            return { table: String(tableName || ''), columns: [], rows: [], limit: 0, offset: 0, error: error.message };
        }
    }
}

module.exports = OptimizedDatabaseService;
// Export singleton
const optimizedDatabaseService = new OptimizedDatabaseService();
module.exports = optimizedDatabaseService;
