/**
 * Mail Monitor - Service de base de données
 * 
 * Copyright (c) 2025 Tanguy Raingeard. Tous droits réservés.
 * 
 * Service de base de données pour Mail Monitor
 * Gestion complète des emails et événements de monitoring
 */

const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

class DatabaseService {
  constructor() {
    this.db = null;
    this.dbPath = path.join(__dirname, '../../data/emails.db');
    this.isInitialized = false;
    this._isInitializing = false;
    this.appSettings = null;
  }

  /**
   * Charge la configuration de l'application depuis la base de données
   */
  async loadAppSettings() {
    if (this.appSettings) return this.appSettings;
    
    try {
      // S'assurer que la base de données est initialisée
      if (!this.isInitialized) {
        await this.initialize();
      }
      
      // Charger depuis la base de données
      const sql = 'SELECT config_key as key, config_value as value FROM app_config';
      const rows = await new Promise((resolve, reject) => {
        this.db.all(sql, [], (error, rows) => {
          if (error) reject(error);
          else resolve(rows || []);
        });
      });
      
      this.appSettings = {
        monitoring: {
          treatReadEmailsAsProcessed: false,
          scanInterval: 30000,
          autoStart: true
        }
      };
      
      // Appliquer les valeurs de la BDD
      rows.forEach(row => {
        if (row.key === 'treatReadEmailsAsProcessed') {
          this.appSettings.monitoring.treatReadEmailsAsProcessed = row.value === 'true';
        } else if (row.key === 'scanInterval') {
          this.appSettings.monitoring.scanInterval = parseInt(row.value) || 30000;
        } else if (row.key === 'autoStart') {
          this.appSettings.monitoring.autoStart = row.value === 'true';
        }
      });
      
    } catch (error) {
      console.warn('⚠️ Erreur chargement configuration BDD, utilisation des valeurs par défaut:', error.message);
      this.appSettings = {
        monitoring: {
          treatReadEmailsAsProcessed: false,
          scanInterval: 30000,
          autoStart: true
        }
      };
    }
    
    return this.appSettings;
  }

  /**
   * Sauvegarde un paramètre d'application dans la base de données
   */
  async saveAppConfig(configKey, configValue) {
    if (!this.isInitialized) {
      await this.initialize();
    }
    
    return new Promise((resolve, reject) => {
      const sql = `
        INSERT OR REPLACE INTO app_config 
        (config_key, config_value, config_type, description, updated_at) 
        VALUES (?, ?, 'json', 'Configuration application', CURRENT_TIMESTAMP)
      `;
      
      const value = typeof configValue === 'object' ? JSON.stringify(configValue) : String(configValue);
      
      this.db.run(sql, [configKey, value], function(error) {
        if (error) {
          console.error(`❌ Erreur sauvegarde config ${configKey}:`, error);
          reject(error);
        } else {
          console.log(`✅ Configuration sauvegardée: ${configKey} = ${value}`);
          resolve(this.changes);
        }
      });
    });
  }

  /**
   * S'assurer que la base de données est prête
   */
  async ensureDatabase() {
    if (!this.isInitialized) {
      await this.initialize();
    }
    return this.db;
  }

  /**
   * Initialise la connexion à la base de données
   */
  async initialize() {
    // Protection contre les multiples initialisations simultanées
    if (this.isInitialized) {
      // Log supprimé pour éviter le spam
      // console.log('✅ Base de données déjà initialisée - skip');
      return;
    }
    
    // Protection contre les initialisations parallèles
    if (this._isInitializing) {
      console.log('⏳ Initialisation en cours - attente...');
      // Attendre que l'initialisation en cours se termine
      while (this._isInitializing) {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
      return;
    }

    try {
      this._isInitializing = true;
      console.log('🔧 Initialisation de la base de données...');
      await this.connectStep();
      await this.createTablesStep();
      await this.createIndexesStep();
      
      // Corriger les chemins corrompus après l'initialisation
      await this.fixCorruptedPaths();
      
      // Nettoyer les doublons dans les configurations
      try {
        const duplicatesRemoved = await this.cleanupDuplicateFolders();
        if (duplicatesRemoved > 0) {
          console.log(`🧹 ${duplicatesRemoved} doublons supprimés`);
        }
      } catch (cleanupError) {
        console.log('⚠️ Erreur nettoyage doublons (ignorée):', cleanupError.message);
      }
      
      this.isInitialized = true;
      console.log('✅ Base de données initialisée avec succès');
    } catch (error) {
      console.error('❌ Erreur initialisation base de données:', error);
      throw error;
    } finally {
      this._isInitializing = false;
    }
  }

  /**
   * Étape de connexion
   */
  async connectStep() {
    return new Promise((resolve, reject) => {
      this.db = new sqlite3.Database(this.dbPath, (error) => {
        if (error) {
          console.error('❌ Erreur connexion SQLite:', error);
          reject(error);
        } else {
          console.log(`🔗 Connexion SQLite établie: ${this.dbPath}`);
          resolve();
        }
      });
    });
  }

  /**
   * Création des tables
   */
  async createTablesStep() {
    const createEmailsTable = `
      CREATE TABLE IF NOT EXISTS emails (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        outlook_id TEXT UNIQUE,
        subject TEXT NOT NULL,
        sender_name TEXT,
        sender_email TEXT,
        recipient_email TEXT,
        received_time DATETIME,
        sent_time DATETIME,
        treated_time DATETIME,
        folder_name TEXT,
        category TEXT,
        is_read BOOLEAN DEFAULT 0,
        is_replied BOOLEAN DEFAULT 0,
        has_attachment BOOLEAN DEFAULT 0,
        body_preview TEXT,
        importance INTEGER DEFAULT 1,
        size_kb INTEGER,
        event_type TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
      )
    `;

    const createEventsTable = `
      CREATE TABLE IF NOT EXISTS email_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email_id INTEGER,
        event_type TEXT NOT NULL,
        event_date DATETIME DEFAULT CURRENT_TIMESTAMP,
        details TEXT,
        FOREIGN KEY (email_id) REFERENCES emails (id)
      )
    `;

    const createFoldersTable = `
      CREATE TABLE IF NOT EXISTS monitored_folders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        folder_name TEXT UNIQUE NOT NULL,
        folder_path TEXT,
        category TEXT,
        is_active BOOLEAN DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
      )
    `;

    // Tables de configuration
    const createConfigTable = `
      CREATE TABLE IF NOT EXISTS app_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        config_key TEXT UNIQUE NOT NULL,
        config_value TEXT NOT NULL,
        config_type TEXT DEFAULT 'json',
        description TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
      )
    `;

    const createFolderConfigTable = `
      CREATE TABLE IF NOT EXISTS folder_configurations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        folder_path TEXT UNIQUE NOT NULL,
        category TEXT NOT NULL,
        folder_name TEXT,
        is_active BOOLEAN DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
      )
    `;

    const createMetricsHistoryTable = `
      CREATE TABLE IF NOT EXISTS metrics_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        metric_date DATE NOT NULL,
        metric_type TEXT NOT NULL,
        metric_data TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
      )
    `;

    return new Promise((resolve, reject) => {
      this.db.serialize(() => {
        this.db.run(createEmailsTable, (error) => {
          if (error) {
            console.error('❌ Erreur création table emails:', error);
            reject(error);
            return;
          }
        });

        this.db.run(createEventsTable, (error) => {
          if (error) {
            console.error('❌ Erreur création table events:', error);
            reject(error);
            return;
          }
        });

        this.db.run(createFoldersTable, (error) => {
          if (error) {
            console.error('❌ Erreur création table folders:', error);
            reject(error);
            return;
          }
        });

        this.db.run(createConfigTable, (error) => {
          if (error) {
            console.error('❌ Erreur création table config:', error);
            reject(error);
            return;
          }
        });

        this.db.run(createFolderConfigTable, (error) => {
          if (error) {
            console.error('❌ Erreur création table folder_config:', error);
            reject(error);
            return;
          }
        });

        this.db.run(createMetricsHistoryTable, (error) => {
          if (error) {
            console.error('❌ Erreur création table metrics_history:', error);
            reject(error);
            return;
          }
          console.log('✅ Tables créées avec succès');
          
          // Migration : ajouter treated_time aux bases existantes
          this.migrateAddTreatedTime(() => {
            // Migration : ajouter les nouvelles colonnes pour le service unifié
            this.migrateAddUnifiedColumns(async () => {
              // Migration : migrer les fichiers JSON vers la BD
              try {
                await this.migrateJsonToDatabase();
                resolve();
              } catch (error) {
                console.error('❌ Erreur migration JSON:', error);
                resolve(); // Continuer même en cas d'erreur de migration
              }
            });
          });
        });
      });
    });
  }

  /**
   * Migration pour ajouter la colonne treated_time
   */
  migrateAddTreatedTime(callback) {
    this.db.run('ALTER TABLE emails ADD COLUMN treated_time DATETIME', (error) => {
      if (error && !error.message.includes('duplicate column name')) {
        console.error('❌ Erreur migration treated_time:', error);
      } else {
        console.log('✅ Migration treated_time effectuée');
      }
      callback();
    });
  }

  /**
   * Migration pour ajouter les colonnes nécessaires au service unifié
   */
  migrateAddUnifiedColumns(callback) {
    const migrations = [
      // Ajouter les colonnes pour compatibilité avec le service unifié
      'ALTER TABLE emails ADD COLUMN entry_id TEXT',
      'ALTER TABLE emails ADD COLUMN sender TEXT',
      'ALTER TABLE emails ADD COLUMN size INTEGER DEFAULT 0',
      'ALTER TABLE emails ADD COLUMN folder_path TEXT',
      'ALTER TABLE emails ADD COLUMN folder_type TEXT',
      'ALTER TABLE emails ADD COLUMN is_treated BOOLEAN DEFAULT 0',
      'ALTER TABLE emails ADD COLUMN deleted_at DATETIME'
    ];

    let completed = 0;
    const total = migrations.length;

    migrations.forEach((migration) => {
      this.db.run(migration, (error) => {
        if (error && !error.message.includes('duplicate column name')) {
          console.error('❌ Erreur migration:', migration, error);
        }
        completed++;
        if (completed === total) {
          console.log('✅ Migrations colonnes unifiées effectuées');
          callback();
        }
      });
    });
  }

  /**
   * Migration des fichiers JSON vers la base de données
   */
  async migrateJsonToDatabase() {
    console.log('🔄 Migration des configurations JSON vers la base de données...');
    
    const result = {
      folders: false,
      settings: false,
      weekly: false,
      history: false
    };
    
    try {
      // Migration de la configuration des dossiers
      const foldersPath = path.join(__dirname, '../../data/folders-config.json');
      if (fs.existsSync(foldersPath)) {
        const foldersData = JSON.parse(fs.readFileSync(foldersPath, 'utf8'));
        await this.migrateFoldersConfigAsync(foldersData);
        result.folders = true;
        console.log('✅ Configuration dossiers migrée vers BD');
        
        // Sauvegarder le fichier JSON comme backup puis le supprimer
        try {
          const backupPath = foldersPath + '.backup';
          fs.copyFileSync(foldersPath, backupPath);
          fs.unlinkSync(foldersPath);
          console.log('✅ Fichier JSON folders-config supprimé (backup créé)');
        } catch (err) {
          console.log('⚠️ Impossible de supprimer le fichier JSON folders-config:', err.message);
        }
      }

      // Migration des paramètres
      const settingsPath = path.join(__dirname, '../../data/settings.json');
      if (fs.existsSync(settingsPath)) {
        const settingsData = JSON.parse(fs.readFileSync(settingsPath, 'utf8'));
        await this.migrateSettingsAsync(settingsData);
        result.settings = true;
        console.log('✅ Paramètres migrés vers BD');
        
        // Sauvegarder le fichier JSON comme backup puis le supprimer
        try {
          const backupPath = settingsPath + '.backup';
          fs.copyFileSync(settingsPath, backupPath);
          fs.unlinkSync(settingsPath);
          console.log('✅ Fichier JSON settings supprimé (backup créé)');
        } catch (err) {
          console.log('⚠️ Impossible de supprimer le fichier JSON settings:', err.message);
        }
      }

      // Migration des métriques hebdomadaires
      const weeklyPath = path.join(__dirname, '../../data/weekly-metrics.json');
      if (fs.existsSync(weeklyPath)) {
        const weeklyData = JSON.parse(fs.readFileSync(weeklyPath, 'utf8'));
        if (Array.isArray(weeklyData)) {
          for (const week of weeklyData) {
            await this.saveWeeklyMetrics(week);
          }
        }
        result.weekly = true;
        console.log('✅ Métriques hebdomadaires migrées vers BD');
        
        // Sauvegarder le fichier JSON comme backup puis le supprimer
        try {
          const backupPath = weeklyPath + '.backup';
          fs.copyFileSync(weeklyPath, backupPath);
          fs.unlinkSync(weeklyPath);
          console.log('✅ Fichier JSON weekly-metrics supprimé (backup créé)');
        } catch (err) {
          console.log('⚠️ Impossible de supprimer le fichier JSON weekly-metrics:', err.message);
        }
      }

      // Migration de l'historique
      const historyPath = path.join(__dirname, '../../data/metrics-history.json');
      if (fs.existsSync(historyPath)) {
        const historyData = JSON.parse(fs.readFileSync(historyPath, 'utf8'));
        if (Array.isArray(historyData)) {
          for (const entry of historyData) {
            await this.saveHistoricalData(entry);
          }
        }
        result.history = true;
        console.log('✅ Historique migré');
      }

      console.log('✅ Migration JSON vers BD terminée');
      return result;

    } catch (error) {
      console.error('❌ Erreur migration JSON:', error);
      throw error;
    }
  }

  // Versions asynchrones des méthodes de migration
  async migrateFoldersConfigAsync(data) {
    return new Promise((resolve, reject) => {
      this.migrateFoldersConfig(data, (err) => {
        if (err) reject(err);
        else resolve();
      });
    });
  }

  async migrateSettingsAsync(data) {
    return new Promise((resolve, reject) => {
      this.migrateSettings(data, (err) => {
        if (err) reject(err);
        else resolve();
      });
    });
  }

  /**
   * Migration de la configuration des dossiers
   */
  migrateFoldersConfig(data, callback) {
    if (!data.folderCategories) {
      callback();
      return;
    }

    const stmt = this.db.prepare(`
      INSERT OR REPLACE INTO folder_configurations 
      (folder_path, category, folder_name, is_active) 
      VALUES (?, ?, ?, 1)
    `);

    Object.entries(data.folderCategories).forEach(([folderPath, config]) => {
      stmt.run(folderPath, config.category, config.name || 'Unknown');
    });

    stmt.finalize((error) => {
      if (error) {
        console.error('❌ Erreur migration folders config:', error);
      } else {
        console.log('✅ Configuration dossiers migrée vers BD');
      }
      callback();
    });
  }

  /**
   * Migration des paramètres
   */
  migrateSettings(data, callback) {
    const stmt = this.db.prepare(`
      INSERT OR REPLACE INTO app_config 
      (config_key, config_value, config_type, description) 
      VALUES (?, ?, 'json', ?)
    `);

    Object.entries(data).forEach(([key, value]) => {
      stmt.run(key, JSON.stringify(value), `Configuration ${key}`);
    });

    stmt.finalize((error) => {
      if (error) {
        console.error('❌ Erreur migration settings:', error);
      } else {
        console.log('✅ Paramètres migrés vers BD');
      }
      callback();
    });
  }

  /**
   * Migration des métriques hebdomadaires
   */
  migrateWeeklyMetrics(data, callback) {
    if (!Array.isArray(data)) {
      callback();
      return;
    }

    const stmt = this.db.prepare(`
      INSERT OR REPLACE INTO metrics_history 
      (metric_date, metric_type, metric_data) 
      VALUES (?, 'weekly', ?)
    `);

    data.forEach((metric) => {
      const date = metric.timestamp ? new Date(metric.timestamp).toISOString().split('T')[0] : new Date().toISOString().split('T')[0];
      stmt.run(date, JSON.stringify(metric));
    });

    stmt.finalize((error) => {
      if (error) {
        console.error('❌ Erreur migration weekly metrics:', error);
      } else {
        console.log('✅ Métriques hebdomadaires migrées vers BD');
      }
      callback();
    });
  }

  /**
   * Création des index pour optimiser les performances
   */
  async createIndexesStep() {
    const indexes = [
      'CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails(received_time)',
      'CREATE INDEX IF NOT EXISTS idx_emails_treated_time ON emails(treated_time)',
      'CREATE INDEX IF NOT EXISTS idx_emails_sender_email ON emails(sender_email)',
      'CREATE INDEX IF NOT EXISTS idx_emails_folder_name ON emails(folder_name)',
      'CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category)',
      'CREATE INDEX IF NOT EXISTS idx_emails_is_read ON emails(is_read)',
      'CREATE INDEX IF NOT EXISTS idx_events_email_id ON email_events(email_id)',
      'CREATE INDEX IF NOT EXISTS idx_events_type ON email_events(event_type)',
      'CREATE INDEX IF NOT EXISTS idx_events_date ON email_events(event_date)'
    ];

    return new Promise((resolve, reject) => {
      let completed = 0;
      const total = indexes.length;

      indexes.forEach((indexSQL) => {
        this.db.run(indexSQL, (error) => {
          if (error) {
            console.error('❌ Erreur création index:', error);
            reject(error);
            return;
          }
          completed++;
          if (completed === total) {
            console.log('✅ Index créés avec succès');
            resolve();
          }
        });
      });
    });
  }

  /**
   * Statistiques pour la vue d'ensemble
   */
  async getEmailCountByDate(date) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT COUNT(*) as count 
        FROM emails 
        WHERE DATE(received_time) = ? AND event_type = 'received'
      `;
      
      this.db.get(sql, [date], (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row ? row.count : 0);
        }
      });
    });
  }

  async getTreatedEmailCountByDate(date) {
    // Protection d'init supprimée - la DB est déjà initialisée par l'application
    
    const settings = await this.loadAppSettings();
    const treatReadAsProcessed = settings.monitoring?.treatReadEmailsAsProcessed || false;
    
    return new Promise((resolve, reject) => {
      let sql;
      let params;
      
      if (treatReadAsProcessed) {
        // Mode souple : compter les emails traités OU lus
        sql = `
          SELECT COUNT(*) as count 
          FROM emails 
          WHERE DATE(treated_time) = ? 
            AND (event_type = 'treated' OR (event_type = 'read' AND is_read = 1))
        `;
        params = [date];
      } else {
        // Mode strict : compter seulement les emails explicitement traités (supprimés)
        sql = `
          SELECT COUNT(*) as count 
          FROM emails 
          WHERE DATE(treated_time) = ? AND event_type = 'treated'
        `;
        params = [date];
      }
      
      this.db.get(sql, params, (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row ? row.count : 0);
        }
      });
    });
  }

  async getSentEmailCountByDate(date) {
    // SUPPRIMÉ - Plus besoin de compter les mails envoyés
    // On compte maintenant les mails traités (supprimés)
    return this.getTreatedEmailCountByDate(date);
  }

  async getTotalEmailCount() {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = 'SELECT COUNT(*) as count FROM emails';
      
      this.db.get(sql, (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row ? row.count : 0);
        }
      });
    });
  }

  async getUnreadEmailCount() {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = 'SELECT COUNT(*) as count FROM emails WHERE is_read = 0';
      
      this.db.get(sql, (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row ? row.count : 0);
        }
      });
    });
  }

  /**
   * Récupère un email par son ID
   */
  async getEmailById(emailId) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = 'SELECT * FROM emails WHERE id = ?';
      
      this.db.get(sql, [emailId], (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row || null);
        }
      });
    });
  }

  /**
   * Récupère un email par son Entry ID ou Outlook ID
   */
  async getEmailByEntryId(entryId, folderPath = null) {
    return new Promise((resolve, reject) => {
      let sql = 'SELECT * FROM emails WHERE outlook_id = ? OR entry_id = ?';
      let params = [entryId, entryId];
      
      if (folderPath) {
        sql += ' AND folder_path = ?';
        params.push(folderPath);
      }
      
      this.db.get(sql, params, (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row || null);
        }
      });
    });
  }

  /**
   * Récupère tous les emails de la base
   */
  async getAllEmails() {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = 'SELECT * FROM emails ORDER BY received_time DESC';
      
      this.db.all(sql, (error, rows) => {
        if (error) {
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Récupère les emails depuis une date donnée
   */
  async getEmailsSince(cutoffDate) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = 'SELECT * FROM emails WHERE received_time >= ? ORDER BY received_time DESC';
      
      this.db.all(sql, [cutoffDate.toISOString()], (error, rows) => {
        if (error) {
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Met à jour un email avec de nouveaux champs
   */
  async updateEmail(emailId, updateFields) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const fields = Object.keys(updateFields);
      const values = Object.values(updateFields);
      
      const setClause = fields.map(field => `${field} = ?`).join(', ');
      const sql = `UPDATE emails SET ${setClause} WHERE id = ?`;
      
      values.push(emailId); // Ajouter l'ID à la fin
      
      this.db.run(sql, values, function(error) {
        if (error) {
          reject(error);
        } else {
          resolve(this.changes);
        }
      });
    });
  }

  async getAverageResponseTime() {
    // Protection d'init supprimée - la DB est déjà initialisée par l'application
    
    const settings = await this.loadAppSettings();
    const treatReadAsProcessed = settings.monitoring?.treatReadEmailsAsProcessed || false;
    
    return new Promise((resolve, reject) => {
      let sql;
      
      if (treatReadAsProcessed) {
        // Mode souple : calculer le temps entre réception et lecture/suppression
        sql = `
          SELECT AVG(
            (julianday(treated_time) - julianday(received_time)) * 24 * 60
          ) as avg_minutes
          FROM emails
          WHERE (event_type = 'treated' OR (event_type = 'read' AND is_read = 1))
          AND received_time IS NOT NULL
          AND treated_time IS NOT NULL
          AND treated_time > received_time
        `;
      } else {
        // Mode strict : calculer seulement le temps entre réception et suppression
        sql = `
          SELECT AVG(
            (julianday(treated_time) - julianday(received_time)) * 24 * 60
          ) as avg_minutes
          FROM emails
          WHERE event_type = 'treated'
          AND received_time IS NOT NULL
          AND treated_time IS NOT NULL
          AND treated_time > received_time
        `;
      }
      
      this.db.get(sql, (error, row) => {
        if (error) {
          reject(error);
        } else {
          // Retourner en heures avec 1 décimale
          const avgMinutes = row && row.avg_minutes ? row.avg_minutes : 0;
          const avgHours = avgMinutes / 60;
          resolve(avgHours > 0 ? avgHours.toFixed(1) : "0.0");
        }
      });
    });
  }

  /**
   * Récupération des emails récents (compatible avec l'ancienne API)
   */
  async getRecentEmails(limit = 20) {
    // ⚠️ SUPPRESSION de la vérification d'initialisation - gérée par le service appelant
    
    try {
      console.log(`📧 Récupération de ${limit} emails récents...`);
      
      const emails = await new Promise((resolve) => {
        const sql = `
          SELECT 
            id, subject, sender_name, sender_email, 
            folder_name, category, event_type,
            received_time, created_at, is_read
          FROM emails 
          WHERE deleted_at IS NULL
          ORDER BY COALESCE(received_time, created_at) DESC 
          LIMIT ?
        `;
        
        this.db.all(sql, [limit], (error, rows) => {
          if (error) {
            console.error('❌ Erreur récupération emails récents:', error);
            resolve([]);
          } else {
            console.log(`📧 ${rows.length} emails récents récupérés`);
            resolve(rows || []);
          }
        });
      });
      
      return emails;
      
    } catch (error) {
      console.error('❌ Erreur récupération emails récents:', error);
      return [];
    }
  }

  /**
   * Récupération des statistiques complètes
   */
  async getDatabaseStats() {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    try {
      const [totalCount, unreadCount, avgResponseTime] = await Promise.all([
        this.getTotalEmailCount(),
        this.getUnreadEmailCount(),
        this.getAverageResponseTime()
      ]);

      // Statistiques par jour (7 derniers jours)
      const dailyStats = [];
      for (let i = 6; i >= 0; i--) {
        const date = new Date();
        date.setDate(date.getDate() - i);
        const dateStr = date.toISOString().split('T')[0];
        
        const [received, treated] = await Promise.all([
          this.getEmailCountByDate(dateStr),
          this.getTreatedEmailCountByDate(dateStr)
        ]);
        
        dailyStats.push({
          date: dateStr,
          received,
          treated  // Changé de 'sent' à 'treated'
        });
      }

      return {
        summary: {
          totalEmails: totalCount,
          unreadEmails: unreadCount,
          averageResponseTime: avgResponseTime
        },
        dailyStats
      };
    } catch (error) {
      console.error('❌ Erreur récupération stats database:', error);
      return {
        summary: {
          totalEmails: 0,
          unreadEmails: 0,
          averageResponseTime: "0.0"
        },
        dailyStats: []
      };
    }
  }

  /**
   * Récupération de l'historique des emails
   */
  async getEmailHistory(limit = 50) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          id, subject, sender_name, sender_email, 
          folder_name, category, event_type,
          received_time, created_at, is_read
        FROM emails 
        ORDER BY created_at DESC 
        LIMIT ?
      `;
      
      this.db.all(sql, [limit], (error, rows) => {
        if (error) {
          console.error('❌ Erreur récupération historique:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Purge des anciens emails
   */
  async purgeOldEmails(daysToKeep = 30) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        DELETE FROM emails 
        WHERE created_at < datetime('now', '-' || ? || ' days')
      `;
      
      this.db.run(sql, [daysToKeep], function(error) {
        if (error) {
          console.error('❌ Erreur purge emails:', error);
          reject(error);
        } else {
          console.log(`🗑️ ${this.changes} emails supprimés (plus de ${daysToKeep} jours)`);
          resolve(this.changes);
        }
      });
    });
  }

  /**
   * Statistiques par catégorie
   */
  async getCategoryStats() {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          category,
          COUNT(*) as count,
          SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread
        FROM emails 
        WHERE category IS NOT NULL
        GROUP BY category
        ORDER BY count DESC
      `;
      
      this.db.all(sql, (error, rows) => {
        if (error) {
          console.error('❌ Erreur stats catégories:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Statistiques par dossier
   */
  async getFolderStats() {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          folder_name,
          COUNT(*) as count,
          SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread,
          MAX(received_time) as last_email
        FROM emails 
        WHERE folder_name IS NOT NULL
        GROUP BY folder_name
        ORDER BY count DESC
      `;
      
      this.db.all(sql, (error, rows) => {
        if (error) {
          console.error('❌ Erreur stats dossiers:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Vérifie si un email existe déjà dans la base de données
   */
  async emailExists(entryId, receivedTime) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT COUNT(*) as count 
        FROM emails 
        WHERE outlook_id = ? OR (subject = ? AND received_time = ?)
      `;
      
      this.db.get(sql, [entryId, entryId, receivedTime], (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row.count > 0);
        }
      });
    });
  }

  /**
   * Enregistre un email depuis Outlook dans la base de données
   */
  async saveEmailFromOutlook(emailData, category = 'general') {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    // Vérifier si l'email existe déjà
    const exists = await this.emailExists(emailData.EntryID, emailData.ReceivedTime);
    if (exists) {
      console.log(`📧 Email déjà existant: ${emailData.Subject}`);
      return null;
    }
    
    return new Promise((resolve, reject) => {
      const sql = `
        INSERT INTO emails (
          outlook_id, subject, sender_name, sender_email, recipient_email,
          received_time, size_kb, is_read, importance,
          folder_name, category, has_attachment,
          created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
      `;
      
      const params = [
        emailData.EntryID,
        emailData.Subject || '(Sans objet)',
        emailData.SenderName || 'Inconnu',
        emailData.SenderEmailAddress || '',
        emailData.RecipientEmailAddress || '',
        emailData.ReceivedTime,
        Math.round((emailData.Size || 0) / 1024), // Convertir en KB
        emailData.UnRead ? 0 : 1, // UnRead = true signifie is_read = 0
        emailData.Importance || 1,
        emailData.FolderName,
        category,
        emailData.HasAttachments || false
      ];
      
      this.db.run(sql, params, function(error) {
        if (error) {
          console.error('❌ Erreur enregistrement email Outlook:', error);
          reject(error);
        } else {
          console.log(`📝 Email enregistré: ${emailData.Subject} (ID: ${this.lastID})`);
          resolve(this.lastID);
        }
      });
    });
  }

  /**
   * Met à jour le statut d'un email existant (méthode unifiée)
   */
  async updateEmailStatus(outlookId, updateData) {
    // Protection d'init supprimée - la DB est déjà initialisée par l'application
    
    return new Promise((resolve, reject) => {
      // Gestion des deux signatures: updateEmailStatus(id, boolean) ou updateEmailStatus(id, object)
      let updateFields = [];
      let values = [];
      
      if (typeof updateData === 'boolean') {
        // Ancienne signature: updateEmailStatus(outlookId, isRead)
        updateFields.push('is_read = ?');
        values.push(updateData ? 1 : 0);
      } else if (typeof updateData === 'object' && updateData !== null) {
        // Nouvelle signature: updateEmailStatus(entryId, statusUpdate)
        if (updateData.is_read !== undefined) {
          updateFields.push('is_read = ?');
          values.push(updateData.is_read ? 1 : 0);
        }
        
        if (updateData.size !== undefined) {
          updateFields.push('size = ?');
          values.push(updateData.size);
        }
        
        if (updateData.folder_path !== undefined) {
          updateFields.push('folder_path = ?');
          values.push(updateData.folder_path);
        }
        
        if (updateData.size_kb !== undefined) {
          updateFields.push('size_kb = ?');
          values.push(updateData.size_kb);
        }
      }
      
      if (updateFields.length === 0) {
        resolve(0);
        return;
      }
      
      updateFields.push('updated_at = CURRENT_TIMESTAMP');
      values.push(outlookId);
      
      const sql = `UPDATE emails SET ${updateFields.join(', ')} WHERE outlook_id = ?`;
      
      this.db.run(sql, values, function(error) {
        if (error) {
          console.error('❌ Erreur mise à jour statut email:', error);
          reject(error);
        } else {
          if (this.changes > 0) {
            console.log(`📝 Statut email mis à jour: ${outlookId} (${this.changes} changements)`);
          }
          resolve(this.changes);
        }
      });
    });
  }

  /**
   * Traite et enregistre une liste d'emails depuis Outlook (pour scan initial seulement)
   */
  async processFolderEmails(folderEmails, category = 'general') {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    let processed = 0;
    let skipped = 0;
    let errors = 0;

    console.log(`📥 Traitement de ${folderEmails.Emails.length} emails du dossier ${folderEmails.FolderName}`);

    for (const email of folderEmails.Emails) {
      try {
        const result = await this.saveEmailFromOutlook(email, category);
        if (result) {
          processed++;
        } else {
          skipped++;
        }
      } catch (error) {
        console.error(`❌ Erreur traitement email ${email.Subject}:`, error);
        errors++;
      }
    }

    console.log(`✅ Traitement terminé: ${processed} nouveaux, ${skipped} ignorés, ${errors} erreurs`);

    return {
      total: folderEmails.Emails.length,
      processed,
      skipped,
      errors,
      folderName: folderEmails.FolderName,
      folderPath: folderEmails.FolderPath
    };
  }

  /**
   * NOUVEAU: Traite les changements d'emails détectés par le polling intelligent
   */
  async processPollingEmailChange(emailUpdateData) {
    try {
      console.log(`🔄 [DATABASE] Traitement changement polling: ${emailUpdateData.subject}`);
      
      // Vérifier d'abord si l'email existe en base
      const existingEmail = await this.getEmailByOutlookId(emailUpdateData.messageId);
      
      if (!existingEmail) {
        console.log(`⚠️ [DATABASE] Email non trouvé en base: ${emailUpdateData.messageId} - ${emailUpdateData.subject}`);
        // Si l'email n'existe pas, on pourrait le créer ici ou l'ignorer
        return { updated: false, reason: 'Email non trouvé en base' };
      }

      // Préparer les données de mise à jour
      const updateData = {};
      
      // Mise à jour du statut lu/non lu si c'est le changement détecté
      if (emailUpdateData.changes && emailUpdateData.changes.includes('MarkedRead')) {
        updateData.is_read = true;
        console.log(`📖 [DATABASE] Marquage comme lu: ${emailUpdateData.subject}`);
      } else if (emailUpdateData.changes && emailUpdateData.changes.includes('MarkedUnread')) {
        updateData.is_read = false;
        console.log(`📬 [DATABASE] Marquage comme non lu: ${emailUpdateData.subject}`);
      } else if (emailUpdateData.isRead !== undefined) {
        updateData.is_read = emailUpdateData.isRead;
        console.log(`📝 [DATABASE] Mise à jour statut lecture: ${emailUpdateData.subject} -> ${emailUpdateData.isRead ? 'Lu' : 'Non lu'}`);
      }

      // Si aucune modification détectée, pas de mise à jour nécessaire
      if (Object.keys(updateData).length === 0) {
        console.log(`ℹ️ [DATABASE] Aucune modification à appliquer: ${emailUpdateData.subject}`);
        return { updated: false, reason: 'Aucune modification détectée' };
      }

      // Effectuer la mise à jour
      const changes = await this.updateEmailStatus(emailUpdateData.messageId, updateData);
      
      if (changes > 0) {
        console.log(`✅ [DATABASE] Email mis à jour: ${emailUpdateData.subject} (${changes} changements)`);
        return { 
          updated: true, 
          changes: changes,
          emailData: { ...existingEmail, ...updateData }
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
   * Synchronise les emails d'un dossier avec la base de données
   */
  async syncFolderEmails(folderEmails, category = 'general') {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    let processed = 0;
    let updated = 0;
    let skipped = 0;
    let errors = 0;

    console.log(`� Synchronisation de ${folderEmails.Emails.length} emails du dossier ${folderEmails.FolderName}`);

    for (const email of folderEmails.Emails) {
      try {
        // Vérifier si l'email existe déjà
        const exists = await this.emailExists(email.EntryID, email.ReceivedTime);
        
        if (exists) {
          // Mettre à jour le statut si l'email existe
          const changes = await this.updateEmailStatus(email.EntryID, !email.UnRead);
          if (changes > 0) {
            updated++;
          } else {
            skipped++;
          }
        } else {
          // Créer un nouvel email
          const result = await this.saveEmailFromOutlook(email, category);
          if (result) {
            processed++;
          } else {
            skipped++;
          }
        }
      } catch (error) {
        console.error(`❌ Erreur synchronisation email ${email.Subject}:`, error);
        errors++;
      }
    }

    console.log(`✅ Synchronisation terminée: ${processed} nouveaux, ${updated} mis à jour, ${skipped} inchangés, ${errors} erreurs`);

    return {
      total: folderEmails.Emails.length,
      processed,
      updated,
      skipped,
      errors,
      folderName: folderEmails.FolderName,
      folderPath: folderEmails.FolderPath
    };
  }

  /**
   * Traite un scan complet d'un dossier (premier monitoring)
   */
  async processFullFolderScan(folderEmails, category = 'general') {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    let processed = 0;
    let updated = 0;
    let skipped = 0;
    let errors = 0;

    console.log(`📊 Scan complet: traitement de ${folderEmails.Emails.length} emails du dossier ${folderEmails.FolderName}`);

    for (const email of folderEmails.Emails) {
      try {
        // Vérifier si l'email existe déjà
        const exists = await this.emailExists(email.EntryID, email.ReceivedTime);
        
        if (exists) {
          // Mettre à jour le statut de l'email existant
          const changes = await this.updateEmailStatus(email.EntryID, !email.UnRead);
          if (changes > 0) {
            updated++;
            console.log(`🔄 Email mis à jour: ${email.Subject}`);
          } else {
            skipped++;
            console.log(`📧 Email déjà existant: ${email.Subject}`);
          }
        } else {
          // Créer un nouvel email avec statut complet
          const result = await this.saveEmailFromOutlook(email, category);
          if (result) {
            processed++;
            console.log(`📧 Nouvel email: ${email.Subject}`);
          } else {
            skipped++;
          }
        }
      } catch (error) {
        console.error(`❌ Erreur traitement email ${email.Subject}:`, error);
        errors++;
      }
    }

    console.log(`✅ Scan complet terminé: ${processed} nouveaux, ${updated} mis à jour, ${skipped} ignorés, ${errors} erreurs`);

    return {
      total: folderEmails.Emails.length,
      processed,
      updated,
      skipped,
      errors,
      folderName: folderEmails.FolderName,
      folderPath: folderEmails.FolderPath
    };
  }

  /**
   * Synchronise les changements incrémentaux (monitoring de routine) avec événements complets
   */
  async syncIncrementalChanges(folderEmails, category = 'general') {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    let newEmails = 0;
    let statusChanges = 0;
    let treatedEmails = 0;
    let errors = 0;

    console.log(`🔄 Synchronisation incrémentale: ${folderEmails.Emails.length} emails récents du dossier ${folderEmails.FolderName}`);

    // 1. Traiter les emails présents dans Outlook avec logging complet
    const currentOutlookIds = new Set();
    
    for (const email of folderEmails.Emails) {
      try {
        currentOutlookIds.add(email.EntryID);
        
        // Utiliser la nouvelle méthode avec événements
        const emailId = await this.saveEmailFromOutlookWithEvents(email, category);
        
        if (emailId) {
          // Vérifier si c'est un nouvel email en regardant sa date de création
          const emailRecord = await this.getEmailByOutlookId(email.EntryID);
          const createdRecently = emailRecord && new Date(emailRecord.created_at) > new Date(Date.now() - 60000); // Créé il y a moins d'1 minute
          
          if (createdRecently) {
            newEmails++;
            console.log(`📧 Nouvel email détecté: ${email.Subject}`);
          } else {
            statusChanges++;
            console.log(`� Email mis à jour: ${email.Subject}`);
          }
        }
      } catch (error) {
        console.error(`❌ Erreur synchronisation email ${email.Subject}:`, error);
        errors++;
      }
    }

    // 2. Détecter les emails supprimés (traités) de ce dossier
    try {
      const treatedCount = await this.detectTreatedEmails(folderEmails.FolderName, currentOutlookIds);
      treatedEmails = treatedCount;
      
      if (treatedEmails > 0) {
        console.log(`🗑️ ${treatedEmails} emails traités (supprimés) détectés dans ${folderEmails.FolderName}`);
      }
    } catch (error) {
      console.error('❌ Erreur détection emails traités:', error);
      errors++;
    }

    const totalChanges = newEmails + statusChanges + treatedEmails;
    
    if (totalChanges > 0) {
      console.log(`✅ Synchronisation terminée: ${newEmails} nouveaux emails, ${statusChanges} changements de statut`);
    } else {
      console.log(`✅ Aucun changement détecté dans ${folderEmails.FolderName}`);
    }

    return {
      total: folderEmails.Emails.length,
      newEmails,
      statusChanges,
      changes: totalChanges,
      errors,
      folderName: folderEmails.FolderName,
      folderPath: folderEmails.FolderPath
    };
  }

  /**
   * Détecte les emails qui ont été supprimés/traités
   */
  async detectTreatedEmails(folderName, currentOutlookIds) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      // Récupérer tous les emails de ce dossier qui ne sont pas encore marqués comme traités
      const sql = `
        SELECT id, outlook_id, subject, received_time
        FROM emails 
        WHERE folder_name = ? 
        AND event_type = 'received' 
        AND treated_time IS NULL
        AND outlook_id IS NOT NULL
      `;
      
      this.db.all(sql, [folderName], (error, rows) => {
        if (error) {
          reject(error);
          return;
        }
        
        let treatedCount = 0;
        const treatedPromises = [];
        
        // Vérifier chaque email stocké
        for (const row of rows) {
          // Si l'email n'est plus présent dans Outlook, il a été traité (supprimé)
          if (!currentOutlookIds.has(row.outlook_id)) {
            const promise = this.markEmailAsTreated(row.id, row.subject);
            treatedPromises.push(promise);
            treatedCount++;
          }
        }
        
        // Attendre que tous les emails soient marqués comme traités
        Promise.all(treatedPromises)
          .then(() => {
            resolve(treatedCount);
          })
          .catch(reject);
      });
    });
  }

  /**
   * Marque un email comme traité (méthode unifiée)
   */
  async markEmailAsTreated(identifier, subject = null) {
    // Protection d'init supprimée - la DB est déjà initialisée par l'application
    
    return new Promise((resolve, reject) => {
      let sql, params;
      
      // Déterminer si c'est un ID numérique (BDD) ou un EntryID (string)
      if (typeof identifier === 'number' || /^\d+$/.test(identifier)) {
        // Mise à jour par ID de base de données
        sql = `
          UPDATE emails 
          SET treated_time = CURRENT_TIMESTAMP, 
              event_type = 'treated',
              is_treated = 1,
              updated_at = CURRENT_TIMESTAMP
          WHERE id = ?
        `;
        params = [identifier];
      } else {
        // Mise à jour par Outlook ID / Entry ID
        sql = `
          UPDATE emails 
          SET treated_time = CURRENT_TIMESTAMP, 
              event_type = 'treated',
              is_treated = 1,
              updated_at = CURRENT_TIMESTAMP
          WHERE outlook_id = ?
        `;
        params = [identifier];
      }
      
      this.db.run(sql, params, function(error) {
        if (error) {
          console.error(`❌ Erreur marquage email traité ${subject || identifier}:`, error);
          reject(error);
        } else {
          if (subject) {
            console.log(`✅ Email marqué comme traité: ${subject}`);
          } else {
            console.log(`✅ Email marqué comme traité (ID: ${identifier})`);
          }
          resolve(this.changes);
        }
      });
    });
  }

  /**
   * Enregistre une activité email pour le monitoring
   */
  async logEmailActivity(emailData) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        INSERT INTO emails (
          subject, folder_name, category, event_type, 
          received_time, created_at
        ) VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
      `;
      
      const params = [
        `Email surveillé dans ${emailData.folderName}`,
        emailData.folderName,
        emailData.category,
        emailData.action,
        emailData.timestamp.toISOString()
      ];
      
      this.db.run(sql, params, function(error) {
        if (error) {
          console.error('❌ Erreur enregistrement activité email:', error);
          reject(error);
        } else {
          console.log(`📝 Activité email enregistrée: ${emailData.action} dans ${emailData.folderName}`);
          resolve(this.lastID);
        }
      });
    });
  }

  /**
   * SYSTÈME COMPLET DE LOGGING DES ÉVÉNEMENTS EMAIL
   */

  /**
   * Enregistre un événement email dans la table email_events
   */
  async logEmailEvent(emailId, eventType, details = null) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        INSERT INTO email_events (email_id, event_type, event_date, details)
        VALUES (?, ?, CURRENT_TIMESTAMP, ?)
      `;
      
      this.db.run(sql, [emailId, eventType, details], function(error) {
        if (error) {
          console.error(`❌ Erreur log événement ${eventType}:`, error);
          reject(error);
        } else {
          console.log(`📝 Événement logué: ${eventType} pour email ID ${emailId}`);
          resolve(this.lastID);
        }
      });
    });
  }

  /**
   * Sauvegarde un email depuis Outlook avec logging complet des événements
   */
  async saveEmailFromOutlookWithEvents(outlookEmail, category = 'general') {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    try {
      // 1. Vérifier si l'email existe déjà
      const existingEmail = await this.getEmailByOutlookId(outlookEmail.EntryID);
      
      if (existingEmail) {
        // Email existant - vérifier les changements d'état
        await this.updateEmailStateWithEvents(existingEmail, outlookEmail);
        return existingEmail.id;
      } else {
        // Nouvel email - l'insérer avec événement "arrived"
        const emailId = await this.insertNewEmailWithEvent(outlookEmail, category);
        return emailId;
      }
    } catch (error) {
      console.error('❌ Erreur sauvegarde email avec événements:', error);
      throw error;
    }
  }

  /**
   * Récupère un email par son ID Outlook
   */
  async getEmailByOutlookId(outlookId) {
    // Protection d'init supprimée - la DB est déjà initialisée par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT * FROM emails WHERE outlook_id = ?
      `;
      
      this.db.get(sql, [outlookId], (error, row) => {
        if (error) {
          reject(error);
        } else {
          resolve(row || null);
        }
      });
    });
  }

  /**
   * Alias pour compatibilité avec le nouveau service
   */
  async getEmailByEntryId(entryId) {
    return this.getEmailByOutlookId(entryId);
  }

  /**
   * Insère un nouvel email dans la base de données
   */
  async insertEmail(emailRecord) {
    // Protection d'init supprimée - la DB est déjà initialisée par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        INSERT INTO emails (
          outlook_id, subject, sender_name, sender_email, recipient_email,
          received_time, is_read, size_kb, folder_name, folder_type, 
          is_treated, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `;
      
      const values = [
        emailRecord.entry_id || emailRecord.outlook_id,
        emailRecord.subject,
        emailRecord.sender_name || emailRecord.sender,
        emailRecord.sender_email || emailRecord.sender,
        emailRecord.recipient_email || emailRecord.recipient || '',
        emailRecord.received_time,
        emailRecord.is_read ? 1 : 0,
        emailRecord.size,
        emailRecord.folder_name || emailRecord.folder_path,
        emailRecord.folder_type,
        emailRecord.is_treated ? 1 : 0,
        emailRecord.created_at
      ];
      
      this.db.run(sql, values, function(error) {
        if (error) {
          reject(error);
        } else {
          resolve({ id: this.lastID, ...emailRecord });
        }
      });
    });
  }

  /**
   * Marque un email comme supprimé
   */
  async markEmailAsDeleted(outlookId) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        UPDATE emails 
        SET is_deleted = 1, deleted_at = CURRENT_TIMESTAMP 
        WHERE outlook_id = ?
      `;
      
      this.db.run(sql, [outlookId], function(error) {
        if (error) {
          reject(error);
        } else {
          resolve({ changes: this.changes });
        }
      });
    });
  }

  /**
   * Insérer ou mettre à jour les statistiques hebdomadaires
   */
  async insertOrUpdateWeeklyStats(weeklyRecord) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      // D'abord, créer la table si elle n'existe pas
      const createTableSql = `
        CREATE TABLE IF NOT EXISTS weekly_stats (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          week_identifier TEXT UNIQUE NOT NULL,
          week_number INTEGER,
          year INTEGER,
          arrivals_count INTEGER DEFAULT 0,
          treatments_count INTEGER DEFAULT 0,
          stock_debut INTEGER DEFAULT 0,
          stock_fin INTEGER DEFAULT 0,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
      `;
      
      this.db.run(createTableSql, (createError) => {
        if (createError && !createError.message.includes('already exists')) {
          reject(createError);
          return;
        }
        
        // Insérer ou mettre à jour
        const upsertSql = `
          INSERT OR REPLACE INTO weekly_stats (
            week_identifier, week_number, year, arrivals_count, 
            treatments_count, stock_debut, stock_fin, updated_at
          ) VALUES (?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        `;
        
        this.db.run(upsertSql, [
          weeklyRecord.week_identifier,
          weeklyRecord.week_number,
          weeklyRecord.year,
          weeklyRecord.arrivals_count,
          weeklyRecord.treatments_count,
          weeklyRecord.stock_debut,
          weeklyRecord.stock_fin
        ], function(error) {
          if (error) {
            reject(error);
          } else {
            resolve(this.lastID);
          }
        });
      });
    });
  }

  /**
   * Récupérer les statistiques hebdomadaires
   */
  async getWeeklyStats(limit = 10) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT * FROM weekly_stats 
        ORDER BY year DESC, week_number DESC 
        LIMIT ?
      `;
      
      this.db.all(sql, [limit], (error, rows) => {
        if (error) {
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * S'assurer que la table weekly_stats existe
   */
  async ensureWeeklyStatsTable() {
    return new Promise((resolve, reject) => {
      const createTableSql = `
        CREATE TABLE IF NOT EXISTS weekly_stats (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          week_identifier TEXT UNIQUE NOT NULL,
          week_number INTEGER,
          year INTEGER,
          arrivals_count INTEGER DEFAULT 0,
          treatments_count INTEGER DEFAULT 0,
          stock_debut INTEGER DEFAULT 0,
          stock_fin INTEGER DEFAULT 0,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
      `;
      
      this.db.run(createTableSql, (error) => {
        if (error && !error.message.includes('already exists')) {
          console.error('❌ Erreur création table weekly_stats:', error);
          reject(error);
        } else {
          console.log('✅ Table weekly_stats créée/vérifiée avec succès');
          resolve();
        }
      });
    });
  }

  /**
   * Insère un nouvel email avec événement "arrived"
   */
  async insertNewEmailWithEvent(outlookEmail, category) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    const sql = `
      INSERT INTO emails (
        outlook_id, subject, sender_name, sender_email, recipient_email,
        received_time, folder_name, category, is_read, 
        has_attachment, body_preview, importance, event_type,
        created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
    `;
    
    const params = [
      outlookEmail.EntryID,
      outlookEmail.Subject || 'Sans objet',
      outlookEmail.SenderName || 'Inconnu',
      outlookEmail.SenderEmailAddress || '',
      outlookEmail.RecipientEmailAddress || '',
      outlookEmail.ReceivedTime ? new Date(outlookEmail.ReceivedTime).toISOString() : null,
      outlookEmail.Parent?.Name || 'Boîte de réception',
      category,
      !outlookEmail.UnRead, // is_read = inverse de UnRead
      outlookEmail.Attachments?.Count > 0 || false,
      outlookEmail.Body ? outlookEmail.Body.substring(0, 500) : '',
      outlookEmail.Importance || 1,
      'received'
    ];
    
    try {
      // Insérer l'email
      const emailId = await new Promise((resolve, reject) => {
        this.db.run(sql, params, function(error) {
          if (error) {
            reject(error);
          } else {
            resolve(this.lastID);
          }
        });
      });
      
      // Logger l'événement "arrived"
      await this.logEmailEvent(emailId, 'arrived', `Email reçu dans ${outlookEmail.Parent?.Name || 'Boîte de réception'}`);
      
      console.log(`📧 Nouvel email inséré avec ID ${emailId}: ${outlookEmail.Subject}`);
      return emailId;
      
    } catch (error) {
      console.error('❌ Erreur insertion nouvel email:', error);
      throw error;
    }
  }

  /**
   * Met à jour les états d'un email avec logging des événements
   */
  async updateEmailStateWithEvents(existingEmail, outlookEmail) {
    const changes = [];
    const events = [];
    const settings = await this.loadAppSettings();
    const treatReadAsProcessed = settings.monitoring?.treatReadEmailsAsProcessed || false;
    
    // Détecter les changements d'état
    const newIsRead = !outlookEmail.UnRead;
    const newFolderName = outlookEmail.Parent?.Name || 'Boîte de réception';
    
    // Changement de statut lu/non lu
    if (existingEmail.is_read !== newIsRead) {
      changes.push(`is_read = ${newIsRead ? 1 : 0}`);
      
      if (newIsRead) {
        events.push({
          type: 'read',
          details: `Email marqué comme lu`
        });
        
        // Si le paramètre est activé, marquer comme traité quand lu
        if (treatReadAsProcessed && !existingEmail.treated_time) {
          changes.push(`treated_time = CURRENT_TIMESTAMP`);
          changes.push(`event_type = 'read'`);
          events.push({
            type: 'processed_by_read',
            details: `Email considéré comme traité (paramètre: emails lus = traités)`
          });
          console.log(`📖 Email "${outlookEmail.Subject}" marqué comme traité (lu)`);
        }
      } else {
        events.push({
          type: 'unread',
          details: `Email marqué comme non lu`
        });
      }
    }
    
    // Changement de dossier
    if (existingEmail.folder_name !== newFolderName) {
      changes.push(`folder_name = '${newFolderName.replace(/'/g, "''")}'`);
      events.push({
        type: 'moved',
        details: `Email déplacé de '${existingEmail.folder_name}' vers '${newFolderName}'`
      });
    }
    
    // Appliquer les changements s'il y en a
    if (changes.length > 0) {
      changes.push('updated_at = CURRENT_TIMESTAMP');
      
      await new Promise((resolve, reject) => {
        const sql = `UPDATE emails SET ${changes.join(', ')} WHERE id = ?`;
        
        this.db.run(sql, [existingEmail.id], (error) => {
          if (error) {
            console.error('❌ Erreur mise à jour email:', error);
            reject(error);
          } else {
            resolve();
          }
        });
      });
      
      // Logger tous les événements
      for (const event of events) {
        try {
          await this.logEmailEvent(existingEmail.id, event.type, event.details);
        } catch (error) {
          console.error(`❌ Erreur log événement ${event.type}:`, error);
        }
      }
      
      console.log(`📝 Email ID ${existingEmail.id} mis à jour avec ${events.length} événement(s)`);
    }
  }

  /**
   * Marque un email comme traité (supprimé) avec événement
   */
  async markEmailAsTreated(emailId, subject) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    const sql = `
      UPDATE emails 
      SET treated_time = CURRENT_TIMESTAMP, 
          event_type = 'treated',
          updated_at = CURRENT_TIMESTAMP
      WHERE id = ?
    `;
    
    try {
      // Mettre à jour l'email
      const changes = await new Promise((resolve, reject) => {
        this.db.run(sql, [emailId], function(error) {
          if (error) {
            reject(error);
          } else {
            resolve(this.changes);
          }
        });
      });
      
      // Logger l'événement "treated"
      await this.logEmailEvent(emailId, 'treated', `Email supprimé/traité: ${subject}`);
      
      console.log(`✅ Email marqué comme traité: ${subject}`);
      return changes;
      
    } catch (error) {
      console.error(`❌ Erreur marquage email traité ${subject}:`, error);
      throw error;
    }
  }

  /**
   * Récupère l'historique des événements pour un email
   */
  async getEmailEvents(emailId) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          ee.event_type,
          ee.event_date,
          ee.details,
          e.subject,
          e.folder_name
        FROM email_events ee
        JOIN emails e ON ee.email_id = e.id
        WHERE ee.email_id = ?
        ORDER BY ee.event_date DESC
      `;
      
      this.db.all(sql, [emailId], (error, rows) => {
        if (error) {
          console.error('❌ Erreur récupération événements email:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Récupère les statistiques d'événements
   */
  async getEventStats(days = 7) {
    // Protection d'init supprim�e - la DB est d�j� initialis�e par l'application
    
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          event_type,
          COUNT(*) as count,
          DATE(event_date) as date
        FROM email_events 
        WHERE event_date >= datetime('now', '-${days} days')
        GROUP BY event_type, DATE(event_date)
        ORDER BY event_date DESC, event_type
      `;
      
      this.db.all(sql, (error, rows) => {
        if (error) {
          console.error('❌ Erreur stats événements:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Récupère les statistiques des emails
   */
  async getEmailStats() {
    // ⚠️ SUPPRESSION de la vérification d'initialisation - gérée par le service appelant
    
    try {
      // Version simplifiée sans promesse complexe
      console.log('📊 Récupération stats emails (version simplifiée)...');
      
      const today = new Date().toISOString().split('T')[0];
      
      // Requête simple count total
      const totalResult = await new Promise((resolve) => {
        this.db.get('SELECT COUNT(*) as count FROM emails WHERE deleted_at IS NULL', (err, row) => {
          resolve(err ? { count: 0 } : row);
        });
      });
      
      // Requête simple count non lus
      const unreadResult = await new Promise((resolve) => {
        this.db.get('SELECT COUNT(*) as count FROM emails WHERE is_read = 0 AND deleted_at IS NULL', (err, row) => {
          resolve(err ? { count: 0 } : row);
        });
      });
      
      // Requête simple count aujourd'hui
      const todayResult = await new Promise((resolve) => {
        this.db.get(`SELECT COUNT(*) as count FROM emails WHERE DATE(received_time) = ? AND deleted_at IS NULL`, [today], (err, row) => {
          resolve(err ? { count: 0 } : row);
        });
      });
      
      // Requête simple count traités aujourd'hui
      const treatedResult = await new Promise((resolve) => {
        this.db.get(`SELECT COUNT(*) as count FROM emails WHERE is_treated = 1 AND DATE(treated_time) = ? AND deleted_at IS NULL`, [today], (err, row) => {
          resolve(err ? { count: 0 } : row);
        });
      });
      
      const stats = {
        totalEmails: totalResult.count || 0,
        unreadTotal: unreadResult.count || 0,
        emailsToday: todayResult.count || 0,
        treatedToday: treatedResult.count || 0
      };
      
      console.log('📊 Stats simplifiées calculées:', stats);
      return stats;
      
    } catch (error) {
      console.error('❌ Erreur récupération stats emails:', error);
      return { 
        totalEmails: 0, 
        unreadTotal: 0, 
        emailsToday: 0, 
        treatedToday: 0 
      };
    }
  }

  // ==================== MÉTHODES DE CONFIGURATION ====================

  /**
   * Récupérer la configuration des dossiers
   */
  async getFoldersConfiguration() {
    return new Promise((resolve, reject) => {
      // Vérifier que la base de données est initialisée
      if (!this.db) {
        console.error('❌ Base de données non initialisée pour getFoldersConfiguration');
        resolve([]); // Retourner un tableau vide plutôt qu'erreur
        return;
      }

      const sql = `
        SELECT folder_path, category, folder_name, is_active 
        FROM folder_configurations 
        WHERE is_active = 1
        ORDER BY rowid DESC
      `;
      
      console.log('🔍 Exécution de la requête getFoldersConfiguration...');
      
      this.db.all(sql, [], (error, rows) => {
        if (error) {
          console.error('❌ Erreur lecture config dossiers:', error);
          resolve([]); // Retourner un tableau vide plutôt qu'erreur
        } else {
          console.log(`✅ ${rows.length} configurations trouvées`);
          
          // Éliminer les doublons en gardant la dernière configuration par chemin
          const uniqueConfigs = new Map();
          
          rows.forEach(row => {
            let cleanPath = row.folder_path;
            if (cleanPath.startsWith('\\\\\\\\')) {
              cleanPath = cleanPath.substring(2);
            } else if (cleanPath.startsWith('\\\\')) {
              cleanPath = cleanPath.substring(2);
            }
            
            // Ne garder que la première occurrence (la plus récente grâce à ORDER BY rowid DESC)
            if (!uniqueConfigs.has(cleanPath)) {
              uniqueConfigs.set(cleanPath, {
                path: cleanPath,
                name: row.folder_name,
                category: row.category,
                enabled: row.is_active === 1
              });
            }
          });
          
          const configArray = Array.from(uniqueConfigs.values());
          console.log(`🔍 Configuration nettoyée: ${configArray.length} dossiers uniques`);
          resolve(configArray);
        }
      });
    });
  }

  /**
   * Nettoyer les doublons dans la table folder_configurations
   */
  async cleanupDuplicateFolders() {
    return new Promise((resolve, reject) => {
      console.log('🧹 Nettoyage des doublons dans folder_configurations...');
      
      // Garder seulement la configuration la plus récente pour chaque chemin
      const cleanupSql = `
        DELETE FROM folder_configurations 
        WHERE rowid NOT IN (
          SELECT MAX(rowid) 
          FROM folder_configurations 
          GROUP BY folder_path
        )
      `;
      
      this.db.run(cleanupSql, [], function(error) {
        if (error) {
          console.error('❌ Erreur nettoyage doublons:', error);
          reject(error);
        } else {
          console.log(`✅ ${this.changes} doublons supprimés`);
          resolve(this.changes);
        }
      });
    });
  }

  /**
   * Corriger les chemins corrompus dans la base de données
   */
  async fixCorruptedPaths() {
    return new Promise((resolve, reject) => {
      console.log('🔧 Correction des chemins corrompus...');
      
      // Trouver les chemins qui commencent par \\\\
      const findSql = "SELECT rowid, folder_path FROM folder_configurations WHERE folder_path LIKE '\\\\\\\\%'";
      
      this.db.all(findSql, [], (error, rows) => {
        if (error) {
          console.error('❌ Erreur recherche chemins corrompus:', error);
          reject(error);
          return;
        }
        
        if (rows.length === 0) {
          console.log('✅ Aucun chemin corrompu trouvé');
          resolve();
          return;
        }
        
        console.log(`🔧 ${rows.length} chemins corrompus trouvés, correction...`);
        
        // Corriger chaque chemin
        const updateStmt = this.db.prepare("UPDATE folder_configurations SET folder_path = ? WHERE rowid = ?");
        
        rows.forEach(row => {
          const cleanPath = row.folder_path.substring(2); // Supprimer les 2 premiers \\
          console.log(`🔧 Correction: ${row.folder_path} → ${cleanPath}`);
          updateStmt.run(cleanPath, row.rowid);
        });
        
        updateStmt.finalize((finalizeError) => {
          if (finalizeError) {
            console.error('❌ Erreur finalisation correction chemins:', finalizeError);
            reject(finalizeError);
          } else {
            console.log('✅ Chemins corrompus corrigés');
            resolve();
          }
        });
      });
    });
  }

  /**
   * Sauvegarder la configuration des dossiers
   */
  async saveFoldersConfiguration(foldersConfig) {
    return new Promise((resolve, reject) => {
      this.db.serialize(() => {
        // Désactiver toutes les configurations existantes
        this.db.run('UPDATE folder_configurations SET is_active = 0', (error) => {
          if (error) {
            console.error('❌ Erreur désactivation configs:', error);
            reject(error);
            return;
          }

          // Convertir en tableau si c'est un objet, et éliminer les doublons
          let configArray = [];
          if (Array.isArray(foldersConfig)) {
            configArray = foldersConfig;
          } else if (typeof foldersConfig === 'object') {
            configArray = Object.entries(foldersConfig).map(([path, config]) => ({
              path: path,
              name: config.name || 'Unknown',
              category: config.category || 'mails_simples'
            }));
          }

          // Éliminer les doublons par chemin
          const uniqueConfigs = new Map();
          configArray.forEach(config => {
            if (config.path && !uniqueConfigs.has(config.path)) {
              uniqueConfigs.set(config.path, config);
            }
          });

          // Insérer ou mettre à jour les configurations uniques
          const stmt = this.db.prepare(`
            INSERT OR REPLACE INTO folder_configurations 
            (folder_path, category, folder_name, is_active, updated_at) 
            VALUES (?, ?, ?, 1, CURRENT_TIMESTAMP)
          `);

          Array.from(uniqueConfigs.values()).forEach(config => {
            stmt.run(config.path, config.category, config.name);
          });

          stmt.finalize((error) => {
            if (error) {
              console.error('❌ Erreur sauvegarde config dossiers:', error);
              reject(error);
            } else {
              console.log(`✅ Configuration dossiers sauvegardée (${uniqueConfigs.size} dossiers uniques)`);
              resolve();
            }
          });
        });
      });
    });
  }

  /**
   * Récupérer un paramètre de configuration
   */
  async getAppConfig(configKey) {
    return new Promise((resolve, reject) => {
      const sql = 'SELECT config_value, config_type FROM app_config WHERE config_key = ?';
      
      this.db.get(sql, [configKey], (error, row) => {
        if (error) {
          console.error('❌ Erreur lecture config:', error);
          reject(error);
        } else if (row) {
          try {
            const value = row.config_type === 'json' ? JSON.parse(row.config_value) : row.config_value;
            resolve(value);
          } catch (parseError) {
            console.error('❌ Erreur parsing config:', parseError);
            resolve(row.config_value);
          }
        } else {
          resolve(null);
        }
      });
    });
  }

  /**
   * Sauvegarder un paramètre de configuration
   */
  async setAppConfig(configKey, configValue, description = null) {
    return new Promise((resolve, reject) => {
      const isObject = typeof configValue === 'object';
      const valueToStore = isObject ? JSON.stringify(configValue) : configValue;
      const configType = isObject ? 'json' : 'string';

      const sql = `
        INSERT OR REPLACE INTO app_config 
        (config_key, config_value, config_type, description, updated_at) 
        VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
      `;

      this.db.run(sql, [configKey, valueToStore, configType, description], (error) => {
        if (error) {
          console.error('❌ Erreur sauvegarde config:', error);
          reject(error);
        } else {
          console.log(`✅ Configuration ${configKey} sauvegardée`);
          resolve();
        }
      });
    });
  }

  /**
   * Ajouter une configuration de dossier
   */
  async addFolderConfiguration(folderPath, category, folderName) {
    return new Promise((resolve, reject) => {
      if (!this.db) {
        console.error('❌ Base de données non initialisée pour addFolderConfiguration');
        reject(new Error('Base de données non initialisée'));
        return;
      }

      // D'abord vérifier si le dossier existe déjà
      const checkSql = 'SELECT folder_path FROM folder_configurations WHERE folder_path = ?';
      
      this.db.get(checkSql, [folderPath], (err, row) => {
        if (err) {
          console.error('❌ Erreur vérification dossier existant:', err);
          reject(err);
          return;
        }
        
        if (row) {
          // Dossier existe, le réactiver
          const updateSql = 'UPDATE folder_configurations SET category = ?, folder_name = ?, is_active = 1 WHERE folder_path = ?';
          this.db.run(updateSql, [category, folderName, folderPath], function(error) {
            if (error) {
              console.error('❌ Erreur mise à jour config dossier:', error);
              reject(error);
            } else {
              console.log(`✅ Configuration mise à jour pour: ${folderPath}`);
              resolve(true);
            }
          });
        } else {
          // Nouveau dossier, l'insérer
          const insertSql = 'INSERT INTO folder_configurations (folder_path, category, folder_name, is_active) VALUES (?, ?, ?, 1)';
          this.db.run(insertSql, [folderPath, category, folderName], function(error) {
            if (error) {
              console.error('❌ Erreur ajout config dossier:', error);
              reject(error);
            } else {
              console.log(`✅ Configuration ajoutée pour: ${folderPath}`);
              resolve(true);
            }
          });
        }
      });
    });
  }

  /**
   * Mettre à jour la catégorie d'un dossier
   */
  async updateFolderCategory(folderPath, category) {
    return new Promise((resolve, reject) => {
      if (!this.db) {
        console.error('❌ Base de données non initialisée pour updateFolderCategory');
        reject(new Error('Base de données non initialisée'));
        return;
      }

      const sql = 'UPDATE folder_configurations SET category = ? WHERE folder_path = ? AND is_active = 1';
      
      console.log(`🔄 Mise à jour de la catégorie pour: ${folderPath} -> ${category}`);
      
      this.db.run(sql, [category, folderPath], function(error) {
        if (error) {
          console.error('❌ Erreur mise à jour catégorie:', error);
          reject(error);
        } else {
          if (this.changes > 0) {
            console.log(`✅ Catégorie mise à jour pour: ${folderPath}`);
            resolve(true);
          } else {
            console.log(`⚠️ Aucun dossier actif trouvé pour: ${folderPath}`);
            resolve(false);
          }
        }
      });
    });
  }

  /**
   * Supprimer un dossier de la configuration
   */
  async deleteFolderConfiguration(folderPath) {
    return new Promise((resolve, reject) => {
      if (!this.db) {
        console.error('❌ Base de données non initialisée pour deleteFolderConfiguration');
        reject(new Error('Base de données non initialisée'));
        return;
      }

      const sql = 'UPDATE folder_configurations SET is_active = 0 WHERE folder_path = ?';
      
      console.log(`🗑️ Suppression de la configuration pour: ${folderPath}`);
      
      this.db.run(sql, [folderPath], function(error) {
        if (error) {
          console.error('❌ Erreur suppression config dossier:', error);
          reject(error);
        } else {
          console.log(`✅ Configuration supprimée pour: ${folderPath} (${this.changes} lignes affectées)`);
          resolve(this.changes > 0);
        }
      });
    });
  }

  /**
   * Sauvegarder les métriques hebdomadaires
   */
  async saveWeeklyMetrics(weeklyData) {
    return new Promise((resolve, reject) => {
      const date = new Date().toISOString().split('T')[0];
      const sql = `
        INSERT OR REPLACE INTO metrics_history 
        (metric_date, metric_type, metric_data) 
        VALUES (?, 'weekly', ?)
      `;

      this.db.run(sql, [date, JSON.stringify(weeklyData)], (error) => {
        if (error) {
          console.error('❌ Erreur sauvegarde métriques:', error);
          reject(error);
        } else {
          resolve();
        }
      });
    });
  }

  /**
   * Récupérer les métriques hebdomadaires
   */
  async getWeeklyMetrics(limit = 30) {
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT metric_data, metric_date 
        FROM metrics_history 
        WHERE metric_type = 'weekly' 
        ORDER BY metric_date DESC 
        LIMIT ?
      `;

      this.db.all(sql, [limit], (error, rows) => {
        if (error) {
          console.error('❌ Erreur lecture métriques:', error);
          reject(error);
        } else {
          const metrics = rows.map(row => {
            try {
              return JSON.parse(row.metric_data);
            } catch (parseError) {
              console.error('❌ Erreur parsing métriques:', parseError);
              return null;
            }
          }).filter(metric => metric !== null);
          resolve(metrics);
        }
      });
    });
  }

  /**
   * Sauvegarder l'historique des métriques quotidiennes
   */
  async saveHistoricalData(historicalData) {
    return new Promise((resolve, reject) => {
      const date = historicalData.date ? new Date(historicalData.date).toISOString().split('T')[0] : new Date().toISOString().split('T')[0];
      const sql = `
        INSERT OR REPLACE INTO metrics_history 
        (metric_date, metric_type, metric_data) 
        VALUES (?, 'daily', ?)
      `;

      this.db.run(sql, [date, JSON.stringify(historicalData)], (error) => {
        if (error) {
          console.error('❌ Erreur sauvegarde historique:', error);
          reject(error);
        } else {
          resolve();
        }
      });
    });
  }

  /**
   * Récupérer l'historique des métriques
   */
  async getHistoricalData(limit = 30) {
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT metric_data 
        FROM metrics_history 
        WHERE metric_type = 'daily' 
        ORDER BY metric_date DESC 
        LIMIT ?
      `;

      this.db.all(sql, [limit], (error, rows) => {
        if (error) {
          console.error('❌ Erreur lecture historique:', error);
          reject(error);
        } else {
          const history = rows.map(row => {
            try {
              return JSON.parse(row.metric_data);
            } catch (parseError) {
              console.error('❌ Erreur parsing historique:', parseError);
              return null;
            }
          }).filter(item => item !== null);
          resolve(history);
        }
      });
    });
  }

  /**
   * Récupère les statistiques par catégorie
   */
  async getStatsByCategory() {
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          folder_path,
          COUNT(*) as total,
          SUM(CASE WHEN is_read = 1 THEN 1 ELSE 0 END) as read,
          SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread
        FROM emails 
        WHERE folder_path IS NOT NULL
        GROUP BY folder_path
        ORDER BY total DESC
      `;

      this.db.all(sql, [], (error, rows) => {
        if (error) {
          console.error('❌ Erreur stats par catégorie:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Récupère un résumé des métriques
   */
  async getMetricsSummary() {
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          COUNT(*) as total_emails,
          SUM(CASE WHEN is_read = 1 THEN 1 ELSE 0 END) as read_emails,
          SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread_emails,
          COUNT(DISTINCT folder_path) as total_folders,
          MAX(received_time) as latest_email,
          MIN(received_time) as oldest_email
        FROM emails
      `;

      this.db.get(sql, [], (error, row) => {
        if (error) {
          console.error('❌ Erreur résumé métriques:', error);
          reject(error);
        } else {
          resolve(row || {
            total_emails: 0,
            read_emails: 0,
            unread_emails: 0,
            total_folders: 0,
            latest_email: null,
            oldest_email: null
          });
        }
      });
    });
  }

  /**
   * Récupère la distribution par dossiers
   */
  async getFolderDistribution() {
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          folder_path,
          COUNT(*) as count,
          ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM emails), 2) as percentage
        FROM emails 
        WHERE folder_path IS NOT NULL
        GROUP BY folder_path
        ORDER BY count DESC
        LIMIT 10
      `;

      this.db.all(sql, [], (error, rows) => {
        if (error) {
          console.error('❌ Erreur distribution dossiers:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Récupère l'évolution hebdomadaire
   */
  async getWeeklyEvolution() {
    return new Promise((resolve, reject) => {
      const sql = `
        SELECT 
          DATE(received_time, 'weekday 0', '-6 days') as week_start,
          COUNT(*) as email_count,
          SUM(CASE WHEN is_read = 1 THEN 1 ELSE 0 END) as read_count,
          SUM(CASE WHEN is_read = 0 THEN 1 ELSE 0 END) as unread_count
        FROM emails 
        WHERE received_time >= DATE('now', '-8 weeks')
        GROUP BY DATE(received_time, 'weekday 0', '-6 days')
        ORDER BY week_start DESC
        LIMIT 8
      `;

      this.db.all(sql, [], (error, rows) => {
        if (error) {
          console.error('❌ Erreur évolution hebdomadaire:', error);
          reject(error);
        } else {
          resolve(rows || []);
        }
      });
    });
  }

  /**
   * Ferme la connexion à la base de données
   */
  async close() {
    if (this.db) {
      return new Promise((resolve) => {
        this.db.close((error) => {
          if (error) {
            console.error('❌ Erreur fermeture base de données:', error);
          } else {
            console.log('🔒 Base de données fermée');
          }
          resolve();
        });
      });
    }
  }
}

// Instance singleton
const databaseService = new DatabaseService();

module.exports = databaseService;
