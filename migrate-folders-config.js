const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');

console.log('🔧 Migration Table Folders Configuration');

const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log('\n=== VÉRIFICATION DES TABLES ===');
const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
console.log('Tables existantes:', tables.map(t => t.name));

// Vérifier s'il y a des données dans folder_configurations
const existingConfig = db.prepare('SELECT * FROM folder_configurations').all();
console.log('\nConfiguration existante dans folder_configurations:', existingConfig);

console.log('\n=== MIGRATION VERS folders_config ===');

if (existingConfig.length > 0) {
    console.log('📋 Migration des données existantes...');
    
    // Migrer les données de folder_configurations vers folders_config
    existingConfig.forEach(config => {
        try {
            const insertStmt = db.prepare('INSERT OR REPLACE INTO folders_config (folder_name, is_active) VALUES (?, ?)');
            insertStmt.run(config.folder_name || config.folder_path, config.is_active || 1);
            console.log(`✅ Migré: ${config.folder_name || config.folder_path}`);
        } catch (error) {
            console.log(`❌ Erreur migration ${config.folder_name}:`, error.message);
        }
    });
} else {
    console.log('📋 Aucune donnée à migrer, ajout configuration par défaut...');
    
    const insertStmt = db.prepare('INSERT OR REPLACE INTO folders_config (folder_name, is_active) VALUES (?, ?)');
    insertStmt.run('Inbox', 1);
    insertStmt.run('SentMail', 1);
    console.log('✅ Configuration par défaut ajoutée');
}

console.log('\n=== NETTOYAGE ===');
// Supprimer les anciens dossiers test
db.prepare('DELETE FROM folders_config WHERE folder_name IN (?, ?, ?, ?)').run('testA', 'test', 'test-1', 'test-c');

console.log('\n=== CONFIGURATION FINALE ===');
const finalConfig = db.prepare('SELECT * FROM folders_config').all();
console.log('Configuration folders_config:', finalConfig);

db.close();
console.log('\n🎉 Migration terminée ! Redémarrez l\'application.');
process.exit(0);
