const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');

console.log('🔧 Fix Configuration des Dossiers');

// Configuration avec Electron pour better-sqlite3
const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log('\n=== VÉRIFICATION DE LA BASE ===');
// Vérifier les tables existantes
const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
console.log('Tables existantes:', tables.map(t => t.name));

// Créer la table folders_config si elle n'existe pas
const createTableSQL = `
CREATE TABLE IF NOT EXISTS folders_config (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_name TEXT UNIQUE NOT NULL,
    is_active INTEGER DEFAULT 1,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)`;

db.exec(createTableSQL);
console.log('✅ Table folders_config créée/vérifiée');

console.log('\n=== CONFIGURATION ACTUELLE ===');
const currentFolders = db.prepare('SELECT * FROM folders_config').all();
console.log('Dossiers configurés:', currentFolders);

console.log('\n=== SUPPRESSION DES DOSSIERS INEXISTANTS ===');
const deleteStmt = db.prepare('DELETE FROM folders_config WHERE folder_name IN (?, ?, ?, ?)');
const deleted = deleteStmt.run('testA', 'test', 'test-1', 'test-c');
console.log('✅ Supprimés:', deleted.changes, 'dossiers test');

console.log('\n=== AJOUT DES VRAIS DOSSIERS OUTLOOK ===');
const insertFolder = db.prepare('INSERT OR REPLACE INTO folders_config (folder_name, is_active) VALUES (?, ?)');

// Ajouter les dossiers Outlook standards
const realFolders = [
    { name: 'Inbox', active: 1 },
    { name: 'SentMail', active: 1 }
];

realFolders.forEach(folder => {
    try {
        insertFolder.run(folder.name, folder.active);
        console.log(`✅ Ajouté: ${folder.name}`);
    } catch (error) {
        console.log(`❌ Erreur ${folder.name}:`, error.message);
    }
});

console.log('\n=== NOUVELLE CONFIGURATION ===');
const newFolders = db.prepare('SELECT * FROM folders_config').all();
console.log('Configuration finale:', newFolders);

db.close();
console.log('\n🎉 Configuration corrigée ! Redémarrez l\'application.');
process.exit(0);
