const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');

console.log('🔍 Inspection de la structure de la base de données');

const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log('\n=== STRUCTURE DE LA TABLE folder_configurations ===');
const schema = db.prepare("PRAGMA table_info(folder_configurations)").all();
console.log('Colonnes disponibles:');
schema.forEach(col => {
    console.log(`  - ${col.name} (${col.type}) ${col.notnull ? 'NOT NULL' : ''} ${col.pk ? 'PRIMARY KEY' : ''}`);
});

console.log('\n=== DONNÉES ACTUELLES ===');
const currentFolders = db.prepare('SELECT * FROM folder_configurations').all();
console.log('Dossiers configurés:', currentFolders);

db.close();
process.exit(0);
