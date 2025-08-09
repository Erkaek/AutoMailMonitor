const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

// Trouver la base de données
const dbPaths = [
    path.join(__dirname, 'data', 'emails.db'),
    path.join(__dirname, 'emails.db')
];

let dbPath = null;
for (const testPath of dbPaths) {
    if (fs.existsSync(testPath)) {
        dbPath = testPath;
        break;
    }
}

console.log('🔍 Inspection des tables existantes:', dbPath);

try {
    const db = new Database(dbPath);
    
    // Lister toutes les tables
    const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
    console.log('\n📊 Tables dans la base de données:');
    tables.forEach(table => {
        console.log(`\n📋 Structure de ${table.name}:`);
        const schema = db.prepare(`PRAGMA table_info(${table.name})`).all();
        schema.forEach(col => {
            console.log(`  - ${col.name}: ${col.type} (${col.notnull ? 'NOT NULL' : 'NULL'}) ${col.pk ? '[PK]' : ''}`);
        });
    });
    
    db.close();
    
} catch (error) {
    console.error('❌ Erreur:', error);
    process.exit(1);
}
