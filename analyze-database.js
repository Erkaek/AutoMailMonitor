const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

console.log('🧹 NETTOYAGE ET REORGANISATION DE LA BASE DE DONNÉES');
console.log('='.repeat(60));

// Chemins des bases de données
const rootDbPath = path.join(__dirname, 'emails.db');
const dataDbPath = path.join(__dirname, 'data', 'emails.db');

// Vérifier les deux bases
console.log('\n📊 ÉTAT ACTUEL:');

if (fs.existsSync(rootDbPath)) {
    const stats = fs.statSync(rootDbPath);
    console.log(`📄 Racine: emails.db - ${stats.size} bytes`);
} else {
    console.log('📄 Racine: emails.db - N\'EXISTE PAS');
}

if (fs.existsSync(dataDbPath)) {
    const stats = fs.statSync(dataDbPath);
    console.log(`📁 Data: data/emails.db - ${stats.size} bytes`);
} else {
    console.log('📁 Data: data/emails.db - N\'EXISTE PAS');
}

// Supprimer la base vide à la racine
if (fs.existsSync(rootDbPath)) {
    const stats = fs.statSync(rootDbPath);
    if (stats.size === 0) {
        fs.unlinkSync(rootDbPath);
        console.log('🗑️ Base vide supprimée: emails.db');
    }
}

// Analyser la vraie base de données
console.log('\n🔍 ANALYSE DE LA VRAIE BASE (data/emails.db):');

const db = new Database(dataDbPath);

// Lister toutes les tables
const tables = db.prepare(`
    SELECT name, sql 
    FROM sqlite_master 
    WHERE type='table' AND name NOT LIKE 'sqlite_%'
    ORDER BY name
`).all();

console.log(`\n📋 TABLES TROUVÉES (${tables.length} tables):`);
tables.forEach((table, index) => {
    console.log(`${index + 1}. ${table.name}`);
});

// Analyser le contenu de chaque table
console.log('\n📊 CONTENU DES TABLES:');

for (const table of tables) {
    try {
        const count = db.prepare(`SELECT COUNT(*) as count FROM ${table.name}`).get();
        console.log(`  ${table.name}: ${count.count} enregistrements`);
        
        // Pour certaines tables importantes, montrer la structure
        if (['emails', 'weekly_stats', 'folder_mappings', 'app_settings'].includes(table.name)) {
            const columns = db.prepare(`PRAGMA table_info(${table.name})`).all();
            console.log(`    Colonnes: ${columns.map(col => col.name).join(', ')}`);
        }
    } catch (error) {
        console.log(`  ${table.name}: ERREUR - ${error.message}`);
    }
}

// Identifier les tables inutiles
const usefulTables = [
    'emails',           // Table principale des emails
    'weekly_stats',     // Statistiques hebdomadaires VBA
    'folder_mappings',  // Mapping des dossiers
    'app_settings',     // Paramètres de l'app
    'folder_configurations', // Config des dossiers surveillés
    'monitored_folders' // Dossiers surveillés
];

const uselessTables = tables.filter(t => !usefulTables.includes(t.name));

if (uselessTables.length > 0) {
    console.log('\n🗑️ TABLES POTENTIELLEMENT INUTILES:');
    uselessTables.forEach(table => {
        const count = db.prepare(`SELECT COUNT(*) as count FROM ${table.name}`).get();
        console.log(`  - ${table.name} (${count.count} enregistrements)`);
    });
    
    console.log('\n❓ Voulez-vous supprimer ces tables ? (Créez un fichier "cleanup-confirm.txt" pour confirmer)');
}

// Analyser la table emails spécifiquement
console.log('\n📧 ANALYSE DÉTAILLÉE - TABLE EMAILS:');
const emailColumns = db.prepare(`PRAGMA table_info(emails)`).all();
console.log('Colonnes présentes:');
emailColumns.forEach(col => {
    console.log(`  - ${col.name}: ${col.type} ${col.notnull ? '(NOT NULL)' : ''} ${col.pk ? '[PK]' : ''}`);
});

// Chercher les doublons potentiels
const duplicateColumns = emailColumns.filter(col => 
    ['sender', 'sender_name', 'sender_email'].includes(col.name) ||
    ['size', 'size_kb'].includes(col.name) ||
    ['is_treated', 'treated_time'].includes(col.name)
);

if (duplicateColumns.length > 0) {
    console.log('\n⚠️ COLONNES POTENTIELLEMENT REDONDANTES:');
    duplicateColumns.forEach(col => {
        const hasData = db.prepare(`SELECT COUNT(*) as count FROM emails WHERE ${col.name} IS NOT NULL`).get();
        console.log(`  - ${col.name}: ${hasData.count} valeurs non-nulles`);
    });
}

// Statistiques sur les emails
const emailStats = db.prepare(`
    SELECT 
        COUNT(*) as total,
        COUNT(CASE WHEN week_identifier IS NOT NULL THEN 1 END) as with_week,
        COUNT(CASE WHEN received_time IS NOT NULL THEN 1 END) as with_received_time,
        COUNT(CASE WHEN folder_name IS NOT NULL THEN 1 END) as with_folder,
        MIN(received_time) as oldest,
        MAX(received_time) as newest
    FROM emails
`).get();

console.log('\n📈 STATISTIQUES EMAILS:');
console.log(`  Total: ${emailStats.total}`);
console.log(`  Avec week_identifier: ${emailStats.with_week}`);
console.log(`  Avec received_time: ${emailStats.with_received_time}`);
console.log(`  Avec folder_name: ${emailStats.with_folder}`);
if (emailStats.oldest) {
    console.log(`  Plus ancien: ${emailStats.oldest}`);
    console.log(`  Plus récent: ${emailStats.newest}`);
}

db.close();

console.log('\n✅ ANALYSE TERMINÉE');
console.log('\n💡 RECOMMANDATIONS:');
console.log('1. Utilisez UNIQUEMENT data/emails.db (la base principale)');
console.log('2. Supprimez les tables inutiles identifiées ci-dessus');
console.log('3. Nettoyez les colonnes redondantes dans la table emails');
console.log('4. Configurez l\'application pour utiliser data/emails.db');
