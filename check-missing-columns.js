const fs = require('fs');
const path = require('path');

console.log('🧹 NETTOYAGE DES RÉFÉRENCES À sender_name');
console.log('==========================================');

// Fichiers à corriger et leurs corrections
const corrections = [
    {
        file: 'src/services/optimizedDatabaseService.js',
        description: 'Supprimer les références à sender_name dans les mappings'
    },
    {
        file: 'src/services/unifiedMonitoringService.js',
        description: 'Remplacer sender_name par sender_email'
    },
    {
        file: 'src/services/databaseService.js',
        description: 'Supprimer les références à sender_name'
    },
    {
        file: 'src/server/graphOutlookConnector.js',
        description: 'Utiliser sender_email au lieu de sender_name'
    }
];

console.log('\n📋 Fichiers à corriger:');
corrections.forEach((correction, index) => {
    console.log(`  ${index + 1}. ${correction.file}`);
    console.log(`     ${correction.description}`);
});

console.log('\n🔍 Recherche de colonnes manquantes dans la BDD...');

// Colonnes mentionnées dans le CREATE TABLE du service mais absentes de la BDD réelle
const serviceCreateTable = `
CREATE TABLE IF NOT EXISTS emails (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    outlook_id TEXT UNIQUE,
    entry_id TEXT,
    subject TEXT NOT NULL,
    sender_name TEXT,
    sender_email TEXT,
    recipient_email TEXT,
    received_time DATETIME,
    sent_time DATETIME,
    treated_time DATETIME,
    folder_name TEXT,
    folder_path TEXT,
    category TEXT,
    is_read BOOLEAN DEFAULT 0,
    is_replied BOOLEAN DEFAULT 0,
    is_treated BOOLEAN DEFAULT 0,
    has_attachment BOOLEAN DEFAULT 0,
    body_preview TEXT,
    importance INTEGER DEFAULT 1,
    size_bytes INTEGER,
    week_identifier TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    is_deleted INTEGER DEFAULT 0,
    deleted_at TEXT
)
`;

// Colonnes réellement présentes dans la BDD
const actualColumns = [
    'id', 'outlook_id', 'subject', 'sender_email', 'received_time', 
    'sent_time', 'folder_name', 'category', 'is_read', 'importance', 
    'week_identifier', 'created_at', 'updated_at'
];

// Extraction des colonnes du CREATE TABLE du service
const serviceColumns = serviceCreateTable
    .match(/^\s*(\w+)\s+/gm)
    ?.map(match => match.trim().split(/\s+/)[0])
    .filter(col => col && col !== 'CREATE' && col !== 'TABLE' && col !== 'IF' && col !== 'NOT' && col !== 'EXISTS' && col !== 'emails') || [];

console.log('\n📊 COMPARAISON DES COLONNES:');
console.log('\n✅ Colonnes présentes dans la BDD:');
actualColumns.forEach(col => console.log(`  - ${col}`));

console.log('\n❌ Colonnes dans le code CREATE TABLE mais ABSENTES de la BDD:');
const missingColumns = serviceColumns.filter(col => !actualColumns.includes(col) && col !== 'sender_name');
missingColumns.forEach(col => console.log(`  - ${col}`));

console.log('\n⚠️ Colonnes à SUPPRIMER du code (n\'existent plus en BDD):');
const obsoleteColumns = serviceColumns.filter(col => col === 'sender_name');
obsoleteColumns.forEach(col => console.log(`  - ${col}`));

if (missingColumns.length > 0) {
    console.log('\n📝 SQL pour ajouter les colonnes manquantes:');
    missingColumns.forEach(col => {
        let sqlType = 'TEXT';
        if (col.includes('time') || col.includes('date')) sqlType = 'DATETIME';
        if (col.includes('is_') || col.includes('has_')) sqlType = 'BOOLEAN DEFAULT 0';
        if (col.includes('size') || col.includes('bytes')) sqlType = 'INTEGER';
        
        console.log(`  ALTER TABLE emails ADD COLUMN ${col} ${sqlType};`);
    });
}

console.log('\n💡 RECOMMANDATIONS FINALES:');
console.log('============================');

if (missingColumns.length > 0) {
    console.log(`\n1. 🔧 AJOUTER ${missingColumns.length} colonne(s) manquante(s) à la BDD`);
    console.log('   Exécuter les commandes SQL ci-dessus');
}

if (obsoleteColumns.length > 0) {
    console.log(`\n2. 🧹 NETTOYER ${obsoleteColumns.length} référence(s) obsolète(s) dans le code`);
    console.log('   Remplacer toutes les références à sender_name par sender_email');
}

console.log('\n3. ✅ VÉRIFIER que les INSERT/UPDATE utilisent les bonnes colonnes');
console.log('4. 🔍 TESTER l\'application après les corrections');

console.log('\n✅ Analyse terminée !');
