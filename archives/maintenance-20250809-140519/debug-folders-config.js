const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');

console.log('🔍 DEBUG: Vérification configuration des dossiers');

const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log('\n=== DONNÉES BRUTES DE LA BDD ===');
const rawFolders = db.prepare('SELECT * FROM folder_configurations').all();
rawFolders.forEach((folder, index) => {
    console.log(`${index + 1}. ID: ${folder.id}`);
    console.log(`   📍 folder_path: "${folder.folder_path}"`);
    console.log(`   📝 folder_name: "${folder.folder_name}"`);
    console.log(`   🏷️ category: "${folder.category}"`);
    console.log('');
});

console.log('\n=== SIMULATION LOGIQUE DU SERVICE ===');
// Simuler la logique du unifiedMonitoringService
const processedFolders = rawFolders.filter(folder => 
    folder && 
    (folder.folder_path || folder.folder_name || folder.path) &&
    (folder.folder_name !== 'folderCategories') &&
    folder.category
).map(folder => {
    const processed = {
        path: folder.folder_path || folder.folder_name || folder.path,
        category: folder.category,
        name: folder.folder_name || folder.name,
        enabled: true
    };
    
    console.log(`✅ Dossier traité:`);
    console.log(`   🎯 path (envoyé à PowerShell): "${processed.path}"`);
    console.log(`   📝 name (affiché dans logs): "${processed.name}"`);
    console.log(`   🏷️ category: "${processed.category}"`);
    console.log('');
    
    return processed;
});

console.log(`\n📊 Résultat: ${processedFolders.length} dossiers configurés pour monitoring`);

db.close();
process.exit(0);
