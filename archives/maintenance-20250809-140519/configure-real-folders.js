const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');

console.log('🔧 Configuration FINALE des dossiers Outlook à monitorer');

const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log('\n=== AJOUT DES VRAIS DOSSIERS À MONITORER ===');
const insertFolder = db.prepare(`
    INSERT INTO folder_configurations 
    (folder_path, category, folder_name, created_at, updated_at) 
    VALUES (?, ?, ?, datetime('now'), datetime('now'))
`);

// Les vrais dossiers selon l'utilisateur avec leurs catégories
const realFolders = [
    {
        path: 'erkaekanon@outlook.com\\Boîte de réception\\testA',
        name: 'testA', 
        category: 'Règlements'
    },
    {
        path: 'erkaekanon@outlook.com\\Boîte de réception\\test',
        name: 'test',
        category: 'Déclarations' 
    },
    {
        path: 'erkaekanon@outlook.com\\Boîte de réception\\test\\test-1',
        name: 'test-1',
        category: 'Déclarations'
    },
    {
        path: 'erkaekanon@outlook.com\\Boîte de réception\\testA\\test-c', 
        name: 'test-c',
        category: 'mails_simples'
    }
];

realFolders.forEach(folder => {
    try {
        insertFolder.run(folder.path, folder.category, folder.name);
        console.log(`✅ Ajouté: ${folder.name} (${folder.category})`);
        console.log(`   📍 Chemin: ${folder.path}`);
    } catch (error) {
        console.log(`❌ Erreur ${folder.name}:`, error.message);
    }
});

console.log('\n=== CONFIGURATION FINALE ===');
const finalFolders = db.prepare('SELECT * FROM folder_configurations ORDER BY id').all();
console.log(`📊 ${finalFolders.length} dossiers configurés pour le monitoring:`);
finalFolders.forEach((folder, index) => {
    console.log(`${index + 1}. 📁 ${folder.folder_name} [${folder.category}]`);
    console.log(`   📍 ${folder.folder_path}`);
    console.log('');
});

db.close();
console.log('🎉 Configuration terminée !');
console.log('🚀 Le monitoring va maintenant surveiller ces dossiers Outlook spécifiques.');
process.exit(0);
