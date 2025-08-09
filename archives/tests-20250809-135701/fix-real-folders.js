const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');

console.log('🔧 Configuration des VRAIS dossiers Outlook à monitorer');

const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log('\n=== CONFIGURATION ACTUELLE ===');
const currentFolders = db.prepare('SELECT * FROM folder_configurations').all();
console.log('Dossiers configurés:', currentFolders);

console.log('\n=== SUPPRESSION DE TOUS LES DOSSIERS ACTUELS ===');
const deleteAll = db.prepare('DELETE FROM folder_configurations');
deleteAll.run();
console.log('✅ Configuration nettoyée');

console.log('\n=== AJOUT DES VRAIS DOSSIERS À MONITORER ===');
const insertFolder = db.prepare('INSERT INTO folder_configurations (folder_name, is_active) VALUES (?, ?)');

// Les vrais dossiers à monitorer selon l'utilisateur
const realFolders = [
    'erkaekanon@outlook.com\\Boîte de réception\\testA',
    'erkaekanon@outlook.com\\Boîte de réception\\test', 
    'erkaekanon@outlook.com\\Boîte de réception\\test\\test-1',
    'erkaekanon@outlook.com\\Boîte de réception\\testA\\test-c'
];

realFolders.forEach(folderPath => {
    try {
        insertFolder.run(folderPath, 1);
        console.log(`✅ Ajouté: ${folderPath}`);
    } catch (error) {
        console.log(`❌ Erreur ${folderPath}:`, error.message);
    }
});

console.log('\n=== CONFIGURATION FINALE ===');
const finalFolders = db.prepare('SELECT * FROM folder_configurations').all();
finalFolders.forEach(folder => {
    console.log(`📁 [${folder.is_active ? 'ACTIF' : 'INACTIF'}] ${folder.folder_name}`);
});

db.close();
console.log('\n🎉 Configuration corrigée avec les VRAIS dossiers Outlook !');
console.log('💡 L\'application va maintenant monitorer ces sous-dossiers spécifiques.');
process.exit(0);
