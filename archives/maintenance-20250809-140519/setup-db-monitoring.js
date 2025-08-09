// Script pour configurer directement la base de données
const Database = require('better-sqlite3');
const path = require('path');

async function setupInboxMonitoring() {
    try {
        console.log('🔧 Configuration du monitoring pour le dossier Inbox...');
        
        // Ouvrir la base de données
        const dbPath = path.join(__dirname, 'data', 'emails.db');
        console.log('📂 Ouverture de la base:', dbPath);
        
        const db = new Database(dbPath);
        
        // Vérifier si la table existe
        console.log('🔍 Vérification de la structure de la base...');
        const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
        console.log('Tables existantes:', tables.map(t => t.name));
        
        // Créer la table si elle n'existe pas
        const createTableSQL = `
            CREATE TABLE IF NOT EXISTS folder_configurations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                folder_path TEXT UNIQUE NOT NULL,
                category TEXT NOT NULL,
                folder_name TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        `;
        
        db.exec(createTableSQL);
        console.log('✅ Table folder_configurations assurée');
        
        // Vérifier les dossiers actuels
        const existingFolders = db.prepare('SELECT * FROM folder_configurations').all();
        console.log('📁 Dossiers actuellement configurés:', existingFolders.length);
        
        if (existingFolders.length > 0) {
            existingFolders.forEach(folder => {
                console.log(`  - ${folder.folder_path} (${folder.category})`);
            });
        }
        
        // Ajouter le dossier Inbox s'il n'existe pas
        const inboxExists = existingFolders.some(f => f.folder_path.includes('Inbox'));
        
        if (!inboxExists) {
            console.log('📧 Ajout du dossier Inbox pour monitoring...');
            
            const insertFolder = db.prepare(`
                INSERT OR REPLACE INTO folder_configurations (folder_path, category, folder_name, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            `);
            
            const result = insertFolder.run('\\\\Inbox', 'inbox', 'Boîte de réception');
            console.log('✅ Dossier Inbox ajouté:', result);
            
            // Vérifier après ajout
            const newFolders = db.prepare('SELECT * FROM folder_configurations').all();
            console.log('📁 Configuration mise à jour:');
            newFolders.forEach(folder => {
                console.log(`  - ${folder.folder_path} (${folder.category}) [${folder.folder_name}]`);
            });
        } else {
            console.log('✅ Dossier Inbox déjà configuré');
        }
        
        // Fermer la base
        db.close();
        console.log('✅ Configuration terminée avec succès');
        console.log('');
        console.log('🔄 Redémarrez l\'application AutoMailMonitor pour que les changements prennent effet');
        
    } catch (error) {
        console.error('❌ Erreur configuration:', error.message);
        console.error('Stack:', error.stack);
    }
}

// Exécuter si lancé directement
if (require.main === module) {
    setupInboxMonitoring();
}

module.exports = { setupInboxMonitoring };
