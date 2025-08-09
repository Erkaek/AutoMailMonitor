// Script pour ajouter un dossier de test à la configuration de monitoring
const path = require('path');

// Simuler une initialisation basique de la base de données
async function addTestFolderToConfig() {
    try {
        console.log('🔧 Ajout d\'un dossier de test à la configuration...');
        
        // Créer un script SQL simple
        const sqlCommands = [
            `-- Configuration du dossier de test`,
            `INSERT OR REPLACE INTO folder_configurations (folder_path, category, folder_name, updated_at)`,
            `VALUES ('\\\\erkaekanon@outlook.com\\Boîte de réception\\test', 'test', 'Dossier Test', CURRENT_TIMESTAMP);`,
            ``,
            `-- Vérifier la configuration`,
            `SELECT * FROM folder_configurations;`
        ];
        
        console.log('📝 Commandes SQL générées :');
        sqlCommands.forEach(cmd => console.log(cmd));
        
        console.log('\n✅ Configuration SQL prête');
        console.log('📋 Pour appliquer :');
        console.log('1. Arrêter l\'application AutoMailMonitor');
        console.log('2. Exécuter ces commandes SQL dans la base data/emails.db');
        console.log('3. Redémarrer l\'application');
        
        return sqlCommands;
        
    } catch (error) {
        console.error('❌ Erreur génération config:', error.message);
        throw error;
    }
}

// Exécuter si lancé directement
if (require.main === module) {
    addTestFolderToConfig();
}

module.exports = { addTestFolderToConfig };
