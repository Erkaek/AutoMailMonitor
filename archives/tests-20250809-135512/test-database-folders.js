const path = require('path');
const dbService = require('./src/services/optimizedDatabaseService');

async function testDatabaseFolders() {
    try {
        console.log('🔍 Test de la configuration des dossiers...');
        
        // Le service est un singleton déjà initialisé
        
        // Vérifier la configuration des dossiers
        const foldersConfig = dbService.getFoldersConfiguration();
        console.log('\n📁 Configuration des dossiers:');
        console.log('Nombre de dossiers configurés:', foldersConfig.length);
        
        if (foldersConfig.length > 0) {
            foldersConfig.forEach((folder, index) => {
                console.log(`  ${index + 1}. ${folder.folder_path} (${folder.category})`);
            });
        } else {
            console.log('  ❌ Aucun dossier configuré !');
            console.log('\n🔧 Ajout du dossier Inbox par défaut...');
            
            // Ajouter le dossier Inbox par défaut
            const result = dbService.addFolderConfiguration(
                '\\\\Inbox',
                'inbox',
                'Boîte de réception'
            );
            console.log('Résultat ajout:', result);
            
            // Vérifier après ajout
            const newConfig = dbService.getFoldersConfiguration();
            console.log('\n📁 Nouvelle configuration:');
            newConfig.forEach((folder, index) => {
                console.log(`  ${index + 1}. ${folder.folder_path} (${folder.category})`);
            });
        }
        
        // Statistiques générales
        const stats = dbService.getGeneralStats();
        console.log('\n📊 Statistiques générales:');
        console.log('Total emails:', stats.totalEmails);
        console.log('Non lus:', stats.unreadTotal);
        console.log('Emails aujourd\'hui:', stats.emailsToday);
        
        console.log('\n✅ Test terminé avec succès');
        
    } catch (error) {
        console.error('❌ Erreur test database:', error.message);
        console.error('Stack:', error.stack);
    }
}

testDatabaseFolders();
