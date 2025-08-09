// Script de test pour valider le système de rechargement dynamique des dossiers
const UnifiedMonitoringService = require('./src/services/unifiedMonitoringService');
const optimizedDatabaseService = require('./src/services/optimizedDatabaseService');

async function testDynamicConfigReload() {
    try {
        console.log('🧪 Test du système de rechargement dynamique des dossiers...\n');
        
        // 1. Initialiser les services
        console.log('📊 Initialisation de la base de données...');
        await optimizedDatabaseService.initialize();
        
        console.log('🔧 Initialisation du service de monitoring...');
        const monitoringService = new UnifiedMonitoringService();
        await monitoringService.initialize();
        
        // 2. Vérifier la configuration actuelle
        console.log('\n📋 Configuration actuelle:');
        const currentConfig = await optimizedDatabaseService.getFoldersConfiguration();
        console.log(`Dossiers configurés: ${currentConfig.length}`);
        currentConfig.forEach((folder, index) => {
            console.log(`  ${index + 1}. ${folder.folder_path} (${folder.category})`);
        });
        
        // 3. Tester le rechargement manuel
        console.log('\n🔄 Test du rechargement manuel...');
        const reloadResult = await monitoringService.reloadFoldersConfiguration();
        console.log('Résultat:', reloadResult);
        
        // 4. Ajouter un nouveau dossier pour tester la détection
        console.log('\n➕ Ajout d\'un dossier de test...');
        const testFolderPath = '\\\\test-dynamic';
        
        try {
            optimizedDatabaseService.addFolderConfiguration(testFolderPath, 'test', 'Test Dynamique');
            console.log(`✅ Dossier ${testFolderPath} ajouté en base`);
        } catch (error) {
            console.log(`⚠️ Dossier déjà existant ou erreur: ${error.message}`);
        }
        
        // 5. Tester la détection automatique du changement
        console.log('\n👁️ Test de la détection automatique...');
        await monitoringService.checkConfigurationChanges();
        
        // 6. Vérifier la nouvelle configuration
        console.log('\n📋 Nouvelle configuration:');
        const newConfig = await optimizedDatabaseService.getFoldersConfiguration();
        console.log(`Dossiers configurés: ${newConfig.length}`);
        newConfig.forEach((folder, index) => {
            console.log(`  ${index + 1}. ${folder.folder_path} (${folder.category})`);
        });
        
        // 7. Tester le hash de configuration
        console.log('\n🔐 Test du système de hash:');
        const hash1 = monitoringService.calculateConfigHash(currentConfig);
        const hash2 = monitoringService.calculateConfigHash(newConfig);
        console.log(`Hash config initiale: ${hash1}`);
        console.log(`Hash config modifiée: ${hash2}`);
        console.log(`Changement détecté: ${hash1 !== hash2 ? '✅ OUI' : '❌ NON'}`);
        
        // 8. Nettoyer
        console.log('\n🧹 Nettoyage...');
        try {
            optimizedDatabaseService.deleteFolderConfiguration(testFolderPath);
            console.log('✅ Dossier de test supprimé');
        } catch (error) {
            console.log(`⚠️ Erreur nettoyage: ${error.message}`);
        }
        
        console.log('\n✅ Test terminé avec succès !');
        
    } catch (error) {
        console.error('❌ Erreur durant le test:', error.message);
        console.error('Stack:', error.stack);
    }
}

// Exécuter si lancé directement
if (require.main === module) {
    testDynamicConfigReload().then(() => {
        console.log('\n🏁 Fin du test');
        process.exit(0);
    }).catch(error => {
        console.error('💥 Test échoué:', error);
        process.exit(1);
    });
}

module.exports = { testDynamicConfigReload };
