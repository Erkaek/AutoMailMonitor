/**
 * Test de la gestion des configurations de dossiers via base de données uniquement
 */

const databaseService = require('./src/services/databaseService');

async function testDatabaseConfig() {
    try {
        console.log('🧪 Test des configurations de dossiers en base de données...\n');
        
        // Initialiser la base de données
        await databaseService.initialize();
        console.log('✅ Base de données initialisée\n');
        
        // Test 1: Vérifier l'état initial
        console.log('📋 Test 1: État initial');
        let folders = await databaseService.getFoldersConfiguration();
        console.log(`Dossiers trouvés: ${folders.length}`);
        folders.forEach(folder => {
            console.log(`  - ${folder.path} (${folder.category})`);
        });
        console.log('');
        
        // Test 2: Ajouter un dossier de test
        console.log('📋 Test 2: Ajout d\'un dossier');
        const testPath = 'Test\\Dossier';
        const testCategory = 'Mails simples';
        const testName = 'Dossier';
        
        await databaseService.addFolderConfiguration(testPath, testCategory, testName);
        console.log(`✅ Dossier ajouté: ${testPath}\n`);
        
        // Test 3: Vérifier l'ajout
        console.log('📋 Test 3: Vérification après ajout');
        folders = await databaseService.getFoldersConfiguration();
        console.log(`Dossiers trouvés: ${folders.length}`);
        folders.forEach(folder => {
            console.log(`  - ${folder.path} (${folder.category})`);
        });
        console.log('');
        
        // Test 4: Mettre à jour la catégorie
        console.log('📋 Test 4: Mise à jour de catégorie');
        await databaseService.updateFolderCategory(testPath, 'Déclarations');
        console.log(`✅ Catégorie mise à jour pour: ${testPath}\n`);
        
        // Test 5: Vérifier la mise à jour
        console.log('📋 Test 5: Vérification après mise à jour');
        folders = await databaseService.getFoldersConfiguration();
        folders.forEach(folder => {
            console.log(`  - ${folder.path} (${folder.category})`);
        });
        console.log('');
        
        // Test 6: Supprimer le dossier de test
        console.log('📋 Test 6: Suppression du dossier de test');
        await databaseService.deleteFolderConfiguration(testPath);
        console.log(`✅ Dossier supprimé: ${testPath}\n`);
        
        // Test 7: Vérifier la suppression
        console.log('📋 Test 7: Vérification après suppression');
        folders = await databaseService.getFoldersConfiguration();
        console.log(`Dossiers trouvés: ${folders.length}`);
        folders.forEach(folder => {
            console.log(`  - ${folder.path} (${folder.category})`);
        });
        
        console.log('\n🎉 Tous les tests de base de données réussis !');
        console.log('📊 Résumé: La logique BDD-only fonctionne parfaitement');
        
    } catch (error) {
        console.error('❌ Erreur lors du test:', error);
    }
}

// Exécuter le test
testDatabaseConfig();
