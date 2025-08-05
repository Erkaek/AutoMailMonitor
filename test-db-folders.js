/**
 * Test simple pour vérifier les dossiers en BDD
 */

const path = require('path');
const databaseService = require('./src/services/databaseService');

async function testFolders() {
  try {
    console.log('🔧 Initialisation du service de base de données...');
    await databaseService.initialize();
    
    console.log('📁 Récupération des dossiers configurés...');
    const folders = await databaseService.getFoldersConfiguration();
    
    console.log(`✅ ${folders.length} dossiers trouvés:`);
    folders.forEach((folder, index) => {
      console.log(`  ${index + 1}. ${folder.path} (${folder.category}) - ${folder.name}`);
    });
    
    if (folders.length === 0) {
      console.log('⚠️ Aucun dossier configuré trouvé en base de données');
      
      // Tester l'ajout d'un dossier de test
      console.log('🧪 Test d\'ajout d\'un dossier...');
      await databaseService.addFolderConfiguration('TestPath\\TestFolder', 'Mails simples', 'Test Folder');
      
      const foldersAfter = await databaseService.getFoldersConfiguration();
      console.log(`✅ Après ajout: ${foldersAfter.length} dossiers trouvés`);
    }
    
    process.exit(0);
  } catch (error) {
    console.error('❌ Erreur:', error);
    process.exit(1);
  }
}

testFolders();
