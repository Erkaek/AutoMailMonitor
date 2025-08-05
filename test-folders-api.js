const { ipcMain, app, BrowserWindow } = require('electron');

// Simuler l'appel IPC api-folders-tree
async function testFoldersAPI() {
  console.log('🧪 Test de l\'API api-folders-tree...');
  
  try {
    // Importer les dépendances
    const outlookConnector = require('./src/server/outlookConnector');
    const databaseService = require('./src/services/databaseService');
    
    console.log('📊 Initialisation de la base de données...');
    await databaseService.initialize();
    
    console.log('📁 Récupération des configurations de dossiers...');
    const foldersConfig = await databaseService.getFoldersConfiguration();
    console.log(`✅ ${foldersConfig.length} dossiers configurés trouvés en BDD`);
    
    // Afficher chaque configuration
    foldersConfig.forEach((config, index) => {
      console.log(`   ${index + 1}. ${config.path} (${config.category}) - ${config.name}`);
    });
    
    console.log('📁 Test de connexion Outlook...');
    console.log('   Outlook connecté:', outlookConnector.isOutlookConnected);
    
    if (outlookConnector.isOutlookConnected) {
      console.log('📁 Récupération de la structure Outlook...');
      try {
        const allFolders = await outlookConnector.getFolderStructure();
        console.log(`   ${allFolders.length} dossiers trouvés dans Outlook`);
      } catch (outlookError) {
        console.log('   ⚠️ Erreur récupération structure Outlook:', outlookError.message);
      }
    }
    
    // Simuler la logique de l'API
    const monitoredFolders = [];
    
    foldersConfig.forEach(config => {
      monitoredFolders.push({
        path: config.path,
        name: config.name,
        isMonitored: true,
        category: config.category || 'Mails simples',
        emailCount: 0, // On ne peut pas récupérer le vrai count sans Outlook
        parentPath: getParentPath(config.path)
      });
    });
    
    // Calculer les statistiques
    const stats = calculateFolderStats(monitoredFolders);
    
    console.log(`📊 Résultat final:`);
    console.log(`   Dossiers monitorés: ${monitoredFolders.length}`);
    console.log(`   Stats:`, stats);
    
    if (monitoredFolders.length === 0) {
      console.log('❌ PROBLÈME: Aucun dossier dans le résultat final !');
    } else {
      console.log('✅ L\'API devrait fonctionner correctement');
    }
    
  } catch (error) {
    console.error('❌ Erreur dans le test:', error);
  }
}

function getParentPath(folderPath) {
  const parts = folderPath.split('\\');
  if (parts.length <= 1) return null;
  return parts.slice(0, -1).join('\\');
}

function calculateFolderStats(folders) {
  const stats = {
    total: folders.length,
    active: 0,
    declarations: 0,
    reglements: 0,
    simples: 0
  };

  folders.forEach(folder => {
    if (folder.isMonitored) {
      stats.active++;
      
      switch (folder.category) {
        case 'Déclarations':
        case 'declarations':
          stats.declarations++;
          break;
        case 'Règlements':
        case 'reglements':
          stats.reglements++;
          break;
        case 'Mails simples':
        case 'mails_simples':
          stats.simples++;
          break;
      }
    }
  });

  return stats;
}

// Lancer le test
testFoldersAPI().then(() => {
  console.log('🏁 Test terminé');
}).catch((error) => {
  console.error('💥 Erreur test:', error);
});
