const { ipcMain } = require('electron');

// Simuler l'appel IPC api-folders-tree et log complet
async function debugFoldersAPI() {
  console.log('🔍 DEBUG COMPLET de l\'API api-folders-tree...');
  
  try {
    const outlookConnector = require('./src/server/outlookConnector');
    const databaseService = require('./src/services/databaseService');
    
    await databaseService.initialize();
    
    // Récupérer les configurations de dossiers
    const foldersConfig = await databaseService.getFoldersConfiguration();
    console.log('📊 Configurations BDD:', JSON.stringify(foldersConfig, null, 2));
    
    // Récupérer la structure Outlook (si connecté)
    if (outlookConnector.isOutlookConnected) {
      try {
        const allFolders = await outlookConnector.getFolderStructure();
        console.log('📊 Structure Outlook count:', allFolders.length);
      } catch (error) {
        console.log('⚠️ Erreur Outlook:', error.message);
      }
    }
    
    // Fonction helper
    function findFolderInStructure(folders, targetPath) {
      for (const folder of folders) {
        if (folder.FolderPath === targetPath) {
          return folder;
        }
        if (folder.SubFolders && folder.SubFolders.length > 0) {
          const found = findFolderInStructure(folder.SubFolders, targetPath);
          if (found) return found;
        }
      }
      return null;
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
    
    // Créer la liste des dossiers monitorés (reproduire la logique exacte de l'API)
    const monitoredFolders = [];
    const allFolders = outlookConnector.isOutlookConnected ? await outlookConnector.getFolderStructure() : [];

    foldersConfig.forEach(config => {
      // Chercher le dossier dans la structure Outlook pour obtenir le nombre d'emails
      const outlookFolder = findFolderInStructure(allFolders, config.path);

      monitoredFolders.push({
        path: config.path,
        name: config.name,
        isMonitored: true,
        category: config.category || 'Mails simples',
        emailCount: outlookFolder ? outlookFolder.Count || 0 : 0,
        parentPath: getParentPath(config.path)
      });
    });

    // Calculer les statistiques
    const stats = calculateFolderStats(monitoredFolders);
    
    const result = {
      folders: monitoredFolders,
      stats: stats,
      timestamp: new Date().toISOString()
    };
    
    console.log('📊 RÉSULTAT FINAL API:', JSON.stringify(result, null, 2));
    
  } catch (error) {
    console.error('❌ Erreur debug:', error);
  }
}

// Lancer le debug
debugFoldersAPI().then(() => {
  console.log('🏁 Debug terminé');
}).catch((error) => {
  console.error('💥 Erreur debug:', error);
});
