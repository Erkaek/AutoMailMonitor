const databaseService = require('../src/services/optimizedDatabaseService');

async function checkFolders() {
  try {
    await databaseService.initialize();
    
    console.log('Dossiers testboitepartagee dans la DB:');
    const results = databaseService.db.prepare('SELECT folderPath FROM folders WHERE folderPath LIKE ?').all('%testboitepartagee%');
    results.forEach(r => console.log(' -', r.folderPath));
    
    await databaseService.close();
  } catch (error) {
    console.error('Erreur:', error.message);
  }
}

checkFolders();
