const Database = require('better-sqlite3');
const path = require('path');

try {
  const dbPath = path.join(__dirname, 'data', 'emails.db');
  console.log('📂 Ouverture de la base de données:', dbPath);
  
  const db = new Database(dbPath, { readonly: true });
  
  console.log('\n📋 Dossiers configurés dans la BDD:');
  console.log('=====================================');
  
  try {
    const folders = db.prepare('SELECT * FROM folders_config').all();
    
    if (folders.length === 0) {
      console.log('❌ Aucun dossier configuré dans la table folders_config');
    } else {
      folders.forEach((folder, index) => {
        console.log(`${index + 1}. Dossier: "${folder.folderPath}"`);
        console.log(`   - Catégorie: ${folder.category}`);
        console.log(`   - Surveillé: ${folder.isActive ? 'OUI' : 'NON'}`);
        console.log(`   - Créé: ${new Date(folder.createdAt).toLocaleString()}`);
        console.log('');
      });
    }
  } catch (error) {
    console.log('❌ Erreur lecture table folders_config:', error.message);
    
    // Essayer de voir toutes les tables
    console.log('\n📋 Tables disponibles:');
    const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
    tables.forEach(table => console.log(`  - ${table.name}`));
  }
  
  db.close();
  
} catch (error) {
  console.error('❌ Erreur:', error.message);
}
