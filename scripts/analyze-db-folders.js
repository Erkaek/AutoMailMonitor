const Database = require('better-sqlite3');
const path = require('path');

const dbPath = path.join(__dirname, '..', 'data', 'emails.db');
const db = new Database(dbPath, { readonly: true });

try {
  // Récupérer tous les folder_name de testboitepartagee
  const folders = db.prepare(`
    SELECT DISTINCT folder_name, COUNT(*) as email_count
    FROM emails 
    WHERE folder_name LIKE '%testboitepartagee%' 
    ORDER BY folder_name
  `).all();
  
  console.log('=== DOSSIERS testboitepartagee DANS LA DB ===');
  folders.forEach(row => {
    console.log(`${row.folder_name} (${row.email_count} emails)`);
  });
  
  // Analyser la structure
  console.log('\n=== STRUCTURE DEDUITE ===');
  const structure = {};
  folders.forEach(row => {
    const parts = row.folder_name.split('\\');
    if (parts.length >= 3) {
      const mailbox = parts[0];
      const parentFolder = parts[1]; 
      const childFolder = parts[2];
      
      if (!structure[mailbox]) structure[mailbox] = {};
      if (!structure[mailbox][parentFolder]) structure[mailbox][parentFolder] = [];
      
      if (!structure[mailbox][parentFolder].includes(childFolder)) {
        structure[mailbox][parentFolder].push(childFolder);
      }
    }
  });
  
  console.log(JSON.stringify(structure, null, 2));
  
} catch (error) {
  console.error('Erreur:', error.message);
} finally {
  db.close();
}
