// Script pour voir tous les chemins de dossiers configurés
const Database = require('better-sqlite3');
const path = require('path');

const dbPath = path.join(__dirname, 'data', 'emails.db');

try {
    console.log("=== VERIFICATION DES CHEMINS DE DOSSIERS ===");
    
    const db = new Database(dbPath, { 
        fileMustExist: true 
    });
    
    console.log("✅ Base de données ouverte");
    
    // Voir tous les chemins dans folder_configurations
    console.log("\n=== TOUS LES CHEMINS CONFIGURES ===");
    const folders = db.prepare(`
        SELECT id, folder_path, folder_name, category, created_at, updated_at
        FROM folder_configurations 
        ORDER BY id
    `).all();
    
    folders.forEach((folder, index) => {
        console.log(`${index + 1}. [ID:${folder.id}] "${folder.folder_path}"`);
        console.log(`   Nom: "${folder.folder_name}" - Catégorie: ${folder.category}`);
        console.log(`   Créé: ${folder.created_at} - MAJ: ${folder.updated_at}`);
        console.log('');
    });
    
    console.log(`✅ ${folders.length} dossiers configurés trouvés`);
    
    // Vérifier s'il y a encore des caractères corrompus
    const corruptedFolders = folders.filter(f => 
        f.folder_path.includes('├«') || 
        f.folder_path.includes('├®') ||
        f.folder_path.includes('Ã®') ||
        f.folder_path.includes('Ã©')
    );
    
    if (corruptedFolders.length > 0) {
        console.log("\n⚠️  CHEMINS AVEC CARACTERES CORROMPUS DETECTES:");
        corruptedFolders.forEach((folder, index) => {
            console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name})`);
        });
    } else {
        console.log("\n✅ Aucun caractère corrompu détecté - tous les chemins semblent corrects!");
    }
    
    db.close();
    
} catch (error) {
    console.error("❌ Erreur:", error.message);
}
