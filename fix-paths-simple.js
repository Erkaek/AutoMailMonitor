// Script simple pour corriger les chemins
const Database = require('better-sqlite3');
const path = require('path');

const dbPath = path.join(__dirname, 'data', 'emails.db');

try {
    console.log("=== CORRECTION DES CHEMINS DE DOSSIERS ===");
    
    const db = new Database(dbPath, { 
        verbose: console.log,
        fileMustExist: true 
    });
    
    console.log("✅ Base de données ouverte");
    
    // Voir les chemins actuels
    console.log("\n=== CHEMINS ACTUELS ===");
    const folders = db.prepare(`
        SELECT folder_path, folder_name, category 
        FROM folder_config 
        ORDER BY folder_path
    `).all();
    
    folders.forEach((folder, index) => {
        console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name}) - ${folder.category}`);
    });
    
    // Corrections
    console.log("\n=== APPLICATION DES CORRECTIONS ===");
    const corrections = [
        {
            old: 'erkaekanon@outlook.com\\Bo├«te de r├®ception\\testA',
            new: 'erkaekanon@outlook.com\\Boîte de réception\\testA'
        },
        {
            old: 'erkaekanon@outlook.com\\Bo├«te de r├®ception\\test',
            new: 'erkaekanon@outlook.com\\Boîte de réception\\test'
        },
        {
            old: 'erkaekanon@outlook.com\\Bo├«te de r├®ception\\test\\test-1',
            new: 'erkaekanon@outlook.com\\Boîte de réception\\test\\test-1'
        },
        {
            old: 'erkaekanon@outlook.com\\Bo├«te de r├®ception\\testA\\test-c',
            new: 'erkaekanon@outlook.com\\Boîte de réception\\testA\\test-c'
        }
    ];
    
    const updateStmt = db.prepare(`
        UPDATE folder_config 
        SET folder_path = ? 
        WHERE folder_path = ?
    `);
    
    let correctionCount = 0;
    for (const correction of corrections) {
        const result = updateStmt.run(correction.new, correction.old);
        if (result.changes > 0) {
            console.log(`✅ Corrigé: "${correction.old}" -> "${correction.new}"`);
            correctionCount++;
        } else {
            console.log(`⚠️  Aucun changement pour: "${correction.old}"`);
        }
    }
    
    // Voir les chemins après correction
    console.log("\n=== CHEMINS APRES CORRECTION ===");
    const foldersAfter = db.prepare(`
        SELECT folder_path, folder_name, category 
        FROM folder_config 
        ORDER BY folder_path
    `).all();
    
    foldersAfter.forEach((folder, index) => {
        console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name}) - ${folder.category}`);
    });
    
    console.log(`\n✅ ${correctionCount} chemins corrigés dans la base de données`);
    
    db.close();
    
} catch (error) {
    console.error("❌ Erreur:", error.message);
}
