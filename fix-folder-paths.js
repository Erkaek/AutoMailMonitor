// Script pour corriger les chemins des dossiers dans la BDD
const path = require('path');
const Database = require('better-sqlite3');

// Connexion à la base de données
const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

console.log("=== VERIFICATION CHEMINS ACTUELS ===");

// Lire les chemins actuels
const folders = db.prepare(`
    SELECT folder_path, folder_name, category 
    FROM folder_config 
    ORDER BY folder_path
`).all();

console.log("Chemins actuels dans la BDD:");
folders.forEach((folder, index) => {
    console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name}) - ${folder.category}`);
});

console.log("\n=== CORRECTION DES CHEMINS ===");

// Mapping des corrections
const corrections = {
    'erkaekanon@outlook.com\\Bo├«te de r├®ception\\testA': 'erkaekanon@outlook.com\\Boîte de réception\\testA',
    'erkaekanon@outlook.com\\Bo├«te de r├®ception\\test': 'erkaekanon@outlook.com\\Boîte de réception\\test',
    'erkaekanon@outlook.com\\Bo├«te de r├®ception\\test\\test-1': 'erkaekanon@outlook.com\\Boîte de réception\\test\\test-1',
    'erkaekanon@outlook.com\\Bo├«te de r├®ception\\testA\\test-c': 'erkaekanon@outlook.com\\Boîte de réception\\testA\\test-c'
};

// Appliquer les corrections
const updateStmt = db.prepare(`
    UPDATE folder_config 
    SET folder_path = ? 
    WHERE folder_path = ?
`);

let correctionCount = 0;
for (const [oldPath, newPath] of Object.entries(corrections)) {
    const result = updateStmt.run(newPath, oldPath);
    if (result.changes > 0) {
        console.log(`✅ Corrigé: "${oldPath}" -> "${newPath}"`);
        correctionCount++;
    } else {
        console.log(`⚠️  Aucun changement pour: "${oldPath}"`);
    }
}

console.log(`\n=== VERIFICATION APRES CORRECTION ===`);

// Vérifier les nouveaux chemins
const foldersAfter = db.prepare(`
    SELECT folder_path, folder_name, category 
    FROM folder_config 
    ORDER BY folder_path
`).all();

console.log("Nouveaux chemins dans la BDD:");
foldersAfter.forEach((folder, index) => {
    console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name}) - ${folder.category}`);
});

console.log(`\n✅ ${correctionCount} chemins corrigés dans la base de données`);

db.close();
