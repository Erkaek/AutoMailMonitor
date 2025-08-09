// Script pour corriger les chemins BDD en utilisant l'app elle-même
const path = require('path');

// Import du service de base de données optimisé
const optimizedDatabaseService = require('./src/services/optimizedDatabaseService');

async function fixFolderPaths() {
    try {
        console.log("=== CORRECTION DES CHEMINS DE DOSSIERS ===");
        
        // Initialiser la base de données
        await optimizedDatabaseService.initialize();
        console.log("✅ Base de données initialisée");
        
        // Mapping des corrections
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
        
        console.log("\n=== CHEMINS ACTUELS ===");
        const currentFolders = optimizedDatabaseService.getFoldersConfig();
        currentFolders.forEach((folder, index) => {
            console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name}) - ${folder.category}`);
        });
        
        console.log("\n=== APPLICATION DES CORRECTIONS ===");
        let correctionCount = 0;
        
        for (const correction of corrections) {
            try {
                // Vérifier si le chemin existe
                const folder = currentFolders.find(f => f.folder_path === correction.old);
                if (folder) {
                    // Appliquer la correction directement via une requête SQL
                    const db = optimizedDatabaseService.getDatabase();
                    const stmt = db.prepare(`
                        UPDATE folder_config 
                        SET folder_path = ? 
                        WHERE folder_path = ?
                    `);
                    
                    const result = stmt.run(correction.new, correction.old);
                    if (result.changes > 0) {
                        console.log(`✅ Corrigé: "${correction.old}" -> "${correction.new}"`);
                        correctionCount++;
                    }
                } else {
                    console.log(`⚠️  Chemin non trouvé: "${correction.old}"`);
                }
            } catch (error) {
                console.error(`❌ Erreur correction "${correction.old}":`, error.message);
            }
        }
        
        console.log("\n=== CHEMINS APRES CORRECTION ===");
        const correctedFolders = optimizedDatabaseService.getFoldersConfig();
        correctedFolders.forEach((folder, index) => {
            console.log(`${index + 1}. "${folder.folder_path}" (${folder.folder_name}) - ${folder.category}`);
        });
        
        console.log(`\n✅ ${correctionCount} chemins corrigés dans la base de données`);
        
    } catch (error) {
        console.error("❌ Erreur lors de la correction:", error);
    }
}

// Exécuter la correction
fixFolderPaths();
