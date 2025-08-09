// Script pour explorer la structure de la BDD
const Database = require('better-sqlite3');
const path = require('path');

const dbPath = path.join(__dirname, 'data', 'emails.db');

try {
    console.log("=== EXPLORATION DE LA BASE DE DONNEES ===");
    
    const db = new Database(dbPath, { 
        fileMustExist: true 
    });
    
    console.log("✅ Base de données ouverte");
    
    // Lister toutes les tables
    console.log("\n=== TABLES DISPONIBLES ===");
    const tables = db.prepare(`
        SELECT name FROM sqlite_master 
        WHERE type='table' 
        ORDER BY name
    `).all();
    
    tables.forEach((table, index) => {
        console.log(`${index + 1}. ${table.name}`);
    });
    
    // Explorer chaque table pour trouver les configurations de dossiers
    console.log("\n=== EXPLORATION DES TABLES ===");
    for (const table of tables) {
        console.log(`\n--- Table: ${table.name} ---`);
        try {
            // Obtenir la structure
            const columns = db.prepare(`PRAGMA table_info(${table.name})`).all();
            console.log("Colonnes:", columns.map(c => `${c.name} (${c.type})`).join(", "));
            
            // Obtenir quelques échantillons de données
            const sampleData = db.prepare(`SELECT * FROM ${table.name} LIMIT 3`).all();
            if (sampleData.length > 0) {
                console.log("Échantillon de données:");
                sampleData.forEach((row, i) => {
                    console.log(`  ${i + 1}:`, JSON.stringify(row, null, 2));
                });
            } else {
                console.log("Aucune donnée dans cette table");
            }
        } catch (error) {
            console.log(`Erreur exploration ${table.name}:`, error.message);
        }
    }
    
    db.close();
    
} catch (error) {
    console.error("❌ Erreur:", error.message);
}
