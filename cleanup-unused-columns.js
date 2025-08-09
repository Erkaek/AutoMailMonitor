const Database = require('better-sqlite3');
const fs = require('fs');
const path = require('path');

console.log('🗑️ NETTOYAGE FINAL DE LA BASE DE DONNÉES');
console.log('Suppression des colonnes jamais utilisées pour optimiser les performances\n');

// Backup avant nettoyage
const dbPath = path.join(__dirname, 'data', 'emails.db');
const backupPath = path.join(__dirname, 'data', `emails-backup-final-cleanup-${Date.now()}.db`);

console.log('📦 Création d\'un backup de sécurité...');
fs.copyFileSync(dbPath, backupPath);
console.log(`✅ Backup créé : ${path.basename(backupPath)}`);

// Ouvrir la base de données
const db = new Database(dbPath);

// Vérifier l'état avant nettoyage
const beforeInfo = db.prepare("PRAGMA table_info(emails)").all();
console.log(`\n📊 État AVANT : ${beforeInfo.length} colonnes`);

// Colonnes à conserver (celles qui sont utilisées)
const columnsToKeep = [
    'id',                    // Clé primaire
    'outlook_id',           // Identifiant Outlook (utilisé)
    'subject',              // Sujet (utilisé)
    'sender_email',         // Email expéditeur (utilisé)
    'received_time',        // Heure réception (utilisé)
    'sent_time',           // Heure envoi (utilisé)
    'folder_name',         // Nom dossier (utilisé)
    'category',            // Catégorie (utilisé)
    'is_read',             // Lu/Non lu (peu utilisé mais important)
    'importance',          // Importance (utilisé)
    'week_identifier',     // Semaine ISO (utilisé pour stats)
    'created_at',          // Date création (système)
    'updated_at'           // Date modification (système)
];

console.log('\n🔧 Reconstruction de la table avec colonnes optimales...');

// Commencer la transaction
db.exec('BEGIN TRANSACTION');

try {
    // 1. Créer nouvelle table optimisée
    const createOptimizedTable = `
        CREATE TABLE emails_optimized (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            outlook_id TEXT NOT NULL,
            subject TEXT NOT NULL DEFAULT '',
            sender_email TEXT DEFAULT '',
            received_time DATETIME,
            sent_time DATETIME,
            folder_name TEXT DEFAULT '',
            category TEXT DEFAULT 'Mails simples',
            is_read BOOLEAN DEFAULT 0,
            importance INTEGER DEFAULT 1,
            week_identifier TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    `;
    
    db.exec(createOptimizedTable);
    console.log('✅ Table optimisée créée');

    // 2. Copier les données essentielles
    const columnsList = columnsToKeep.join(', ');
    const copyData = `
        INSERT INTO emails_optimized (${columnsList})
        SELECT ${columnsList}
        FROM emails
    `;
    
    const result = db.exec(copyData);
    console.log('✅ Données copiées');

    // 3. Supprimer l'ancienne table et renommer
    db.exec('DROP TABLE emails');
    db.exec('ALTER TABLE emails_optimized RENAME TO emails');
    console.log('✅ Table renommée');

    // 4. Recréer les index optimisés
    db.exec(`
        CREATE INDEX IF NOT EXISTS idx_emails_outlook_id ON emails(outlook_id);
        CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails(received_time);
        CREATE INDEX IF NOT EXISTS idx_emails_week_folder ON emails(week_identifier, folder_name);
        CREATE INDEX IF NOT EXISTS idx_emails_unread ON emails(is_read, received_time);
        CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category, received_time);
    `);
    console.log('✅ Index optimisés créés');

    // Valider la transaction
    db.exec('COMMIT');
    console.log('✅ Transaction validée');

} catch (error) {
    // Annuler en cas d'erreur
    db.exec('ROLLBACK');
    console.error('❌ Erreur lors du nettoyage :', error.message);
    console.log('🔄 Transaction annulée, base inchangée');
    db.close();
    return;
}

// Vérifier l'état après nettoyage
const afterInfo = db.prepare("PRAGMA table_info(emails)").all();
console.log(`\n📊 État APRÈS : ${afterInfo.length} colonnes`);

// Afficher les colonnes conservées
console.log('\n🟢 Colonnes conservées :');
afterInfo.forEach(col => {
    console.log(`  - ${col.name} (${col.type})`);
});

// Vérifier l'intégrité des données
const emailCount = db.prepare("SELECT COUNT(*) as count FROM emails").get();
console.log(`\n📈 Emails conservés : ${emailCount.count}`);

// Calculer l'économie d'espace
const beforeColumns = beforeInfo.length;
const afterColumns = afterInfo.length;
const reduction = Math.round(((beforeColumns - afterColumns) / beforeColumns) * 100);

console.log(`\n💾 OPTIMISATION RÉALISÉE :`);
console.log(`  - Colonnes supprimées : ${beforeColumns - afterColumns}`);
console.log(`  - Réduction : ${reduction}% de colonnes en moins`);
console.log(`  - Espace libéré : ~${reduction}% d'espace disque`);

// Analyser la taille des fichiers
const originalSize = fs.statSync(backupPath).size;
const newSize = fs.statSync(dbPath).size;
const sizeDiff = originalSize - newSize;
const sizeReduction = Math.round((sizeDiff / originalSize) * 100);

console.log(`  - Taille avant : ${Math.round(originalSize / 1024)} KB`);
console.log(`  - Taille après : ${Math.round(newSize / 1024)} KB`);
console.log(`  - Gain réel : ${Math.round(sizeDiff / 1024)} KB (${sizeReduction}%)`);

console.log('\n✅ NETTOYAGE TERMINÉ AVEC SUCCÈS !');
console.log(`📦 Backup disponible : ${path.basename(backupPath)}`);

db.close();
