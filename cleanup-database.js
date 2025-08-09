const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

console.log('🧹 NETTOYAGE COMPLET DE LA BASE DE DONNÉES');
console.log('='.repeat(50));

const dataDbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dataDbPath);

console.log('\n🗑️ SUPPRESSION DES TABLES INUTILES...');

// Tables à supprimer (vides ou obsolètes)
const tablesToDrop = ['app_config', 'email_events', 'metrics_history'];

for (const table of tablesToDrop) {
    try {
        const count = db.prepare(`SELECT COUNT(*) as count FROM ${table}`).get();
        console.log(`  Suppression de ${table} (${count.count} enregistrements)...`);
        db.exec(`DROP TABLE IF EXISTS ${table}`);
        console.log(`  ✅ ${table} supprimée`);
    } catch (error) {
        console.log(`  ❌ Erreur suppression ${table}: ${error.message}`);
    }
}

console.log('\n🔧 NETTOYAGE DE LA TABLE EMAILS...');

// D'abord, sauvegarder les données importantes
const emailsData = db.prepare(`
    SELECT 
        id, outlook_id, subject, sender_name, sender_email, recipient_email,
        received_time, sent_time, treated_time, folder_name, category,
        is_read, is_replied, has_attachment, body_preview, importance,
        COALESCE(size_kb, size) as size_bytes,
        event_type, created_at, updated_at, folder_type, entry_id,
        COALESCE(is_treated, CASE WHEN treated_time IS NOT NULL THEN 1 ELSE 0 END) as is_treated,
        folder_path, week_identifier
    FROM emails
    WHERE deleted_at IS NULL AND (is_deleted IS NULL OR is_deleted = 0)
`).all();

console.log(`  📧 ${emailsData.length} emails valides trouvés`);

// Créer une nouvelle table emails propre
console.log('  🏗️ Création d\'une nouvelle table emails optimisée...');

db.exec(`DROP TABLE IF EXISTS emails_clean`);
db.exec(`
    CREATE TABLE emails_clean (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        outlook_id TEXT UNIQUE,
        entry_id TEXT,
        subject TEXT NOT NULL,
        sender_name TEXT,
        sender_email TEXT,
        recipient_email TEXT,
        received_time DATETIME,
        sent_time DATETIME,
        treated_time DATETIME,
        folder_name TEXT,
        folder_path TEXT,
        category TEXT,
        is_read BOOLEAN DEFAULT 0,
        is_replied BOOLEAN DEFAULT 0,
        is_treated BOOLEAN DEFAULT 0,
        has_attachment BOOLEAN DEFAULT 0,
        body_preview TEXT,
        importance INTEGER DEFAULT 1,
        size_bytes INTEGER,
        week_identifier TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
`);

// Insérer les données nettoyées
console.log('  📥 Insertion des données nettoyées...');
const insertClean = db.prepare(`
    INSERT INTO emails_clean (
        outlook_id, entry_id, subject, sender_name, sender_email, recipient_email,
        received_time, sent_time, treated_time, folder_name, folder_path, category,
        is_read, is_replied, is_treated, has_attachment, body_preview, importance,
        size_bytes, week_identifier, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
`);

let inserted = 0;
for (const email of emailsData) {
    try {
        insertClean.run(
            email.outlook_id,
            email.entry_id,
            email.subject,
            email.sender_name,
            email.sender_email,
            email.recipient_email,
            email.received_time,
            email.sent_time,
            email.treated_time,
            email.folder_name,
            email.folder_path,
            email.category,
            email.is_read,
            email.is_replied,
            email.is_treated,
            email.has_attachment,
            email.body_preview,
            email.importance,
            email.size_bytes,
            email.week_identifier,
            email.created_at,
            email.updated_at
        );
        inserted++;
    } catch (error) {
        console.log(`  ⚠️ Erreur insertion email ${email.id}: ${error.message}`);
    }
}

console.log(`  ✅ ${inserted} emails insérés dans la table propre`);

// Remplacer l'ancienne table
console.log('  🔄 Remplacement de l\'ancienne table...');
db.exec(`DROP TABLE emails`);
db.exec(`ALTER TABLE emails_clean RENAME TO emails`);

// Créer les index optimisés
console.log('  🔗 Création des index optimisés...');
db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_outlook_id ON emails(outlook_id)`);
db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails(received_time)`);
db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_folder ON emails(folder_name)`);
db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_week ON emails(week_identifier)`);
db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category)`);
db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_treated ON emails(is_treated, treated_time)`);

console.log('\n📊 VÉRIFICATION DE LA TABLE WEEKLY_STATS...');

// Vérifier et nettoyer weekly_stats
const weeklyStatsCount = db.prepare(`SELECT COUNT(*) as count FROM weekly_stats`).get();
console.log(`  📈 ${weeklyStatsCount.count} enregistrements dans weekly_stats`);

if (weeklyStatsCount.count > 0) {
    const weeklyData = db.prepare(`SELECT * FROM weekly_stats LIMIT 3`).all();
    console.log('  📋 Exemples de données:');
    weeklyData.forEach((row, index) => {
        console.log(`    ${index + 1}. ${row.week_identifier} - ${row.folder_type}: ${row.emails_received}/${row.emails_treated}`);
    });
}

console.log('\n📁 VÉRIFICATION DE FOLDER_MAPPINGS...');

// Vérifier folder_mappings
const folderMappingsCount = db.prepare(`SELECT COUNT(*) as count FROM folder_mappings`).get();
console.log(`  📂 ${folderMappingsCount.count} mappings de dossiers`);

if (folderMappingsCount.count > 0) {
    const mappings = db.prepare(`SELECT original_folder_path, mapped_category FROM folder_mappings WHERE is_active = 1`).all();
    console.log('  📋 Mappings actifs:');
    mappings.forEach((mapping, index) => {
        console.log(`    ${index + 1}. ${mapping.original_folder_path} → ${mapping.mapped_category}`);
    });
}

console.log('\n🔧 CONFIGURATION DE L\'APPLICATION...');

// Vérifier que l'application utilise le bon chemin
const configFiles = [
    'src/services/optimizedDatabaseService.js',
    'src/main/index.js'
];

for (const configFile of configFiles) {
    const fullPath = path.join(__dirname, configFile);
    if (fs.existsSync(fullPath)) {
        console.log(`  📄 Vérifié: ${configFile} existe`);
    } else {
        console.log(`  ❌ Manquant: ${configFile}`);
    }
}

// Statistiques finales
console.log('\n📊 STATISTIQUES FINALES:');

const finalTables = db.prepare(`
    SELECT name 
    FROM sqlite_master 
    WHERE type='table' AND name NOT LIKE 'sqlite_%'
    ORDER BY name
`).all();

console.log(`  📋 Tables restantes: ${finalTables.map(t => t.name).join(', ')}`);

const finalEmailCount = db.prepare(`SELECT COUNT(*) as count FROM emails`).get();
console.log(`  📧 Emails: ${finalEmailCount.count}`);

const finalWeeklyCount = db.prepare(`SELECT COUNT(*) as count FROM weekly_stats`).get();
console.log(`  📊 Stats hebdomadaires: ${finalWeeklyCount.count}`);

const finalMappingsCount = db.prepare(`SELECT COUNT(*) as count FROM folder_mappings`).get();
console.log(`  📁 Mappings dossiers: ${finalMappingsCount.count}`);

// Calculer la taille de la base
const dbStats = fs.statSync(dataDbPath);
console.log(`  💾 Taille base: ${(dbStats.size / 1024).toFixed(1)} KB`);

db.close();

console.log('\n🎉 NETTOYAGE TERMINÉ !');
console.log('\n✅ La base de données est maintenant propre et optimisée');
console.log('📍 Emplacement: data/emails.db');
console.log('\n🔧 PROCHAINES ÉTAPES:');
console.log('1. Reconstruire better-sqlite3 pour Electron: npx electron-rebuild -f -w better-sqlite3');
console.log('2. Tester l\'application: npm start');
console.log('3. Vérifier le suivi hebdomadaire dans l\'interface');
