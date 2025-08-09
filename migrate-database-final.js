const Database = require('better-sqlite3');
const path = require('path');

console.log('🚀 MIGRATION COMPLÈTE - OPTIMISATION BASE DE DONNÉES');
console.log('=====================================================');

try {
    const dbPath = './data/emails.db';
    const db = new Database(dbPath);
    
    console.log('📊 1. Analyse de la structure actuelle...');
    
    // Vérifier la structure actuelle
    const currentSchema = db.prepare("PRAGMA table_info(emails)").all();
    console.log(`   Structure actuelle: ${currentSchema.length} colonnes`);
    
    // Sauvegarder les données existantes
    console.log('💾 2. Sauvegarde des données existantes...');
    const existingEmails = db.prepare(`
        SELECT id, outlook_id, subject, sender_email, received_time, 
               folder_name, category, is_read, week_identifier,
               created_at, updated_at, NULL as deleted_at
        FROM emails
    `).all();
    
    console.log(`   ${existingEmails.length} emails sauvegardés`);
    
    // Créer la nouvelle structure
    console.log('🔧 3. Création de la nouvelle structure...');
    
    // Supprimer l'ancienne table
    db.exec('DROP TABLE IF EXISTS emails_old');
    db.exec('ALTER TABLE emails RENAME TO emails_old');
    
    // Créer la nouvelle table optimisée
    db.exec(`
        CREATE TABLE emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            outlook_id TEXT UNIQUE NOT NULL,
            subject TEXT NOT NULL,
            sender_email TEXT,
            received_time DATETIME,
            folder_name TEXT,
            category TEXT,
            is_read BOOLEAN DEFAULT 0,
            is_treated BOOLEAN DEFAULT 0,
            deleted_at DATETIME NULL,
            week_identifier TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    `);
    
    // Index optimisés
    db.exec(`
        CREATE INDEX IF NOT EXISTS idx_emails_outlook_id ON emails(outlook_id);
        CREATE INDEX IF NOT EXISTS idx_emails_folder_name ON emails(folder_name);
        CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails(received_time);
        CREATE INDEX IF NOT EXISTS idx_emails_is_read ON emails(is_read);
        CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category);
        CREATE INDEX IF NOT EXISTS idx_emails_week_identifier ON emails(week_identifier);
        CREATE INDEX IF NOT EXISTS idx_emails_deleted_at ON emails(deleted_at);
        CREATE INDEX IF NOT EXISTS idx_emails_is_treated ON emails(is_treated);
    `);
    
    console.log('   ✅ Nouvelle structure créée avec index optimisés');
    
    // Migrer les données
    console.log('📦 4. Migration des données...');
    
    const insertStmt = db.prepare(`
        INSERT INTO emails 
        (id, outlook_id, subject, sender_email, received_time, folder_name, 
         category, is_read, is_treated, deleted_at, week_identifier, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    
    let migratedCount = 0;
    const transaction = db.transaction(() => {
        for (const email of existingEmails) {
            try {
                insertStmt.run(
                    email.id,
                    email.outlook_id,
                    email.subject,
                    email.sender_email,
                    email.received_time,
                    email.folder_name,
                    email.category,
                    email.is_read ? 1 : 0,
                    0, // is_treated par défaut à false
                    email.deleted_at,
                    email.week_identifier,
                    email.created_at,
                    email.updated_at
                );
                migratedCount++;
            } catch (error) {
                console.log(`   ⚠️  Erreur migration email ID ${email.id}: ${error.message}`);
            }
        }
    });
    
    transaction();
    
    console.log(`   ✅ ${migratedCount}/${existingEmails.length} emails migrés`);
    
    // Supprimer l'ancienne table
    db.exec('DROP TABLE emails_old');
    
    // Vérification finale
    console.log('🔍 5. Vérification finale...');
    const finalCount = db.prepare('SELECT COUNT(*) as count FROM emails').get().count;
    const finalSchema = db.prepare("PRAGMA table_info(emails)").all();
    
    console.log(`   Structure finale: ${finalSchema.length} colonnes`);
    console.log(`   Emails conservés: ${finalCount}`);
    
    // Afficher la nouvelle structure
    console.log('\n📋 NOUVELLE STRUCTURE:');
    finalSchema.forEach(col => {
        console.log(`   ${col.cid.toString().padEnd(2)} | ${col.name.padEnd(20)} | ${col.type.padEnd(10)} | ${col.notnull ? 'NOT NULL' : 'NULL'}`);
    });
    
    console.log('\n✅ MIGRATION TERMINÉE AVEC SUCCÈS !');
    console.log('=====================================================');
    console.log('🎯 Structure optimisée pour la logique VBA :');
    console.log('   - Arrivées = week_identifier');
    console.log('   - Traités = deleted_at OU is_read (selon config)');
    console.log('   - 13 colonnes essentielles seulement');
    
    db.close();
    
} catch (error) {
    console.error('❌ Erreur lors de la migration:', error.message);
    console.error(error.stack);
}
