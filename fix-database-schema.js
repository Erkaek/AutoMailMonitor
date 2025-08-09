const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

// Vérifier les deux emplacements possibles
const dbPaths = [
    path.join(__dirname, 'data', 'emails.db'),
    path.join(__dirname, 'emails.db')
];

let dbPath = null;
for (const testPath of dbPaths) {
    if (fs.existsSync(testPath)) {
        dbPath = testPath;
        break;
    }
}

if (!dbPath) {
    console.log('❌ Aucune base de données trouvée dans:', dbPaths);
    process.exit(1);
}

console.log('🔍 Examen de la base de données:', dbPath);

try {
    const db = new Database(dbPath);
    
    // Vérifier la structure actuelle de la table emails
    console.log('\n📋 Structure actuelle de la table emails:');
    const emailsSchema = db.prepare("PRAGMA table_info(emails)").all();
    emailsSchema.forEach(col => {
        console.log(`  - ${col.name}: ${col.type} (${col.notnull ? 'NOT NULL' : 'NULL'}) ${col.pk ? '[PK]' : ''}`);
    });
    
    // Vérifier si la colonne week_identifier existe
    const hasWeekIdentifier = emailsSchema.find(col => col.name === 'week_identifier');
    
    if (!hasWeekIdentifier) {
        console.log('\n⚠️ La colonne week_identifier n\'existe pas. Ajout...');
        
        // Ajouter la colonne week_identifier à la table emails
        db.exec(`ALTER TABLE emails ADD COLUMN week_identifier TEXT`);
        
        console.log('✅ Colonne week_identifier ajoutée');
    } else {
        console.log('✅ La colonne week_identifier existe déjà');
    }
    
    // Vérifier si les nouvelles tables existent
    const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
    console.log('\n📊 Tables existantes:');
    tables.forEach(table => {
        console.log(`  - ${table.name}`);
    });
    
    const hasWeeklyStats = tables.find(t => t.name === 'weekly_stats');
    const hasFolderMappings = tables.find(t => t.name === 'folder_mappings');
    
    if (!hasWeeklyStats) {
        console.log('\n📅 Création de la table weekly_stats...');
        db.exec(`
            CREATE TABLE weekly_stats (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_identifier TEXT NOT NULL,
                week_number INTEGER NOT NULL,
                week_year INTEGER NOT NULL,
                week_start_date DATE NOT NULL,
                week_end_date DATE NOT NULL,
                folder_type TEXT NOT NULL,
                emails_received INTEGER DEFAULT 0,
                emails_treated INTEGER DEFAULT 0,
                manual_adjustments INTEGER DEFAULT 0,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(week_identifier, folder_type)
            )
        `);
        console.log('✅ Table weekly_stats créée');
    } else {
        console.log('✅ Table weekly_stats existe déjà');
    }
    
    if (!hasFolderMappings) {
        console.log('\n📁 Création de la table folder_mappings...');
        db.exec(`
            CREATE TABLE folder_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                original_folder_path TEXT NOT NULL,
                mapped_category TEXT NOT NULL,
                display_name TEXT,
                is_active BOOLEAN DEFAULT 1,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(original_folder_path)
            )
        `);
        console.log('✅ Table folder_mappings créée');
    } else {
        console.log('✅ Table folder_mappings existe déjà');
    }
    
    // Créer les index
    console.log('\n🔗 Création des index...');
    try {
        db.exec(`CREATE INDEX IF NOT EXISTS idx_weekly_stats_week ON weekly_stats(week_identifier, folder_type)`);
        db.exec(`CREATE INDEX IF NOT EXISTS idx_weekly_stats_date ON weekly_stats(week_start_date, week_end_date)`);
        db.exec(`CREATE INDEX IF NOT EXISTS idx_folder_mappings_path ON folder_mappings(original_folder_path)`);
        console.log('✅ Index créés');
    } catch (error) {
        console.log('⚠️ Erreur création index (peut-être déjà existants):', error.message);
    }
    
    // Mettre à jour les emails existants avec week_identifier
    console.log('\n🔄 Mise à jour des emails existants avec week_identifier...');
    
    const emailsToUpdate = db.prepare(`
        SELECT id, received_time 
        FROM emails 
        WHERE week_identifier IS NULL AND received_time IS NOT NULL
    `).all();
    
    console.log(`📧 ${emailsToUpdate.length} emails à mettre à jour`);
    
    const updateStmt = db.prepare(`UPDATE emails SET week_identifier = ? WHERE id = ?`);
    
    // Fonction pour calculer la semaine ISO (équivalent WorksheetFunction.IsoWeekNum)
    function getISOWeekInfo(date) {
        const d = new Date(date);
        
        // Cloner la date pour éviter les modifications
        const target = new Date(d.getTime());
        
        // ISO 8601: Une semaine commence le lundi
        // Ajuster au jeudi de la même semaine (garantit la bonne année ISO)
        const dayOfWeek = (target.getDay() + 6) % 7; // Convertir dimanche=0 en lundi=0
        target.setDate(target.getDate() - dayOfWeek + 3); // Jeudi de la semaine
        
        // Le 1er janvier de l'année ISO du jeudi
        const jan1 = new Date(target.getFullYear(), 0, 1);
        
        // Calculer le numéro de semaine
        const weekNum = Math.ceil((((target - jan1) / 86400000) + jan1.getDay() + 1) / 7);
        
        // Date de début de semaine (lundi)
        const startOfWeek = new Date(d.getTime());
        startOfWeek.setDate(d.getDate() - dayOfWeek);
        
        // Date de fin de semaine (dimanche) 
        const endOfWeek = new Date(startOfWeek.getTime());
        endOfWeek.setDate(startOfWeek.getDate() + 6);
        
        const weekYear = target.getFullYear();
        const weekIdentifier = `S${weekNum.toString().padStart(2, '0')}-${weekYear}`;
        
        return {
            weekNumber: weekNum,
            weekYear: weekYear,
            weekIdentifier: weekIdentifier,
            weekStartDate: startOfWeek.toISOString().split('T')[0],
            weekEndDate: endOfWeek.toISOString().split('T')[0],
            dateRange: `${startOfWeek.toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit' })} au ${endOfWeek.toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' })}`
        };
    }
    
    let updated = 0;
    for (const email of emailsToUpdate) {
        try {
            const weekInfo = getISOWeekInfo(email.received_time);
            updateStmt.run(weekInfo.weekIdentifier, email.id);
            updated++;
        } catch (error) {
            console.log(`⚠️ Erreur mise à jour email ${email.id}:`, error.message);
        }
    }
    
    console.log(`✅ ${updated} emails mis à jour avec week_identifier`);
    
    // Vérification finale
    console.log('\n🔍 Vérification finale...');
    const finalCheck = db.prepare("PRAGMA table_info(emails)").all();
    const hasWeekCol = finalCheck.find(col => col.name === 'week_identifier');
    
    if (hasWeekCol) {
        const countWithWeek = db.prepare("SELECT COUNT(*) as count FROM emails WHERE week_identifier IS NOT NULL").get();
        console.log(`✅ Base de données corrigée ! ${countWithWeek.count} emails ont un week_identifier`);
    } else {
        console.log('❌ Problème: la colonne week_identifier n\'a pas été ajoutée');
    }
    
    db.close();
    console.log('\n🎉 Migration terminée avec succès !');
    
} catch (error) {
    console.error('❌ Erreur:', error);
    process.exit(1);
}
