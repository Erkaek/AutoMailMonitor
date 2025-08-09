const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

// Trouver la base de données
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

console.log('🔧 Correction des tables de suivi hebdomadaire:', dbPath);

try {
    const db = new Database(dbPath);
    
    // Sauvegarder les données existantes de weekly_stats si elles existent
    console.log('💾 Sauvegarde des données existantes...');
    let existingData = [];
    try {
        existingData = db.prepare("SELECT * FROM weekly_stats").all();
        console.log(`📊 ${existingData.length} entrées trouvées dans weekly_stats`);
    } catch (error) {
        console.log('⚠️ Pas de données à sauvegarder:', error.message);
    }
    
    // Supprimer les anciennes tables
    console.log('🗑️ Suppression des anciennes tables...');
    try {
        db.exec("DROP TABLE IF EXISTS weekly_stats");
        console.log('✅ Ancienne table weekly_stats supprimée');
    } catch (error) {
        console.log('⚠️ Erreur suppression weekly_stats:', error.message);
    }
    
    try {
        db.exec("DROP TABLE IF EXISTS folder_mappings");
        console.log('✅ Ancienne table folder_mappings supprimée');
    } catch (error) {
        console.log('⚠️ Erreur suppression folder_mappings:', error.message);
    }
    
    // Recréer les tables avec la bonne structure
    console.log('🏗️ Création des nouvelles tables...');
    
    // Table weekly_stats avec la structure correcte
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
    console.log('✅ Nouvelle table weekly_stats créée');
    
    // Table folder_mappings avec la structure correcte
    db.exec(`
        CREATE TABLE IF NOT EXISTS folder_mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_folder_path TEXT NOT NULL,
            mapped_category TEXT NOT NULL,
            display_name TEXT,
            is_active BOOLEAN DEFAULT 1,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(original_folder_path)
        )
    `);
    console.log('✅ Nouvelle table folder_mappings créée');
    
    // Créer les index
    console.log('🔗 Création des index...');
    db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_week ON emails(week_identifier)`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_weekly_stats_week ON weekly_stats(week_identifier, folder_type)`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_weekly_stats_date ON weekly_stats(week_start_date, week_end_date)`);
    db.exec(`CREATE INDEX IF NOT EXISTS idx_folder_mappings_path ON folder_mappings(original_folder_path)`);
    console.log('✅ Index créés');
    
    // Fonction pour calculer la semaine ISO
    function getISOWeekInfo(date) {
        const d = new Date(date);
        const target = new Date(d.getTime());
        const dayOfWeek = (target.getDay() + 6) % 7;
        target.setDate(target.getDate() - dayOfWeek + 3);
        const jan1 = new Date(target.getFullYear(), 0, 1);
        const weekNum = Math.ceil((((target - jan1) / 86400000) + jan1.getDay() + 1) / 7);
        const startOfWeek = new Date(d.getTime());
        startOfWeek.setDate(d.getDate() - dayOfWeek);
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
    
    // Initialiser les statistiques de la semaine actuelle
    console.log('📊 Initialisation des statistiques hebdomadaires...');
    
    const currentWeek = getISOWeekInfo(new Date());
    console.log('📅 Semaine actuelle:', currentWeek.weekIdentifier, '-', currentWeek.dateRange);
    
    const insertWeeklyStats = db.prepare(`
        INSERT OR REPLACE INTO weekly_stats 
        (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type, emails_received, emails_treated)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `);
    
    // Analyser les emails existants et créer les statistiques
    const emailsByWeekAndCategory = db.prepare(`
        SELECT 
            week_identifier,
            COALESCE(category, 'Sans catégorie') as category,
            COUNT(*) as received_count,
            SUM(CASE WHEN is_treated = 1 OR is_read = 1 THEN 1 ELSE 0 END) as treated_count
        FROM emails 
        WHERE week_identifier IS NOT NULL
        GROUP BY week_identifier, COALESCE(category, 'Sans catégorie')
    `).all();
    
    console.log(`📧 Trouvé ${emailsByWeekAndCategory.length} groupes d'emails à traiter`);
    
    for (const group of emailsByWeekAndCategory) {
        const weekInfo = getISOWeekInfo(new Date()); // Pour la structure, on prend la semaine actuelle
        
        // Extraire les infos de la semaine depuis week_identifier
        const weekMatch = group.week_identifier.match(/S(\d+)-(\d+)/);
        if (weekMatch) {
            const weekNum = parseInt(weekMatch[1]);
            const weekYear = parseInt(weekMatch[2]);
            
            // Calculer les dates de début/fin pour cette semaine
            const firstDayOfYear = new Date(weekYear, 0, 1);
            const daysToFirstMonday = (8 - firstDayOfYear.getDay()) % 7;
            const firstMonday = new Date(weekYear, 0, 1 + daysToFirstMonday);
            const weekStart = new Date(firstMonday.getTime() + (weekNum - 1) * 7 * 24 * 60 * 60 * 1000);
            const weekEnd = new Date(weekStart.getTime() + 6 * 24 * 60 * 60 * 1000);
            
            insertWeeklyStats.run(
                group.week_identifier,
                weekNum,
                weekYear,
                weekStart.toISOString().split('T')[0],
                weekEnd.toISOString().split('T')[0],
                group.category,
                group.received_count,
                group.treated_count
            );
            
            console.log(`  ✅ ${group.week_identifier} - ${group.category}: ${group.received_count} reçus, ${group.treated_count} traités`);
        }
    }
    
    // Ajouter quelques mappings de dossiers par défaut
    console.log('\n📁 Ajout des mappings de dossiers par défaut...');
    
    const insertFolderMapping = db.prepare(`
        INSERT OR IGNORE INTO folder_mappings (original_folder_path, mapped_category, display_name)
        VALUES (?, ?, ?)
    `);
    
    const defaultMappings = [
        ['Boîte de réception', 'Général', 'Boîte de réception'],
        ['Inbox', 'Général', 'Boîte de réception'],
        ['Personnel', 'Personnel', 'Personnel'],
        ['Travail', 'Professionnel', 'Travail'],
        ['Work', 'Professionnel', 'Travail'],
        ['Newsletter', 'Newsletter', 'Newsletters'],
        ['Promotion', 'Publicité', 'Promotions'],
        ['Spam', 'Spam', 'Courriers indésirables'],
        ['Junk', 'Spam', 'Courriers indésirables']
    ];
    
    for (const [path, category, name] of defaultMappings) {
        insertFolderMapping.run(path, category, name);
    }
    
    console.log('✅ Mappings de dossiers ajoutés');
    
    // Vérification finale
    const totalStats = db.prepare(`SELECT COUNT(*) as count FROM weekly_stats`).get();
    const totalMappings = db.prepare(`SELECT COUNT(*) as count FROM folder_mappings`).get();
    
    console.log(`\n📈 Total des statistiques hebdomadaires: ${totalStats.count} entrées`);
    console.log(`📁 Total des mappings de dossiers: ${totalMappings.count} entrées`);
    
    db.close();
    console.log('\n🎉 Tables de suivi hebdomadaire entièrement configurées !');
    
} catch (error) {
    console.error('❌ Erreur:', error);
    process.exit(1);
}
