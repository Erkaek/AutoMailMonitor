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

if (!dbPath) {
    console.log('❌ Aucune base de données trouvée');
    process.exit(1);
}

console.log('🔧 Création des index finaux pour:', dbPath);

try {
    const db = new Database(dbPath);
    
    // Vérifier que week_identifier existe
    const emailsSchema = db.prepare("PRAGMA table_info(emails)").all();
    const hasWeekIdentifier = emailsSchema.find(col => col.name === 'week_identifier');
    
    if (!hasWeekIdentifier) {
        console.log('❌ Colonne week_identifier manquante');
        process.exit(1);
    }
    
    console.log('✅ Colonne week_identifier trouvée');
    
    // Créer les index manquants
    console.log('🔗 Création des index...');
    
    try {
        db.exec(`CREATE INDEX IF NOT EXISTS idx_emails_week ON emails(week_identifier)`);
        console.log('✅ Index idx_emails_week créé');
    } catch (error) {
        console.log('⚠️ Index idx_emails_week:', error.message);
    }
    
    try {
        db.exec(`CREATE INDEX IF NOT EXISTS idx_weekly_stats_week ON weekly_stats(week_identifier, folder_type)`);
        console.log('✅ Index idx_weekly_stats_week créé');
    } catch (error) {
        console.log('⚠️ Index idx_weekly_stats_week:', error.message);
    }
    
    try {
        db.exec(`CREATE INDEX IF NOT EXISTS idx_weekly_stats_date ON weekly_stats(week_start_date, week_end_date)`);
        console.log('✅ Index idx_weekly_stats_date créé');
    } catch (error) {
        console.log('⚠️ Index idx_weekly_stats_date:', error.message);
    }
    
    try {
        db.exec(`CREATE INDEX IF NOT EXISTS idx_folder_mappings_path ON folder_mappings(original_folder_path)`);
        console.log('✅ Index idx_folder_mappings_path créé');
    } catch (error) {
        console.log('⚠️ Index idx_folder_mappings_path:', error.message);
    }
    
    // Tester les nouvelles fonctions
    console.log('\n🧪 Test des fonctions de suivi hebdomadaire...');
    
    // Fonction ISO Week
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
    
    // Test des statistiques par semaine
    const currentWeek = getISOWeekInfo(new Date());
    console.log('📅 Semaine actuelle:', currentWeek.weekIdentifier, '-', currentWeek.dateRange);
    
    // Compter les emails de la semaine actuelle
    const emailsThisWeek = db.prepare(`
        SELECT COUNT(*) as count, folder_name, category
        FROM emails 
        WHERE week_identifier = ?
        GROUP BY folder_name, category
    `).all(currentWeek.weekIdentifier);
    
    console.log(`📧 Emails cette semaine (${currentWeek.weekIdentifier}):`);
    emailsThisWeek.forEach(item => {
        console.log(`  - ${item.folder_name || 'Sans dossier'} (${item.category || 'Sans catégorie'}): ${item.count} emails`);
    });
    
    // Initialiser quelques données de test dans weekly_stats
    console.log('\n📊 Initialisation des statistiques hebdomadaires...');
    
    const insertWeeklyStats = db.prepare(`
        INSERT OR REPLACE INTO weekly_stats 
        (week_identifier, week_number, week_year, week_start_date, week_end_date, folder_type, emails_received, emails_treated)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `);
    
    // Exemple de données pour la semaine actuelle
    const categories = ['Personnel', 'Professionnel', 'Newsletter', 'Urgent'];
    
    for (const category of categories) {
        const emailsInCategory = db.prepare(`
            SELECT COUNT(*) as count FROM emails 
            WHERE week_identifier = ? AND (category = ? OR folder_name LIKE ?)
        `).get(currentWeek.weekIdentifier, category, `%${category}%`);
        
        if (emailsInCategory.count > 0) {
            insertWeeklyStats.run(
                currentWeek.weekIdentifier,
                currentWeek.weekNumber,
                currentWeek.weekYear,
                currentWeek.weekStartDate,
                currentWeek.weekEndDate,
                category,
                emailsInCategory.count,
                Math.floor(emailsInCategory.count * 0.8) // 80% traités par défaut
            );
            console.log(`  ✅ ${category}: ${emailsInCategory.count} reçus, ${Math.floor(emailsInCategory.count * 0.8)} traités`);
        }
    }
    
    // Vérification finale
    const totalWeeklyStats = db.prepare(`SELECT COUNT(*) as count FROM weekly_stats`).get();
    console.log(`\n📈 Total des statistiques hebdomadaires: ${totalWeeklyStats.count} entrées`);
    
    db.close();
    console.log('\n🎉 Base de données entièrement préparée pour le suivi hebdomadaire !');
    
} catch (error) {
    console.error('❌ Erreur:', error);
    process.exit(1);
}
