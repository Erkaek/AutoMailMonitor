/**
 * Test de l'API weekly-current-stats pour diagnostiquer le problème
 */

const path = require('path');
const optimizedDatabaseService = require('./src/services/optimizedDatabaseService');

async function testWeeklyStats() {
    console.log('🔍 Test des statistiques hebdomadaires...');
    
    try {
        // Test direct du service de base de données
        console.log('\n1. Test getCurrentWeekStats()...');
        const currentWeekStats = optimizedDatabaseService.getCurrentWeekStats();
        console.log('✅ Résultat getCurrentWeekStats:', JSON.stringify(currentWeekStats, null, 2));
        
        // Test des informations de semaine ISO
        console.log('\n2. Test getISOWeekInfo()...');
        const weekInfo = optimizedDatabaseService.getISOWeekInfo();
        console.log('✅ Résultat getISOWeekInfo:', JSON.stringify(weekInfo, null, 2));
        
        // Test de l'historique des semaines
        console.log('\n3. Test getWeeklyStats(5)...');
        const weeklyHistory = optimizedDatabaseService.getWeeklyStats(null, 5);
        console.log('✅ Résultat getWeeklyStats:', JSON.stringify(weeklyHistory, null, 2));
        
        // Vérifier s'il y a des données dans la table emails
        console.log('\n4. Vérification des données emails...');
        const db = optimizedDatabaseService.db;
        const totalEmails = db.prepare('SELECT COUNT(*) as count FROM emails').get();
        console.log('📊 Total emails dans la BD:', totalEmails.count);
        
        // Vérifier les emails de cette semaine
        const thisWeekEmails = db.prepare(`
            SELECT 
                folder_path,
                DATE(creation_date) as date,
                COUNT(*) as count
            FROM emails 
            WHERE strftime('%Y-%W', creation_date) = strftime('%Y-%W', 'now')
            GROUP BY folder_path, DATE(creation_date)
            ORDER BY date DESC
        `).all();
        console.log('📅 Emails de cette semaine par dossier/jour:', thisWeekEmails);
        
    } catch (error) {
        console.error('❌ Erreur lors du test:', error);
        console.error('Stack:', error.stack);
    }
}

testWeeklyStats();
