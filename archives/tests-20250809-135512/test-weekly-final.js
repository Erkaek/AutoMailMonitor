const path = require('path');
const { app } = require('electron');

// Simulation pour test sans Electron
if (!app) {
    console.log('🔍 Test exécuté hors Electron - simulation des APIs...');
}

// Import du service optimisé
const optimizedDatabaseService = require('./src/services/optimizedDatabaseService');

console.log('🔧 Initialisation du service de base de données...');
await optimizedDatabaseService.initialize();

console.log('🔍 TEST FINAL - Statistiques hebdomadaires');

async function runTest() {

try {
    // 1. Test de base des statistiques
    console.log('\n📊 1. Test getWeeklyStats basique:');
    
    // Test avec limite
    console.log('Test avec limite (4 semaines):');
    const statsWithLimit = optimizedDatabaseService.getWeeklyStats(4);
    console.log('Nombre de résultats:', Array.isArray(statsWithLimit) ? statsWithLimit.length : 'Non array');
    console.log('Données:', JSON.stringify(statsWithLimit, null, 2));
    
    // 2. Test avec identifiant spécifique
    console.log('\n📊 2. Test avec identifiant de semaine:');
    const currentDate = new Date();
    const getISOWeek = (date) => {
        const target = new Date(date.valueOf());
        const dayNr = (date.getDay() + 6) % 7;
        target.setDate(target.getDate() - dayNr + 3);
        const firstThursday = target.valueOf();
        target.setMonth(0, 1);
        if (target.getDay() !== 4) {
            target.setMonth(0, 1 + ((4 - target.getDay()) + 7) % 7);
        }
        return 1 + Math.ceil((firstThursday - target) / 604800000);
    };
    
    const currentWeek = getISOWeek(currentDate);
    const currentYear = currentDate.getFullYear();
    const currentWeekId = `${currentYear}-W${currentWeek.toString().padStart(2, '0')}`;
    
    console.log(`Semaine actuelle: ${currentWeekId}`);
    const statsWithId = optimizedDatabaseService.getWeeklyStats(currentWeekId);
    console.log('Données pour semaine actuelle:', JSON.stringify(statsWithId, null, 2));

    // 3. Test update current week stats
    console.log('\n📊 3. Test updateCurrentWeekStats:');
    optimizedDatabaseService.updateCurrentWeekStats();
    
    // Vérifier après update
    const updatedStats = optimizedDatabaseService.getWeeklyStats(4);
    console.log('Stats après updateCurrentWeekStats:', JSON.stringify(updatedStats, null, 2));

    console.log('\n✅ Test terminé avec succès !');

} catch (error) {
    console.error('❌ Erreur:', error);
    console.error('Stack:', error.stack);
}
}

// Exécuter le test
runTest().catch(console.error);
