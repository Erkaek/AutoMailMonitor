// Test simple pour voir ce que retourne l'API api-weekly-current-stats
const { ipcRenderer } = require('electron');

async function testWeeklyCurrentStats() {
    try {
        console.log('🧪 Test de l\'API api-weekly-current-stats...');
        
        const result = await ipcRenderer.invoke('api-weekly-current-stats');
        
        console.log('📊 Résultat de l\'API:', JSON.stringify(result, null, 2));
        
        if (result.success) {
            console.log('✅ API fonctionne');
            console.log('📅 WeekInfo:', result.weekInfo);
            console.log('📈 Categories:', result.categories);
        } else {
            console.log('❌ Erreur API:', result.error);
        }
        
    } catch (error) {
        console.error('💥 Erreur lors du test:', error);
    }
}

// Attendre que la page soit chargée
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', testWeeklyCurrentStats);
} else {
    testWeeklyCurrentStats();
}
