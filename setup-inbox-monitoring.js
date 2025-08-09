// Script pour ajouter automatiquement la configuration du dossier Inbox
const { ipcRenderer } = require('electron');

async function configurerDossierInbox() {
    try {
        console.log('🔧 Configuration automatique du dossier Inbox...');
        
        // Ajouter le dossier Inbox pour monitoring
        const result = await ipcRenderer.invoke('api-folders-add', {
            folderPath: '\\Inbox',
            category: 'inbox'
        });
        
        console.log('✅ Dossier Inbox configuré:', result);
        
        // Vérifier la configuration
        const folders = await ipcRenderer.invoke('api-settings-folders-load');
        console.log('📁 Dossiers configurés:', folders);
        
        return result;
        
    } catch (error) {
        console.error('❌ Erreur configuration dossier:', error);
        throw error;
    }
}

// Si exécuté directement (pas dans Electron)
if (typeof window === 'undefined') {
    console.log('⚠️ Ce script doit être exécuté dans le contexte Electron');
    console.log('📝 Ajouter manuellement via l\'interface ou utiliser l\'API REST');
}

module.exports = { configurerDossierInbox };
