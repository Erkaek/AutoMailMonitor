// Script pour configurer facilement le dossier de test
const optimizedDatabaseService = require('./src/services/optimizedDatabaseService');

async function setupTestFolder() {
    try {
        console.log('🔧 Configuration du dossier de test pour monitoring...\n');
        
        // 1. Initialiser la base de données
        await optimizedDatabaseService.initialize();
        console.log('✅ Base de données initialisée');
        
        // 2. Vérifier la configuration actuelle
        const currentFolders = optimizedDatabaseService.getFoldersConfiguration();
        console.log(`📁 Dossiers actuellement configurés: ${currentFolders.length}`);
        
        if (currentFolders.length > 0) {
            console.log('Dossiers existants:');
            currentFolders.forEach((folder, index) => {
                console.log(`  ${index + 1}. ${folder.folder_path} (${folder.category})`);
            });
        }
        
        // 3. Demander quel dossier ajouter
        console.log('\n🎯 Dossiers de test détectés précédemment:');
        console.log('  - \\\\erkaekanon@outlook.com\\\\test (9 emails, 9 non lus)');
        console.log('  - \\\\erkaekanon@outlook.com\\\\test-1 (5 emails, 5 non lus)');
        console.log('  - \\\\erkaekanon@outlook.com\\\\test-11 (5 emails, 4 non lus)');
        console.log('  - \\\\erkaekanon@outlook.com\\\\azert (6 emails, 6 non lus)');
        
        // 4. Ajouter le dossier de test principal
        const testFolderPath = '\\\\erkaekanon@outlook.com\\\\test';
        
        try {
            const result = optimizedDatabaseService.addFolderConfiguration(
                testFolderPath, 
                'test', 
                'Dossier de test'
            );
            console.log(`\n✅ Dossier principal configuré: ${testFolderPath}`);
            console.log('Résultat:', result);
        } catch (error) {
            console.log(`⚠️ Dossier déjà configuré ou erreur: ${error.message}`);
        }
        
        // 5. Vérifier la nouvelle configuration
        const newFolders = optimizedDatabaseService.getFoldersConfiguration();
        console.log(`\n📁 Configuration mise à jour (${newFolders.length} dossiers):`);
        newFolders.forEach((folder, index) => {
            console.log(`  ${index + 1}. ${folder.folder_path} (${folder.category}) [${folder.folder_name}]`);
        });
        
        console.log('\n🔄 IMPORTANT: Redémarrez AutoMailMonitor pour que les changements prennent effet !');
        console.log('💡 Ou utilisez l\'API api-folders-reload-config depuis l\'interface');
        
    } catch (error) {
        console.error('❌ Erreur configuration:', error.message);
        console.error('Stack:', error.stack);
    }
}

// Fonction pour ajouter un dossier personnalisé
async function addCustomFolder(folderPath, category = 'custom', name = null) {
    try {
        await optimizedDatabaseService.initialize();
        
        const result = optimizedDatabaseService.addFolderConfiguration(
            folderPath,
            category,
            name || `Dossier ${category}`
        );
        
        console.log(`✅ Dossier ajouté: ${folderPath}`);
        return result;
    } catch (error) {
        console.error(`❌ Erreur ajout dossier: ${error.message}`);
        throw error;
    }
}

// Exécuter si lancé directement
if (require.main === module) {
    const args = process.argv.slice(2);
    
    if (args.length >= 1) {
        // Mode ajout dossier personnalisé
        const [folderPath, category, name] = args;
        addCustomFolder(folderPath, category, name).then(() => {
            console.log('✅ Dossier personnalisé ajouté');
            process.exit(0);
        }).catch(error => {
            console.error('💥 Erreur:', error);
            process.exit(1);
        });
    } else {
        // Mode configuration automatique
        setupTestFolder().then(() => {
            console.log('✅ Configuration terminée');
            process.exit(0);
        }).catch(error => {
            console.error('💥 Configuration échouée:', error);
            process.exit(1);
        });
    }
}

module.exports = { setupTestFolder, addCustomFolder };
