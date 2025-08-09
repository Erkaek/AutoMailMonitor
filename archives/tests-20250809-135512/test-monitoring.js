// Test du monitoring en temps réel des dossiers
const outlookConnector = require('./src/server/outlookConnector');

async function testRealTimeMonitoring() {
    try {
        console.log("🧪 Test du monitoring en temps réel");
        
        // Configurer les listeners d'événements
        outlookConnector.on('newEmailDetected', (data) => {
            console.log(`🆕 NOUVEL EMAIL: ${data.subject} (${data.folderPath})`);
        });
        
        outlookConnector.on('emailStatusChanged', (data) => {
            console.log(`📝 STATUT CHANGÉ: ${data.subject} -> ${data.isRead ? 'Lu' : 'Non lu'}`);
        });
        
        outlookConnector.on('emailDeleted', (data) => {
            console.log(`🗑️ EMAIL SUPPRIMÉ: ${data.subject}`);
        });
        
        outlookConnector.on('folderCountChanged', (data) => {
            console.log(`📊 NOMBRE CHANGÉ: ${data.folderPath} (${data.oldCount} -> ${data.newCount})`);
        });
        
        // Démarrer le monitoring pour le dossier "test"
        const testPath = "erkaekanon@outlook.com\\Boîte de réception\\test";
        console.log(`\n🎧 Démarrage monitoring pour: ${testPath}`);
        
        const result = await outlookConnector.startFolderMonitoring(testPath);
        console.log("📋 Résultat monitoring:", JSON.stringify(result, null, 2));
        
        if (result.success) {
            console.log("\n✅ Monitoring actif !");
            console.log("📝 Instructions:");
            console.log("   1. Ajoutez un email au dossier 'test' dans Outlook");
            console.log("   2. Marquez un email comme lu/non lu");
            console.log("   3. Supprimez un email");
            console.log("   4. Les changements seront détectés toutes les 30 secondes");
            console.log("\n⏰ Surveillance en cours... (Ctrl+C pour arrêter)");
            
            // Garder le script en vie pour observer les événements
            process.on('SIGINT', async () => {
                console.log("\n🛑 Arrêt du monitoring...");
                await outlookConnector.stopFolderMonitoring(testPath);
                console.log("✅ Monitoring arrêté");
                process.exit(0);
            });
            
            // Maintenir le processus actif
            await new Promise(() => {}); // Infinite wait
        }
        
    } catch (error) {
        console.error("❌ Erreur:", error.message);
    }
}

testRealTimeMonitoring();
