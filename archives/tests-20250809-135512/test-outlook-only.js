// Test minimal de l'OutlookConnector sans la base de données
const outlookConnector = require('./src/server/outlookConnector');

async function testOutlookOnly() {
    try {
        console.log("🧪 Test minimal OutlookConnector");
        
        // Test de récupération d'emails pour "test"
        console.log("\n📧 Test récupération emails du dossier 'test'");
        const testPath = "erkaekanon@outlook.com\\Boîte de réception\\test";
        
        const result = await outlookConnector.getFolderEmails(testPath);
        console.log("📋 Résultat:", JSON.stringify(result, null, 2));
        
        // Test de récupération d'emails pour "testA"
        console.log("\n📧 Test récupération emails du dossier 'testA'");
        const testAPath = "erkaekanon@outlook.com\\Boîte de réception\\testA";
        
        const resultA = await outlookConnector.getFolderEmails(testAPath);
        console.log("📋 Résultat:", JSON.stringify(resultA, null, 2));
        
    } catch (error) {
        console.error("❌ Erreur:", error.message);
    }
}

testOutlookOnly();
