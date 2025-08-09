const Database = require('better-sqlite3');

console.log('📊 Vérification finale du schéma de la base de données...');

try {
    const db = new Database('./data/emails.db');
    
    // Structure de la table emails
    console.log('\n=== Structure de la table emails ===');
    const emailsInfo = db.prepare('PRAGMA table_info(emails)').all();
    emailsInfo.forEach(col => {
        console.log(`${col.cid}: ${col.name} - ${col.type} ${col.notnull ? '(NOT NULL)' : ''} ${col.pk ? '(PRIMARY KEY)' : ''}`);
    });
    
    // Test d'une requête simple
    console.log('\n=== Test de requête ===');
    const sampleEmails = db.prepare('SELECT * FROM emails LIMIT 3').all();
    console.log(`✅ ${sampleEmails.length} emails trouvés dans la base`);
    
    if (sampleEmails.length > 0) {
        console.log('\n=== Colonnes disponibles dans le premier email ===');
        console.log(Object.keys(sampleEmails[0]).join(', '));
    }
    
    db.close();
    console.log('\n✅ Vérification terminée avec succès');
} catch (error) {
    console.error('❌ Erreur:', error.message);
}
