const Database = require('better-sqlite3');

console.log('🔍 Vérification du schéma final de la base de données...');

try {
    const db = new Database('./data/emails.db', { readonly: true });
    
    // Récupérer le schéma de la table emails
    const tableInfo = db.prepare("PRAGMA table_info(emails)").all();
    
    console.log('\n📋 Colonnes actuelles de la table emails:');
    console.log('Index | Nom de la colonne | Type | Nullable | Défaut | Clé primaire');
    console.log('------|-------------------|------|----------|---------|---------------');
    
    tableInfo.forEach(col => {
        console.log(`${col.cid.toString().padEnd(5)} | ${col.name.padEnd(17)} | ${col.type.padEnd(4)} | ${col.notnull ? 'Non' : 'Oui'} | ${(col.dflt_value || 'NULL').toString().padEnd(7)} | ${col.pk ? 'Oui' : 'Non'}`);
    });
    
    console.log(`\n📊 Total: ${tableInfo.length} colonnes`);
    
    // Vérifier si event_type existe encore
    const hasEventType = tableInfo.some(col => col.name === 'event_type');
    if (hasEventType) {
        console.log('❌ La colonne event_type existe encore ! Il faut la supprimer.');
    } else {
        console.log('✅ La colonne event_type a été correctement supprimée.');
    }
    
    db.close();
    
} catch (error) {
    console.error('❌ Erreur lors de la vérification:', error.message);
}
