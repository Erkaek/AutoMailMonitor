const fs = require('fs');
const path = require('path');

console.log('🔧 Correction du OptimizedDatabaseService pour le nouveau schéma...');

const servicePath = path.join(__dirname, 'src', 'services', 'optimizedDatabaseService.js');

try {
    let content = fs.readFileSync(servicePath, 'utf8');
    
    console.log('📝 Corrections en cours...');
    
    // 1. Corriger la requête d'insertion - remplacer "sender" par "sender_name"
    content = content.replace(
        /INSERT OR REPLACE INTO emails \s*\(\s*outlook_id,\s*subject,\s*sender,\s*recipient_email/g,
        'INSERT OR REPLACE INTO emails \n                (outlook_id, subject, sender_name, recipient_email'
    );
    
    // 2. Corriger les paramètres dans les statements
    content = content.replace(
        /emailData\.sender \|\| ''/g,
        "emailData.sender_name || ''"
    );
    
    content = content.replace(
        /email\.sender \|\| ''/g,
        "email.sender_name || ''"
    );
    
    // 3. Corriger les mappings de colonnes
    content = content.replace(
        /'sender': 'sender'/g,
        "'sender_name': 'sender_name'"
    );
    
    // 4. Ajouter sender_email dans les propriétés retournées
    content = content.replace(
        /sender: emailData\.sender_name \|\| ''/g,
        "sender_name: emailData.sender_name || '', sender_email: emailData.sender_email || ''"
    );
    
    // 5. Corriger les commentaires
    content = content.replace(
        /compatible schéma existant/g,
        'compatible schéma nettoyé (sender_name, sender_email)'
    );
    
    // Sauvegarder
    fs.writeFileSync(servicePath, content);
    
    console.log('✅ OptimizedDatabaseService corrigé avec succès');
    console.log('   - sender → sender_name');
    console.log('   - Ajout de sender_email dans les mappings');
    console.log('   - Mise à jour des statements SQL');
    
} catch (error) {
    console.error('❌ Erreur lors de la correction:', error.message);
    process.exit(1);
}
