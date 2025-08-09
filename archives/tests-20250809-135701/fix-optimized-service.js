const fs = require('fs');
const path = require('path');

console.log('🔧 Mise à jour du service optimizedDatabaseService.js pour la nouvelle structure BD...');

const servicePath = path.join(__dirname, 'src/services/optimizedDatabaseService.js');
let content = fs.readFileSync(servicePath, 'utf8');

console.log('📁 Lecture du fichier service...');

// Liste des colonnes supprimées et leurs remplacements
const replacements = [
    // Colonnes supprimées - les retirer des CREATE TABLE
    {
        from: `                sender_name TEXT,`,
        to: ''
    },
    {
        from: `                sender_email TEXT,`,
        to: ''
    },
    {
        from: `                size_kb INTEGER,`,
        to: ''
    },
    
    // Requêtes INSERT - retirer les colonnes supprimées
    {
        from: `                (outlook_id, subject, sender_name, sender_email, recipient_email, received_time, sent_time, 
                 treated_time, folder_name, category, is_read, is_replied, has_attachment, body_preview, 
                 importance, size_kb, event_type, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`,
        to: `                (outlook_id, subject, sender, recipient_email, received_time, sent_time, 
                 treated_time, folder_name, category, is_read, is_replied, has_attachment, body_preview, 
                 importance, event_type, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`
    },
    
    // Valeurs INSERT - retirer les colonnes supprimées
    {
        from: `            emailData.sender_name || emailData.sender || '',
            emailData.sender_email || '',`,
        to: `            emailData.sender || '',`
    },
    {
        from: `            emailData.size_kb || 0,`,
        to: ''
    },
    
    // Valeurs INSERT pour emails en batch
    {
        from: `                    email.sender_name || email.sender || '',
                    email.sender_email || '',`,
        to: `                    email.sender || '',`
    },
    {
        from: `                    email.size_kb || 0,`,
        to: ''
    },
    
    // Mappage des colonnes - retirer les références aux colonnes supprimées
    {
        from: `                'sender_name': 'sender_name',
                'sender_email': 'sender_email'`,
        to: `                'sender': 'sender'`
    }
];

// Appliquer les remplacements
replacements.forEach((replacement, index) => {
    if (content.includes(replacement.from)) {
        content = content.replace(replacement.from, replacement.to);
        console.log(`✅ Remplacement ${index + 1} effectué`);
    } else {
        console.log(`⚠️  Remplacement ${index + 1} non trouvé`);
    }
});

// Nettoyer les lignes vides multiples
content = content.replace(/\n\n\n+/g, '\n\n');

// Écrire le fichier mis à jour
fs.writeFileSync(servicePath, content, 'utf8');

console.log('✅ Service optimizedDatabaseService.js mis à jour avec succès !');
console.log('🎯 Structure adaptée à la nouvelle BD nettoyée');
