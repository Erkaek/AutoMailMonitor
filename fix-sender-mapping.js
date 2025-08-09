const fs = require('fs');
const path = require('path');

console.log('🔧 Correction du mapping sender_name dans unifiedMonitoringService...');

const servicePath = path.join(__dirname, 'src', 'services', 'unifiedMonitoringService.js');

// Lire le contenu du service
let content = fs.readFileSync(servicePath, 'utf8');

// Corriger la ligne 948 - ajouter sender_name
content = content.replace(
    /sender_email: processedEmail\.senderEmail,/g,
    `sender_name: processedEmail.senderName,
                sender_email: processedEmail.senderEmail,`
);

// Corriger la ligne 1433 - ajouter sender_name
content = content.replace(
    /sender_email: emailData\.senderEmail \|\| emailData\.SenderEmailAddress,/g,
    `sender_name: emailData.senderName || emailData.SenderName,
                sender_email: emailData.senderEmail || emailData.SenderEmailAddress,`
);

// Corriger les autres endroits où on crée des emailRecord
content = content.replace(
    /(\s+)(sender_email: [^,\n]+),(\s+)/g,
    '$1sender_name: emailData.senderName || emailData.SenderName || processedEmail.senderName || \'\',$3$1sender_email: emailData.senderEmail || emailData.SenderEmailAddress || processedEmail.senderEmail || \'\',$3'
);

// Écrire le fichier corrigé
fs.writeFileSync(servicePath, content, 'utf8');

console.log('✅ Mapping sender_name ajouté dans unifiedMonitoringService');

// Maintenant corriger l'optimizedDatabaseService pour utiliser les bonnes colonnes
console.log('🔧 Correction de l\'INSERT dans optimizedDatabaseService...');

const optimizedServicePath = path.join(__dirname, 'src', 'services', 'optimizedDatabaseService.js');
let optimizedContent = fs.readFileSync(optimizedServicePath, 'utf8');

// Ajouter entry_id dans les INSERT statements (colonne présente en BDD mais oubliée)
optimizedContent = optimizedContent.replace(
    /\(outlook_id, subject, sender_name, recipient_email/g,
    '(outlook_id, entry_id, subject, sender_name, recipient_email'
);

optimizedContent = optimizedContent.replace(
    /VALUES \(\?, \?, \?, \?,/g,
    'VALUES (?, ?, ?, ?, ?,'
);

// Ajouter entry_id dans les appels à insertEmail
optimizedContent = optimizedContent.replace(
    /emailData\.outlook_id \|\| emailData\.id \|\| '',/g,
    `emailData.outlook_id || emailData.id || '',
            emailData.entry_id || emailData.outlook_id || emailData.id || '',`
);

optimizedContent = optimizedContent.replace(
    /email\.outlook_id \|\| email\.id \|\| '',/g,
    `email.outlook_id || email.id || '',
                    email.entry_id || email.outlook_id || email.id || '',`
);

fs.writeFileSync(optimizedServicePath, optimizedContent, 'utf8');

console.log('✅ INSERT corrigé avec entry_id dans optimizedDatabaseService');
console.log('📝 Services mis à jour pour utiliser sender_name et entry_id');
