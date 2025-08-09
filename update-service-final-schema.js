const fs = require('fs');
const path = require('path');

console.log('🔧 Mise à jour du service optimizedDatabaseService avec le schéma final...');

const servicePath = path.join(__dirname, 'src', 'services', 'optimizedDatabaseService.js');

// Lire le contenu du service
let content = fs.readFileSync(servicePath, 'utf8');

// Remplacer la définition de la table emails par le schéma optimisé
const newTableDefinition = `
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                outlook_id TEXT NOT NULL,
                subject TEXT NOT NULL DEFAULT '',
                sender_email TEXT DEFAULT '',
                received_time DATETIME,
                sent_time DATETIME,
                folder_name TEXT DEFAULT '',
                category TEXT DEFAULT 'Mails simples',
                is_read BOOLEAN DEFAULT 0,
                importance INTEGER DEFAULT 1,
                week_identifier TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )`;

// Trouver et remplacer la définition CREATE TABLE emails
content = content.replace(
    /CREATE TABLE IF NOT EXISTS emails \([^}]+\)/s,
    newTableDefinition
);

// Mettre à jour la requête INSERT pour correspondre au nouveau schéma
const newInsertQuery = `
                INSERT OR REPLACE INTO emails 
                (outlook_id, subject, sender_email, received_time, sent_time, 
                 folder_name, category, is_read, importance, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)`;

content = content.replace(
    /INSERT OR REPLACE INTO emails[^V]+VALUES[^)]+\)/s,
    newInsertQuery
);

// Mettre à jour les appels aux statements pour correspondre aux nouvelles colonnes
content = content.replace(
    /emailData\.outlook_id \|\| emailData\.id \|\| '',\s*emailData\.entry_id[^,]*,\s*/g,
    'emailData.outlook_id || emailData.id || \'\',\n            '
);

content = content.replace(
    /emailData\.sender_name \|\| '',/g,
    '// sender_name supprimé'
);

content = content.replace(
    /emailData\.recipient_email[^,]*,/g,
    '// recipient_email supprimé'
);

content = content.replace(
    /emailData\.sent_time[^,]*,\s*emailData\.treated_time[^,]*,/g,
    'emailData.sent_time || emailData.received_time || new Date().toISOString(),'
);

content = content.replace(
    /emailData\.folder_name[^,]*,\s*emailData\.category[^,]*,\s*emailData\.is_read[^,]*,\s*emailData\.is_replied[^,]*,\s*emailData\.has_attachment[^,]*,\s*emailData\.body_preview[^,]*,/g,
    `emailData.folder_name || emailData.folder_path || '',
            emailData.category || 'Mails simples',
            emailData.is_read ? 1 : 0,`
);

// Simplifier les batch inserts aussi
content = content.replace(
    /email\.outlook_id \|\| email\.id \|\| '',\s*email\.entry_id[^,]*,\s*/g,
    'email.outlook_id || email.id || \'\',\n                    '
);

content = content.replace(
    /email\.sender_name \|\| '',/g,
    '// sender_name supprimé'
);

content = content.replace(
    /email\.recipient_email[^,]*,/g,
    '// recipient_email supprimé'
);

content = content.replace(
    /email\.sent_time[^,]*,\s*email\.treated_time[^,]*,/g,
    'email.sent_time || email.received_time || new Date().toISOString(),'
);

content = content.replace(
    /email\.folder_name[^,]*,\s*email\.category[^,]*,\s*email\.is_read[^,]*,\s*email\.is_replied[^,]*,\s*email\.has_attachment[^,]*,\s*email\.body_preview[^,]*,/g,
    `email.folder_name || email.folder_path || '',
                    email.category || 'Mails simples',
                    email.is_read ? 1 : 0,`
);

// Mettre à jour les index pour correspondre au nouveau schéma
const newIndexes = `
        this.db.exec(\`
            CREATE INDEX IF NOT EXISTS idx_emails_outlook_id ON emails(outlook_id);
            CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails(received_time);
            CREATE INDEX IF NOT EXISTS idx_emails_week_folder ON emails(week_identifier, folder_name);
            CREATE INDEX IF NOT EXISTS idx_emails_unread ON emails(is_read, received_time);
            CREATE INDEX IF NOT EXISTS idx_emails_category ON emails(category, received_time);
        \`);`;

content = content.replace(
    /this\.db\.exec\(`[^`]*CREATE INDEX[^`]*`\);/s,
    newIndexes
);

// Nettoyer les commentaires de debug et références obsolètes
content = content.replace(/\/\/ sender_name supprimé\s*,?\s*/g, '');
content = content.replace(/\/\/ recipient_email supprimé\s*,?\s*/g, '');
content = content.replace(/,\s*,/g, ',');
content = content.replace(/,\s*\)/g, ')');

// Mettre à jour les messages de log
content = content.replace(
    /compatible schéma nettoyé \(sender_name, sender_email\)/g,
    'schéma optimisé final (13 colonnes)'
);

// Écrire le fichier corrigé
fs.writeFileSync(servicePath, content, 'utf8');

console.log('✅ Service optimizedDatabaseService mis à jour avec le schéma final');
console.log('📝 Alignement avec la base de données optimisée (13 colonnes)');
