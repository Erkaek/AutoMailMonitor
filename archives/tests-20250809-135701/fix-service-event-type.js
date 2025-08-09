const fs = require('fs');
const path = require('path');

console.log('🔧 Suppression des références à event_type dans optimizedDatabaseService.js...');

const servicePath = path.join(__dirname, 'src', 'services', 'optimizedDatabaseService.js');

// Lire le contenu du service
let content = fs.readFileSync(servicePath, 'utf8');

// Supprimer event_type de la définition de table CREATE TABLE
content = content.replace(/,\s*event_type TEXT,?/g, ',');

// Supprimer event_type de la liste des colonnes INSERT
content = content.replace(/,\s*event_type(?=\s*\))/g, '');
content = content.replace(/event_type,\s*/g, '');

// Supprimer les paramètres event_type dans les INSERT statements
content = content.replace(/,\s*emailData\.event_type \|\| ''/g, '');
content = content.replace(/emailData\.event_type \|\| '',?\s*/g, '');
content = content.replace(/,\s*email\.event_type \|\| ''/g, '');
content = content.replace(/email\.event_type \|\| '',?\s*/g, '');

// Nettoyer les virgules en double et les espaces
content = content.replace(/,\s*,/g, ',');
content = content.replace(/,\s*\)/g, ')');

// Écrire le fichier corrigé
fs.writeFileSync(servicePath, content, 'utf8');

console.log('✅ Références à event_type supprimées du service');
console.log('📝 Service optimizedDatabaseService.js mis à jour');
