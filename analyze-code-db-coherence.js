const Database = require('better-sqlite3');
const fs = require('fs');
const path = require('path');

console.log('🔍 Analyse de cohérence Code ↔ Base de données');

// Ouvrir la base de données
const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

// 1. Récupérer le schéma réel de la table emails
console.log('\n📊 1. SCHÉMA RÉEL DE LA TABLE EMAILS :');
const columns = db.prepare("PRAGMA table_info(emails)").all();
const dbColumns = columns.map(col => col.name);
console.log('Colonnes dans la BDD :', dbColumns.join(', '));

// 2. Analyser le service pour voir quelles colonnes sont utilisées
console.log('\n🔧 2. ANALYSE DU SERVICE optimizedDatabaseService.js :');
const servicePath = path.join(__dirname, 'src', 'services', 'optimizedDatabaseService.js');
const serviceContent = fs.readFileSync(servicePath, 'utf8');

// Extraire les colonnes du CREATE TABLE dans le service
const createTableMatch = serviceContent.match(/CREATE TABLE IF NOT EXISTS emails \(\s*([\s\S]+?)\s*\)/);
let serviceColumns = [];
if (createTableMatch) {
    const tableContent = createTableMatch[1];
    const columnLines = tableContent.split(',').map(line => line.trim());
    serviceColumns = columnLines
        .filter(line => line && !line.startsWith('PRIMARY KEY') && !line.startsWith('FOREIGN KEY'))
        .map(line => {
            const match = line.match(/^(\w+)/);
            return match ? match[1] : null;
        })
        .filter(col => col);
}
console.log('Colonnes définies dans le service :', serviceColumns.join(', '));

// 3. Extraire les colonnes utilisées dans les INSERT statements
console.log('\n📝 3. COLONNES UTILISÉES DANS LES INSERT :');
const insertMatches = serviceContent.match(/INSERT.*INTO emails[^(]*\(([^)]+)\)/g);
let insertColumns = [];
if (insertMatches) {
    insertMatches.forEach((match, index) => {
        const columnsMatch = match.match(/\(([^)]+)\)/);
        if (columnsMatch) {
            const cols = columnsMatch[1].split(',').map(col => col.trim());
            console.log(`INSERT ${index + 1} :`, cols.join(', '));
            insertColumns = insertColumns.concat(cols);
        }
    });
}

// 4. Analyser les données réelles pour voir quelles colonnes sont remplies
console.log('\n📈 4. ANALYSE DES DONNÉES RÉELLES :');
const sampleEmails = db.prepare("SELECT * FROM emails LIMIT 5").all();
if (sampleEmails.length > 0) {
    const realData = {};
    sampleEmails.forEach(email => {
        Object.keys(email).forEach(col => {
            if (!realData[col]) realData[col] = { filled: 0, empty: 0, types: new Set() };
            if (email[col] !== null && email[col] !== '') {
                realData[col].filled++;
                realData[col].types.add(typeof email[col]);
            } else {
                realData[col].empty++;
            }
        });
    });
    
    console.log('Utilisation des colonnes dans les données réelles :');
    Object.entries(realData).forEach(([col, stats]) => {
        const fillRate = Math.round((stats.filled / sampleEmails.length) * 100);
        console.log(`  ${col}: ${fillRate}% rempli (${stats.filled}/${sampleEmails.length}) - Types: ${Array.from(stats.types).join(', ')}`);
    });
}

// 5. Rechercher dans tout le code quelles colonnes sont référencées
console.log('\n🔎 5. RECHERCHE DES RÉFÉRENCES DANS LE CODE :');
const codeReferences = {};

// Analyser le service
const allMatches = serviceContent.match(/\b(sender_name|sender_email|entry_id|outlook_id|subject|received_time|sent_time|treated_time|folder_name|folder_path|category|is_read|is_replied|is_treated|has_attachment|body_preview|importance|size_bytes|week_identifier|created_at|updated_at|is_deleted|deleted_at|recipient_email)\b/g);
if (allMatches) {
    allMatches.forEach(match => {
        if (!codeReferences[match]) codeReferences[match] = 0;
        codeReferences[match]++;
    });
}

console.log('Références dans le service :');
Object.entries(codeReferences)
    .sort(([,a], [,b]) => b - a)
    .forEach(([col, count]) => {
        console.log(`  ${col}: ${count} fois`);
    });

// 6. Identifier les incohérences
console.log('\n❌ 6. INCOHÉRENCES DÉTECTÉES :');

// Colonnes dans la BDD mais pas dans le service
const dbNotInService = dbColumns.filter(col => !serviceColumns.includes(col));
if (dbNotInService.length > 0) {
    console.log('🟡 Colonnes dans la BDD mais non définies dans le service :', dbNotInService.join(', '));
}

// Colonnes dans le service mais pas dans la BDD
const serviceNotInDb = serviceColumns.filter(col => !dbColumns.includes(col));
if (serviceNotInDb.length > 0) {
    console.log('🔴 Colonnes définies dans le service mais absentes de la BDD :', serviceNotInDb.join(', '));
}

// Colonnes référencées dans le code mais pas dans la BDD
const codeNotInDb = Object.keys(codeReferences).filter(col => !dbColumns.includes(col));
if (codeNotInDb.length > 0) {
    console.log('🔴 Colonnes référencées dans le code mais absentes de la BDD :', codeNotInDb.join(', '));
}

// Colonnes dans la BDD mais jamais référencées
const dbNotInCode = dbColumns.filter(col => !Object.keys(codeReferences).includes(col));
if (dbNotInCode.length > 0) {
    console.log('🟡 Colonnes dans la BDD mais jamais référencées dans le code :', dbNotInCode.join(', '));
}

// 7. Vérifier les problèmes spécifiques mentionnés
console.log('\n🎯 7. VÉRIFICATION DES PROBLÈMES SPÉCIFIQUES :');

// Vérifier sender_name vs sender_email
if (sampleEmails.length > 0) {
    const senderNameFilled = sampleEmails.filter(e => e.sender_name && e.sender_name.trim() !== '').length;
    const senderEmailFilled = sampleEmails.filter(e => e.sender_email && e.sender_email.trim() !== '').length;
    
    console.log(`📧 sender_name rempli dans ${senderNameFilled}/${sampleEmails.length} emails`);
    console.log(`📧 sender_email rempli dans ${senderEmailFilled}/${sampleEmails.length} emails`);
    
    if (senderNameFilled === 0 && senderEmailFilled > 0) {
        console.log('⚠️  PROBLÈME : sender_name jamais rempli mais sender_email oui !');
    }
}

console.log('\n✅ Analyse terminée');

db.close();
