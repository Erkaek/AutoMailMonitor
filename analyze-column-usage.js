const Database = require('better-sqlite3');
const path = require('path');

console.log('🗑️ Identification et nettoyage des colonnes inutilisées');

// Ouvrir la base de données
const dbPath = path.join(__dirname, 'data', 'emails.db');
const db = new Database(dbPath);

// Analyser l'utilisation réelle des colonnes
console.log('\n📊 Analyse de l\'utilisation des colonnes :');

const emails = db.prepare("SELECT * FROM emails").all();
console.log(`Analysant ${emails.length} emails...`);

const columnUsage = {};

emails.forEach(email => {
    Object.keys(email).forEach(col => {
        if (!columnUsage[col]) {
            columnUsage[col] = { 
                filled: 0, 
                empty: 0, 
                unique_values: new Set(),
                sample_values: []
            };
        }
        
        if (email[col] !== null && email[col] !== '' && email[col] !== 0) {
            columnUsage[col].filled++;
            columnUsage[col].unique_values.add(email[col]);
            if (columnUsage[col].sample_values.length < 3) {
                columnUsage[col].sample_values.push(email[col]);
            }
        } else {
            columnUsage[col].empty++;
        }
    });
});

// Calculer les taux d'utilisation
console.log('\n📈 Taux d\'utilisation des colonnes :');
Object.entries(columnUsage).forEach(([col, stats]) => {
    const fillRate = Math.round((stats.filled / emails.length) * 100);
    const uniqueCount = stats.unique_values.size;
    const samples = stats.sample_values.slice(0, 2).join(', ');
    
    if (fillRate === 0) {
        console.log(`🔴 ${col}: ${fillRate}% (JAMAIS UTILISÉ)`);
    } else if (fillRate < 10) {
        console.log(`🟡 ${col}: ${fillRate}% (TRÈS PEU UTILISÉ) - ${uniqueCount} valeurs uniques - Ex: ${samples}`);
    } else {
        console.log(`🟢 ${col}: ${fillRate}% - ${uniqueCount} valeurs uniques - Ex: ${samples}`);
    }
});

// Identifier les colonnes candidates à la suppression
const unusedColumns = [];
const underusedColumns = [];

Object.entries(columnUsage).forEach(([col, stats]) => {
    const fillRate = (stats.filled / emails.length) * 100;
    
    if (fillRate === 0 && !['id', 'created_at', 'updated_at'].includes(col)) {
        unusedColumns.push(col);
    } else if (fillRate < 5 && !['id', 'created_at', 'updated_at', 'treated_time', 'deleted_at'].includes(col)) {
        underusedColumns.push(col);
    }
});

console.log('\n🗑️ RECOMMANDATIONS DE NETTOYAGE :');

if (unusedColumns.length > 0) {
    console.log('❌ Colonnes jamais utilisées (à supprimer) :', unusedColumns.join(', '));
}

if (underusedColumns.length > 0) {
    console.log('⚠️ Colonnes très peu utilisées (à examiner) :', underusedColumns.join(', '));
}

// Vérifier la cohérence des données
console.log('\n🔍 VÉRIFICATION DE LA COHÉRENCE :');

// Vérifier sender_name vs sender_email
const emailsWithSenderEmail = emails.filter(e => e.sender_email && e.sender_email.trim() !== '').length;
const emailsWithSenderName = emails.filter(e => e.sender_name && e.sender_name.trim() !== '').length;

console.log(`📧 Emails avec sender_email: ${emailsWithSenderEmail}/${emails.length}`);
console.log(`📧 Emails avec sender_name: ${emailsWithSenderName}/${emails.length}`);

if (emailsWithSenderEmail > 0 && emailsWithSenderName === 0) {
    console.log('⚠️ PROBLÈME: Tous les emails ont sender_email mais aucun n\'a sender_name !');
}

// Vérifier entry_id vs outlook_id
const emailsWithEntryId = emails.filter(e => e.entry_id && e.entry_id.trim() !== '').length;
const emailsWithOutlookId = emails.filter(e => e.outlook_id && e.outlook_id.trim() !== '').length;

console.log(`🆔 Emails avec entry_id: ${emailsWithEntryId}/${emails.length}`);
console.log(`🆔 Emails avec outlook_id: ${emailsWithOutlookId}/${emails.length}`);

// Vérifier folder_name vs folder_path
const emailsWithFolderName = emails.filter(e => e.folder_name && e.folder_name.trim() !== '').length;
const emailsWithFolderPath = emails.filter(e => e.folder_path && e.folder_path.trim() !== '').length;

console.log(`📁 Emails avec folder_name: ${emailsWithFolderName}/${emails.length}`);
console.log(`📁 Emails avec folder_path: ${emailsWithFolderPath}/${emails.length}`);

console.log('\n✅ Analyse terminée');

db.close();
