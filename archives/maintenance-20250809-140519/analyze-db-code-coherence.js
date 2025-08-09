const Database = require('better-sqlite3');
const fs = require('fs');
const path = require('path');

console.log('🔍 ANALYSE DE COHÉRENCE BASE DE DONNÉES <-> CODE\n');

// Reconstruction de better-sqlite3 si nécessaire
console.log('🔧 Vérification de better-sqlite3...');
try {
    const db = new Database('./data/emails.db', { readonly: true });
    db.close();
    console.log('✅ Better-sqlite3 opérationnel\n');
} catch (error) {
    if (error.message.includes('NODE_MODULE_VERSION')) {
        console.log('⚠️ Reconstruction de better-sqlite3 nécessaire...');
        const { execSync } = require('child_process');
        try {
            execSync('npm rebuild better-sqlite3', { stdio: 'inherit' });
            console.log('✅ Better-sqlite3 reconstruit\n');
        } catch (rebuildError) {
            console.error('❌ Échec de la reconstruction:', rebuildError.message);
            process.exit(1);
        }
    } else {
        throw error;
    }
}

// === 1. ANALYSE DE LA BASE DE DONNÉES ===
console.log('📊 ANALYSE DE LA BASE DE DONNÉES');
console.log('=====================================');

const db = new Database('./data/emails.db', { readonly: true });

// Colonnes de la table emails
const emailsColumns = db.prepare("PRAGMA table_info(emails)").all();
console.log('\n📋 Colonnes de la table emails:');
emailsColumns.forEach(col => {
    console.log(`  - ${col.name} (${col.type})`);
});

// Autres tables
const tables = db.prepare("SELECT name FROM sqlite_master WHERE type='table'").all();
console.log('\n📦 Tables présentes:');
tables.forEach(table => {
    console.log(`  - ${table.name}`);
});

db.close();

// === 2. ANALYSE DU CODE ===
console.log('\n\n💻 ANALYSE DU CODE');
console.log('==================');

// Fonctions utilitaires
function searchInFile(filePath, pattern) {
    try {
        const content = fs.readFileSync(filePath, 'utf8');
        const matches = content.match(pattern);
        return matches || [];
    } catch (error) {
        return [];
    }
}

function searchInDirectory(dir, filePattern, searchPattern) {
    const results = [];
    
    function walkDir(currentDir) {
        const files = fs.readdirSync(currentDir);
        files.forEach(file => {
            const filePath = path.join(currentDir, file);
            const stat = fs.statSync(filePath);
            
            if (stat.isDirectory() && !file.startsWith('.') && file !== 'node_modules') {
                walkDir(filePath);
            } else if (stat.isFile() && filePattern.test(file)) {
                const matches = searchInFile(filePath, searchPattern);
                if (matches.length > 0) {
                    results.push({
                        file: filePath.replace(process.cwd() + '\\', ''),
                        matches: [...new Set(matches)] // Supprime les doublons
                    });
                }
            }
        });
    }
    
    walkDir(dir);
    return results;
}

// Recherche des colonnes utilisées dans le code
const columnNames = emailsColumns.map(col => col.name);
const usedColumns = {};

console.log('\n🔍 Recherche des colonnes dans le code...');

columnNames.forEach(columnName => {
    // Pattern pour trouver les utilisations de colonnes
    const patterns = [
        new RegExp(`${columnName}`, 'gi'),
        new RegExp(`'${columnName}'`, 'gi'),
        new RegExp(`"${columnName}"`, 'gi'),
        new RegExp(`\\.${columnName}`, 'gi')
    ];
    
    let allMatches = [];
    patterns.forEach(pattern => {
        const results = searchInDirectory('./src', /\.(js|html)$/, pattern);
        allMatches = allMatches.concat(results);
    });
    
    if (allMatches.length > 0) {
        usedColumns[columnName] = allMatches;
    }
});

// === 3. RAPPORT DE COHÉRENCE ===
console.log('\n\n📋 RAPPORT DE COHÉRENCE');
console.log('========================');

// Colonnes jamais utilisées
const unusedColumns = columnNames.filter(col => !usedColumns[col]);
if (unusedColumns.length > 0) {
    console.log('\n❌ COLONNES JAMAIS UTILISÉES DANS LE CODE:');
    unusedColumns.forEach(col => {
        console.log(`  - ${col} (peut être supprimée)`);
    });
} else {
    console.log('\n✅ Toutes les colonnes de la BDD sont utilisées dans le code');
}

// Colonnes utilisées
console.log('\n✅ COLONNES UTILISÉES DANS LE CODE:');
Object.keys(usedColumns).forEach(col => {
    console.log(`  - ${col} (${usedColumns[col].length} fichier(s))`);
});

// === 4. ANALYSE SPÉCIFIQUE sender_name vs sender_email ===
console.log('\n\n🔍 ANALYSE SPÉCIFIQUE: sender_name vs sender_email');
console.log('==================================================');

const senderNameUsage = usedColumns['sender_name'] || [];
const senderEmailUsage = usedColumns['sender_email'] || [];

console.log(`\n📊 Utilisation de sender_name: ${senderNameUsage.length} fichier(s)`);
if (senderNameUsage.length > 0) {
    senderNameUsage.forEach(usage => {
        console.log(`  - ${usage.file}`);
    });
}

console.log(`\n📊 Utilisation de sender_email: ${senderEmailUsage.length} fichier(s)`);
if (senderEmailUsage.length > 0) {
    senderEmailUsage.forEach(usage => {
        console.log(`  - ${usage.file}`);
    });
}

// Vérification dans les données réelles
const dbCheck = new Database('./data/emails.db', { readonly: true });
const emailsSample = dbCheck.prepare("SELECT sender_email FROM emails LIMIT 5").all();
dbCheck.close();

console.log('\n📋 Échantillon de données réelles:');
emailsSample.forEach((email, index) => {
    console.log(`  Email ${index + 1}:`);
    console.log(`    - sender_email: "${email.sender_email || 'NULL'}"`);
});

// === 5. RECOMMANDATIONS ===
console.log('\n\n💡 RECOMMANDATIONS');
console.log('===================');

if (unusedColumns.length > 0) {
    console.log('\n🗑️ COLONNES À SUPPRIMER:');
    unusedColumns.forEach(col => {
        console.log(`  ALTER TABLE emails DROP COLUMN ${col};`);
    });
}

// Analyse sender_name vs sender_email
const hasValidSenderEmails = emailsSample.some(email => email.sender_email && email.sender_email.includes('@'));

if (hasValidSenderEmails) {
    console.log('\n📧 RECOMMANDATION SENDER:');
    console.log('  - Seuls les sender_email sont présents et utilisés');
    console.log('  - La base de données est correctement optimisée');
}

console.log('\n✅ Analyse terminée !');
