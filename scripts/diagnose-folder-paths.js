#!/usr/bin/env node
/**
 * Diagnostic: Afficher tous les chemins exacts de folder_configurations
 * Aide à identifier quels chemins sont enregistrés et pourquoi ils échouent
 */

const path = require('path');
const Database = require('better-sqlite3');

// Même logique que l'application: env var ou data/emails.db
const resolveDbPath = () => {
  const envPath = process.env.MAILMONITOR_DB_PATH || process.env.AUTO_MAIL_MONITOR_DB_PATH;
  if (envPath) return path.isAbsolute(envPath) ? envPath : path.resolve(process.cwd(), envPath);
  return path.join(__dirname, '../data/emails.db');
};

const DB_PATH = resolveDbPath();

if (!require('fs').existsSync(DB_PATH)) {
  console.error(`❌ BDD non trouvée: ${DB_PATH}`);
  console.error('Indiquez le chemin via MAILMONITOR_DB_PATH ou placez le fichier dans data/emails.db');
  process.exit(1);
}

const db = new Database(DB_PATH);

console.log('🔍 Diagnostic: Chemins enregistrés dans folder_configurations\n');

const rows = db.prepare(`
  SELECT rowid, folder_path, folder_name, category, store_id, entry_id 
  FROM folder_configurations
  ORDER BY category ASC, folder_path ASC
`).all();

if (rows.length === 0) {
  console.log('❌ Aucune configuration trouvée.');
  db.close();
  process.exit(0);
}

console.log(`📋 Total: ${rows.length} dossiers\n`);

// Grouper par catégorie
const byCategory = {};
rows.forEach(row => {
  const cat = row.category || '(sans catégorie)';
  if (!byCategory[cat]) byCategory[cat] = [];
  byCategory[cat].push(row);
});

Object.keys(byCategory).sort().forEach(cat => {
  console.log(`\n📂 ${cat} (${byCategory[cat].length}):`);
  byCategory[cat].forEach((row, i) => {
    const path = (row.folder_path || '').trim();
    const hasBackslash = path.includes('\\') || path.includes('/');
    const hasIds = row.store_id && row.entry_id;
    const validity = (hasBackslash || hasIds) ? '✅' : '❌';
    
    const details = [];
    if (path) details.push(`path="${path}"`);
    if (row.folder_name) details.push(`name="${row.folder_name}"`);
    if (row.store_id) details.push(`store_id="${row.store_id}"`);
    if (row.entry_id) details.push(`entry_id="${row.entry_id}"`);
    
    console.log(`  ${validity} ${i + 1}. ${details.join(', ')}`);
  });
});

console.log('\n📊 Résumé:');
const valid = rows.filter(r => {
  const p = (r.folder_path || '').trim();
  const hasBackslash = p.includes('\\') || p.includes('/');
  const hasIds = r.store_id && r.entry_id;
  return hasBackslash || hasIds;
});
const invalid = rows.filter(r => {
  const p = (r.folder_path || '').trim();
  const hasBackslash = p.includes('\\') || p.includes('/');
  const hasIds = r.store_id && r.entry_id;
  return !hasBackslash && !hasIds && p;
});

console.log(`  ✅ ${valid.length} chemins valides (avec antislash ou IDs)`);
console.log(`  ❌ ${invalid.length} chemins invalides (orphelins)`);
if (invalid.length > 0) {
  console.log(`     ${invalid.map(r => `"${r.folder_path}"`).join(', ')}`);
}

db.close();
