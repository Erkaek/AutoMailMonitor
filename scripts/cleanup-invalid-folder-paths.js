#!/usr/bin/env node
/**
 * Script de nettoyage: Supprimer les chemins de dossiers invalides (trop courts)
 * 
 * Problème: Après le parser, des chemins orphelins comme "11- Tanguy" (sans boîte-mère)
 * restent dans folder_configurations. Ces chemins ne peuvent jamais être résolus par Outlook.
 * 
 * Solution: Identifier et supprimer les chemins trop courts (pas d'antislash, pas d'IDs)
 */

const path = require('path');
const Database = require('better-sqlite3');

const DB_PATH = path.join(__dirname, '../data.db');

if (!require('fs').existsSync(DB_PATH)) {
  console.error(`❌ BDD non trouvée: ${DB_PATH}`);
  process.exit(1);
}

const db = new Database(DB_PATH);

console.log('🔍 Scanning folder_configurations pour les chemins invalides...\n');

const rows = db.prepare(`
  SELECT rowid, folder_path, folder_name, store_id, entry_id 
  FROM folder_configurations
  ORDER BY folder_path ASC
`).all();

const invalid = rows.filter(row => {
  const p = (row.folder_path || '').trim();
  const hasBackslash = p.includes('\\') || p.includes('/');
  const hasIds = row.store_id && row.entry_id;
  return !hasBackslash && !hasIds && p.length > 0;
});

if (invalid.length === 0) {
  console.log('✅ Aucun chemin invalide détecté.');
  db.close();
  process.exit(0);
}

console.log(`⚠️  Trouvé ${invalid.length} chemins invalides (orphelins):\n`);
invalid.forEach((row, i) => {
  console.log(`  ${i + 1}. "${row.folder_path}" (folder_name="${row.folder_name}")`);
});

console.log(`\n🗑️  Suppression des ${invalid.length} chemins invalides...`);

try {
  const stmt = db.prepare('DELETE FROM folder_configurations WHERE folder_path = ?');
  let deleted = 0;
  
  invalid.forEach(row => {
    const info = stmt.run(row.folder_path);
    deleted += info.changes;
  });
  
  console.log(`✅ ${deleted} lignes supprimées.\n`);
  
  // Vérification post-nettoyage
  const remaining = db.prepare(`
    SELECT COUNT(*) as cnt FROM folder_configurations
  `).get();
  
  console.log(`ℹ️  Dossiers monitorés restants: ${remaining.cnt}`);
  
} catch (e) {
  console.error(`❌ Erreur lors de la suppression: ${e.message}`);
  process.exit(1);
} finally {
  db.close();
}

console.log('\n✨ Nettoyage terminé. Redémarrez l\'application pour recharger les dossiers.');
