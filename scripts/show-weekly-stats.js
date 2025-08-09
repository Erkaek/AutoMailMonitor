// Affiche un aperçu de la table weekly_stats (data/emails.db)
// Usage: npm run db:weekly [annee]

const path = require('path');
const Database = require('better-sqlite3');

const yearArg = process.argv[2] ? parseInt(process.argv[2], 10) : null;

const dbPath = path.join(__dirname, '..', 'data', 'emails.db');
const db = new Database(dbPath, { readonly: true });

const hasTable = db
  .prepare("SELECT name FROM sqlite_master WHERE type='table' AND name='weekly_stats'")
  .get();
if (!hasTable) {
  console.error('Table weekly_stats introuvable dans', dbPath);
  process.exit(1);
}

let query = `
  SELECT week_year, week_number, week_start_date, week_end_date, folder_type,
         emails_received, emails_treated, manual_adjustments,
         (emails_received + manual_adjustments) as total_received
  FROM weekly_stats`;
const params = [];
if (yearArg) {
  query += ' WHERE week_year = ?';
  params.push(yearArg);
}
query += ' ORDER BY week_year DESC, week_number DESC, folder_type ASC';

const rows = db.prepare(query).all(...params);
if (!rows.length) {
  console.log('Aucune ligne weekly_stats trouvée', yearArg ? `pour ${yearArg}` : '');
  process.exit(0);
}

console.log(`DB: ${dbPath}`);
console.log(`Total lignes: ${rows.length}`);
console.log('--- Aperçu (premières 15 lignes) ---');
for (let i = 0; i < Math.min(15, rows.length); i++) {
  const r = rows[i];
  console.log(`${r.week_year}-W${String(r.week_number).padStart(2,'0')} (${r.week_start_date} -> ${r.week_end_date}) | ${r.folder_type} | recu=${r.emails_received} traité=${r.emails_treated} total_recu=${r.total_received}`);
}
