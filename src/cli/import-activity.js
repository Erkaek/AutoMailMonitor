#!/usr/bin/env node
/**
 * CLI d'import activité (.xlsb)
 * Options:
 *  --input <path> (obligatoire)
 *  --weeks "1-10,12,14-20" (optionnel)
 *  --out-csv (écrit build/activity_<year>.csv)
 */
const path = require('path');
const fs = require('fs');
const { importActivityFromXlsb, writeCsv } = require('../importers/activityXlsbImporter');
const optimizedDatabaseService = require('../services/optimizedDatabaseService');

function parseArgs(argv) {
  const args = { _: [] };
  for (let i = 2; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--input' || a === '-i') { args.input = argv[++i]; }
    else if (a === '--weeks' || a === '-w') { args.weeks = argv[++i]; }
    else if (a === '--out-csv') { args.outCsv = true; }
    else { args._.push(a); }
  }
  return args;
}

async function main() {
  const args = parseArgs(process.argv);
  if (!args.input) {
    console.error('Usage: import-activity --input <file.xlsb> [--weeks "1-10,12"] [--out-csv]');
    process.exit(1);
  }

  await optimizedDatabaseService.initialize();

  const { rows, skippedWeeks, year } = importActivityFromXlsb(args.input, { weeks: args.weeks });
  // Plus de backfill partiel depuis activity_weekly (supprimé). L'import est calculé depuis le fichier.

  // Agréger et écrire dans weekly_stats
  const mapCategoryToFolderType = (cat) => {
    if (!cat) return 'Mails simples';
    if (cat === 'MailSimple') return 'Mails simples';
    if (cat === 'Reglements') return 'Règlements';
    if (cat === 'Declarations') return 'Déclarations';
    return 'Mails simples';
  };
  const getISOWeekInfo = (y, w) => {
    const simple = new Date(Date.UTC(y, 0, 1 + (w - 1) * 7));
    const day = simple.getUTCDay();
    const diff = (day <= 4 ? day : day - 7) - 1; // Monday=1
    const start = new Date(simple);
    start.setUTCDate(simple.getUTCDate() - diff);
    start.setUTCHours(0,0,0,0);
    const end = new Date(start);
    end.setUTCDate(start.getUTCDate() + 6);
    const startStr = start.toISOString().slice(0,10);
    const endStr = end.toISOString().slice(0,10);
    const weekId = `${y}-W${String(w).padStart(2,'0')}`;
    return { startStr, endStr, weekId };
  };
  const agg = new Map();
  for (const r of rows) {
    const folderType = mapCategoryToFolderType(r.category);
    const key = `${r.year}-${r.week_number}-${folderType}`;
    const { startStr, endStr, weekId } = getISOWeekInfo(r.year, r.week_number);
    if (!agg.has(key)) {
      agg.set(key, {
        week_identifier: weekId,
        week_number: r.week_number,
        week_year: r.year,
        week_start_date: startStr,
        week_end_date: endStr,
        folder_type: folderType,
        emails_received: 0,
    emails_treated: 0,
    manual_adjustments: 0
      });
    }
    const a = agg.get(key);
    a.emails_received += (r.recu || 0);
  a.emails_treated += (r.traite || 0);
  a.manual_adjustments += (r.traite_adg || 0);
  }
  if (agg.size > 0 && optimizedDatabaseService.upsertWeeklyStatsBatch) {
    optimizedDatabaseService.upsertWeeklyStatsBatch(Array.from(agg.values()));
  }

  // CSV optionnel
  let csvPath = null;
  if (args.outCsv) {
    csvPath = writeCsv(rows, year, path.join(process.cwd(), 'build'));
  }

  const inserted = rows.length;
  console.log(`✅ Import terminé: ${inserted} lignes source, ${agg.size} semaines agrégées, ${skippedWeeks} semaine(s) ignorée(s)${csvPath ? `, CSV: ${csvPath}` : ''}`);
}

if (require.main === module) {
  main().catch(err => { console.error('❌ Import échoué:', err); process.exit(1); });
}
