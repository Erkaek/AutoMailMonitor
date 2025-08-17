/**
 * Importeur d'activité .xlsb avec coordonnées fixes
 * - Lit les feuilles S1..S52
 * - Récupère C7/D7/E7, C10/D10/E10, C13/D13/E13 et (en S1) M7/M10/M13
 * - Calcule stock_debut/stock_fin en rolling
 * - Ignore les semaines vides
 */

const path = require('path');
const fs = require('fs');
const { execFileSync, spawnSync } = require('child_process');

const CATS = [
  { key: 'Declarations', row: 7 },
  { key: 'Reglements', row: 10 },
  { key: 'MailSimple', row: 13 }
];

function toInt(n) {
  const v = Number(n);
  return Number.isFinite(v) ? Math.trunc(v) : 0;
}

function readCell(sheet, addr) {
  const cell = sheet[addr];
  if (!cell) return 0;
  const v = cell.v ?? cell.w;
  return toInt(v);
}

function getMondayISO(year, week) {
  // ISO week: Monday as start
  const simple = new Date(Date.UTC(year, 0, 1 + (week - 1) * 7));
  const dayOfWeek = simple.getUTCDay();
  const ISOweekStart = new Date(simple);
  const diff = (dayOfWeek <= 4 ? dayOfWeek : dayOfWeek - 7) - 1; // Monday=1
  ISOweekStart.setUTCDate(simple.getUTCDate() - diff);
  // Force to Monday
  ISOweekStart.setUTCHours(0, 0, 0, 0);
  return ISOweekStart.toISOString().slice(0, 10);
}

function inferYearFromFilename(filePath) {
  const base = path.basename(filePath);
  const m = base.match(/(20\d{2})/);
  if (m) {
    const y = parseInt(m[1]);
    if (y >= 2000 && y <= 2099) return y;
  }
  return new Date().getFullYear();
}

function parseWeeksFilter(filterStr) {
  // e.g. "1-10,12,14-20" -> Set of weeks
  if (!filterStr) return null; // null means all
  const set = new Set();
  const parts = String(filterStr).split(',').map(s => s.trim()).filter(Boolean);
  for (const p of parts) {
    if (p.includes('-')) {
      const [a, b] = p.split('-').map(x => parseInt(x.trim(), 10));
      if (Number.isInteger(a) && Number.isInteger(b)) {
        const start = Math.min(a, b);
        const end = Math.max(a, b);
        for (let k = start; k <= end; k++) set.add(k);
      }
    } else {
      const n = parseInt(p, 10);
      if (Number.isInteger(n)) set.add(n);
    }
  }
  return set;
}

/**
 * Calcule les lignes d'activité à partir d'un classeur XLSX (déjà chargé)
 * workbook: objet renvoyé par XLSX.read
 * options: { year?, weeksFilter?: Set<number> }
 * Retourne { rows, skippedWeeks }
 */
function computeFromWorkbook(workbook, options = {}) {
  const { year: yOpt, weeksFilter } = options;
  const year = yOpt || new Date().getFullYear();

  let initialStocks = { Declarations: 0, Reglements: 0, MailSimple: 0 };

  const rows = [];
  let skippedWeeks = 0;

  // Pre-read S1 initial stocks
  const s1 = workbook.Sheets['S1'];
  if (s1) {
    initialStocks = {
      Declarations: readCell(s1, 'M7'),
      Reglements: readCell(s1, 'M10'),
      MailSimple: readCell(s1, 'M13')
    };
  }

  // Track running stocks per category
  const stockDebut = { ...initialStocks };

  for (let k = 1; k <= 52; k++) {
    if (weeksFilter && !weeksFilter.has(k)) continue;
    const sheetName = `S${k}`;
    const sh = workbook.Sheets[sheetName];
    if (!sh) {
      skippedWeeks++;
      continue;
    }

    // Read values for all three categories
    const weekValues = {};
    let allZero = true;
    for (const cat of CATS) {
      const recu = readCell(sh, `C${cat.row}`);
      const traite = readCell(sh, `D${cat.row}`);
      const traite_adg = readCell(sh, `E${cat.row}`);
      if (recu || traite || traite_adg) allZero = false;
      weekValues[cat.key] = { recu, traite, traite_adg };
    }

    // If week is empty and (if k=1) initial stocks are zero -> skip entirely
    const s1InitialZero = (k !== 1) || (initialStocks.Declarations === 0 && initialStocks.Reglements === 0 && initialStocks.MailSimple === 0);
    if (allZero && s1InitialZero) {
      skippedWeeks++;
      continue;
    }

    // For each category compute stock_debut/fin with rolling
    for (const cat of CATS) {
      const catKey = cat.key;
      const { recu, traite, traite_adg } = weekValues[catKey];
      const sd = k === 1 ? (stockDebut[catKey] ?? 0) : (stockDebut[catKey] ?? 0);
      const sf = sd + recu - (traite + traite_adg);
      const week_start_date = getMondayISO(year, k);

      rows.push({
        year,
        week_number: k,
        week_start_date,
        category: catKey,
        recu, traite, traite_adg,
        stock_debut: sd,
        stock_fin: sf
      });

      // Update rolling stock for next week
      stockDebut[catKey] = sf;
    }
  }

  return { rows, skippedWeeks };
}

/**
 * Importe depuis un fichier .xlsb et retourne les lignes calculées
 * options: { year?, weeks?: string }
 */
function importActivityFromXlsb(filePath, options = {}) {
  if (!filePath || !fs.existsSync(filePath)) {
    throw new Error('Fichier d\'entrée introuvable');
  }
  const year = options.year || inferYearFromFilename(filePath);

  // 1) Try secure PowerShell + Excel COM path (no JS parsing of untrusted binary)
  try {
    const psPath = path.join(__dirname, '..', '..', 'scripts', 'import-xlsb.ps1');
    if (fs.existsSync(psPath)) {
      const args = ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', psPath, '-Path', filePath, '-Year', String(year)];
      if (options.weeks) {
        args.push('-Weeks', String(options.weeks));
      }
      const res = spawnSync('powershell.exe', args, {
        encoding: 'utf8',
        windowsHide: true,
        timeout: 60_000
      });
      if (res.status === 0 && res.stdout) {
        const parsed = JSON.parse(res.stdout.trim());
        if (parsed && Array.isArray(parsed.rows)) return parsed;
      } else if (res.stderr) {
        // Fall through to sandboxed JS fallback
        // console.warn('[XLSB PS] stderr:', res.stderr);
      }
    }
  } catch (_) {
    // ignore and fallback
  }

  // 2) Sandbox fallback: run parsing in a separate Node process to isolate any prototype pollution
  //    Can be disabled with XLSB_IMPORT_DISABLE_JS=1 to require Excel COM path only
  if (String(process.env.XLSB_IMPORT_DISABLE_JS || '').toLowerCase() === '1') {
    throw new Error('Échec import XLSB via PowerShell et fallback JS désactivé (XLSB_IMPORT_DISABLE_JS=1)');
  }
  try {
    const reader = path.join(__dirname, '..', '..', 'scripts', 'xlsb-safe-reader.js');
    const args = [reader, filePath, String(options.weeks || ''), String(year)];
    const out = execFileSync(process.execPath, args, { encoding: 'utf8', env: { ...process.env, ELECTRON_RUN_AS_NODE: '1' } });
    const parsed = JSON.parse(out.trim());
    if (parsed && Array.isArray(parsed.rows)) return parsed;
  } catch (e) {
    // Last resort: propagate error with context
    throw new Error('Échec import XLSB (PS et sandbox): ' + e.message);
  }
}

function toCsv(rows) {
  const header = ['year','week_number','week_start_date','category','recu','traite','traite_adg','stock_debut','stock_fin'];
  const lines = [header.join(',')];
  for (const r of rows) {
    lines.push([
      r.year,
      r.week_number,
      r.week_start_date,
      r.category,
      r.recu,
      r.traite,
      r.traite_adg,
      r.stock_debut,
      r.stock_fin
    ].join(','));
  }
  return lines.join('\n');
}

function writeCsv(rows, year, outDir = path.join(process.cwd(), 'build')) {
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, `activity_${year}.csv`);
  fs.writeFileSync(outPath, toCsv(rows), 'utf8');
  return outPath;
}

module.exports = {
  importActivityFromXlsb,
  computeFromWorkbook,
  parseWeeksFilter,
  writeCsv
};
