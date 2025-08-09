/**
 * Debug script: computes carry baseline before week 3 and verifies S2→S3 stock for "Déclarations".
 * It finds the latest year in the DB that has weeks 2 and 3 for Déclarations,
 * then prints received/treated/adjustments for both weeks, baseline before S3,
 * and computed stockEnd for S3.
 */
const dbService = require('../src/services/optimizedDatabaseService');

function pickDeclarationsFolderFilter() {
  // Accept both localized and internal identifiers
  return `(
    LOWER(folder_type) IN ('declarations')
    OR folder_type LIKE 'Déclaration%'
    OR folder_type LIKE 'DÃ©claration%'
  )`;
}

async function main() {
  await dbService.initialize();

  const whereDecl = pickDeclarationsFolderFilter();
  const yearRow = dbService.db.prepare(
    `SELECT week_year
     FROM weekly_stats
     WHERE week_number IN (2,3) AND ${whereDecl}
     GROUP BY week_year
     HAVING SUM(CASE WHEN week_number=2 THEN 1 ELSE 0 END)>0
        AND SUM(CASE WHEN week_number=3 THEN 1 ELSE 0 END)>0
     ORDER BY week_year DESC
     LIMIT 1`
  ).get();

  if (!yearRow) {
    console.log('No suitable year found with weeks 2 and 3 for Déclarations.');
    return;
  }
  const year = yearRow.week_year;

  const rows = dbService.db.prepare(
    `SELECT week_number,
            SUM(COALESCE(emails_received,0)) AS rec,
            SUM(COALESCE(emails_treated,0)) AS trt,
            SUM(COALESCE(manual_adjustments,0)) AS adj
     FROM weekly_stats
     WHERE week_year = ? AND week_number IN (2,3) AND ${whereDecl}
     GROUP BY week_number
     ORDER BY week_number ASC`
  ).all(year);

  const s2 = rows.find(r => r.week_number === 2) || { rec: 0, trt: 0, adj: 0 };
  const s3 = rows.find(r => r.week_number === 3) || { rec: 0, trt: 0, adj: 0 };

  const carry = dbService.getCarryBeforeWeek(year, 3);
  const baseline = carry?.declarations || 0; // stock fin S2 attendu

  const startS3 = baseline;
  const endS3 = Math.max(0, startS3 + (s3.rec || 0) - ((s3.trt || 0) + (s3.adj || 0)));

  console.log('--- Debug S2→S3 Déclarations ---');
  console.log('Year:', year);
  console.log('S2:', s2);
  console.log('S3:', s3);
  console.log('Baseline before S3 (end of S2):', baseline);
  console.log('Computed S3 end stock:', endS3);
}

main().catch(e => {
  console.error('Debug error:', e);
  process.exitCode = 1;
});
