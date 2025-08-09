/**
 * Test script: verifies weekly carry-over (stock) across two weeks.
 * Scenario (Déclarations):
 *  - Week S31: start=?, received=20, treated=0 => end = start + 20
 *  - Week S32: start=end(S31), received=10, treated=15 => end = start + 10 - 15
 * The script uses the same DB service as the app and restores previous data.
 */

const path = require('path');
const dbService = require('../src/services/optimizedDatabaseService');

// Choose a year unlikely to collide; if existing data exists, we properly compute baseline
const YEAR = 2038;
const W1 = 31;
const W2 = 32;
const CATEGORY = 'Déclarations';

function getISOWeekDates(year, week) {
  // Monday-Sunday of ISO week
  const simple = new Date(Date.UTC(year, 0, 1 + (week - 1) * 7));
  const day = simple.getUTCDay();
  const diff = (day <= 4 ? day : day - 7) - 1; // Monday=1
  const start = new Date(simple);
  start.setUTCDate(simple.getUTCDate() - diff);
  start.setUTCHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setUTCDate(start.getUTCDate() + 6);
  const startStr = start.toISOString().slice(0, 10);
  const endStr = end.toISOString().slice(0, 10);
  const weekId = `S${week}-${year}`;
  return { startStr, endStr, weekId };
}

async function main() {
  await dbService.initialize();

  const { startStr: s31, endStr: e31, weekId: id31 } = getISOWeekDates(YEAR, W1);
  const { startStr: s32, endStr: e32, weekId: id32 } = getISOWeekDates(YEAR, W2);

  const readExisting = (id) => dbService.db.prepare(
    `SELECT * FROM weekly_stats WHERE week_identifier = ? AND folder_type = ?`
  ).all(id, CATEGORY);
  const existing31 = readExisting(id31);
  const existing32 = readExisting(id32);

  // Seed test data
  const rows = [
    {
      week_identifier: id31,
      week_number: W1,
      week_year: YEAR,
      week_start_date: s31,
      week_end_date: e31,
      folder_type: CATEGORY,
      emails_received: 20,
      emails_treated: 0,
      manual_adjustments: 0,
      created_at: null
    },
    {
      week_identifier: id32,
      week_number: W2,
      week_year: YEAR,
      week_start_date: s32,
      week_end_date: e32,
      folder_type: CATEGORY,
      emails_received: 10,
      emails_treated: 15,
      manual_adjustments: 0,
      created_at: null
    }
  ];

  try {
    dbService.upsertWeeklyStatsBatch(rows);

    // Compute baseline carry before oldest week
    const carry = dbService.getCarryBeforeWeek(YEAR, W1);
    const baseline = carry?.declarations || 0;

    // Fetch both weeks and compute rolling stock according to IPC logic
    const all = dbService.db.prepare(
      `SELECT * FROM weekly_stats WHERE week_identifier IN (?, ?) AND folder_type = ? ORDER BY week_year ASC, week_number ASC`
    ).all(id31, id32, CATEGORY);

    let running = baseline;
    const results = [];
    for (const r of all) {
      const rec = r.emails_received || 0;
      const trt = r.emails_treated || 0;
      const adj = r.manual_adjustments || 0;
      const start = running;
      const end = Math.max(0, start + rec - (trt + adj));
      running = end;
      results.push({ week: r.week_identifier, start, rec, trt, adj, end });
    }

    // Expected per the scenario (with baseline accounted):
    // end(S31) = baseline + 20 - 0
    // end(S32) = end(S31) + 10 - 15
    const expected31 = Math.max(0, baseline + 20 - 0);
    const expected32 = Math.max(0, expected31 + 10 - 15);

    const r31 = results.find(r => r.week === id31);
    const r32 = results.find(r => r.week === id32);

    console.log('--- Carry-over test (Déclarations) ---');
    console.log(`Baseline before S${W1}-${YEAR}:`, baseline);
    console.log('Computed:', results);
    console.log('Expectations:', { [`${id31}`]: expected31, [`${id32}`]: expected32 });

    const pass = r31 && r32 && r31.end === expected31 && r32.end === expected32;
    console.log(pass ? 'PASS ✅' : 'FAIL ❌');
    process.exitCode = pass ? 0 : 1;
  } catch (e) {
    console.error('Test error:', e);
    process.exitCode = 2;
  } finally {
    // Restore previous state
    const tx = dbService.db.transaction(() => {
      // Clear our rows
      dbService.db.prepare(
        `DELETE FROM weekly_stats WHERE week_identifier IN (?, ?) AND folder_type = ?`
      ).run(id31, id32, CATEGORY);

      // Restore previous, if any
      const restore = (arr) => {
        if (!arr || arr.length === 0) return;
        for (const r of arr) {
          dbService.upsertWeeklyStats({
            week_identifier: r.week_identifier,
            week_number: r.week_number,
            week_year: r.week_year,
            week_start_date: r.week_start_date,
            week_end_date: r.week_end_date,
            folder_type: r.folder_type,
            emails_received: r.emails_received,
            emails_treated: r.emails_treated,
            manual_adjustments: r.manual_adjustments || 0,
            created_at: r.created_at || null
          });
        }
      };
      restore(existing31);
      restore(existing32);
    });
    tx();
  }
}

main();
