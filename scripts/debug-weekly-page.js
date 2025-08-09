#!/usr/bin/env node
/**
 * Debug weekly page transformation (carry-aware)
 * Usage: node scripts/debug-weekly-page.js [page] [pageSize]
 */
const db = require('../src/services/optimizedDatabaseService');

(async () => {
  try {
    const page = parseInt(process.argv[2] || '1', 10);
  const pageSize = parseInt(process.argv[3] || '5', 10);

    if (!db.isInitialized) await db.initialize();

    const result = db.getWeeklyHistoryPage(page, pageSize);
    const weeklyStats = result.rows;

    // Group by week
    const weeklyGroups = {};
    weeklyStats.forEach(row => {
      const weekKey = `S${row.week_number} - ${row.week_year}`;
      if (!weeklyGroups[weekKey]) {
        weeklyGroups[weekKey] = {
          weekDisplay: weekKey,
          week_number: row.week_number,
          week_year: row.week_year,
          dateRange: `${row.week_start_date || ''} -> ${row.week_end_date || ''}`,
          categories: {
            'Déclarations': { received: 0, treated: 0, adjustments: 0 },
            'Règlements': { received: 0, treated: 0, adjustments: 0 },
            'Mails simples': { received: 0, treated: 0, adjustments: 0 }
          }
        };
      }
      let category = row.folder_type || 'Mails simples';
      if (category === 'mails_simples') category = 'Mails simples';
      else if (category === 'declarations') category = 'Déclarations';
      else if (category === 'reglements') category = 'Règlements';

      const received = row.emails_received || 0;
      const treated = row.emails_treated || 0;
      const adjustments = row.manual_adjustments || 0;

      if (weeklyGroups[weekKey].categories[category]) {
        weeklyGroups[weekKey].categories[category] = {
          received,
          treated,
          adjustments,
          stockEndWeek: Math.max(0, (received || 0) - ((treated || 0) + (adjustments || 0)))
        };
      }
    });

    const sortedWeeks = Object.keys(weeklyGroups).sort((a, b) => {
      const aMatch = a.match(/S(\d+) - (\d+)/);
      const bMatch = b.match(/S(\d+) - (\d+)/);
      if (aMatch && bMatch) {
        const aYear = parseInt(aMatch[2]);
        const bYear = parseInt(bMatch[2]);
        if (aYear !== bYear) return bYear - aYear;
        return parseInt(bMatch[1]) - parseInt(aMatch[1]);
      }
      return 0;
    });

    let startYear = null, startWeek = null;
    if (sortedWeeks.length) {
      const m = sortedWeeks[sortedWeeks.length - 1].match(/S(\d+) - (\d+)/);
      if (m) { startWeek = parseInt(m[1], 10); startYear = parseInt(m[2], 10); }
    }

    const carryInitial = (startYear && startWeek)
      ? db.getCarryBeforeWeek(startYear, startWeek)
      : { declarations: 0, reglements: 0, mails_simples: 0 };

    const running = {
      'Déclarations': carryInitial.declarations || 0,
      'Règlements': carryInitial.reglements || 0,
      'Mails simples': carryInitial.mails_simples || 0
    };

    const transformed = [];
    for (let i = sortedWeeks.length - 1; i >= 0; i--) {
      const wk = sortedWeeks[i];
      const weekData = weeklyGroups[wk];
      const cats = ['Déclarations', 'Règlements', 'Mails simples'].map(name => {
        const rec = weekData.categories[name].received || 0;
        const trt = weekData.categories[name].treated || 0;
        const adj = weekData.categories[name].adjustments || 0;
        const start = running[name] || 0;
        const end = Math.max(0, start + rec - (trt + adj));
        running[name] = end;
        return { name, received: rec, treated: trt, adjustments: adj, stockStart: start, stockEndWeek: end };
      });
      transformed.push({ weekDisplay: weekData.weekDisplay, week_number: weekData.week_number, week_year: weekData.week_year, dateRange: weekData.dateRange, categories: cats });
    }
    transformed.reverse();

    console.log('=== Weekly Page Debug ===');
    for (const w of transformed) {
      const total = w.categories.reduce((s, c) => s + (c.stockEndWeek || 0), 0);
      console.log(`${w.weekDisplay} [${w.dateRange}] → Total=${total}`);
      for (const c of w.categories) {
        console.log(`  - ${c.name}: start=${c.stockStart}, rec=${c.received}, trt=${c.treated}, adj=${c.adjustments} → end=${c.stockEndWeek}`);
      }
    }
  } catch (e) {
    console.error('Error:', e);
    process.exit(1);
  } finally {
    process.exit(0);
  }
})();
