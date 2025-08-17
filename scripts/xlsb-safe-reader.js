#!/usr/bin/env node
// Minimal, isolated XLSB reader: runs in a separate Node process to sandbox any prototype pollution
// Input: argv[2]=filePath, argv[3]=weeks ("" for all), argv[4]=year
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function toInt(n){const v=Number(n);return Number.isFinite(v)?Math.trunc(v):0}
function readCell(sheet,addr){const cell=sheet[addr];if(!cell) return 0;const v=(cell.v??cell.w);return toInt(v)}
function getMondayISO(year,week){const simple=new Date(Date.UTC(year,0,1+(week-1)*7));const dayOfWeek=simple.getUTCDay();const diff=((dayOfWeek<=4?dayOfWeek:dayOfWeek-7)-1);const ISOweekStart=new Date(simple);ISOweekStart.setUTCDate(simple.getUTCDate()-diff);ISOweekStart.setUTCHours(0,0,0,0);return ISOweekStart.toISOString().slice(0,10)}
function parseWeeksFilter(s){if(!s) return null; const set=new Set(); for(const p of String(s).split(',').map(x=>x.trim()).filter(Boolean)){ if(p.includes('-')){ const [a,b]=p.split('-').map(n=>parseInt(n,10)); if(Number.isInteger(a)&&Number.isInteger(b)){ const start=Math.min(a,b); const end=Math.max(a,b); for(let k=start;k<=end;k++) set.add(k); } } else { const n=parseInt(p,10); if(Number.isInteger(n)) set.add(n); } } return set }

const filePath=process.argv[2];
const weeksArg=process.argv[3];
const year=parseInt(process.argv[4],10) || new Date().getFullYear();
if(!filePath || !fs.existsSync(filePath)){
  console.error('Input file not found');
  process.exit(2);
}
const weeksFilter=parseWeeksFilter(weeksArg);
const wb=XLSX.read(fs.readFileSync(filePath),{type:'buffer',WTF:false});
let initial={Declarations:0,Reglements:0,MailSimple:0};
const s1=wb.Sheets['S1'];
if(s1){initial={Declarations:readCell(s1,'M7'),Reglements:readCell(s1,'M10'),MailSimple:readCell(s1,'M13')}}
const stock={...initial};
const rows=[]; let skipped=0;
for(let k=1;k<=52;k++){
  if(weeksFilter && !weeksFilter.has(k)) continue;
  const sh=wb.Sheets['S'+k];
  if(!sh){skipped++; continue}
  const cats=[{key:'Declarations',row:7},{key:'Reglements',row:10},{key:'MailSimple',row:13}];
  const weekValues={}; let allZero=true;
  for(const cat of cats){ const recu=readCell(sh,'C'+cat.row); const traite=readCell(sh,'D'+cat.row); const traite_adg=readCell(sh,'E'+cat.row); if(recu||traite||traite_adg) allZero=false; weekValues[cat.key]={recu,traite,traite_adg}; }
  const s1Zero=(k!==1) || (initial.Declarations===0 && initial.Reglements===0 && initial.MailSimple===0);
  if(allZero && s1Zero){ skipped++; continue }
  for(const cat of [{key:'Declarations',row:7},{key:'Reglements',row:10},{key:'MailSimple',row:13}]){
    const vals=weekValues[cat.key]; const sd=stock[cat.key]??0; const sf=sd+vals.recu-(vals.traite+vals.traite_adg); const week_start_date=getMondayISO(year,k);
    rows.push({year,week_number:k,week_start_date,category:cat.key,recu:vals.recu,traite:vals.traite,traite_adg:vals.traite_adg,stock_debut:sd,stock_fin:sf});
    stock[cat.key]=sf;
  }
}
const payload=JSON.stringify({rows,skippedWeeks:skipped,year});
process.stdout.write(payload);
