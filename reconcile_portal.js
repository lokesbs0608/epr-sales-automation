"use strict";
// reconcile_portal.js
// Reconcile our report sheets against the portal export (ground truth):
//   1. BACKFILL: rows that have NO EPR in our sheet but DO exist on the portal
//      (matched by Quantity+Date, fallback rounded-Amount+Date) -> fill the real EPR.
//   2. DETECT duplicates: report every EPR that the PORTAL itself assigned to 2+
//      different sales (server-side collision) so they can be fixed manually.
//
// Usage:
//   node reconcile_portal.js          # dry run (report only)
//   node reconcile_portal.js --apply  # write the backfilled EPRs into the sheets

const ExcelJS = require("exceljs");
const fs = require("fs");
const SAN = require("./sanitize");

const PORTAL = "/home/lokesh-b-s/Downloads/PORTAL ENTRIES PCL BAZPUR.xlsx";
const APPLY = process.argv.includes("--apply");
const V = /^\d{10,}$/;

function ct(v){return SAN.cellText(v);}
function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
function col(ws,names){let c=null;ws.getRow(1).eachCell((cell,n)=>{if(c===null&&names.some(x=>nh(x)===nh(ct(cell.value))))c=n;});return c;}
function dkey(v){if(v instanceof Date)return v.toISOString().slice(0,10);const s=ct(v);let m=s.match(/(\d{4})-(\d{2})-(\d{2})/);if(m)return m[0];m=s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/);if(m)return m[3]+"-"+String(m[2]).padStart(2,"0")+"-"+String(m[1]).padStart(2,"0");return s;}
function num(v){const n=Number(String(ct(v)).replace(/[,\s₹]/g,""));return Number.isFinite(n)?n:NaN;}

const JOBS=[
  { file:"PP_FINAL_REPORT.xlsx",  sheet:"PP BAAZPUR" },
  { file:"PET_FINAL_REPORT.xlsx", sheet:"PET BAAZPUR" },
];

(async()=>{
  // ---- load portal ----
  const pwb=new ExcelJS.Workbook(); await pwb.xlsx.readFile(PORTAL);
  const pws=pwb.worksheets[0]; // col3=qty col4=amount col5=date col7=EPR
  const byQ=new Map(), byA=new Map();
  const eprCount=new Map();
  for(let r=2;r<=pws.rowCount;r++){
    const e=ct(pws.getRow(r).getCell(7).value); if(!V.test(e))continue;
    eprCount.set(e,(eprCount.get(e)||0)+1);
    const q=num(pws.getRow(r).getCell(3).value).toFixed(4)+"|"+dkey(pws.getRow(r).getCell(5).value);
    const a=Math.round(num(pws.getRow(r).getCell(4).value))+"|"+dkey(pws.getRow(r).getCell(5).value);
    if(!byQ.has(q))byQ.set(q,new Set()); byQ.get(q).add(e);
    if(!byA.has(a))byA.set(a,new Set()); byA.get(a).add(e);
  }
  const portalDupEprs=new Set([...eprCount].filter(([,c])=>c>1).map(([e])=>e));
  console.log(`Portal: ${eprCount.size} unique EPRs; ${portalDupEprs.size} EPRs duplicated on the portal itself.\n`);

  for(const job of JOBS){
    if(!fs.existsSync(job.file)){ console.log(`SKIP ${job.file}`); continue; }
    const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(job.file);
    const ws=wb.getWorksheet(job.sheet)||wb.worksheets[0];
    const ec=col(ws,["EPR Invoice Number"]), ic=col(ws,["E-Invoice Number*","E-Invoice Number"]);
    const amtC=col(ws,["Principal Amount(₹)*"]), dateC=col(ws,["Sales date*"]), qtyC=col(ws,["Quantity Sold(MT)"]);

    let backfilled=0, ambiguousBlank=0, stillMissing=0;
    const bfLog=[];
    for(let r=2;r<=ws.rowCount;r++){
      const inv=ic?ct(ws.getRow(r).getCell(ic).value):"";
      if(!inv) continue;
      const e=ct(ws.getRow(r).getCell(ec).value);
      if(V.test(e)) continue; // already has EPR
      const d=dkey(ws.getRow(r).getCell(dateC).value);
      const qk=num(ws.getRow(r).getCell(qtyC).value).toFixed(4)+"|"+d;
      const ak=Math.round(num(ws.getRow(r).getCell(amtC).value))+"|"+d;
      let set=byQ.get(qk); if(!set) set=byA.get(ak);
      if(!set){ stillMissing++; continue; } // genuinely not on portal -> not submitted
      if(set.size!==1){ ambiguousBlank++; bfLog.push(`row ${r} ${inv}: AMBIGUOUS portal match (${[...set].join(",")}) - not filled`); continue; }
      const real=[...set][0];
      bfLog.push(`row ${r} ${inv}: BLANK -> ${real}` + (portalDupEprs.has(real)?"  (note: this EPR is ALSO duplicated on portal)":""));
      if(APPLY){ ws.getRow(r).getCell(ec).value=real; ws.getRow(r).commit(); }
      backfilled++;
    }
    if(APPLY && backfilled){ const tmp=job.file+".tmp"; await wb.xlsx.writeFile(tmp); fs.renameSync(tmp, job.file); }

    // duplicates present in OUR sheet (after backfill, in memory)
    const byEpr=new Map();
    for(let r=2;r<=ws.rowCount;r++){const e=ct(ws.getRow(r).getCell(ec).value);if(V.test(e)){if(!byEpr.has(e))byEpr.set(e,[]);byEpr.get(e).push(r-1);}}
    const dupInSheet=[...byEpr.entries()].filter(([,a])=>a.length>1);

    console.log(`==== ${job.file} ====`);
    if(bfLog.length){ console.log("Backfill candidates (blank in sheet, present on portal):"); bfLog.forEach(l=>console.log("  "+l)); }
    console.log(`\n  backfilled=${backfilled}  ambiguous=${ambiguousBlank}  notOnPortal(not submitted)=${stillMissing}` + (APPLY?"  [WRITTEN]":"  [DRY RUN]"));
    console.log(`  duplicate EPRs in this sheet: ${dupInSheet.length}`);
    dupInSheet.forEach(([e,rows])=>console.log(`     EPR ${e} on rows ${rows.join(",")}`+(portalDupEprs.has(e)?"  (duplicated on portal too)":"  (sheet-only dup)")));
    console.log("");
  }
})().catch(e=>{console.error("FATAL:",e?.stack||e);process.exit(1);});
