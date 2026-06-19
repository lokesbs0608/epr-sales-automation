"use strict";
// repair_epr_from_portal.js
// Fix mis-stamped EPRs (from the early cross-window read-back race) using the
// portal export as ground truth. The portal export has no E-Invoice number, so
// we match each of OUR rows to a portal row by Quantity+Date (fallback rounded
// Amount+Date) and take the portal's real EPR (portal "Invoice No" col = EPR).
//
// Then we verify uniqueness: every E-Invoice -> one EPR, every EPR -> one E-Invoice.
//
// Usage:
//   node repair_epr_from_portal.js            # dry run (report only)
//   node repair_epr_from_portal.js --apply    # write fixes into the report sheets

const ExcelJS = require("exceljs");
const fs = require("fs");
const SAN = require("./sanitize");

const PORTAL = "/home/lokesh-b-s/Downloads/PORTAL ENTRIES PCL BAZPUR.xlsx";
const APPLY = process.argv.includes("--apply");
const VALID_EPR = /^\d{10,}$/;

function ct(v){return SAN.cellText(v);}
function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
function col(ws,names){let c=null;ws.getRow(1).eachCell((cell,n)=>{if(c===null&&names.some(x=>nh(x)===nh(ct(cell.value))))c=n;});return c;}
function dkey(v){
  if(v instanceof Date)return v.toISOString().slice(0,10);
  const s=ct(v);let m=s.match(/(\d{4})-(\d{2})-(\d{2})/);if(m)return m[0];
  m=s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/);
  if(m)return m[3]+"-"+String(m[2]).padStart(2,"0")+"-"+String(m[1]).padStart(2,"0");
  return s;
}
function num(v){const n=Number(String(ct(v)).replace(/[,\s₹]/g,""));return Number.isFinite(n)?n:NaN;}

const JOBS=[
  { file:"PP_FINAL_REPORT.xlsx",  sheet:"PP BAAZPUR" },
  { file:"PET_FINAL_REPORT.xlsx", sheet:"PET BAAZPUR" },
];

(async()=>{
  // ---- load portal: build qty+date and roundedAmt+date -> Set(EPR) ----
  const pwb=new ExcelJS.Workbook(); await pwb.xlsx.readFile(PORTAL);
  const pws=pwb.worksheets[0];
  const byQtyDate=new Map(), byAmtDate=new Map();
  for(let r=2;r<=pws.rowCount;r++){
    const epr=ct(pws.getRow(r).getCell(7).value); if(!VALID_EPR.test(epr))continue;
    const q=num(pws.getRow(r).getCell(3).value), d=dkey(pws.getRow(r).getCell(5).value), a=num(pws.getRow(r).getCell(4).value);
    const k1=q.toFixed(4)+"|"+d; if(!byQtyDate.has(k1))byQtyDate.set(k1,new Set()); byQtyDate.get(k1).add(epr);
    const k2=Math.round(a)+"|"+d; if(!byAmtDate.has(k2))byAmtDate.set(k2,new Set()); byAmtDate.get(k2).add(epr);
  }
  const portalEprs=new Set(); for(const s of byQtyDate.values())for(const e of s)portalEprs.add(e);
  console.log(`Portal rows loaded. unique EPRs=${portalEprs.size}\n`);

  for(const job of JOBS){
    if(!fs.existsSync(job.file)){ console.log(`SKIP ${job.file} (not found)`); continue; }
    const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(job.file);
    const ws=wb.getWorksheet(job.sheet)||wb.worksheets[0];
    const ec=col(ws,["EPR Invoice Number"]), ic=col(ws,["E-Invoice Number*","E-Invoice Number"]);
    const amtC=col(ws,["Principal Amount(₹)*"]), dateC=col(ws,["Sales date*"]), qtyC=col(ws,["Quantity Sold(MT)"]);

    // suspect rows = EPR shared by >1 row
    const byEpr=new Map();
    for(let r=2;r<=ws.rowCount;r++){const e=ct(ws.getRow(r).getCell(ec).value);if(VALID_EPR.test(e)){if(!byEpr.has(e))byEpr.set(e,[]);byEpr.get(e).push(r);}}
    const suspect=[]; for(const [e,rows] of byEpr){ if(rows.length>1) suspect.push(...rows); }

    // repair each suspect row from the portal
    let fixed=0, same=0, unmatched=0;
    const log=[];
    for(const r of suspect){
      const cur=ct(ws.getRow(r).getCell(ec).value);
      const inv=ic?ct(ws.getRow(r).getCell(ic).value):"";
      const d=dkey(ws.getRow(r).getCell(dateC).value);
      const qk=num(ws.getRow(r).getCell(qtyC).value).toFixed(4)+"|"+d;
      const ak=Math.round(num(ws.getRow(r).getCell(amtC).value))+"|"+d;
      let set=byQtyDate.get(qk); if(!set||set.size!==1) set=byAmtDate.get(ak);
      if(!set||set.size!==1){ unmatched++; log.push(`row ${r} ${inv}: NO UNIQUE PORTAL MATCH (kept ${cur})`); continue; }
      const real=[...set][0];
      if(real===cur){ same++; log.push(`row ${r} ${inv}: already correct (${cur})`); continue; }
      log.push(`row ${r} ${inv}: ${cur} -> ${real}  *FIX*`);
      if(APPLY){ ws.getRow(r).getCell(ec).value=real; ws.getRow(r).commit(); }
      fixed++;
    }

    // write
    if(APPLY && fixed){ const tmp=job.file+".tmp"; await wb.xlsx.writeFile(tmp); fs.renameSync(tmp, job.file); }

    // ---- uniqueness check (after fix in memory) ----
    const eprNow=new Map(), invNow=new Map();
    for(let r=2;r<=ws.rowCount;r++){
      const e=ct(ws.getRow(r).getCell(ec).value), iv=ic?ct(ws.getRow(r).getCell(ic).value):"";
      if(VALID_EPR.test(e)){ if(!eprNow.has(e))eprNow.set(e,[]); eprNow.get(e).push({r,iv}); }
      if(iv){ if(!invNow.has(iv))invNow.set(iv,new Set()); invNow.get(iv).add(e); }
    }
    const dupEpr=[...eprNow.entries()].filter(([,a])=>a.length>1);
    const dupInv=[...invNow.entries()].filter(([,s])=>[...s].filter(x=>VALID_EPR.test(x)).length>1);

    console.log(`==== ${job.file} (${job.sheet}) ====`);
    console.log(log.join("\n"));
    console.log(`\n  suspect=${suspect.length}  fixed=${fixed}  alreadyCorrect=${same}  unmatched=${unmatched}` + (APPLY?"  [WRITTEN]":"  [DRY RUN]"));
    console.log(`  AFTER: duplicate EPRs=${dupEpr.length}  duplicate E-Invoices(>1 EPR)=${dupInv.length}`);
    if(dupEpr.length) dupEpr.forEach(([e,a])=>console.log(`     EPR ${e} still on rows ${a.map(x=>x.r).join(",")}`));
    if(dupInv.length) dupInv.forEach(([iv,s])=>console.log(`     E-Invoice ${iv} on EPRs ${[...s].join(",")}`));
    console.log("");
  }
})().catch(e=>{console.error("FATAL:",e?.stack||e);process.exit(1);});
