"use strict";
// verify_against_portal.js — full audit: for EVERY row, what EPR does the portal
// assign for that exact sale (Amount + Qty + Date)? Compare to our sheet's EPR.
// Flags: rows whose EPR is on the WRONG row (mismatch), rows we can fix, dupes.
//   node verify_against_portal.js          # report only
//   node verify_against_portal.js --apply  # rewrite EPR column from portal truth

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
  // ---- load portal: build several keys -> Set(EPR). Strongest first. ----
  const pwb=new ExcelJS.Workbook(); await pwb.xlsx.readFile(PORTAL);
  const pws=pwb.worksheets[0]; // 3=qty 4=amount 5=date 7=EPR
  const kQAD=new Map();  // qty(4dp)|amount(2dp)|date  -> Set(EPR)  (strongest)
  const kQD=new Map();   // qty(4dp)|date
  const kAD=new Map();   // round(amount)|date
  const add=(m,k,e)=>{ if(!m.has(k))m.set(k,new Set()); m.get(k).add(e); };
  for(let r=2;r<=pws.rowCount;r++){
    const e=ct(pws.getRow(r).getCell(7).value); if(!V.test(e))continue;
    const q=num(pws.getRow(r).getCell(3).value), a=num(pws.getRow(r).getCell(4).value), d=dkey(pws.getRow(r).getCell(5).value);
    add(kQAD, q.toFixed(4)+"|"+a.toFixed(2)+"|"+d, e);
    add(kQD,  q.toFixed(4)+"|"+d, e);
    add(kAD,  Math.round(a)+"|"+d, e);
  }
  // unique portal EPR -> for "is this EPR even on portal"
  const allEpr=new Set(); for(const m of [kQAD]) for(const s of m.values()) for(const e of s) allEpr.add(e);

  // best unique EPR for a sheet row (try strongest key, fall back); null if not unique
  function lookup(q,a,d){
    for(const m of [kQAD, kQD, kAD]){
      const k = m===kQAD ? q.toFixed(4)+"|"+a.toFixed(2)+"|"+d
              : m===kQD  ? q.toFixed(4)+"|"+d
              :            Math.round(a)+"|"+d;
      const s=m.get(k);
      if(s && s.size===1) return [...s][0];
    }
    return null;
  }

  for(const job of JOBS){
    if(!fs.existsSync(job.file)){ console.log("SKIP "+job.file); continue; }
    const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(job.file);
    const ws=wb.getWorksheet(job.sheet)||wb.worksheets[0];
    const ec=col(ws,["EPR Invoice Number"]), ic=col(ws,["E-Invoice Number*","E-Invoice Number"]);
    const amtC=col(ws,["Principal Amount(₹)*"]), dC=col(ws,["Sales date*"]), qtyC=col(ws,["Quantity Sold(MT)"]);

    let match=0, mismatchFixed=0, filledBlank=0, ambiguous=0, notOnPortal=0;
    const changes=[];
    for(let r=2;r<=ws.rowCount;r++){
      const inv=ic?ct(ws.getRow(r).getCell(ic).value):""; if(!inv) continue;
      const cur=ct(ws.getRow(r).getCell(ec).value);
      const q=num(ws.getRow(r).getCell(qtyC).value), a=num(ws.getRow(r).getCell(amtC).value), d=dkey(ws.getRow(r).getCell(dC).value);
      const real=lookup(q,a,d);
      if(real===null){
        if(!V.test(cur)) notOnPortal++; else ambiguous++; // can't verify uniquely
        continue;
      }
      if(real===cur){ match++; continue; }
      // portal has a DIFFERENT unique EPR for this exact sale -> our row was wrong
      if(V.test(cur)) { changes.push(`row ${r-1} ${inv}: ${cur} -> ${real}  (WRONG EPR corrected)`); mismatchFixed++; }
      else            { changes.push(`row ${r-1} ${inv}: BLANK -> ${real}`); filledBlank++; }
      if(APPLY){ ws.getRow(r).getCell(ec).value=real; ws.getRow(r).commit(); }
    }
    if(APPLY && (mismatchFixed+filledBlank)){ const tmp=job.file+".tmp"; await wb.xlsx.writeFile(tmp); fs.renameSync(tmp,job.file); }

    // remaining duplicate EPRs in sheet
    const byEpr=new Map();
    for(let r=2;r<=ws.rowCount;r++){const e=ct(ws.getRow(r).getCell(ec).value);if(V.test(e)){if(!byEpr.has(e))byEpr.set(e,[]);byEpr.get(e).push(r-1);}}
    const dups=[...byEpr.entries()].filter(([,a])=>a.length>1);

    console.log("================ "+job.file+" ================");
    if(changes.length){ console.log("Corrections:"); changes.forEach(c=>console.log("  "+c)); console.log(""); }
    console.log("  matched-correct="+match+"  WRONG-EPR-fixed="+mismatchFixed+"  blank-filled="+filledBlank+"  ambiguous="+ambiguous+"  notOnPortal(not submitted)="+notOnPortal + (APPLY?"  [WRITTEN]":"  [DRY RUN]"));
    console.log("  duplicate EPRs remaining in sheet: "+dups.length+(dups.length?"  -> "+dups.map(([e,rs])=>e+"@"+rs.join("/")).join("  "):""));
    console.log("");
  }
})().catch(e=>{console.error("FATAL:",e?.stack||e);process.exit(1);});
