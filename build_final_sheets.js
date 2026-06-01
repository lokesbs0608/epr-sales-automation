"use strict";
// build_final_sheets.js
// ONE clean sheet per product from the ORIGINAL file, EPR + Status placed
// RIGHT AFTER the data columns (no empty gap), so EPR is immediately visible.
// All rows included; each row gets a Status (Filled / PENDING / exact error).
// Output: PP_FINAL_REPORT.xlsx, PET_FINAL_REPORT.xlsx

const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const SAN = require("./sanitize");

function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
function ct(v){return SAN.cellText(v);}

// Keep the worker's EXACT error message verbatim. Only strip ANSI color codes
// and the giant Playwright "Call log:" retry dump (everything after it),
// preserving the real error text exactly as the worker reported it.
function cleanReason(raw){
  let s=String(raw||"").replace(/?\[[0-9;]*m/g,""); // strip ANSI color codes
  const cut=s.search(/\n?\s*Call log:/i);                  // drop verbose retry log
  if(cut>=0) s=s.slice(0,cut);
  s=s.replace(/\s*\n\s*/g," ").replace(/\s{2,}/g," ").trim(); // collapse to one line
  return s; // e.g. "Failed: Submit disabled: some required fields still missing."
}
function findCol(ws, names){let c=null;ws.getRow(1).eachCell((cell,n)=>{if(c===null&&names.some(x=>nh(x)===nh(ct(cell.value))))c=n;});return c;}
const INVOICE_NAMES=["E-Invoice Number*","E-Invoice Number","INVOICE No","Invoice Number"];
const EPR_NAMES=["EPR Invoice Number"];
const VALID_EPR=/^\d{10,}$/;

// Every field the portal requires for a Registered sale. Used to pinpoint which
// field is missing/invalid when the portal says "some required fields missing".
const REQUIRED_FIELDS=[
  "Quantity Sold(MT)","Registration Type*","Entity Type*","Name of the Entity *",
  "Address*","State*","District*","GST No. of Seller *","Buyer GST","HSN CODE",
  "E-Invoice Number*","Bank Account No*","IFSC Code*","Principal Amount(₹)*",
  "GST & Other Charges(₹)*","Sales date*",
];

// Inspect the row's own data and list which required fields are empty / zero.
// Returns "" if nothing obviously wrong (then we keep the generic portal message).
function whichFieldsMissing(srcWs, rowNum, colIndex){
  const missing=[];
  for(const name of REQUIRED_FIELDS){
    const col=colIndex[nh(name)];
    if(!col) continue; // column not in sheet -> can't check
    const v=ct(srcWs.getRow(rowNum).getCell(col).value);
    if(v==="") missing.push(name.replace(/\*$/,"").trim());
    else if(nh(name)==="quantity sold(mt)" && Number(v)===0) missing.push("Quantity is 0");
  }
  return missing.join(", ");
}

const JOBS=[
  // prefix matches the main chunks AND any pending re-run chunks for the product
  { name:"PP",  src:"POLYPLEX BAAZPUR Automation READY (PP).xlsx",  sheet:"PP BAAZPUR",  prefix:"config_pp_",  out:"PP_FINAL_REPORT.xlsx"  },
  { name:"PET", src:"POLYPLEX BAAZPUR AUTOMATION READY (PET).xlsx", sheet:"PET BAAZPUR", prefix:"config_pet_", out:"PET_FINAL_REPORT.xlsx" },
];

// For each invoice, capture: the valid EPR (if any), and the exact failure
// reason from the chunk's Status column (so the report shows WHY it failed).
async function buildEprMap(prefix){
  const map=new Map(); // invKey -> { epr, valid, reason }
  for(const f of fs.readdirSync("chunks").filter(f=>f.startsWith(prefix)&&f.endsWith("_output.xlsx")&&!f.includes(".bak")&&!f.includes(".tmp"))){
    let wb;try{wb=new ExcelJS.Workbook();await wb.xlsx.readFile(path.join("chunks",f));}catch(e){continue;}
    const ws=wb.worksheets[0];if(!ws)continue;
    const ic=findCol(ws,INVOICE_NAMES), ec=findCol(ws,EPR_NAMES), sc=findCol(ws,["Status"]);
    if(!ic||!ec)continue;
    for(let r=2;r<=ws.rowCount;r++){
      const inv=ct(ws.getRow(r).getCell(ic).value);
      if(!inv)continue;
      const epr=ct(ws.getRow(r).getCell(ec).value);
      const status=sc?ct(ws.getRow(r).getCell(sc).value):"";
      const key=inv.toUpperCase();
      if(VALID_EPR.test(epr)){
        map.set(key,{epr,valid:true,reason:""}); // valid EPR wins, always
      } else {
        const prev=map.get(key);
        if(prev&&prev.valid)continue; // don't overwrite a good EPR
        // keep the most informative status seen (prefer a "Failed:" message)
        const reason = status || (epr ? "Bad EPR captured: "+epr : "");
        if(!prev || (reason && !/^$|attempt/i.test(reason))) map.set(key,{epr:"",valid:false,reason});
      }
    }
  }
  return map;
}

(async()=>{
  const summary=[];
  for(const job of JOBS){
    if(!fs.existsSync(job.src)){console.log(`SKIP ${job.name}: source not found`);continue;}
    const eprMap=await buildEprMap(job.prefix);

    const srcWb=new ExcelJS.Workbook();
    await srcWb.xlsx.readFile(job.src);
    const srcWs=srcWb.getWorksheet(job.sheet)||srcWb.worksheets[0];

    // Real data columns = the contiguous labeled block at the START, stopping at
    // the first empty header OR at a previously-added "EPR/Status" header.
    let lastDataCol=0;
    for(let c=1;c<=srcWs.columnCount;c++){
      const h=nh(ct(srcWs.getRow(1).getCell(c).value));
      if(h==="" || h==="epr invoice number" || h==="status") break;
      lastDataCol=c;
    }
    const invCol=findCol(srcWs,INVOICE_NAMES);
    const eprOut=lastDataCol+1, statusOut=lastDataCol+2;

    // Map required-field name -> its column number in this sheet (for diagnosing
    // exactly which field is empty when the portal says "fields missing").
    const colIndex={};
    for(let c=1;c<=lastDataCol;c++){ const h=nh(ct(srcWs.getRow(1).getCell(c).value)); if(h) colIndex[h]=c; }

    // Build a NEW workbook with only the real data columns + EPR + Status (no gaps).
    const outWb=new ExcelJS.Workbook();
    const outWs=outWb.addWorksheet(job.sheet);

    const hdr=[];
    for(let c=1;c<=lastDataCol;c++) hdr[c]=ct(srcWs.getRow(1).getCell(c).value);
    hdr[eprOut]="EPR Invoice Number";
    hdr[statusOut]="Status";
    outWs.getRow(1).values = hdr;

    let filled=0, garbage=0, pending=0, blankInv=0, total=0, dst=2;
    for(let r=2;r<=srcWs.rowCount;r++){
      let any=false;for(let c=1;c<=lastDataCol;c++){if(ct(srcWs.getRow(r).getCell(c).value)!==""){any=true;break;}}
      if(!any)continue;
      total++;
      const rowVals=[];
      for(let c=1;c<=lastDataCol;c++) rowVals[c]=srcWs.getRow(r).getCell(c).value;
      const inv=ct(srcWs.getRow(r).getCell(invCol).value);
      let epr="", status="";
      if(!inv){ blankInv++; status="NO INVOICE NO"; }
      else {
        const hit=eprMap.get(inv.toUpperCase());
        if(hit&&hit.valid){ epr=hit.epr; status="Filled"; filled++; }
        else if(hit&&hit.reason){
          status=cleanReason(hit.reason); // EXACT worker error, one line
          // If the portal said "required fields missing", pinpoint which ones
          // from this row's own data and append them to the message.
          if(/required fields|submit disabled/i.test(status)){
            const miss=whichFieldsMissing(srcWs, r, colIndex);
            status = miss
              ? `Failed: missing/empty field(s) in sheet: ${miss}`
              : "Failed: portal disabled Submit (all fields present in sheet - portal rejected, e.g. sales date out of allowed period / value validation). Retry or check on portal.";
          }
          if(/duplicate|bad epr/i.test(status)) garbage++; else pending++;
        }
        else { status="PENDING - not attempted (chunk did not run)"; pending++; }
      }
      rowVals[eprOut]=epr; rowVals[statusOut]=status;
      outWs.getRow(dst++).values = rowVals;
    }

    const tmp=job.out+".tmp";
    await outWb.xlsx.writeFile(tmp);
    fs.renameSync(tmp, job.out);
    summary.push({job:job.name,out:job.out,total,filled,garbage,pending,blankInv,eprCol:eprOut,statusCol:statusOut});
  }

  console.log("\n================ FINAL REPORT SHEETS ================");
  for(const s of summary){
    console.log(`\n${s.job}  ->  ${s.out}`);
    console.log(`  EPR Number is in COLUMN ${s.eprCol}, Status in COLUMN ${s.statusCol} (right after the data)`);
    console.log(`  total rows : ${s.total}`);
    console.log(`  Filled  : ${s.filled}`);
    console.log(`  Failed/Review : ${s.garbage}`);
    console.log(`  Pending : ${s.pending}`);
  }
  console.log("\n=====================================================");
})().catch(e=>{console.error("FATAL:",e?.stack||e);process.exit(1);});
