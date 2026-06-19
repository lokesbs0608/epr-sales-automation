"use strict";
// repair_epr.js
// For rows whose EPR is suspect (shared with another invoice from the early
// cross-window race), RE-FETCH the correct EPR from the portal by searching the
// row's E-Invoice Number (which is always correct). READ-ONLY on the portal -
// no submitting. Writes the correct EPR back into the report sheet.
//
// Usage:
//   node repair_epr.js --file PP_FINAL_REPORT.xlsx  --sheet "PP BAAZPUR"
//   node repair_epr.js --file PET_FINAL_REPORT.xlsx --sheet "PET BAAZPUR"
//   (add --storage storageState_registered.json ; --dry to preview without writing)

const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");
const SAN = require("./sanitize");

function arg(flag, def=null){const i=process.argv.indexOf(flag);return i!==-1&&process.argv[i+1]?process.argv[i+1]:def;}
const FILE = arg("--file");
const SHEET = arg("--sheet");
const STORAGE = path.resolve(__dirname, arg("--storage","storageState_registered.json"));
const DRY = process.argv.includes("--dry");
const URL = "https://eprplastic.cpcb.gov.in/#/epr/details/sales";
const VALID_EPR = /^\d{10,}$/;

if(!FILE){ console.error("Pass --file <report.xlsx> [--sheet NAME]"); process.exit(1); }

function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
function ct(v){return SAN.cellText(v);}
function col(ws,names){let c=null;ws.getRow(1).eachCell((cell,n)=>{if(c===null&&names.some(x=>nh(x)===nh(ct(cell.value))))c=n;});return c;}

async function waitLoader(page){
  for(const sel of [".spinner-border",".ngx-spinner",".loading",".loader",".overlay"]){
    try{const l=page.locator(sel);if(await l.count())await l.first().waitFor({state:"hidden",timeout:15000}).catch(()=>{});}catch{}
  }
}

// Search by E-Invoice Number and return the EPR number shown for that invoice.
async function fetchEprByInvoice(page, eInvoice){
  await page.locator(".modal-dialog").first().waitFor({state:"hidden",timeout:1500}).catch(()=>{});
  await waitLoader(page);
  const search=page.locator('input[name="searchField"]').first();
  await search.waitFor({state:"visible",timeout:30000});
  await search.click(); await search.fill(""); await search.fill(eInvoice);
  await page.locator("button",{hasText:"Search"}).first().click();
  await page.waitForTimeout(1200); await waitLoader(page);

  // the matching row contains the E-Invoice text; read its EPR (15-digit 2026...)
  const row=page.locator("#ScrollableSimpleTableBody tr",{hasText:eInvoice}).first();
  try{ await row.waitFor({state:"visible",timeout:15000}); }
  catch{ return { epr:"", note:"no row found for invoice" }; }
  const text=(await row.innerText().catch(()=>""))||"";
  const m=text.match(/\b(20\d{12,14})\b/); // EPR is a 2026... number
  if(m) return { epr:m[1], note:"" };
  return { epr:"", note:"row found but no EPR number in it" };
}

(async()=>{
  const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(FILE);
  const ws=SHEET?wb.getWorksheet(SHEET):wb.worksheets[0];
  const eprCol=col(ws,["EPR Invoice Number"]);
  const invCol=col(ws,["E-Invoice Number*","E-Invoice Number"]);
  if(!eprCol||!invCol){ console.error("EPR or E-Invoice column not found"); process.exit(1); }

  // find suspect rows: EPR shared by >1 invoice
  const byEpr=new Map();
  for(let r=2;r<=ws.rowCount;r++){const e=ct(ws.getRow(r).getCell(eprCol).value);if(VALID_EPR.test(e)){if(!byEpr.has(e))byEpr.set(e,[]);byEpr.get(e).push(r);}}
  const suspectRows=[];
  for(const [e,rows] of byEpr){ if(rows.length>1) suspectRows.push(...rows); }
  console.log(`Suspect rows (shared EPR): ${suspectRows.length}`);
  if(!suspectRows.length){ console.log("Nothing to repair."); return; }

  const browser=await chromium.launch({headless:false});
  const ctx=await browser.newContext(fs.existsSync(STORAGE)?{storageState:STORAGE}:{});
  const page=await ctx.newPage();
  await page.goto(URL,{waitUntil:"domcontentloaded",timeout:90000});
  // login if needed
  const isLogin=async()=>(await page.locator('input[type="password"]').first().count())>0;
  if(!fs.existsSync(STORAGE)||await isLogin()){
    console.log("Log in manually, wait for the sales table, then press ENTER here...");
    await new Promise(res=>process.stdin.once("data",()=>res()));
    await ctx.storageState({path:STORAGE});
  }
  await page.waitForSelector("#ScrollableSimpleTableBody",{timeout:60000});

  let fixed=0, same=0, failed=0;
  const report=[];
  for(const r of suspectRows){
    const inv=ct(ws.getRow(r).getCell(invCol).value);
    const cur=ct(ws.getRow(r).getCell(eprCol).value);
    let res;
    try{ res=await fetchEprByInvoice(page, inv); }
    catch(e){ res={epr:"",note:String(e?.message||e).split("\n")[0]}; }
    if(!res.epr){ failed++; report.push(`row ${r} ${inv}: COULD NOT FETCH (${res.note}) [kept ${cur}]`); continue; }
    if(res.epr===cur){ same++; report.push(`row ${r} ${inv}: already correct (${cur})`); continue; }
    report.push(`row ${r} ${inv}: ${cur} -> ${res.epr}  *FIXED*`);
    if(!DRY){ ws.getRow(r).getCell(eprCol).value=res.epr; ws.getRow(r).commit(); }
    fixed++;
  }

  if(!DRY && fixed){
    const tmp=FILE+".tmp"; await wb.xlsx.writeFile(tmp); fs.renameSync(tmp, FILE);
  }
  console.log("\n==== REPAIR REPORT ====");
  console.log(report.join("\n"));
  console.log(`\nfixed=${fixed} alreadyCorrect=${same} failed=${failed}` + (DRY?"  (DRY RUN - nothing written)":""));
  await browser.close();
})().catch(e=>{console.error("FATAL:",e?.stack||e);process.exit(1);});
