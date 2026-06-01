// Lists rows with garbage EPR or duplicate EPR (the read-back bug), by invoice.
const ExcelJS = require("exceljs");
const SAN = require("./sanitize");
const fs = require("fs");
function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
const VALID=/^\d{13,16}$/;
async function readWB(f){for(const c of [f,f+".bak"]){try{const w=new ExcelJS.Workbook();await w.xlsx.readFile(c);return w;}catch(e){}}return null;}
(async()=>{
  const garbage=[], byEpr=new Map();
  for(const f of fs.readdirSync("chunks").filter(x=>x.endsWith("_output.xlsx")&&!x.includes(".bak")&&!x.includes(".tmp")).sort()){
    const wb=await readWB("chunks/"+f); if(!wb)continue;
    const ws=wb.worksheets[0];
    let ec=0,ic=0;ws.getRow(1).eachCell((c,n)=>{const l=nh(c.value);if(l.includes("epr invoice")&&!ec)ec=n;if(l.includes("e-invoice")&&!ic)ic=n;});
    for(let r=2;r<=ws.rowCount;r++){
      const e=SAN.cellText(ws.getRow(r).getCell(ec).value);
      const inv=SAN.cellText(ws.getRow(r).getCell(ic).value);
      if(!e)continue;
      const loc=f.replace("config_","").replace("_output.xlsx","")+" row"+r;
      if(!VALID.test(e)) garbage.push({loc,inv,epr:JSON.stringify(e)});
      else { if(!byEpr.has(e))byEpr.set(e,[]); byEpr.get(e).push({loc,inv}); }
    }
  }
  const out=[];
  out.push("======= GARBAGE EPR ROWS ("+garbage.length+") =======");
  garbage.forEach(g=>out.push("  "+g.loc+"  invoice="+g.inv+"  epr="+g.epr));
  out.push("");
  const dups=[...byEpr.entries()].filter(([e,a])=>a.length>1);
  out.push("======= DUPLICATE EPR NUMBERS ("+dups.length+") =======");
  for(const [e,locs] of dups){ out.push("  EPR "+e+" used by "+locs.length+" invoices:"); locs.forEach(l=>out.push("      "+l.loc+"  invoice="+l.inv)); }
  fs.writeFileSync("cleanup_list.txt", out.join("\n")+"\n");
  console.log(out.join("\n"));
  console.log("\nSaved -> cleanup_list.txt");
})();
