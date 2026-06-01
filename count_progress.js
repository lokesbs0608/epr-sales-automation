const ExcelJS = require("exceljs");
const SAN = require("./sanitize");
const fs = require("fs");
function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
const VALID=/^\d{13,16}$/;
async function readWB(f){
  for(const c of [f,f+".bak"]){try{const w=new ExcelJS.Workbook();await w.xlsx.readFile(c);return w;}catch(e){}}
  return null;
}
(async()=>{
  const files=fs.readdirSync("chunks").filter(x=>x.endsWith("_output.xlsx")&&!x.includes(".bak")&&!x.includes(".tmp")).sort();
  let ppGood=0,ppBad=0,petGood=0,petBad=0;
  const ppEpr=new Map(),petEpr=new Map();
  console.log("PER CHUNK:");
  for(const f of files){
    const isPP=f.includes("_pp_");
    const wb=await readWB("chunks/"+f);
    if(!wb){console.log("  "+f+": UNREADABLE");continue;}
    const ws=wb.worksheets[0];let ec=0;ws.getRow(1).eachCell((c,n)=>{if(nh(c.value).includes("epr invoice")&&!ec)ec=n;});
    let g=0,b=0;
    for(let r=2;r<=ws.rowCount;r++){const e=SAN.cellText(ws.getRow(r).getCell(ec).value);if(!e)continue;
      if(VALID.test(e)){g++;if(isPP){ppGood++;ppEpr.set(e,(ppEpr.get(e)||0)+1);}else{petGood++;petEpr.set(e,(petEpr.get(e)||0)+1);}}
      else{b++;if(isPP)ppBad++;else petBad++;}}
    console.log("  "+f.replace("config_","").replace("_output.xlsx","")+": valid="+g+" bad="+b);
  }
  let ppDup=0;for(const c of ppEpr.values())if(c>1)ppDup++;
  let petDup=0;for(const c of petEpr.values())if(c>1)petDup++;
  console.log("");
  console.log("PP  : valid="+ppGood+" unique="+ppEpr.size+" dupEPR="+ppDup+" garbage="+ppBad+"  / 4354 ("+Math.round(ppGood/4354*100)+"%)");
  console.log("PET : valid="+petGood+" unique="+petEpr.size+" dupEPR="+petDup+" garbage="+petBad+"  / 2165 ("+Math.round(petGood/2165*100)+"%)");
  console.log("GRAND valid="+(ppGood+petGood)+" / 6519 ("+Math.round((ppGood+petGood)/6519*100)+"%)  garbage="+(ppBad+petBad));
})();
