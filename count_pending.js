const ExcelJS = require("exceljs");
const SAN = require("./sanitize");
const fs = require("fs");
function nh(s){return String(s||"").trim().replace(/\s+/g," ").toLowerCase();}
const V=/^\d{10,}$/;
(async()=>{
  if(!fs.existsSync("chunks")){console.log("no chunks dir");return;}
  const files=fs.readdirSync("chunks").filter(x=>/config_pp_pending_chunk\d+_output\.xlsx$/.test(x)&&!x.includes(".bak")).sort();
  let tot=0;
  console.log("PER CHUNK (this pending run):");
  for(const f of files){
    let wb;try{wb=new ExcelJS.Workbook();await wb.xlsx.readFile("chunks/"+f);}catch(e){console.log("  "+f+": busy");continue;}
    const ws=wb.worksheets[0];let ec=0;ws.getRow(1).eachCell((c,n)=>{if(nh(c.value).includes("epr invoice")&&!ec)ec=n;});
    let g=0,fail=0;if(ec)for(let r=2;r<=ws.rowCount;r++){const e=SAN.cellText(ws.getRow(r).getCell(ec).value);if(V.test(e))g++;}
    tot+=g;console.log("  "+f.replace("config_pp_pending_","").replace("_output.xlsx","")+": "+g+" filled");
  }
  console.log("THIS RUN filled: "+tot+" / 413");
})();
