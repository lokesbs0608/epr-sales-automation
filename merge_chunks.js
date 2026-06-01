"use strict";
// merge_chunks.js
// Merge all chunk outputs for each job into ONE final Excel per job (PP, PET),
// in chunk order, removing duplicate rows. Dedup key = E-Invoice Number when
// present, else full-row hash. Reports valid/duplicate/garbage EPR counts.
//
// Usage: node merge_chunks.js

const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const SAN = require("./sanitize");

function nh(s) { return String(s || "").trim().replace(/\s+/g, " ").toLowerCase(); }
function ct(v) { return SAN.cellText(v); }
function col(ws, names) {
    let c = null;
    ws.getRow(1).eachCell((cell, n) => { if (c === null && names.some(x => nh(x) === nh(ct(cell.value)))) c = n; });
    return c;
}
const INVOICE_NAMES = ["E-Invoice Number*", "E-Invoice Number", "INVOICE No", "Invoice Number"];
const EPR_NAMES = ["EPR Invoice Number"];
const VALID_EPR = /^\d{13,16}$/;

const JOBS = [
    { name: "PP", sheet: "PP BAAZPUR", prefix: "config_pp_chunk", out: "FINAL_PP.xlsx" },
    { name: "PET", sheet: "PET BAAZPUR", prefix: "config_pet_chunk", out: "FINAL_PET.xlsx" },
];

function chunkFilesFor(prefix) {
    return fs.readdirSync("chunks")
        .filter(f => f.startsWith(prefix) && f.endsWith("_output.xlsx") && !f.includes(".bak") && !f.includes(".tmp"))
        .map(f => {
            const m = f.match(/chunk(\d+)_output\.xlsx$/);
            return { f: path.join("chunks", f), idx: m ? Number(m[1]) : 9999 };
        })
        .sort((a, b) => a.idx - b.idx);
}

function rowKey(row, lastCol, invCol) {
    if (invCol) { const k = ct(row.getCell(invCol).value); if (k) return "I:" + k.toUpperCase(); }
    const parts = [];
    for (let c = 1; c <= lastCol; c++) parts.push(ct(row.getCell(c).value));
    return "H:" + parts.join("|");
}

(async () => {
    for (const job of JOBS) {
        const files = chunkFilesFor(job.prefix);
        if (!files.length) { console.log(`${job.name}: no chunk outputs found, skipping`); continue; }

        const outWb = new ExcelJS.Workbook();
        const outWs = outWb.addWorksheet(job.sheet);
        let headerWritten = false, dst = 2;
        const seen = new Set();
        let kept = 0, dupes = 0, validEpr = 0, garbageEpr = 0, blankEpr = 0;
        const eprCount = new Map();

        for (const { f, idx } of files) {
            // Read the chunk output; if it's mid-write/corrupt, fall back to its
            // .bak (the previous atomic snapshot). This makes merge safe to run
            // even while the orchestrator is still writing chunks.
            let wb = null;
            for (const candidate of [f, f + ".bak"]) {
                if (!fs.existsSync(candidate)) continue;
                try { const w = new ExcelJS.Workbook(); await w.xlsx.readFile(candidate); wb = w; if (candidate !== f) console.log(`  ${job.name} chunk${idx}: using .bak (main was busy/corrupt)`); break; }
                catch (e) { /* try next candidate */ }
            }
            if (!wb) { console.log(`  ${job.name} chunk${idx}: READ FAILED (main + .bak) — skipped`); continue; }
            const ws = wb.worksheets[0];
            if (!ws) continue;
            if (!headerWritten) { outWs.getRow(1).values = ws.getRow(1).values; headerWritten = true; }
            const lastCol = ws.columnCount;
            const invCol = col(ws, INVOICE_NAMES);
            const eprCol = col(ws, EPR_NAMES);
            for (let r = 2; r <= ws.rowCount; r++) {
                const row = ws.getRow(r);
                let any = false;
                for (let c = 1; c <= lastCol; c++) { if (ct(row.getCell(c).value) !== "") { any = true; break; } }
                if (!any) continue;
                const key = rowKey(row, lastCol, invCol);
                if (seen.has(key)) { dupes++; continue; }
                seen.add(key);
                outWs.getRow(dst++).values = row.values;
                kept++;
                if (eprCol) {
                    const e = ct(row.getCell(eprCol).value);
                    if (!e) blankEpr++;
                    else if (VALID_EPR.test(e)) { validEpr++; eprCount.set(e, (eprCount.get(e) || 0) + 1); }
                    else garbageEpr++;
                }
            }
        }
        let eprDup = 0; for (const c of eprCount.values()) if (c > 1) eprDup++;

        const tmp = job.out + ".tmp";
        await outWb.xlsx.writeFile(tmp);
        fs.renameSync(tmp, job.out);

        console.log(`\n=== ${job.name} -> ${job.out} ===`);
        console.log(`  chunks merged   : ${files.length}`);
        console.log(`  unique rows     : ${kept}   (duplicate rows removed: ${dupes})`);
        console.log(`  valid EPR       : ${validEpr}`);
        console.log(`  blank EPR (no submit): ${blankEpr}`);
        console.log(`  garbage EPR     : ${garbageEpr}`);
        console.log(`  duplicate EPR #s: ${eprDup}`);
    }
    console.log("\nDone.");
})().catch(e => { console.error("FATAL:", e?.stack || e); process.exit(1); });
