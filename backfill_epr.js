"use strict";
// backfill_epr.js
// Writes already-generated EPR numbers back into the ORIGINAL source sheets,
// matched by E-Invoice Number, and ensures EPR/Status columns exist. This makes
// the originals resume-safe regardless of chunk size: rows that were already
// submitted carry their EPR and get skipped on the next run (no duplicates).
//
// Sources of truth (in priority): chunks/*_output.xlsx + *_combined_output.xlsx
// Usage: node backfill_epr.js

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
const STATUS_NAMES = ["Status"];

// Targets: original source files + the sheet name.
const TARGETS = [
    { file: "POLYPLEX BAAZPUR Automation READY (PP).xlsx", sheet: "PP BAAZPUR" },
    { file: "POLYPLEX BAAZPUR AUTOMATION READY (PET).xlsx", sheet: "PET BAAZPUR" },
];

// Collect EPR by invoice number from every result file we can find.
async function collectEprMap() {
    const map = new Map(); // invoiceNo(upper) -> { epr, status }
    const files = [];
    if (fs.existsSync("chunks")) {
        for (const f of fs.readdirSync("chunks").filter(x => x.endsWith("_output.xlsx"))) files.push(path.join("chunks", f));
    }
    for (const f of fs.readdirSync(".").filter(x => /_combined_output\.xlsx$/.test(x))) files.push(f);

    for (const f of files) {
        try {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.readFile(f);
            for (const ws of wb.worksheets) {
                const invCol = col(ws, INVOICE_NAMES);
                const eprCol = col(ws, EPR_NAMES);
                const stCol = col(ws, STATUS_NAMES);
                if (!invCol || !eprCol) continue;
                for (let r = 2; r <= ws.rowCount; r++) {
                    const inv = ct(ws.getRow(r).getCell(invCol).value);
                    const epr = ct(ws.getRow(r).getCell(eprCol).value);
                    if (inv && epr) {
                        map.set(inv.toUpperCase(), { epr, status: stCol ? ct(ws.getRow(r).getCell(stCol).value) : "Filled" });
                    }
                }
            }
        } catch (e) { console.log(`  (skip ${f}: ${e.message})`); }
    }
    return map;
}

(async () => {
    const eprMap = await collectEprMap();
    console.log(`Collected ${eprMap.size} EPR numbers from result files.`);

    for (const t of TARGETS) {
        if (!fs.existsSync(t.file)) { console.log(`SKIP (not found): ${t.file}`); continue; }
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(t.file);
        const ws = wb.getWorksheet(t.sheet) || wb.worksheets[0];

        // ensure columns
        SAN.ensureWritebackColumns(ws, nh, ["EPR Invoice Number", "Status"]);
        const invCol = col(ws, INVOICE_NAMES);
        const eprCol = col(ws, EPR_NAMES);
        const stCol = col(ws, STATUS_NAMES);

        let updated = 0, already = 0;
        for (let r = 2; r <= ws.rowCount; r++) {
            const inv = ct(ws.getRow(r).getCell(invCol).value);
            if (!inv) continue;
            const existing = ct(ws.getRow(r).getCell(eprCol).value);
            if (existing) { already++; continue; }
            const hit = eprMap.get(inv.toUpperCase());
            if (hit) {
                ws.getRow(r).getCell(eprCol).value = hit.epr;
                ws.getRow(r).getCell(stCol).value = hit.status || "Filled";
                ws.getRow(r).commit();
                updated++;
            }
        }
        // atomic-ish write
        const tmp = t.file + ".tmp";
        await wb.xlsx.writeFile(tmp);
        fs.renameSync(tmp, t.file);
        console.log(`${t.file}: backfilled ${updated} EPR(s), ${already} already had one.`);
    }
    console.log("Done.");
})().catch(e => { console.error("FATAL:", e?.stack || e); process.exit(1); });
