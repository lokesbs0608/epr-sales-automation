"use strict";

// ---------------------------------------------------------------------------
// sim_worker.js  — SIMULATION worker (no browser / no portal)
//
// It mirrors exactly what the real workers (run_registered.js etc.) do per row
// EXCEPT the Playwright form-filling: it sanitizes the row, then "submits" by
// generating a fake EPR Invoice Number. This lets us exercise the full
// orchestrate.js pipeline (chunking, parallel pools, per-row atomic save,
// logging) against real data without logging into CPCB.
//
// Usage (same contract as a real worker):
//   node sim_worker.js --config chunks/PP_chunk1.config.json
// ---------------------------------------------------------------------------

const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const SAN = require("./sanitize");

function argVal(flag, def = null) {
    const i = process.argv.indexOf(flag);
    return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : def;
}

const CONFIG_PATH = path.resolve(__dirname, argVal("--config", "config.json"));
const CONFIG = JSON.parse(fs.readFileSync(CONFIG_PATH, "utf8"));

const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_log.csv`);
const STATE_DISTRICT_CSV = path.resolve(__dirname, "state_district_mapping.xlsx - State-District Mapping.csv");

function normHeader(s) { return String(s || "").trim().replace(/\s+/g, " ").toLowerCase(); }
function cellText(v) { return SAN.cellText(v); }
function getHeaderMap(ws) {
    const m = new Map();
    ws.getRow(1).eachCell((c, n) => { const k = normHeader(cellText(c.value)); if (k) m.set(k, n); });
    return m;
}
function getVal(row, hm, name) { const c = hm.get(normHeader(name)); return c ? row.getCell(c).value : ""; }
function setVal(row, hm, name, v) { const c = hm.get(normHeader(name)); if (c) row.getCell(c).value = v; }
function isCellEmpty(v) { return cellText(v) === ""; }
function isRowEmpty(row, hm) { for (const c of hm.values()) { if (cellText(row.getCell(c).value) !== "") return false; } return true; }

function csvEscape(v) {
    const s = String(v ?? "");
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
function log(msg) { console.log(`${new Date().toISOString()} ${msg}`); }

function ensureLogHeader() {
    if (fs.existsSync(LOG_PATH)) return;
    const header = ["datetime", "row", "e_invoice_number", "sales_date", "quantity", "state", "district", "epr_invoice_number", "status", "message"].join(",");
    fs.writeFileSync(LOG_PATH, header + "\n");
}
function appendLog(row, hm, { status, epr, message }) {
    ensureLogHeader();
    const data = [
        new Date().toISOString(),
        row.number,
        cellText(getVal(row, hm, "E-Invoice Number*")),
        cellText(getVal(row, hm, "Sales date*")),
        cellText(getVal(row, hm, "Quantity Sold(MT)")) || cellText(getVal(row, hm, "Sum of QTY MT")),
        cellText(getVal(row, hm, "State*")) || cellText(getVal(row, hm, "STATE")),
        cellText(getVal(row, hm, "District*")),
        epr || "",
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(LOG_PATH, data.join(",") + "\n");
}

async function safeWrite(wb) {
    // Atomic: write temp, back up current, rename over. Same as real workers —
    // this is what makes per-chunk writes crash-safe even when many run at once.
    await wb.xlsx.writeFile(EXCEL_TMP);
    try { if (fs.existsSync(OUTPUT_PATH) && fs.statSync(OUTPUT_PATH).size > 0) fs.copyFileSync(OUTPUT_PATH, EXCEL_BAK); } catch {}
    fs.renameSync(EXCEL_TMP, OUTPUT_PATH);
}

// Build a sanitize spec generically from whichever columns this sheet has.
function buildSpec(hm, map) {
    const has = (n) => hm.has(normHeader(n));
    const mandatory = [];
    for (const n of ["Entity Type*", "Name of the Entity *", "GST No. of Seller *", "State*", "District*",
                      "Sales date*", "Quantity Sold(MT)", "ENTITY", "GST NO", "STATE", "CATEGORY", "Sum of QTY MT"]) {
        if (has(n)) mandatory.push(n);
    }
    const dateFields = ["Sales date*", "Sales date", "INVOICE DATE"].filter(has);
    const numberFields = [];
    if (has("Principal Amount(₹)*")) numberFields.push({ name: "Principal Amount(₹)*", decimals: 2 });
    if (has("GST & Other Charges(₹)*")) numberFields.push({ name: "GST & Other Charges(₹)*", decimals: 2 });
    if (has("Sum of GST PAID")) numberFields.push({ name: "Sum of GST PAID", decimals: 2 });
    // NOTE: quantity columns intentionally NOT in numberFields (filled exactly).
    const stateField = has("State*") ? "State*" : (has("STATE") ? "STATE" : null);
    const districtField = has("District*") ? "District*" : null;
    return { mandatory, dateFields, numberFields, stateField, districtField, map };
}

function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

(async () => {
    const tag = OUTPUT_BASENAME;
    log(`SIM start: ${path.basename(CONFIG.inputExcel)} sheet="${SHEET}"`);

    let map = null;
    try { map = SAN.loadStateDistrictMap(STATE_DISTRICT_CSV); } catch (e) { log(`map not loaded: ${e.message}`); }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(EXCEL_PATH);
    const ws = wb.getWorksheet(SHEET) || wb.worksheets[0];
    const hm = getHeaderMap(ws);
    const spec = buildSpec(hm, map);

    let filled = 0, skipped = 0, sanitized = 0;
    const lastRow = ws.rowCount;
    let seq = 0;

    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        if (isRowEmpty(row, hm)) continue;

        // resume: skip if already done
        if (!isCellEmpty(getVal(row, hm, "EPR Invoice Number"))) { continue; }

        const san = SAN.sanitizeRow(row, hm, { getVal, setVal }, spec);
        if (san.notes.length) { sanitized++; log(`Row ${r}: sanitized -> ${san.notes.join("; ")}`); }
        if (san.skip) {
            skipped++;
            const reason = `Skipped: ${san.reason}`;
            log(`Row ${r}: ${reason}`);
            setVal(row, hm, "Status", reason);
            row.commit();
            await safeWrite(wb);
            appendLog(row, hm, { status: "Skipped", epr: "", message: san.reason });
            continue;
        }

        // ---- SIMULATE submit ----
        await sleep(2 + Math.floor((r * 7) % 11)); // tiny deterministic-ish delay
        seq++;
        const epr = `SIM-${tag}-${String(seq).padStart(5, "0")}`;

        setVal(row, hm, "Status", "Filled (SIM)");
        setVal(row, hm, "EPR Invoice Number", epr);
        row.commit();
        await safeWrite(wb); // per-row atomic save
        appendLog(row, hm, { status: "Filled (SIM)", epr, message: "simulated" });
        filled++;
        if (filled % 100 === 0) log(`progress: filled=${filled} skipped=${skipped}`);
    }

    log(`SIM done: filled=${filled} skipped=${skipped} sanitizedRows=${sanitized} -> ${path.basename(OUTPUT_PATH)}`);
    process.exit(0);
})().catch((e) => { console.error(`SIM FATAL: ${e?.stack || e}`); process.exit(1); });
