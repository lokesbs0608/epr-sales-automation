const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const INPUT = path.resolve(__dirname, "Processed_Domestic_Granules PWP Sale For Automation.xlsx");
const JOBS = [
    {
        log: "Processed_Domestic_Granules PWP Sale For Automation_output1_log.csv",
        sheet: "Unregistered",
    },
    {
        log: "Processed_Domestic_Granules PWP Sale For Automation_registered_output_log.csv",
        sheet: "Registered",
    },
];

function parseCSV(text) {
    const rows = [];
    let row = [];
    let field = "";
    let inQuotes = false;
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (inQuotes) {
            if (ch === '"') {
                if (text[i + 1] === '"') { field += '"'; i++; }
                else inQuotes = false;
            } else field += ch;
        } else {
            if (ch === '"') inQuotes = true;
            else if (ch === ",") { row.push(field); field = ""; }
            else if (ch === "\r") { /* skip */ }
            else if (ch === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
            else field += ch;
        }
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows;
}

function normHeader(s) { return String(s || "").trim().replace(/\s+/g, " ").toLowerCase(); }
function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
}

function buildLogMap(csvPath) {
    const text = fs.readFileSync(csvPath, "utf8");
    const rows = parseCSV(text);
    if (!rows.length) return { byInv: new Map(), total: 0, filled: 0, failed: 0 };
    const header = rows[0].map(h => h.trim());
    const idx = {
        datetime: header.indexOf("datetime"),
        row: header.indexOf("row"),
        eInv: header.indexOf("e_invoice_number"),
        epr: header.indexOf("epr_invoice_number"),
        status: header.indexOf("status"),
        message: header.indexOf("message"),
    };
    const byInv = new Map();
    let filled = 0, failed = 0;
    for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r.length || r.every(c => !c)) continue;
        const inv = (r[idx.eInv] || "").trim();
        if (!inv) continue;
        const entry = {
            datetime: r[idx.datetime] || "",
            row: r[idx.row] || "",
            epr: (r[idx.epr] || "").trim(),
            status: (r[idx.status] || "").trim(),
            message: r[idx.message] || "",
        };
        const prev = byInv.get(inv);
        // Prefer a Filled entry over Failed; otherwise latest datetime
        if (!prev) byInv.set(inv, entry);
        else {
            const prevFilled = /filled/i.test(prev.status);
            const curFilled = /filled/i.test(entry.status);
            if (curFilled && !prevFilled) byInv.set(inv, entry);
            else if (curFilled === prevFilled && entry.datetime >= prev.datetime) byInv.set(inv, entry);
        }
    }
    for (const e of byInv.values()) {
        if (/filled/i.test(e.status)) filled++; else if (/fail/i.test(e.status)) failed++;
    }
    return { byInv, total: byInv.size, filled, failed };
}

function getHeaderMap(ws) {
    const map = new Map();
    ws.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
        const key = normHeader(cellText(cell.value));
        if (key) map.set(key, col);
    });
    return map;
}

function ensureColumn(ws, headerMap, headerName) {
    const key = normHeader(headerName);
    if (headerMap.has(key)) return headerMap.get(key);
    // Pick first empty header cell within the used range, else append
    const last = Math.max(ws.columnCount, 17);
    let target = 0;
    for (let c = 1; c <= last + 2; c++) {
        const v = cellText(ws.getRow(1).getCell(c).value);
        if (!v) { target = c; break; }
    }
    if (!target) target = last + 1;
    ws.getRow(1).getCell(target).value = headerName;
    headerMap.set(key, target);
    return target;
}

async function safeWrite(wb, targetPath) {
    const tmp = `${targetPath}.tmp`;
    const bak = `${targetPath}.bak`;
    await wb.xlsx.writeFile(tmp);
    if (fs.existsSync(targetPath)) fs.copyFileSync(targetPath, bak);
    fs.renameSync(tmp, targetPath);
}

(async () => {
    if (!fs.existsSync(INPUT)) throw new Error("Input not found: " + INPUT);

    // Back up once before any write
    const backup = INPUT.replace(/\.xlsx$/, "_BEFORE_MERGE.xlsx");
    if (!fs.existsSync(backup)) fs.copyFileSync(INPUT, backup);
    console.log("Backup:", backup);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(INPUT);

    for (const job of JOBS) {
        const logPath = path.resolve(__dirname, job.log);
        if (!fs.existsSync(logPath)) {
            console.log(`SKIP (no log): ${job.log}`);
            continue;
        }
        const ws = wb.getWorksheet(job.sheet);
        if (!ws) {
            console.log(`SKIP (no sheet "${job.sheet}")`);
            continue;
        }

        const { byInv, total, filled, failed } = buildLogMap(logPath);
        console.log(`\n[${job.sheet}] log unique invoices=${total} filled=${filled} failed=${failed}`);

        const headerMap = getHeaderMap(ws);
        const eprCol = ensureColumn(ws, headerMap, "EPR Invoice Number");
        const statusCol = ensureColumn(ws, headerMap, "Status");
        const invCol = headerMap.get(normHeader("E-Invoice Number*")) ||
                       headerMap.get(normHeader("E-Invoice Number"));
        if (!invCol) {
            console.log(`  !! no invoice column in sheet "${job.sheet}", skipping`);
            continue;
        }

        let updatedFilled = 0, updatedFailed = 0, missingInLog = 0;
        for (let r = 2; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            const inv = cellText(row.getCell(invCol).value);
            if (!inv) continue;
            const entry = byInv.get(inv);
            if (!entry) { missingInLog++; continue; }
            if (entry.epr) {
                row.getCell(eprCol).value = entry.epr;
                row.getCell(statusCol).value = entry.status || "Filled";
                updatedFilled++;
            } else {
                // Preserve prior EPR if the cell already had one
                const existing = cellText(row.getCell(eprCol).value);
                if (!existing) {
                    row.getCell(statusCol).value = entry.status ? `Failed: ${entry.message || entry.status}` : "Failed";
                    updatedFailed++;
                }
            }
            row.commit();
        }
        console.log(`  updated: filled=${updatedFilled} failed=${updatedFailed} no-log-match=${missingInLog}`);
    }

    await safeWrite(wb, INPUT);
    console.log("\nWrote:", INPUT);
})().catch(e => { console.error(e.stack || e.message); process.exit(1); });
