const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const PIBO_URL = "https://eprplastic.cpcb.gov.in/#/epr/pibo-operations/sales";

function getConfigPath() {
    const idx = process.argv.indexOf("--config");
    if (idx !== -1 && process.argv[idx + 1]) {
        return path.resolve(__dirname, process.argv[idx + 1]);
    }
    return path.resolve(__dirname, "config_pibo_upload.json");
}

const CONFIG_PATH = getConfigPath();
const ROOT_DIR = __dirname;

function loadConfig() {
    if (!fs.existsSync(CONFIG_PATH)) {
        throw new Error(`Missing config file: ${CONFIG_PATH}`);
    }
    const raw = fs.readFileSync(CONFIG_PATH, "utf8");
    const cfg = JSON.parse(raw);
    const inputExcel = String(cfg?.inputExcel || "").trim();
    const sheetName = String(cfg?.sheetName || "").trim();
    const outputExcel = String(cfg?.outputExcel || "").trim();
    const maxRowsRaw = cfg?.max_rows;
    const invoicePdfDir = String(cfg?.invoicePdfDir || "").trim();
    const storageState = String(cfg?.storageState || "storageState_pibo.json").trim();
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config must include inputExcel, sheetName, and outputExcel");
    }
    let maxRows = null;
    if (maxRowsRaw !== undefined && maxRowsRaw !== null && String(maxRowsRaw).trim() !== "") {
        const n = Number(maxRowsRaw);
        if (!Number.isFinite(n)) throw new Error("config max_rows must be a number");
        if (n > 0) maxRows = Math.floor(n);
    }
    return { inputExcel, sheetName, outputExcel, maxRows, invoicePdfDir, storageState };
}

const CONFIG = loadConfig();
const STORAGE = path.resolve(__dirname, CONFIG.storageState);
const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const UPLOAD_LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_upload_log.csv`);
const UPLOAD_FILLED_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_upload_filled.csv`);

// ── Utility ──────────────────────────────────────────────────────────────────

function normHeader(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
}

function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
}

function logStep(message, level = 0) {
    const indent = "  ".repeat(level);
    console.log(`${indent}${new Date().toISOString()} ${message}`);
}

function getHeaderMap(ws) {
    const headerRow = ws.getRow(1);
    const map = new Map();
    headerRow.eachCell((cell, colNumber) => {
        const key = normHeader(cellText(cell.value));
        if (key) map.set(key, colNumber);
    });
    return map;
}

function getHeaderList(ws) {
    const headerRow = ws.getRow(1);
    const headers = [];
    headerRow.eachCell((cell) => {
        const key = cellText(cell.value);
        if (key) headers.push(key);
    });
    return headers;
}

function getVal(row, headerMap, headerName) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return "";
    return row.getCell(col).value;
}

function setVal(row, headerMap, headerName, value) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return;
    row.getCell(col).value = value;
}

function isCellEmpty(v) { return cellText(v) === ""; }

function isRowEmpty(row, headerMap) {
    for (const col of headerMap.values()) {
        if (!isCellEmpty(row.getCell(col).value)) return false;
    }
    return true;
}

function ensureColumn(ws, headerMap, headerName) {
    const key = normHeader(headerName);
    if (headerMap.has(key)) return;
    const maxCol = Math.max(0, ...headerMap.values()) + 1;
    ws.getRow(1).getCell(maxCol).value = headerName;
    headerMap.set(key, maxCol);
}

function isSuccessText(text) {
    const t = String(text || "");
    return /success/i.test(t) && !/error/i.test(t);
}

// ── Logging ──────────────────────────────────────────────────────────────────

function ensureUploadLogHeader() {
    if (fs.existsSync(UPLOAD_LOG_PATH)) return;
    const header = ["datetime", "row", "invoice_no", "epr_invoice_number", "status", "message"].join(",");
    fs.writeFileSync(UPLOAD_LOG_PATH, `${header}\n`);
}

function appendUploadLogRow(row, headerMap, { status, message }) {
    ensureUploadLogHeader();
    const data = [
        new Date().toISOString(),
        row.number,
        cellText(getVal(row, headerMap, "INVOICE No")),
        cellText(getVal(row, headerMap, "EPR Invoice Number")),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(UPLOAD_LOG_PATH, `${data.join(",")}\n`);
}

function ensureUploadFilledHeader(headerList) {
    if (fs.existsSync(UPLOAD_FILLED_PATH)) return;
    const header = [...headerList, "datetime", "message"].map(csvEscape).join(",");
    fs.writeFileSync(UPLOAD_FILLED_PATH, `${header}\n`);
}

function appendUploadFilledRow(row, headerMap, headerList, { message }) {
    ensureUploadFilledHeader(headerList);
    const ts = new Date().toISOString();
    const rowData = headerList.map((h) => cellText(getVal(row, headerMap, h)));
    const data = [...rowData, ts, message || ""].map(csvEscape);
    fs.appendFileSync(UPLOAD_FILLED_PATH, `${data.join(",")}\n`);
}

// ── Safe Excel write ─────────────────────────────────────────────────────────

async function safeWriteWorkbook(wb) {
    await wb.xlsx.writeFile(EXCEL_TMP);
    try {
        if (fs.existsSync(OUTPUT_PATH) && fs.statSync(OUTPUT_PATH).size > 0) {
            fs.copyFileSync(OUTPUT_PATH, EXCEL_BAK);
        }
    } catch { }
    fs.renameSync(EXCEL_TMP, OUTPUT_PATH);
}

async function safeWriteWorkbookToPath(wb, targetPath) {
    const tmp = `${targetPath}.tmp`;
    const bak = `${targetPath}.bak`;
    await wb.xlsx.writeFile(tmp);
    try {
        if (fs.existsSync(targetPath) && fs.statSync(targetPath).size > 0) {
            fs.copyFileSync(targetPath, bak);
        }
    } catch { }
    fs.renameSync(tmp, targetPath);
}

async function syncInputWorkbook(wb) {
    if (path.resolve(EXCEL_PATH) === path.resolve(OUTPUT_PATH)) return;
    await safeWriteWorkbookToPath(wb, EXCEL_PATH);
}

// ── Browser helpers ──────────────────────────────────────────────────────────

async function waitForLoaderToFinish(page) {
    const loaders = [".spinner-border", ".loading", ".loader", ".ngx-spinner-overlay", ".ngx-spinner", ".overlay", ".block-ui-wrapper", ".k-i-loading"];
    for (const sel of loaders) {
        try {
            const loc = page.locator(sel);
            if ((await loc.count()) > 0) {
                await loc.first().waitFor({ state: "hidden", timeout: 30000 }).catch(() => { });
            }
        } catch { }
    }
}

async function readToastText(page) {
    const toast = page.locator(".toast, .toaster, .ngx-toastr, .toast-container").first();
    if (!(await toast.count())) return "";
    try { return (await toast.innerText()).trim(); } catch { return ""; }
}

async function readToastTextWithRetry(page, retries = 6, delayMs = 300) {
    for (let i = 0; i < retries; i++) {
        const text = await readToastText(page);
        if (text) return text;
        await page.waitForTimeout(delayMs);
    }
    return "";
}

async function isLoginPage(page) {
    const pwd = page.locator('input[type="password"]').first();
    if (await pwd.count()) return true;
    const loginBtn = page.locator('button:has-text("Login"), button:has-text("Sign In")').first();
    if (await loginBtn.count()) return true;
    return false;
}

async function attemptLogout(page) {
    try {
        const logoutDirect = page.locator('text=/logout/i').first();
        if (await logoutDirect.count()) { await logoutDirect.click().catch(() => { }); await page.waitForTimeout(1000); return true; }
        const toggles = ['button[aria-haspopup="menu"]', 'button[aria-expanded]', '.dropdown-toggle', '.nav-link.dropdown-toggle', '.user-profile', '.profile'];
        for (const sel of toggles) {
            const btn = page.locator(sel).first();
            if (await btn.count()) {
                await btn.click().catch(() => { });
                const logout = page.locator('text=/logout/i').first();
                if (await logout.count()) { await logout.click().catch(() => { }); await page.waitForTimeout(1000); return true; }
            }
        }
    } catch { }
    return false;
}

// ── Table interaction ────────────────────────────────────────────────────────

async function getTableColumnIndex(page, headerText) {
    return await page.evaluate((headerText) => {
        const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
        const ths = Array.from(document.querySelectorAll("#simple_table_header th"));
        const target = norm(headerText);
        const i = ths.findIndex((th) => norm(th.innerText) === target);
        return i >= 0 ? i + 1 : null;
    }, headerText);
}

async function searchByEprInvoice(page, eprInvoiceNumber) {
    logStep(`[search] EPR: ${eprInvoiceNumber}`, 1);
    await page.locator(".modal-dialog").first().waitFor({ state: "hidden", timeout: 2000 }).catch(() => { });
    await waitForLoaderToFinish(page);

    // The search input is right next to the Search button in the table header area
    const searchInput = page.locator('input.form-control.w-60').first();
    await searchInput.waitFor({ state: "visible", timeout: 30000 });
    await searchInput.scrollIntoViewIfNeeded();
    await searchInput.click();
    await searchInput.fill("");
    await searchInput.fill(String(eprInvoiceNumber));
    logStep(`[search] typed: ${eprInvoiceNumber}`, 1);

    // Click the Search button that's next to this input (not in filter modal)
    const searchBtn = page.locator('input.form-control.w-60 + button.btn-primary, button.btn.btn-primary:has-text("Search")').first();
    await searchBtn.click();
    logStep("[search] clicked", 1);
    await page.waitForTimeout(1500);
    await waitForLoaderToFinish(page);

    const row = page.locator("#ScrollableSimpleTableBody tr", { hasText: String(eprInvoiceNumber) }).first();
    await row.waitFor({ state: "visible", timeout: 20000 });
    logStep("[search] row found", 1);
    return row;
}

async function findCorrectRow(page, rows, expectedRecycledPct, expectedGstPaid) {
    const recycledCol = await getTableColumnIndex(page, "Recycled Plastic %");
    const gstPaidCol = await getTableColumnIndex(page, "GST Paid");
    const dateCol = await getTableColumnIndex(page, "Date");

    if (!recycledCol || !gstPaidCol) {
        logStep("Column index not found, using first row", 1);
        return rows.first();
    }

    const expRecycled = cellText(expectedRecycledPct);
    const expGst = cellText(expectedGstPaid);
    const count = await rows.count();
    const matches = [];

    for (let i = 0; i < count; i++) {
        const row = rows.nth(i);
        const recycledText = (await row.locator(`td:nth-child(${recycledCol})`).first().innerText().catch(() => "")).trim();
        const gstText = (await row.locator(`td:nth-child(${gstPaidCol})`).first().innerText().catch(() => "")).trim();

        const recycledMatch = !expRecycled || recycledText === expRecycled || parseFloat(recycledText) === parseFloat(expRecycled);
        const gstMatch = !expGst || gstText === expGst || parseFloat(gstText) === parseFloat(expGst);

        if (recycledMatch && gstMatch) {
            let dateStr = "";
            if (dateCol) {
                dateStr = (await row.locator(`td:nth-child(${dateCol})`).first().innerText().catch(() => "")).trim();
            }
            matches.push({ row, dateStr, index: i });
            logStep(`  match at table row ${i}: recycled=${recycledText} gst=${gstText} date=${dateStr}`, 2);
        }
    }

    if (matches.length === 0) {
        logStep("No exact match, using first row", 1);
        return rows.first();
    }
    if (matches.length === 1) return matches[0].row;

    // Multiple matches: pick latest by date
    matches.sort((a, b) => {
        const da = new Date(a.dateStr || 0);
        const db = new Date(b.dateStr || 0);
        return db.getTime() - da.getTime();
    });
    logStep(`Multiple matches (${matches.length}), using latest by date`, 1);
    return matches[0].row;
}

// ── PDF matching ─────────────────────────────────────────────────────────────

function findPdfByLast4Digits(rootDir, invoiceNo) {
    const inv = String(invoiceNo || "").trim();
    if (!inv) return "";
    const last4 = inv.slice(-4).toLowerCase();
    if (!last4) return "";
    const skip = new Set(["node_modules", ".git", ".vscode"]);
    const stack = [rootDir];
    const matches = [];

    while (stack.length) {
        const dir = stack.pop();
        let entries = [];
        try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { continue; }
        for (const ent of entries) {
            const full = path.join(dir, ent.name);
            if (ent.isDirectory()) { if (!skip.has(ent.name)) stack.push(full); continue; }
            if (!ent.isFile()) continue;
            const nameLower = ent.name.toLowerCase();
            if (nameLower.endsWith(".pdf") && nameLower.includes(last4)) {
                matches.push(full);
            }
        }
    }

    if (matches.length === 0) return "";
    matches.sort((a, b) => a.length - b.length);
    return matches[0];
}

// ── Upload flow ──────────────────────────────────────────────────────────────

async function openUploadModal(page, { eprInvoiceNumber, invoiceNo, rootDir }) {
    if (!eprInvoiceNumber) throw new Error("Missing EPR Invoice Number for upload.");

    const row = await searchByEprInvoice(page, eprInvoiceNumber);

    const statusCol = await getTableColumnIndex(page, "Invoice File Status");
    if (!statusCol) throw new Error("Invoice File Status column not found.");

    const statusCell = row.locator(`td:nth-child(${statusCol})`).first();
    const greenIcon = statusCell.locator(".fa-check-circle.color-active, .color-green, .fa-check").first();
    const redIcon = statusCell.locator(".fa-exclamation-triangle.color-red").first();

    if (await greenIcon.count()) {
        return { status: "already", toast: "Already uploaded (green icon)." };
    }
    if (!(await redIcon.count())) {
        return { status: "skipped", toast: "Invoice File Status not red." };
    }

    await redIcon.click();
    const modal = page.locator('.modal-dialog:has-text("Upload Invoice")').first();
    await modal.waitFor({ state: "visible", timeout: 20000 });

    const filePath = findPdfByLast4Digits(rootDir, invoiceNo);
    if (!filePath) throw new Error(`PDF not found for invoice last4: ${String(invoiceNo).slice(-4)}`);
    if (!fs.existsSync(filePath)) throw new Error(`PDF path does not exist: ${filePath}`);
    logStep(`PDF found: ${path.basename(filePath)}`, 2);

    const fileInput = modal.locator('input[type="file"][name="invoice"]').first();
    await fileInput.setInputFiles(filePath);
    await page.waitForFunction((el) => el && el.files && el.files.length > 0, await fileInput.elementHandle(), { timeout: 15000 });

    return { status: "ready", modal };
}

async function uploadWithRetry(page, opts) {
    const prep = await openUploadModal(page, opts);
    if (prep.status === "already" || prep.status === "skipped") {
        return { status: prep.status, toast: prep.toast, attempts: 0 };
    }

    const modal = prep.modal;
    const uploadBtn = modal.locator("button.btn.btn-primary", { hasText: "Upload" }).first();

    let lastToast = "";
    let lastStatus = "error";
    for (let attempt = 1; attempt <= 2; attempt++) {
        logStep(`Upload attempt ${attempt}`, 1);
        await waitForLoaderToFinish(page);
        try { await uploadBtn.click(); } catch { await uploadBtn.click({ force: true }); }
        logStep("Upload clicked", 1);
        await page.waitForTimeout(1200);
        const toastText = await readToastTextWithRetry(page, 6, 300);
        logStep(`Toast: ${toastText || "(none)"}`, 1);
        lastToast = toastText;
        lastStatus = isSuccessText(toastText) ? "success" : "error";
        if (lastStatus === "success") break;
    }

    // Close modal
    try {
        if (await modal.isVisible().catch(() => false)) {
            const closeBtn = modal.locator("button", { hasText: "Close" }).first();
            if (await closeBtn.count()) { await closeBtn.click(); }
            else {
                const closeIcon = modal.locator("#closeInvoiceUploadPopup, .close").first();
                if (await closeIcon.count()) await closeIcon.click();
            }
            await modal.waitFor({ state: "hidden", timeout: 3000 }).catch(() => { });
        }
    } catch { }

    return { status: lastStatus, toast: lastToast, attempts: 2 };
}

// ── Main ─────────────────────────────────────────────────────────────────────

(async () => {
    const readPath = fs.existsSync(OUTPUT_PATH) ? OUTPUT_PATH : EXCEL_PATH;
    if (!fs.existsSync(readPath)) throw new Error(`Excel not found: ${readPath}`);
    const stat = fs.statSync(readPath);
    if (stat.size < 1000) throw new Error(`Excel looks empty/corrupt (${stat.size} bytes): ${readPath}`);
    console.log(`Reading from: ${readPath}`);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(readPath);
    let ws = wb.getWorksheet(SHEET);
    if (!ws) { wb.eachSheet((sheet) => { if (sheet.name.trim() === SHEET.trim()) ws = sheet; }); }
    if (!ws) throw new Error(`Sheet not found: ${SHEET}`);

    const headerMap = getHeaderMap(ws);
    ensureColumn(ws, headerMap, "Upload Status");
    const headerList = getHeaderList(ws);

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext(fs.existsSync(STORAGE) ? { storageState: STORAGE } : {});
    const page = await context.newPage();

    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" });
    if (!fs.existsSync(STORAGE) || (await isLoginPage(page))) {
        await attemptLogout(page);
        console.log("Login manually in this Playwright window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: STORAGE });
        console.log("Saved session.");
    }

    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(3000);
    await waitForLoaderToFinish(page);

    // Click Fetch to load the table data
    logStep("Clicking Fetch button...");
    const fetchBtn = page.locator("button.btn.btn-primary", { hasText: /^\s*Fetch\s*$/ }).first();
    await fetchBtn.waitFor({ state: "visible", timeout: 30000 });
    await fetchBtn.click();
    await page.waitForTimeout(1500);
    await waitForLoaderToFinish(page);
    logStep("Fetch done, table loaded");

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;
    const pdfRoot = CONFIG.invoicePdfDir ? path.resolve(__dirname, CONFIG.invoicePdfDir) : ROOT_DIR;

    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        if (isRowEmpty(row, headerMap)) { continue; }

        const eprInvoice = cellText(getVal(row, headerMap, "EPR Invoice Number"));
        const invoiceNo = cellText(getVal(row, headerMap, "INVOICE No"));
        const uploadStatus = cellText(getVal(row, headerMap, "Upload Status"));

        if (!eprInvoice) { continue; }
        if (uploadStatus && (uploadStatus.toLowerCase().includes("success") || uploadStatus.toLowerCase().includes("uploaded"))) {
            continue;
        }

        console.log(`Row ${r}: uploading for EPR ${eprInvoice}...`);

        try {
            const result = await uploadWithRetry(page, {
                eprInvoiceNumber: eprInvoice,
                invoiceNo,
                rootDir: pdfRoot,
            });

            const msg = result.toast || result.status;
            if (result.status === "already") {
                setVal(row, headerMap, "Upload Status", "Invoice Uploaded");
            } else if (result.status === "success") {
                setVal(row, headerMap, "Upload Status", "Invoice Uploaded");
            } else if (result.status === "skipped") {
                setVal(row, headerMap, "Upload Status", "Skipped: " + msg);
            } else {
                setVal(row, headerMap, "Upload Status", "Failed: " + msg);
            }
            console.log(`Row ${r}: ${cellText(getVal(row, headerMap, "Upload Status"))}`);

            appendUploadLogRow(row, headerMap, { status: result.status, message: msg });
            appendUploadFilledRow(row, headerMap, headerList, { message: msg });
        } catch (e) {
            const msg = String(e?.message || e);
            console.log(`Row ${r}: Failed ->`, msg);
            setVal(row, headerMap, "Upload Status", "Failed: " + msg);
            appendUploadLogRow(row, headerMap, { status: "Failed", message: msg });
            appendUploadFilledRow(row, headerMap, headerList, { message: msg });
        }

        // Save Excel IMMEDIATELY after each row
        row.commit();
        await safeWriteWorkbook(wb);
        await syncInputWorkbook(wb);
    }

    await browser.close();
    console.log("Done. Updated Excel:", OUTPUT_PATH);
})();
