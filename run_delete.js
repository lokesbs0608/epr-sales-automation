const ExcelJS = require("exceljs");
const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");

const URL = "https://eprplastic.cpcb.gov.in/#/epr/details/sales";

function getConfigPath() {
    const idx = process.argv.indexOf("--config");
    if (idx !== -1 && process.argv[idx + 1]) {
        return path.resolve(__dirname, process.argv[idx + 1]);
    }
    return path.resolve(__dirname, "config_delete.json");
}

// Dry run: do everything (search, find delete icon, open the Confirm Deletion
// modal, verify the Delete button) EXCEPT clicking the final Delete. Nothing is
// actually deleted. Enable with `--dry-run` on the CLI or "dryRun": true in config.
const CLI_DRY_RUN = process.argv.includes("--dry-run");

const CONFIG_PATH = getConfigPath();

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
    const storageState = String(cfg?.storageState || "storageState_delete.json").trim();
    const dryRun = CLI_DRY_RUN || cfg?.dryRun === true;
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config_delete.json must include inputExcel, sheetName, and outputExcel");
    }
    let maxRows = null;
    if (maxRowsRaw !== undefined && maxRowsRaw !== null && String(maxRowsRaw).trim() !== "") {
        const n = Number(maxRowsRaw);
        if (!Number.isFinite(n)) {
            throw new Error("config_delete.json max_rows must be a number when provided");
        }
        if (n > 0) {
            maxRows = Math.floor(n);
        }
    }
    return { inputExcel, sheetName, outputExcel, maxRows, storageState, dryRun };
}

const CONFIG = loadConfig();
const STORAGE = path.resolve(__dirname, CONFIG.storageState);
const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const DELETE_LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_delete_log.csv`);

// Column the script writes the delete outcome into. Prefers "Delete Status",
// falls back to "Status" if that header is not present in the sheet.
const DELETE_STATUS_HEADERS = ["Delete Status", "Status"];

function normHeader(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
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

function getVal(row, headerMap, headerName) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return "";
    return row.getCell(col).value;
}

function setVal(row, headerMap, headerName, value) {
    const col = headerMap.get(normHeader(headerName));
    if (!col) return false;
    row.getCell(col).value = value;
    return true;
}

// Resolve which header column the delete status should be written to.
function resolveStatusHeader(headerMap) {
    for (const h of DELETE_STATUS_HEADERS) {
        if (headerMap.has(normHeader(h))) return h;
    }
    return DELETE_STATUS_HEADERS[0];
}

function setStatus(row, headerMap, statusHeader, value) {
    if (!setVal(row, headerMap, statusHeader, value)) {
        // Header missing in the sheet; still try the fallbacks.
        for (const h of DELETE_STATUS_HEADERS) {
            if (setVal(row, headerMap, h, value)) return;
        }
    }
}

function isCellEmpty(v) {
    return cellText(v) === "";
}

function isRowEmpty(row, headerMap) {
    for (const col of headerMap.values()) {
        if (!isCellEmpty(row.getCell(col).value)) return false;
    }
    return true;
}

function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) {
        return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
}

function ensureDeleteLogHeader() {
    if (fs.existsSync(DELETE_LOG_PATH)) return;
    const header = ["datetime", "row", "epr_invoice_number", "status", "message"].join(",");
    fs.writeFileSync(DELETE_LOG_PATH, `${header}\n`);
}

function appendDeleteLogRow(row, eprInvoice, { status, message }) {
    ensureDeleteLogHeader();
    const ts = new Date().toISOString();
    const data = [ts, row.number, eprInvoice, status, message || ""].map(csvEscape);
    fs.appendFileSync(DELETE_LOG_PATH, `${data.join(",")}\n`);
}

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

async function waitForLoaderToFinish(page) {
    const loaders = [
        ".spinner-border",
        ".loading",
        ".loader",
        ".ngx-spinner-overlay",
        ".ngx-spinner",
        ".overlay",
        ".block-ui-wrapper",
        ".k-i-loading",
    ];
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
    try {
        const text = (await toast.innerText()).trim();
        return text;
    } catch {
        return "";
    }
}

async function readToastTextWithRetry(page, retries = 6, delayMs = 300) {
    for (let i = 0; i < retries; i++) {
        const text = await readToastText(page);
        if (text) return text;
        await page.waitForTimeout(delayMs);
    }
    return "";
}

function isSuccessText(text) {
    const t = String(text || "");
    return /success|deleted/i.test(t) && !/error|fail/i.test(t);
}

function normalizeStatus(text) {
    return String(text || "").trim().toLowerCase();
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
        if (await logoutDirect.count()) {
            await logoutDirect.click().catch(() => { });
            await page.waitForTimeout(1000);
            return true;
        }
        const toggles = [
            'button[aria-haspopup="menu"]',
            'button[aria-expanded]',
            '.dropdown-toggle',
            '.nav-link.dropdown-toggle',
            '.user-profile',
            '.profile',
        ];
        for (const sel of toggles) {
            const btn = page.locator(sel).first();
            if (await btn.count()) {
                await btn.click().catch(() => { });
                const logout = page.locator('text=/logout/i').first();
                if (await logout.count()) {
                    await logout.click().catch(() => { });
                    await page.waitForTimeout(1000);
                    return true;
                }
            }
        }
    } catch { }
    return false;
}

async function getTableColumnIndex(page, headerText) {
    const idx = await page.evaluate((headerText) => {
        const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
        const ths = Array.from(document.querySelectorAll("#simple_table_header th"));
        const target = norm(headerText);
        const i = ths.findIndex((th) => norm(th.innerText) === target);
        return i >= 0 ? i + 1 : null;
    }, headerText);
    return idx;
}

async function searchSalesByEprInvoice(page, eprInvoiceNumber) {
    console.log(`[search] start for EPR ${eprInvoiceNumber}`);
    // Ensure any modal/overlay is gone before searching
    await page.locator(".modal-dialog").first().waitFor({ state: "hidden", timeout: 2000 }).catch(() => { });
    await waitForLoaderToFinish(page);
    const searchInput = page.locator('input[name="searchField"]').first();
    await searchInput.waitFor({ state: "visible", timeout: 30000 });
    await page
        .waitForFunction((el) => !el.disabled && !el.readOnly, await searchInput.elementHandle(), {
            timeout: 5000,
        })
        .catch(() => { });
    console.log("[search] input visible");
    await searchInput.click();
    await searchInput.fill("");
    await searchInput.fill(eprInvoiceNumber);

    const searchBtn = page.locator("button", { hasText: "Search" }).first();
    await searchBtn.click();
    console.log("[search] clicked");
    await page.waitForTimeout(1500);
    await waitForLoaderToFinish(page);

    const row = page.locator("#ScrollableSimpleTableBody tr", { hasText: eprInvoiceNumber }).first();
    await row.waitFor({ state: "visible", timeout: 20000 });
    console.log("[search] row visible");
    return row;
}

// Locate the delete (trash) icon/button inside the row's Action column and
// confirm the deletion in the modal that appears.
async function deleteRowByEprInvoice(page, eprInvoiceNumber) {
    const row = await searchSalesByEprInvoice(page, eprInvoiceNumber);

    // Find the Action column so we click the correct cell's delete control.
    let actionIdx =
        (await getTableColumnIndex(page, "Action")) ||
        (await getTableColumnIndex(page, "Actions"));

    let deleteIcon;
    if (actionIdx) {
        const actionCell = row.locator(`td:nth-child(${actionIdx})`).first();
        deleteIcon = actionCell
            .locator('.fa-trash, .fa-trash-o, .fa-trash-alt, [title*="Delete" i], [aria-label*="Delete" i]')
            .first();
    } else {
        // Fallback: search the whole row for a trash control.
        deleteIcon = row
            .locator('.fa-trash, .fa-trash-o, .fa-trash-alt, [title*="Delete" i], [aria-label*="Delete" i]')
            .first();
    }

    if (!(await deleteIcon.count())) {
        return { status: "skipped", toast: "Delete icon not found in row." };
    }

    await deleteIcon.scrollIntoViewIfNeeded().catch(() => { });
    await deleteIcon.click();

    // Confirm Deletion modal
    const modal = page.locator('#confirmModal, .modal-dialog:has-text("Confirm Deletion")').first();
    await modal.waitFor({ state: "visible", timeout: 20000 });

    const confirmBtn = modal.locator('button.btn-danger', { hasText: "Delete" }).first();
    await confirmBtn.waitFor({ state: "visible", timeout: 10000 });

    if (CONFIG.dryRun) {
        // Validate selectors only; do NOT delete. Close the modal and report.
        const cancelBtn = modal
            .locator('button.btn-secondary, button:has-text("Cancel"), .btn-close')
            .first();
        if (await cancelBtn.count()) {
            await cancelBtn.click().catch(() => { });
        }
        await modal.waitFor({ state: "hidden", timeout: 3000 }).catch(() => { });
        console.log("[dry-run] modal + Delete button found; NOT deleting");
        return { status: "dryrun", toast: "Dry run: row found, modal + Delete button OK (not deleted)." };
    }

    await confirmBtn.click();
    console.log("[delete] confirm clicked");

    await page.waitForTimeout(1200);
    const toastText = await readToastTextWithRetry(page, 6, 300);
    await waitForLoaderToFinish(page);

    // Close the modal if it is still open
    try {
        if (await modal.isVisible().catch(() => false)) {
            const closeBtn = modal
                .locator('button.btn-secondary, button:has-text("Cancel"), .btn-close')
                .first();
            if (await closeBtn.count()) {
                await closeBtn.click().catch(() => { });
            }
            await modal.waitFor({ state: "hidden", timeout: 3000 }).catch(() => { });
        }
    } catch { }

    const status = isSuccessText(toastText) || !toastText ? "success" : "error";
    return { status, toast: toastText || "Deleted" };
}

(async () => {
    if (!fs.existsSync(EXCEL_PATH)) {
        throw new Error(`Excel not found: ${EXCEL_PATH}`);
    }
    const stat = fs.statSync(EXCEL_PATH);
    if (stat.size < 1000) {
        throw new Error(`Excel looks empty/corrupt (${stat.size} bytes): ${EXCEL_PATH}`);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(EXCEL_PATH);

    const ws = wb.getWorksheet(SHEET);
    if (!ws) {
        throw new Error(`Sheet not found: ${SHEET}`);
    }

    const headerMap = getHeaderMap(ws);
    const statusHeader = resolveStatusHeader(headerMap);

    if (CONFIG.dryRun) {
        console.log("*** DRY RUN: rows will be searched and the modal opened, but NOTHING will be deleted. ***");
    }

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext(
        fs.existsSync(STORAGE) ? { storageState: STORAGE } : {}
    );
    const page = await context.newPage();

    await page.goto(URL, { waitUntil: "domcontentloaded" });

    if (!fs.existsSync(STORAGE) || (await isLoginPage(page))) {
        await attemptLogout(page);
        console.log("Login manually in this Playwright window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: STORAGE });
        console.log(`Saved session to ${path.basename(STORAGE)}`);
    }

    await page.goto(URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 });

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;

    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        if (isRowEmpty(row, headerMap)) {
            console.log(`Row ${r}: Skipped (row empty)`);
            continue;
        }

        const eprInvoice = cellText(getVal(row, headerMap, "EPR Invoice Number"));
        const deleteStatusRaw = cellText(getVal(row, headerMap, statusHeader));
        const deleteStatus = normalizeStatus(deleteStatusRaw);

        if (!eprInvoice) {
            console.log(`Row ${r}: Skipped (missing EPR Invoice Number)`);
            continue;
        }
        if (deleteStatus.includes("delet")) {
            console.log(`Row ${r}: Already deleted, skipping`);
            continue;
        }

        try {
            const result = await deleteRowByEprInvoice(page, eprInvoice);
            const message = result.toast || result.status;

            if (result.status === "dryrun") {
                appendDeleteLogRow(row, eprInvoice, { status: "DryRun", message });
                console.log(`Row ${r}: Dry run OK (${eprInvoice}) — not deleted`);
            } else if (result.status === "success") {
                setStatus(row, headerMap, statusHeader, "Deleted");
                appendDeleteLogRow(row, eprInvoice, { status: "Deleted", message });
                console.log(`Row ${r}: Deleted (${eprInvoice})`);
            } else if (result.status === "skipped") {
                setStatus(row, headerMap, statusHeader, "Skipped: " + message);
                appendDeleteLogRow(row, eprInvoice, { status: "Skipped", message });
                console.log(`Row ${r}: Skipped (${message})`);
            } else {
                setStatus(row, headerMap, statusHeader, "Failed: " + message);
                appendDeleteLogRow(row, eprInvoice, { status: "Failed", message });
                console.log(`Row ${r}: Failed (${message})`);
            }
        } catch (e) {
            const msg = String(e?.message || e);
            setStatus(row, headerMap, statusHeader, "Failed: " + msg);
            appendDeleteLogRow(row, eprInvoice, { status: "Failed", message: msg });
            console.log(`Row ${r}: Failed (${msg})`);
        }

        row.commit();
        if (!CONFIG.dryRun) {
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);
        }
    }

    await browser.close();
    if (CONFIG.dryRun) {
        console.log("Dry run complete. No data deleted. See log:", DELETE_LOG_PATH);
    } else {
        console.log("Done. Updated Excel:", EXCEL_PATH);
    }
})();
