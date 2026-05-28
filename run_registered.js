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
    return path.resolve(__dirname, "config.json");
}

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
    const plasticType = String(cfg?.plasticType || "PP").trim();
    const storageState = String(cfg?.storageState || "storageState.json").trim();
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config.json must include inputExcel, sheetName, and outputExcel");
    }
    let maxRows = null;
    if (maxRowsRaw !== undefined && maxRowsRaw !== null && String(maxRowsRaw).trim() !== "") {
        const n = Number(maxRowsRaw);
        if (!Number.isFinite(n)) {
            throw new Error("config.json max_rows must be a number when provided");
        }
        if (n > 0) maxRows = Math.floor(n);
    }
    return { inputExcel, sheetName, outputExcel, maxRows, plasticType, storageState };
}

const CONFIG = loadConfig();
const STORAGE = path.resolve(__dirname, CONFIG.storageState);
const EXCEL_PATH = path.resolve(__dirname, CONFIG.inputExcel);
const SHEET = CONFIG.sheetName;
const OUTPUT_PATH = path.resolve(__dirname, CONFIG.outputExcel);
const EXCEL_TMP = `${OUTPUT_PATH}.tmp`;
const EXCEL_BAK = `${OUTPUT_PATH}.bak`;
const OUTPUT_BASENAME = path.basename(OUTPUT_PATH, path.extname(OUTPUT_PATH));
const LOG_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_log.csv`);
const FILLED_OUTPUT_PATH = path.resolve(__dirname, `${OUTPUT_BASENAME}_filled.csv`);

function normHeader(s) {
    return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object" && v.text) return String(v.text).trim();
    return String(v).trim();
}

function excelDateToISO(v) {
    if (!v) throw new Error("Sales date is empty");
    if (v instanceof Date) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, "0");
        const d = String(v.getDate()).padStart(2, "0");
        return `${y}-${m}-${d}`;
    }
    const s = cellText(v);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    if (/^\d{2}\.\d{2}\.\d{4}$/.test(s)) {
        const [dd, mm, yyyy] = s.split(".");
        return `${yyyy}-${mm}-${dd}`;
    }
    if (/^\d{2}-\d{2}-\d{4}$/.test(s)) {
        const [dd, mm, yyyy] = s.split("-");
        return `${yyyy}-${mm}-${dd}`;
    }
    const parts = s.split("/");
    if (parts.length === 3) {
        const mm = String(Number(parts[0])).padStart(2, "0");
        const dd = String(Number(parts[1])).padStart(2, "0");
        const yyyy = String(Number(parts[2]));
        return `${yyyy}-${mm}-${dd}`;
    }
    throw new Error(`Unsupported Sales date format: ${s}`);
}

function formatQty(v) {
    if (typeof v === "number") return v.toFixed(2);
    const s = cellText(v);
    const n = Number(s);
    if (Number.isFinite(n)) return n.toFixed(2);
    return s;
}

function randDelayMs(minMs = 3000, maxMs = 7000) {
    const min = Math.floor(minMs);
    const max = Math.floor(maxMs);
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) {
        return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
}

function logStep(message, level = 0) {
    const indent = "  ".repeat(level);
    const ts = new Date().toISOString();
    console.log(`${indent}${ts} ${message}`);
}

function ensureLogHeader() {
    if (fs.existsSync(LOG_PATH)) return;
    const header = [
        "datetime",
        "row",
        "e_invoice_number",
        "sales_date",
        "quantity_sold_mt",
        "registration_type",
        "entity_name",
        "seller_gst",
        "buyer_gst",
        "epr_invoice_number",
        "status",
        "message",
    ].join(",");
    fs.writeFileSync(LOG_PATH, `${header}\n`);
}

function appendLogRow(row, headerMap, { status, eprInvoiceNumber, message }) {
    ensureLogHeader();
    const ts = new Date().toISOString();
    const data = [
        ts,
        row.number,
        cellText(getVal(row, headerMap, "E-Invoice Number*")),
        cellText(getVal(row, headerMap, "Sales date*")),
        cellText(getVal(row, headerMap, "Quantity Sold(MT)")),
        cellText(getVal(row, headerMap, "Registration Type*")),
        cellText(getVal(row, headerMap, "Name of the Entity *")),
        cellText(getVal(row, headerMap, "GST No. of Seller *")),
        cellText(getVal(row, headerMap, "Buyer GST")),
        cellText(eprInvoiceNumber),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(LOG_PATH, `${data.join(",")}\n`);
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

function ensureFilledHeader(headerList) {
    if (fs.existsSync(FILLED_OUTPUT_PATH)) return;
    const header = [...headerList, "datetime", "message"].map(csvEscape).join(",");
    fs.writeFileSync(FILLED_OUTPUT_PATH, `${header}\n`);
}

function appendFilledRow(row, headerMap, headerList, { message }) {
    ensureFilledHeader(headerList);
    const ts = new Date().toISOString();
    const rowData = headerList.map((h) => cellText(getVal(row, headerMap, h)));
    const data = [...rowData, ts, message || ""].map(csvEscape);
    fs.appendFileSync(FILLED_OUTPUT_PATH, `${data.join(",")}\n`);
}

async function setAngularDateById(page, id, isoDate) {
    await page.evaluate(
        ({ id, isoDate }) => {
            const el = document.getElementById(id);
            if (!el) throw new Error(`Element #${id} not found`);
            el.value = isoDate;
            el.dispatchEvent(new Event("input", { bubbles: true }));
            el.dispatchEvent(new Event("change", { bubbles: true }));
            el.blur();
        },
        { id, isoDate }
    );
}

async function fillById(page, id, value) {
    const v = cellText(value);
    const loc = page.locator(`#${id}`);
    await loc.waitFor({ state: "visible", timeout: 60000 });
    await loc.scrollIntoViewIfNeeded();
    await loc.click({ timeout: 15000 });
    await loc.fill("");
    if (v) await loc.fill(v);
    await loc.blur();
}

async function fillBySelector(page, selector, value) {
    const v = cellText(value);
    const loc = page.locator(selector).first();
    await loc.waitFor({ state: "visible", timeout: 60000 });
    await loc.scrollIntoViewIfNeeded();
    await loc.click({ timeout: 15000 });
    await loc.fill("");
    if (v) await loc.fill(v);
    await loc.blur();
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
    if (!col) return;
    row.getCell(col).value = value;
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

function buildEprSet(ws, headerMap) {
    const set = new Set();
    const col = headerMap.get(normHeader("EPR Invoice Number"));
    if (!col) return set;
    for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        if (!row.hasValues) continue;
        const v = cellText(row.getCell(col).value);
        if (v) set.add(v);
    }
    return set;
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
    logStep("wait for loader: start", 2);
    for (const sel of loaders) {
        try {
            const loc = page.locator(sel);
            if ((await loc.count()) > 0) {
                await loc.first().waitFor({ state: "hidden", timeout: 30000 }).catch(() => { });
            }
        } catch { }
    }
    logStep("wait for loader: done", 2);
}

async function waitForQtyInput(page, timeoutMs = 8000) {
    const qtyInput = page.locator('input[name="qty_product_sold"]').first();
    if (!(await qtyInput.count())) return false;
    try {
        await qtyInput.waitFor({ state: "visible", timeout: timeoutMs });
        return true;
    } catch {
        return false;
    }
}

async function clickAddNew(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    logStep("click add new: start", 1);
    await addNewBtn.waitFor({ state: "visible", timeout: 60000 });
    await addNewBtn.scrollIntoViewIfNeeded();
    await addNewBtn.click();
    await page.waitForTimeout(150);
    logStep("click add new: done", 1);
}

async function clickAddNewIfVisible(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    if (!(await addNewBtn.count())) return false;
    await addNewBtn.waitFor({ state: "visible", timeout: 5000 }).catch(() => { });
    await addNewBtn.scrollIntoViewIfNeeded().catch(() => { });
    await addNewBtn.click().catch(() => { });
    await page.waitForTimeout(150);
    logStep("click add new (visible): done", 2);
    return true;
}

async function resetToFreshPage(page) {
    logStep("reset to fresh page: start", 1);
    await page.goto(URL, { waitUntil: "domcontentloaded" }).catch(() => { });
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 }).catch(() => { });
    await clickAddNewIfVisible(page);
    await page.waitForTimeout(800);
    logStep("reset to fresh page: done", 1);
}

async function ensureSalesFormReady(page) {
    const qtyInput = page.locator('input[name="qty_product_sold"]').first();
    if (await qtyInput.count()) {
        try {
            await qtyInput.waitFor({ state: "visible", timeout: 2000 });
            return true;
        } catch { }
    }

    const resetBtn = page.locator("button", { hasText: /\bReset\b/i }).first();
    if (await resetBtn.count()) {
        logStep("ensure form ready: reset", 1);
        await clickResetAndConfirm(page).catch(() => { });
        await page.waitForTimeout(1500);
    } else {
        logStep("ensure form ready: add new", 1);
        await clickAddNewIfVisible(page);
    }
    return false;
}

async function selectCat2Row(page, plasticTypeText) {
    logStep("select CAT-II row: start", 1);
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 });
    let catRow = page.locator("tbody#ScrollableSimpleTableBody tr", {
        has: page.locator('span[title="CAT-II"]'),
    });

    if (plasticTypeText) {
        catRow = catRow.filter({
            has: page.locator(`span[title="${plasticTypeText}"]`),
        });
    }

    catRow = catRow.first();
    await catRow.waitFor({ state: "visible", timeout: 20000 });
    const checkbox = catRow.locator('input[type="checkbox"][name="check-box"]').first();
    await checkbox.scrollIntoViewIfNeeded();
    await checkbox.click({ force: true });
    await page.waitForSelector('input[name="qty_product_sold"]', { timeout: 30000 });
    logStep("select CAT-II row: done", 1);
}

async function selectCat2RowWithRetry(page, plasticTypeText, attempts = 3) {
    let lastErr = null;
    for (let i = 0; i < attempts; i++) {
        try {
            logStep(`select CAT-II retry ${i + 1}/${attempts}`, 1);
            await ensureSalesFormReady(page);
            await selectCat2Row(page, plasticTypeText);
            return true;
        } catch (e) {
            lastErr = e;
            logStep(`select CAT-II failed: ${String(e?.message || e)}`, 1);
            await waitForLoaderToFinish(page);
            await page.locator("#refersh_data").first().click().catch(() => { });
            const didReset = await clickResetAndConfirm(page);
            if (!didReset) {
                await clickAddNewIfVisible(page);
            }
            if (i === 1) {
                await resetToFreshPage(page);
            }
    await page.waitForTimeout(300);
        }
    }
    throw lastErr || new Error("Failed to select CAT-II row");
}

async function selectNgSelectByLabel(page, labelText, optionText) {
    const text = cellText(optionText);
    if (!text) throw new Error(`Missing option for ${labelText}`);

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: labelText }) })
        .first();

    logStep(`select ng-select: ${labelText} -> ${text}`, 1);
    await group.waitFor({ state: "visible", timeout: 20000 });
    const ng = group.locator("ng-select").first();
    await ng.scrollIntoViewIfNeeded();
    await ng.click();

    const panel = page.locator(".ng-dropdown-panel");
    await panel.waitFor({ state: "visible", timeout: 20000 });

    const searchInput = panel.locator("input[type='text']").first();
    if (await searchInput.count()) {
        try {
            await searchInput.fill(text);
    await page.waitForTimeout(120);
        } catch { }
    }

    const opt = panel.locator(".ng-option", { hasText: text }).first();
    await opt.waitFor({ state: "visible", timeout: 20000 });
    await opt.click();

    await panel.waitFor({ state: "hidden", timeout: 20000 }).catch(() => { });
}

async function clickSubmitAndConfirm(page) {
    const submit = page.locator('button[type="submit"]', { hasText: "Generate EPR Invoice Number" }).first();
    logStep("submit: start", 1);
    await submit.waitFor({ state: "visible", timeout: 20000 });
    if (await submit.isDisabled()) {
        throw new Error("Submit disabled: some required fields still missing.");
    }
    await submit.click();
    try {
        const confirmBtn = page.locator(".modal-footer button", { hasText: "Confirm" }).first();
        await confirmBtn.waitFor({ state: "visible", timeout: 60000 });
        await confirmBtn.click();
    } catch { }
    logStep("submit: done", 1);
}

async function clickResetAndConfirm(page) {
    const reset = page.locator("button", { hasText: /\bReset\b/i }).first();
    if (!(await reset.count())) return false;
    logStep("reset: start", 1);
    await reset.waitFor({ state: "visible", timeout: 20000 }).catch(() => { });
    await reset.scrollIntoViewIfNeeded().catch(() => { });
    await reset.click().catch(() => { });

    const modal = page.locator(".modal-dialog, .modal-content").first();
    if (await modal.count()) {
        try {
            await modal.waitFor({ state: "visible", timeout: 2000 });
            const confirmBtn = modal.getByRole("button", { name: "Confirm", exact: true }).first();
            if (await confirmBtn.count()) {
                await confirmBtn.click();
            }
        } catch { }
    }
    logStep("reset: done", 1);
    return true;
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

async function readEprInvoiceNumber(page) {
    const input = page.locator("#invoiceNumberCopy").first();
    if (await input.count()) {
        try {
            const val = (await input.inputValue()).trim();
            if (val) return val;
        } catch { }
    }
    const label = page.locator("text=/EPR\\s*Invoice\\s*Number/i").first();
    if (await label.count()) {
        const container = label.locator("xpath=ancestor-or-self::*[self::div or self::span or self::p][1]");
        const text = (await container.innerText().catch(() => "")) || "";
        if (/confirm entered details/i.test(text)) return "";
        const match = text.match(/EPR\\s*Invoice\\s*Number\\s*[:\\-]?\\s*([A-Za-z0-9\\-\\/]+)/i);
        if (match && match[1]) return match[1].trim();
        const sibling = label.locator("xpath=following::span[1] | following::div[1] | following::p[1]").first();
        const sibText = (await sibling.innerText().catch(() => "")).trim();
        if (/confirm entered details/i.test(sibText)) return "";
        if (sibText) return sibText;
    }
    return "";
}

async function waitForEprInvoiceNumber(page, timeoutMs = 20000) {
    const start = Date.now();
    logStep("wait EPR invoice: start", 1);
    while (Date.now() - start < timeoutMs) {
        const modal = page.locator(".modal-dialog, .modal-content").first();
        if (await modal.count()) {
            const confirmBtn = modal.getByRole("button", { name: "Confirm", exact: true }).first();
            if (await confirmBtn.count()) {
                await confirmBtn.click().catch(() => { });
                await modal.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
            }
        }
        const val = await readEprInvoiceNumber(page);
        if (val) return val;
        await page.waitForTimeout(300);
    }
    logStep("wait EPR invoice: timeout", 1);
    return "";
}

async function pickEntityName(page, entityNameValue) {
    const name = cellText(entityNameValue);
    if (!name) throw new Error("Entity name empty");

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: "Name of the Entity" }) })
        .first();

    if (await group.count()) {
        const ng = group.locator("ng-select").first();
        if (await ng.count()) {
            console.log("pick entity name: ng-select");
            await ng.scrollIntoViewIfNeeded();
            await ng.click();

            const panel = page.locator(".ng-dropdown-panel");
            await panel.waitFor({ state: "visible", timeout: 20000 });

            const searchInput = panel.locator("input[type='text']").first();
            if (await searchInput.count()) {
                await searchInput.fill(name);
                await page.waitForTimeout(150);
                await searchInput.press("Enter");
            } else {
                await page.keyboard.type(name);
                await page.waitForTimeout(300);
                await page.keyboard.press("Enter");
            }

            await panel.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
            const selected = group.locator(".ng-value-label").first();
            if (await selected.count()) {
                const selectedText = (await selected.innerText().catch(() => "")).trim();
                if (selectedText) {
                    console.log(`pick entity name: selected "${selectedText}"`);
                    await waitForLoaderToFinish(page);
                    return;
                }
            }

            console.log("pick entity name: enter did not select, fallback to click");
            await selectNgSelectByLabel(page, "Name of the Entity", name);
            await waitForLoaderToFinish(page);
            return;
        }
    }

    const input = page.locator('input[formcontrolname="entity_name"]').first();
    await input.waitFor({ state: "visible", timeout: 30000 });
    await input.scrollIntoViewIfNeeded();
    console.log("pick entity name: input autocomplete");
    await input.click();
    await input.fill(name);

    await page.waitForTimeout(300);

    const suggestion = page.locator(
        'ul li, .dropdown-item, .typeahead-item, .autocomplete-items div'
    ).first();

    if (await suggestion.count()) {
        await suggestion.click();
    } else {
        await input.blur();
    }

    await waitForLoaderToFinish(page);
}

async function waitEntityAutofill(page) {
    await waitForLoaderToFinish(page);
    const addr = page.locator('input[formcontrolname="entity_address"]').first();
    await addr.waitFor({ state: "visible", timeout: 20000 });
    await page
        .waitForFunction(() => {
            const a = document.querySelector('input[formcontrolname="entity_address"]');
            return a && a.value && a.value.trim().length > 3;
        }, { timeout: 30000 })
        .catch(() => { });
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
    const eprSet = buildEprSet(ws, headerMap);
    const headerList = getHeaderList(ws);

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
        console.log("Saved session to storageState.json");
    }

    await page.goto(URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#ScrollableSimpleTableBody", { timeout: 60000 });
    await clickAddNew(page);

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;
    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        let successThisRow = false;
        if (isRowEmpty(row, headerMap)) {
            console.log(`Row ${r}: Skipped (row empty)`);
            continue;
        }

        const status = cellText(getVal(row, headerMap, "Status"));
        if (status.toLowerCase().includes("success") || status.toLowerCase().includes("filled")) {
            continue;
        }

        const eprInvoiceExisting = getVal(row, headerMap, "EPR Invoice Number");
        if (!isCellEmpty(eprInvoiceExisting)) {
            console.log(`Row ${r}: Skipped (EPR Invoice already present)`);
            continue;
        }

        const qtySold = getVal(row, headerMap, "Quantity Sold(MT)");
        const regType = getVal(row, headerMap, "Registration Type*");
        const entityType = getVal(row, headerMap, "Entity Type*");
        const entityName = getVal(row, headerMap, "Name of the Entity *");
        const sellerGst = getVal(row, headerMap, "GST No. of Seller *");
        const buyerGst = getVal(row, headerMap, "Buyer GST");
        const hsn = getVal(row, headerMap, "HSN CODE");
        const invno = getVal(row, headerMap, "E-Invoice Number*");
        const account = getVal(row, headerMap, "Bank Account No*");
        const ifsc = getVal(row, headerMap, "IFSC Code*");
        const principal = getVal(row, headerMap, "Principal Amount(₹)*");
        const gstOther = getVal(row, headerMap, "GST & Other Charges(₹)*");
        const salesDateRaw = getVal(row, headerMap, "Sales date*");

        try {
            console.log(`Row ${r} starting...`);
            await selectCat2RowWithRetry(page, CONFIG.plasticType || "PP");
            if (!(await waitForQtyInput(page, 8000))) {
                throw new Error("qty_product_sold not visible");
            }
            logStep("fill qty: start", 1);
            await fillBySelector(page, 'input[name="qty_product_sold"]', formatQty(qtySold));
            logStep("fill qty: done", 1);

            logStep("select registration type: start", 1);
            await selectNgSelectByLabel(page, "Registration Type", regType);
            logStep("select registration type: done", 1);
            logStep("select entity type: start", 1);
            await selectNgSelectByLabel(page, "Entity Type", entityType);
            logStep("select entity type: done", 1);

            console.log("wait for entity list: 500ms");
            await page.waitForTimeout(500);
            logStep("pick entity name: start", 1);
            await pickEntityName(page, entityName);
            logStep("pick entity name: done", 1);
            logStep("wait entity autofill: start", 1);
            await waitEntityAutofill(page);
            logStep("wait entity autofill: done", 1);

            logStep("fill remaining fields: start", 1);
            await fillById(page, "sellerGst", sellerGst);
            await fillById(page, "buyerGst", buyerGst);
            await fillById(page, "hsnCode", hsn);
            await fillById(page, "invno", invno);
            await fillById(page, "account_number", account);
            await fillById(page, "ifsc_code", ifsc);
            await fillById(page, "amount", principal);
            await fillById(page, "gst", gstOther);
            logStep("fill remaining fields: done", 1);

            logStep("set sales date: start", 1);
            const salesDateISO = excelDateToISO(salesDateRaw);
            await setAngularDateById(page, "salesDate", salesDateISO);
            logStep(`set sales date: done (${salesDateISO})`, 1);

            await clickSubmitAndConfirm(page);
            logStep("post-submit: wait 300ms", 1);
            await page.waitForTimeout(300);
            await waitForLoaderToFinish(page);
            const toastText = await readToastText(page);
            if (toastText) {
                logStep(`toast: ${toastText}`, 1);
            }

            const eprInvoice = await waitForEprInvoiceNumber(page);
            if (eprSet.has(eprInvoice)) {
                throw new Error("Duplicate EPR Invoice Number: " + eprInvoice);
            }
            if (!eprInvoice) {
                throw new Error("EPR Invoice Number not found after submit.");
            }
            logStep(`EPR invoice: ${eprInvoice}`, 1);

            setVal(row, headerMap, "Status", "Filled");
            setVal(row, headerMap, "EPR Invoice Number", eprInvoice);
            eprSet.add(eprInvoice);
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);

            appendLogRow(row, headerMap, {
                status: "Filled",
                eprInvoiceNumber: eprInvoice,
                message: toastText,
            });
            appendFilledRow(row, headerMap, headerList, {
                message: toastText,
            });

            console.log(`Row ${r}: Filled ✅`);
            successThisRow = true;
            const delayMs = randDelayMs(3000, 7000);
            const startTs = new Date().toISOString();
            console.log(`Row ${r}: delay start ${startTs} (${delayMs}ms)`);
            // await page.waitForTimeout(delayMs);
            const endTs = new Date().toISOString();
            console.log(`Row ${r}: delay end ${endTs}`);
        } catch (e) {
            const msg = String(e?.message || e);
            console.log(`Row ${r}: Failed ❌ ->`, msg);
            setVal(row, headerMap, "Status", "Failed: " + msg);
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);
            appendLogRow(row, headerMap, {
                status: "Failed",
                eprInvoiceNumber: "",
                message: msg,
            });
            appendFilledRow(row, headerMap, headerList, {
                message: msg,
            });
        } finally {
            if (page.isClosed()) {
                console.log("Page closed. Stopping.");
                break;
            }
            if (successThisRow) {
                await page.waitForTimeout(300);
            }
            await waitForLoaderToFinish(page);
            const didReset = await clickResetAndConfirm(page);
            if (!didReset) {
                await clickAddNewIfVisible(page);
            }
            await page.waitForTimeout(500);
        }
    }

    await browser.close();
    console.log("Done. Updated Excel:", EXCEL_PATH);
})();
