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
    return path.resolve(__dirname, "config_pibo.json");
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
    const storageState = String(cfg?.storageState || "storageState_pibo.json").trim();
    const financialYear = String(cfg?.financialYear || "").trim();
    const address = String(cfg?.address || "").trim();
    const mobileNumber = String(cfg?.mobileNumber || "").trim();
    if (!inputExcel || !sheetName || !outputExcel) {
        throw new Error("config must include inputExcel, sheetName, and outputExcel");
    }
    let maxRows = null;
    if (maxRowsRaw !== undefined && maxRowsRaw !== null && String(maxRowsRaw).trim() !== "") {
        const n = Number(maxRowsRaw);
        if (!Number.isFinite(n)) {
            throw new Error("config max_rows must be a number when provided");
        }
        if (n > 0) maxRows = Math.floor(n);
    }
    return { inputExcel, sheetName, outputExcel, maxRows, storageState, financialYear, address, mobileNumber };
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

function formatQty(v) {
    if (typeof v === "number") return v.toFixed(2);
    const s = cellText(v);
    const n = Number(s);
    if (Number.isFinite(n)) return n.toFixed(2);
    return s;
}

const CATEGORY_MAP = {
    "cat 1": "CAT I",
    "cat 2": "CAT II",
    "cat 3": "CAT III",
    "cat 4": "CAT IV",
};

function mapCategory(v) {
    const s = cellText(v).toLowerCase();
    return CATEGORY_MAP[s] || cellText(v);
}

// ── Excel helpers ────────────────────────────────────────────────────────────

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

function ensureColumn(ws, headerMap, headerName) {
    const key = normHeader(headerName);
    if (headerMap.has(key)) return;
    const maxCol = Math.max(0, ...headerMap.values()) + 1;
    ws.getRow(1).getCell(maxCol).value = headerName;
    headerMap.set(key, maxCol);
}

// ── Logging ──────────────────────────────────────────────────────────────────

function ensureLogHeader() {
    if (fs.existsSync(LOG_PATH)) return;
    const header = [
        "datetime",
        "row",
        "gst_e_invoice_number",
        "registration_type",
        "entity_name",
        "quantity_tons",
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
        cellText(getVal(row, headerMap, "INVOICE No")),
        cellText(getVal(row, headerMap, "REGISTRATION TYPE")),
        cellText(getVal(row, headerMap, "ENTITY")),
        cellText(getVal(row, headerMap, "Sum of QTY MT")),
        cellText(eprInvoiceNumber),
        status,
        message || "",
    ].map(csvEscape);
    fs.appendFileSync(LOG_PATH, `${data.join(",")}\n`);
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

// ── Form interaction ─────────────────────────────────────────────────────────

async function selectRadioByLabel(page, groupLabel, optionText) {
    const text = cellText(optionText);
    if (!text) throw new Error(`Missing option for ${groupLabel}`);
    logStep(`select radio: ${groupLabel} -> ${text}`, 1);

    // Try clicking a radio label that matches the text
    const radio = page.locator(`label`, { hasText: new RegExp(`^\\s*${text}\\s*$`, "i") }).first();
    if (await radio.count()) {
        await radio.scrollIntoViewIfNeeded();
        await radio.click();
        await page.waitForTimeout(200);
        return;
    }

    // Fallback: find radio input by value
    const input = page.locator(`input[type="radio"][value="${text}" i]`).first();
    if (await input.count()) {
        await input.scrollIntoViewIfNeeded();
        await input.click({ force: true });
        await page.waitForTimeout(200);
        return;
    }

    // Fallback: find by name containing the group + matching value
    const anyRadio = page.locator(`input[type="radio"]`);
    const count = await anyRadio.count();
    for (let i = 0; i < count; i++) {
        const r = anyRadio.nth(i);
        const lbl = await r.evaluate((el) => {
            const id = el.id;
            if (id) {
                const l = document.querySelector(`label[for="${id}"]`);
                if (l) return l.innerText.trim();
            }
            const parent = el.closest("label");
            if (parent) return parent.innerText.trim();
            return "";
        });
        if (lbl.toLowerCase().includes(text.toLowerCase())) {
            await r.scrollIntoViewIfNeeded();
            await r.click({ force: true });
            await page.waitForTimeout(200);
            return;
        }
    }

    throw new Error(`Radio option "${text}" not found for ${groupLabel}`);
}

async function fillByName(page, name, value) {
    const v = cellText(value);
    if (!v) return;
    const loc = page.locator(`input[name="${name}"]`).first();
    await loc.waitFor({ state: "visible", timeout: 30000 });
    await loc.scrollIntoViewIfNeeded();
    await loc.click();
    await loc.fill("");
    await loc.fill(v);
    await loc.blur();
}

async function fillByNameIfEmpty(page, name, value) {
    const v = cellText(value);
    if (!v) return;
    const loc = page.locator(`input[name="${name}"]`).first();
    if (!(await loc.count())) return;
    await loc.waitFor({ state: "visible", timeout: 10000 }).catch(() => { });
    const current = (await loc.inputValue().catch(() => "")).trim();
    if (current) {
        logStep(`${name}: already filled ("${current}"), skipping`, 2);
        return;
    }
    await loc.scrollIntoViewIfNeeded();
    await loc.click();
    await loc.fill(v);
    await loc.blur();
}

async function isNgSelectFilled(page, labelText) {
    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: labelText }) })
        .first();
    if (!(await group.count())) return false;
    const selected = group.locator(".ng-value-label, .ng-value").first();
    if (!(await selected.count())) return false;
    const text = (await selected.innerText().catch(() => "")).trim();
    return text.length > 0;
}

async function selectNgSelectByLabelIfEmpty(page, labelText, optionText) {
    const text = cellText(optionText);
    if (!text) return;
    if (await isNgSelectFilled(page, labelText)) {
        logStep(`${labelText}: already filled, skipping`, 2);
        return;
    }
    await selectNgSelectByLabel(page, labelText, optionText);
}

async function getNgSelectValue(page, labelText) {
    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: labelText }) })
        .first();
    if (!(await group.count())) return "";
    const selected = group.locator(".ng-value-label, .ng-value").first();
    if (!(await selected.count())) return "";
    return (await selected.innerText().catch(() => "")).trim();
}

async function clearAndReselectNgSelect(page, labelText, optionText) {
    const text = cellText(optionText);
    if (!text) return;

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: labelText }) })
        .first();
    await group.waitFor({ state: "visible", timeout: 10000 });
    const ng = group.locator("ng-select").first();

    // Clear current selection
    const clearBtn = ng.locator(".ng-clear-wrapper").first();
    if (await clearBtn.count()) {
        await clearBtn.click();
        await page.waitForTimeout(200);
    }

    await selectNgSelectByLabel(page, labelText, optionText);
}

async function verifyAndFixNgSelect(page, labelText, expectedValue) {
    const expected = cellText(expectedValue);
    if (!expected) return;

    const current = await getNgSelectValue(page, labelText);
    if (current.toLowerCase() === expected.toLowerCase()) {
        logStep(`${labelText}: verified OK ("${current}")`, 2);
        return;
    }

    logStep(`${labelText}: mismatch ("${current}" != "${expected}"), re-selecting`, 1);
    await clearAndReselectNgSelect(page, labelText, expected);
    await waitForLoaderToFinish(page);
}

async function verifyAndFixInput(page, name, expectedValue) {
    const expected = cellText(expectedValue);
    if (!expected) return;

    const loc = page.locator(`input[name="${name}"]`).first();
    if (!(await loc.count())) return;
    const current = (await loc.inputValue().catch(() => "")).trim();
    if (current === expected) {
        logStep(`${name}: verified OK`, 2);
        return;
    }

    logStep(`${name}: mismatch ("${current}" != "${expected}"), overwriting`, 1);
    await loc.scrollIntoViewIfNeeded();
    await loc.click();
    await loc.fill("");
    await loc.fill(expected);
    await loc.blur();
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

let lastMatchedEntityIndex = -1;

async function pickEntityWithStateMatch(page, entityNameValue, expectedState, expectedGst) {
    const name = cellText(entityNameValue);
    const expState = cellText(expectedState).toLowerCase();
    const expGst = cellText(expectedGst).toLowerCase();
    if (!name) throw new Error("Entity name empty");

    const group = page
        .locator(".form-group", { has: page.locator("label", { hasText: "Name of the Entity" }) })
        .first();

    if (!(await group.count())) throw new Error("Name of the Entity field not found on form");

    const ng = group.locator("ng-select").first();
    if (!(await ng.count())) {
        // Fallback: plain input (for unregistered entities)
        const input = group.locator("input").first();
        if (await input.count()) {
            logStep("pick entity name: text input", 1);
            await input.waitFor({ state: "visible", timeout: 30000 });
            await input.scrollIntoViewIfNeeded();
            await input.click();
            await input.fill(name);
            await input.blur();
            await waitForLoaderToFinish(page);
            return;
        }
        throw new Error("Name of the Entity field not found on form");
    }

    // Build try order: last matched index first, then the rest
    const tryOrder = [];
    if (lastMatchedEntityIndex >= 0) tryOrder.push(lastMatchedEntityIndex);
    for (let i = 0; i < 15; i++) {
        if (i !== lastMatchedEntityIndex) tryOrder.push(i);
    }

    for (const attempt of tryOrder) {
        logStep(`pick entity: trying option ${attempt + 1}`, 1);

        // Clear current selection if any
        const clearBtn = ng.locator(".ng-clear-wrapper").first();
        if (await clearBtn.count()) {
            await clearBtn.click();
            await page.waitForTimeout(200);
        }

        // Open dropdown and search
        await ng.scrollIntoViewIfNeeded();
        await ng.click();
        const panel = page.locator(".ng-dropdown-panel");
        await panel.waitFor({ state: "visible", timeout: 20000 });

        const searchInput = panel.locator("input[type='text']").first();
        if (await searchInput.count()) {
            await searchInput.fill(name);
            await page.waitForTimeout(300);
        }

        const options = panel.locator(".ng-option", { hasText: name });
        const count = await options.count();
        if (count === 0) throw new Error(`No entity options found for "${name}"`);
        if (attempt >= count) continue;

        // Click the nth option
        await options.nth(attempt).click();
        await panel.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
        await waitForLoaderToFinish(page);
        await page.waitForTimeout(500);

        // Read auto-filled State and GST
        const filledState = (await getNgSelectValue(page, "State")).toLowerCase();
        const filledGst = (await page.locator('input[name="gst_no"]').first().inputValue().catch(() => "")).trim().toLowerCase();

        logStep(`option ${attempt + 1}/${count}: state="${filledState}" gst="${filledGst}"`, 2);

        const stateMatch = !expState || filledState === expState;
        const gstMatch = !expGst || filledGst === expGst;

        if (stateMatch && gstMatch) {
            lastMatchedEntityIndex = attempt;
            logStep(`entity matched on option ${attempt + 1} (cached for next row)`, 1);
            return;
        }

        logStep(`mismatch (expected state="${expState}" gst="${expGst}"), trying next`, 2);
    }

    throw new Error(`Could not match entity "${name}" with State="${expState}" GST="${expGst}"`);
}

async function waitEntityAutofill(page) {
    await waitForLoaderToFinish(page);
    const addr = page.locator('input[name="address"]').first();
    if (!(await addr.count())) return;
    await addr.waitFor({ state: "visible", timeout: 10000 }).catch(() => { });
    // Give auto-fill a chance to populate
    await page.waitForTimeout(500);
}

// ── Submit & EPR read ────────────────────────────────────────────────────────

async function clickSubmitAndConfirm(page) {
    const submit = page.locator('button[type="submit"]', { hasText: "Generate EPR Invoice Number" }).first();
    logStep("submit: start", 1);
    await submit.waitFor({ state: "visible", timeout: 20000 });
    await submit.scrollIntoViewIfNeeded();
    if (await submit.isDisabled()) {
        throw new Error("Submit disabled: some required fields still missing.");
    }
    await submit.click();

    // Wait for confirmation modal and click Confirm
    const confirmBtn = page
        .locator("#openConfirmation .modal-footer button.btn-primary", { hasText: "Confirm" })
        .first();
    await confirmBtn.waitFor({ state: "visible", timeout: 60000 });
    await confirmBtn.click();
    logStep("submit: confirmed", 1);

    // Wait for confirmation modal to disappear
    await page.waitForTimeout(1000);
    await page.locator("#openConfirmation").waitFor({ state: "hidden", timeout: 10000 }).catch(() => { });
    await waitForLoaderToFinish(page);
    logStep("submit: done", 1);
}

async function submitAndCaptureResult(page) {
    await clickSubmitAndConfirm(page);

    // Wait for toast (success or error) — quick check, don't wait long
    await page.waitForTimeout(500);
    const toastText = await readToastTextWithRetry(page, 3, 300);
    logStep(`toast: ${toastText || "(none)"}`, 1);

    // Try to extract EPR number from toast
    let eprInvoice = "";
    if (toastText) {
        const match = toastText.match(/Invoice\s*Id\s*(?:is\s*)?[:\s]*(\d+)/i);
        if (match && match[1]) eprInvoice = match[1].trim();
    }

    // Also try reading from the input field on the form
    if (!eprInvoice) {
        const selectors = [
            '#invoiceNumberCopy',
            'input[name="invoiceNumberCopy"]',
        ];
        for (const sel of selectors) {
            const input = page.locator(sel).first();
            if (await input.count()) {
                const val = (await input.inputValue().catch(() => "")).trim();
                if (val && /^\d+$/.test(val)) {
                    eprInvoice = val;
                    break;
                }
            }
        }
    }

    // Fallback: poll for EPR input to appear
    if (!eprInvoice) {
        const start = Date.now();
        while (Date.now() - start < 5000) {
            // Check all visible inputs for a numeric value that looks like an EPR number
            const inputs = page.locator('input[readonly], input[disabled]');
            const count = await inputs.count();
            for (let i = 0; i < count; i++) {
                const val = (await inputs.nth(i).inputValue().catch(() => "")).trim();
                if (val && /^\d{10,}$/.test(val)) {
                    eprInvoice = val;
                    break;
                }
            }
            if (eprInvoice) break;
            await page.waitForTimeout(500);
        }
    }

    if (eprInvoice) {
        logStep(`EPR invoice captured: ${eprInvoice}`, 1);
    }

    // Determine success: EPR captured = success, OR toast says success
    const isSuccess = !!eprInvoice || (/success/i.test(toastText) && !/error/i.test(toastText));

    // DON'T refresh here — let the caller save Excel first, then refresh
    return { eprInvoice, toastText, isSuccess };
}

async function closeAllModals(page) {
    try {
        if (page.isClosed()) return;
        logStep("refresh page to close modals", 2);
        await page.reload({ waitUntil: "domcontentloaded" }).catch(() => { });
        await page.waitForTimeout(1500).catch(() => { });
        await waitForLoaderToFinish(page).catch(() => { });
        await page.evaluate(() => {
            document.querySelectorAll(".modal-backdrop, .modal.show").forEach(el => el.remove());
            document.body.classList.remove("modal-open");
            document.body.style.removeProperty("overflow");
            document.body.style.removeProperty("padding-right");
        }).catch(() => { });
        await page.waitForSelector("button", { timeout: 10000 }).catch(() => { });
        await waitForLoaderToFinish(page).catch(() => { });
        logStep("page refreshed, ready for next row", 2);
    } catch { }
}

// ── Navigation / reset ───────────────────────────────────────────────────────

async function clickAddNew(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    logStep("click add new: start", 1);
    await addNewBtn.waitFor({ state: "visible", timeout: 60000 });
    await addNewBtn.scrollIntoViewIfNeeded();
    await addNewBtn.click();
    await page.waitForTimeout(300);
    logStep("click add new: done", 1);
}

async function clickAddNewIfVisible(page) {
    const addNewBtn = page.getByRole("button", { name: "Add New", exact: true }).first();
    if (!(await addNewBtn.count())) return false;
    await addNewBtn.waitFor({ state: "visible", timeout: 5000 }).catch(() => { });
    await addNewBtn.scrollIntoViewIfNeeded().catch(() => { });
    await addNewBtn.click().catch(() => { });
    await page.waitForTimeout(300);
    logStep("click add new (visible): done", 2);
    return true;
}

async function closeModalIfOpen(page) {
    try {
        const modal = page.locator(".modal-dialog").first();
        if (!(await modal.isVisible().catch(() => false))) return;

        const closeBtn = modal.locator("button.close, button[data-bs-dismiss='modal']").first();
        if (await closeBtn.count()) {
            await closeBtn.click().catch(() => { });
            await modal.waitFor({ state: "hidden", timeout: 5000 }).catch(() => { });
            logStep("modal closed", 2);
        }
    } catch { }
}

async function resetToFreshPage(page) {
    logStep("reset to fresh page: start", 1);
    await closeModalIfOpen(page);
    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" }).catch(() => { });
    await page.waitForTimeout(2000);
    await waitForLoaderToFinish(page);
    logStep("reset to fresh page: done", 1);
}

// ── Main ─────────────────────────────────────────────────────────────────────

(async () => {
    // Read from OUTPUT Excel if it exists (has EPR numbers from previous runs)
    // Otherwise read from INPUT Excel
    const readPath = fs.existsSync(OUTPUT_PATH) ? OUTPUT_PATH : EXCEL_PATH;
    if (!fs.existsSync(readPath)) {
        throw new Error(`Excel not found: ${readPath}`);
    }
    const stat = fs.statSync(readPath);
    if (stat.size < 1000) {
        throw new Error(`Excel looks empty/corrupt (${stat.size} bytes): ${readPath}`);
    }
    console.log(`Reading from: ${readPath}`);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(readPath);
    let ws = wb.getWorksheet(SHEET);
    if (!ws) {
        // Try trimmed matching (sheet names may have trailing spaces)
        wb.eachSheet((sheet) => {
            if (sheet.name.trim() === SHEET.trim()) ws = sheet;
        });
    }
    if (!ws) {
        throw new Error(`Sheet not found: ${SHEET}`);
    }

    const headerMap = getHeaderMap(ws);
    ensureColumn(ws, headerMap, "EPR Invoice Number");
    ensureColumn(ws, headerMap, "Status");
    const eprSet = buildEprSet(ws, headerMap);
    const headerList = getHeaderList(ws);

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext(
        fs.existsSync(STORAGE) ? { storageState: STORAGE } : {}
    );
    const page = await context.newPage();

    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" });
    if (!fs.existsSync(STORAGE) || (await isLoginPage(page))) {
        await attemptLogout(page);
        console.log("Login manually in this Playwright window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: STORAGE });
        console.log("Saved session to storageState_pibo.json");
    }

    await page.goto(PIBO_URL, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(3000);
    await waitForLoaderToFinish(page);

    const lastRow = CONFIG.maxRows ? Math.min(ws.rowCount, CONFIG.maxRows) : ws.rowCount;
    for (let r = 2; r <= lastRow; r++) {
        const row = ws.getRow(r);
        let successThisRow = false;
        let pageRefreshed = false;
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

        // Read fields from Excel (actual column names)
        const regType = getVal(row, headerMap, "REGISTRATION TYPE");
        const entityType = getVal(row, headerMap, "ENTITY TYPE");
        const entityName = getVal(row, headerMap, "ENTITY");
        const gst = getVal(row, headerMap, "GST NO");
        const state = getVal(row, headerMap, "STATE");
        const plasticType = getVal(row, headerMap, "PLASTIC TYPE");
        const category = getVal(row, headerMap, "CATEGORY");
        const recycledContent = getVal(row, headerMap, "RECYCLED CONTENT %");
        const quantity = getVal(row, headerMap, "Sum of QTY MT");
        const gstPaid = getVal(row, headerMap, "Sum of GST PAID");
        const invoiceNo = getVal(row, headerMap, "INVOICE No");
        const bankAccount = getVal(row, headerMap, "BANK ACCOUNT NO");
        const ifsc = getVal(row, headerMap, "IFSC");

        // Fields not in Excel
        const financialYear = "2025-26";
        const addressDefault = CONFIG.address;
        const mobileDefault = CONFIG.mobileNumber;

        try {
            console.log(`Row ${r} starting...`);

            // Open form modal
            await clickAddNew(page);
            await page.waitForTimeout(1000);
            await waitForLoaderToFinish(page);
            // Wait for Registration Type dropdown to be fully visible and stable
            const regSelect = page.locator('ng-select[name="registration_type"]').first();
            await regSelect.waitFor({ state: "visible", timeout: 20000 });
            await regSelect.waitFor({ state: "attached", timeout: 5000 }).catch(() => { });
            await page.waitForTimeout(500);

            // 1. Registration Type (dropdown)
            logStep("fill Registration Type", 1);
            await selectNgSelectByLabel(page, "Registration Type", regType);
            await page.waitForTimeout(300);

            // 2. Type — click "Entity Details" radio button
            logStep("select Type: Entity Details", 1);
            const entityDetailsRadio = page.locator('label', { hasText: "Entity Details" }).first();
            if (await entityDetailsRadio.count()) {
                await entityDetailsRadio.click();
                await page.waitForTimeout(300);
            } else {
                // Fallback: click the radio input directly
                const radio = page.locator('input[type="radio"]').nth(1);
                if (await radio.count()) await radio.click({ force: true });
                await page.waitForTimeout(300);
            }
            await waitForLoaderToFinish(page);

            // 3. Entity Type (always fill)
            logStep("fill Entity Type", 1);
            await selectNgSelectByLabel(page, "Entity Type", entityType);
            await page.waitForTimeout(500);
            await waitForLoaderToFinish(page);

            // 3. Name of the Entity — iterate options until State + GST match Excel
            logStep("fill Name of the Entity (matching State + GST)", 1);
            await pickEntityWithStateMatch(page, entityName, state, gst);
            await waitEntityAutofill(page);

            // ── Pre-filled fields: only fill if empty ──

            // 4. Address (pre-filled from entity)
            logStep("check Address", 1);
            await fillByNameIfEmpty(page, "address", addressDefault);

            // 5. Mobile Number (pre-filled from entity)
            logStep("check Mobile Number", 1);
            await fillByNameIfEmpty(page, "mobile_number", mobileDefault);

            // ── Always fill from Excel ──

            // 6. Plastic Material Type
            logStep("fill Plastic Material Type", 1);
            await selectNgSelectByLabel(page, "Plastic Material Type", plasticType);
            await page.waitForTimeout(300);

            // 7. Category of Plastic (CAT 2 → CAT II mapping)
            logStep("fill Category of Plastic", 1);
            await selectNgSelectByLabel(page, "Category of Plastic", mapCategory(category));

            // 8. Financial Year (from config)
            logStep("fill Financial Year", 1);
            await selectNgSelectByLabel(page, "Financial Year", financialYear);

            // 9. GST (already matched during entity selection, verify)
            logStep("verify GST", 1);
            await verifyAndFixInput(page, "gst_no", gst);

            // 10. Bank Account No
            logStep("fill Bank Account No", 1);
            await fillByNameIfEmpty(page, "account_no", bankAccount);

            // 12. IFSC Code
            logStep("fill IFSC Code", 1);
            await fillByNameIfEmpty(page, "ifsc", ifsc);

            // 13. GST Paid / Total GST Paid
            logStep("fill GST Paid", 1);
            await fillByName(page, "gst_paid", gstPaid);

            // 14. GST E-Invoice Number
            logStep("fill GST E-Invoice Number", 1);
            await fillByName(page, "gst_invoice", invoiceNo);

            // 15. Total Plastic Quantity (Tons)
            logStep("fill Quantity", 1);
            await fillByName(page, "quantity", formatQty(quantity));

            // 16. % of Recycled Plastic Content
            logStep("fill Recycled Plastic %", 1);
            await fillByName(page, "recycled_plastic", recycledContent);

            // Submit, confirm, read toast + EPR number
            const result = await submitAndCaptureResult(page);

            // SAVE TO EXCEL IMMEDIATELY (before page refresh which might crash)
            if (result.isSuccess && result.eprInvoice) {
                if (eprSet.has(result.eprInvoice)) {
                    throw new Error("Duplicate EPR Invoice Number: " + result.eprInvoice);
                }
                setVal(row, headerMap, "Status", "Filled");
                setVal(row, headerMap, "EPR Invoice Number", result.eprInvoice);
                eprSet.add(result.eprInvoice);
                console.log(`Row ${r}: Filled (EPR: ${result.eprInvoice})`);
                successThisRow = true;
            } else if (result.isSuccess && !result.eprInvoice) {
                setVal(row, headerMap, "Status", "Success but EPR not captured");
                console.log(`Row ${r}: Success but EPR number not captured`);
                successThisRow = true;
            } else {
                const errMsg = result.toastText || "Unknown error";
                setVal(row, headerMap, "Status", "Failed: " + errMsg);
                console.log(`Row ${r}: Failed -> ${errMsg}`);
            }

            // Save Excel + logs BEFORE refreshing page
            row.commit();
            await safeWriteWorkbook(wb);
            await syncInputWorkbook(wb);
            appendLogRow(row, headerMap, {
                status: successThisRow ? "Filled" : "Failed",
                eprInvoiceNumber: result.eprInvoice || "",
                message: result.toastText || "",
            });
            appendFilledRow(row, headerMap, headerList, {
                message: result.toastText || "",
            });

            // NOW refresh the page (safe to crash here, Excel is already saved)
            pageRefreshed = true;
            await closeAllModals(page);
        } catch (e) {
            const msg = String(e?.message || e);
            console.log(`Row ${r}: Failed ->`, msg);
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
            try {
                if (page.isClosed()) {
                    console.log("Page closed. Stopping.");
                    break;
                }
                if (!pageRefreshed) {
                    await closeAllModals(page);
                }
            } catch {
                console.log("Page closed during cleanup. Stopping.");
                break;
            }
        }
    }

    await browser.close();
    console.log("Done. Updated Excel:", EXCEL_PATH);
})();
