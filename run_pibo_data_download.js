/**
 * run_pibo_data_download.js
 * PIBO data extraction from the CPCB EPR portal into Excel.
 *
 * Flow (all in ONE browser session):
 *   PHASE 1        : Material Procurement Details (multi-year range in one go)
 *   PHASE 2..N     : Sales Details, one financial year at a time
 *
 * For each phase you set the dates + click Search manually in the browser,
 * then press ENTER here — the script paginates, captures the API responses,
 * and writes one Excel file per phase.
 *
 * API responses are detected by SHAPE (json.data.tableData.bodyContent),
 * not by endpoint name, so the same code works for both material and sales.
 * Excel columns are built dynamically from the fields the API returns.
 */

const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

function ask(question) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(question, (answer) => { rl.close(); resolve(answer.trim()); });
  });
}

const CONFIG_PATH = path.join(__dirname, 'config_pibo_data_download.json');
const CONFIG = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));

const storageStatePath = path.join(__dirname, CONFIG.storageState);
const maxPages = CONFIG.max_pages && CONFIG.max_pages > 0 ? CONFIG.max_pages : 500;

// ============================================================
// CHECKPOINT (one file per phase)
// ============================================================
function loadCheckpoint(checkpointFilePath) {
  if (fs.existsSync(checkpointFilePath)) {
    try { return JSON.parse(fs.readFileSync(checkpointFilePath, 'utf-8')); }
    catch (e) { console.error('Checkpoint corrupted, starting fresh:', e.message); }
  }
  return { records: [], capturedIds: [], lastPage: 0 };
}

function saveCheckpoint(checkpointFilePath, cp) {
  fs.writeFileSync(checkpointFilePath, JSON.stringify(cp, null, 2));
}

// ============================================================
// RECORD IDENTITY — try known id fields, fall back to a hash
// ============================================================
function recordId(rec) {
  const known = rec.sales_id || rec.material_id || rec.procurement_id || rec.id || rec.invoice;
  if (known !== undefined && known !== null && known !== '') return String(known);
  return crypto.createHash('md5').update(JSON.stringify(rec)).digest('hex');
}

// ============================================================
// RESPONSE SHAPE HELPERS
//   json.data.tableData.bodyContent → array of records
//   json.data.total_no              → total count
//   json.data.endOfRecords          → boolean last-page flag
// ============================================================
function extractBatch(json) {
  try {
    const rows = json.data.tableData.bodyContent;
    if (Array.isArray(rows)) return rows;
  } catch (_) {}
  return null;
}

function extractTotal(json) {
  try { return json.data.total_no || null; } catch (_) { return null; }
}

function isEndOfRecords(json) {
  try { return json.data.endOfRecords === true; } catch (_) { return false; }
}

// ============================================================
// WAIT FOR NEW BATCH — infinite wait with heartbeat every 15s
// ============================================================
async function waitForNextBatch(getVersion, prevVer, pageNum) {
  const heartbeatMs = 15000;
  let elapsed = 0;
  while (getVersion() === prevVer) {
    await new Promise(r => setTimeout(r, 150));
    elapsed += 150;
    if (elapsed % heartbeatMs < 150) {
      const secs = Math.round(elapsed / 1000);
      console.log(`  [waiting] Page ${pageNum}: still waiting for API response... (${secs}s elapsed)`);
    }
  }
}

// ============================================================
// DYNAMIC EXCEL WRITER
// Columns are the union of fields across all records.
// Nested tables (e.g. product_sold.bodyContent) are exploded
// into one row per item, child columns prefixed with the key.
// ============================================================
function humanize(key) {
  return key
    .replace(/_/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/\b\w/g, c => c.toUpperCase());
}

function isNestedTable(v) {
  return v && typeof v === 'object' && Array.isArray(v.bodyContent);
}

function cellValue(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return v;
}

async function writeExcel(records, outputFile, sheetName) {
  // ---- discover columns ----
  const scalarKeys = [];
  const nestedKeys = [];
  const childKeys = {};
  for (const rec of records) {
    for (const [k, v] of Object.entries(rec)) {
      if (isNestedTable(v)) {
        if (!nestedKeys.includes(k)) { nestedKeys.push(k); childKeys[k] = []; }
        for (const item of v.bodyContent) {
          if (!item || typeof item !== 'object') continue;
          for (const ck of Object.keys(item)) {
            if (!childKeys[k].includes(ck)) childKeys[k].push(ck);
          }
        }
      } else if (!scalarKeys.includes(k)) {
        scalarKeys.push(k);
      }
    }
  }

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName.substring(0, 31));

  const columns = [{ header: 'Sr.No', key: '__srNo', width: 8 }];
  for (const k of scalarKeys) {
    columns.push({ header: humanize(k), key: k, width: Math.min(Math.max(k.length + 6, 14), 45) });
  }
  for (const nk of nestedKeys) {
    for (const ck of childKeys[nk]) {
      columns.push({ header: `${humanize(ck)}`, key: `${nk}.${ck}`, width: Math.min(Math.max(ck.length + 6, 14), 45) });
    }
  }
  sheet.columns = columns;

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF006400' } };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' },
      bottom: { style: 'thin' }, right: { style: 'thin' },
    };
  });
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  let srNo = 1;
  for (const rec of records) {
    const base = { };
    for (const k of scalarKeys) base[k] = cellValue(rec[k]);

    // How many rows this record explodes into = longest nested table (min 1)
    let rowCount = 1;
    for (const nk of nestedKeys) {
      if (isNestedTable(rec[nk])) rowCount = Math.max(rowCount, rec[nk].bodyContent.length);
    }

    for (let i = 0; i < rowCount; i++) {
      const row = { __srNo: srNo++, ...base };
      for (const nk of nestedKeys) {
        const items = isNestedTable(rec[nk]) ? rec[nk].bodyContent : [];
        const item = items[i];
        if (item && typeof item === 'object') {
          for (const ck of childKeys[nk]) row[`${nk}.${ck}`] = cellValue(item[ck]);
        }
      }
      sheet.addRow(row);
    }
  }

  await workbook.xlsx.writeFile(outputFile);
  console.log(`\n  ✔ Excel written: ${path.basename(outputFile)}`);
  console.log(`      Records : ${records.length}`);
  console.log(`      Rows    : ${srNo - 1}`);
}

// ============================================================
// SESSION HELPERS
// ============================================================
async function isLoginPage(page) {
  const pwd = page.locator('input[type="password"]').first();
  if (await pwd.count()) return true;
  const loginBtn = page.locator('button:has-text("Login"), button:has-text("Sign In")').first();
  if (await loginBtn.count()) return true;
  return false;
}

// ============================================================
// CLICK NEXT PAGE
// ============================================================
async function clickNextPage(page) {
  const selectors = [
    'button:has-text("Next")',
    '.table-footer button:has-text("Next")',
  ];
  for (const sel of selectors) {
    try {
      const btn = page.locator(sel).last();
      if (await btn.isVisible().catch(() => false) && !await btn.isDisabled().catch(() => true)) {
        await btn.click();
        return true;
      }
    } catch (_) {}
  }
  return await page.evaluate(() => {
    const btn = Array.from(document.querySelectorAll('button'))
      .find(b => (b.textContent || '').trim() === 'Next' && !b.disabled);
    if (btn) { btn.click(); return true; }
    return false;
  });
}

// ============================================================
// BANNER HELPERS
// ============================================================
function banner(lines) {
  const width = Math.max(...lines.map(l => l.length)) + 4;
  console.log('');
  console.log('╔' + '═'.repeat(width) + '╗');
  for (const l of lines) {
    console.log('║  ' + l.padEnd(width - 2) + '║');
  }
  console.log('╚' + '═'.repeat(width) + '╝');
  console.log('');
}

// ============================================================
// RUN ONE PHASE (procurement OR one sales year)
// ============================================================
async function runPhase(page, intercept, phase, phaseNum, phaseTotal) {
  const checkpointFilePath = path.join(__dirname, phase.checkpointFile);
  const outputExcelPath    = path.join(__dirname, phase.outputExcel);

  banner([
    `PHASE ${phaseNum} of ${phaseTotal}: ${phase.label}`,
    `Date range : ${phase.dateHint}`,
    `Output file: ${phase.outputExcel}`,
  ]);

  // Already done? Offer skip.
  if (fs.existsSync(outputExcelPath) && !fs.existsSync(checkpointFilePath)) {
    const ans = await ask(`  "${phase.outputExcel}" already exists. Skip this phase? (Y/n): `);
    if (ans === '' || ans.toLowerCase().startsWith('y')) {
      console.log(`  Skipped ${phase.label}.\n`);
      return;
    }
  }

  const checkpoint = loadCheckpoint(checkpointFilePath);
  const capturedIds = new Set(checkpoint.capturedIds);
  if (checkpoint.records.length > 0) {
    console.log(`  Resuming: ${checkpoint.records.length} records already captured (last page: ${checkpoint.lastPage}).`);
  }

  await page.goto(phase.portalUrl, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(2000);

  banner([
    'MANUAL STEPS — do these in the browser:',
    '',
    `1. Make sure you are on: ${phase.label}`,
    `2. Enter From Date and To Date:`,
    `      ${phase.dateHint}`,
    '3. Click Fetch / Search.',
    '4. Wait until the first rows are visible.',
    '5. Come back here and press ENTER.',
    '',
    'No data for this period? Type S + ENTER to skip.',
  ]);
  const answer = await ask('  Press ENTER when the first page of results is showing (or S to skip)... ');
  if (answer.toLowerCase().startsWith('s')) {
    console.log(`  Skipped ${phase.label}.\n`);
    return;
  }

  // Wait for first batch of THIS phase
  if (intercept.version() === intercept.phaseBaseline) {
    console.log('  Waiting for page 1 API response...');
    await waitForNextBatch(intercept.version, intercept.phaseBaseline, 1);
  }

  const total = extractTotal(intercept.latestJson());
  if (total) {
    console.log(`\n  Total records: ${total} (~${Math.ceil(total / 25)} pages)\n`);
  }

  // ---- fast-forward past already-captured pages (resume) ----
  if (checkpoint.lastPage > 0) {
    console.log(`  Fast-forwarding past ${checkpoint.lastPage} already-captured page(s)...`);
    for (let i = 0; i < checkpoint.lastPage; i++) {
      const prevVer = intercept.version();
      const ok = await clickNextPage(page);
      if (!ok) { console.log('  Could not click Next during fast-forward — continuing from current page.'); break; }
      await waitForNextBatch(intercept.version, prevVer, i + 2);
    }
    console.log(`  Fast-forward complete. Resuming from page ${checkpoint.lastPage + 1}.\n`);
  }

  // ---- pagination loop ----
  let consecutiveNoNew = 0;
  const startPage = checkpoint.lastPage > 0 ? checkpoint.lastPage + 1 : 1;

  for (let pageNum = startPage; pageNum <= maxPages; pageNum++) {

    if (pageNum !== startPage) {
      const prevVer = intercept.version();
      const ok = await clickNextPage(page);
      if (!ok) {
        console.log(`\n  Next button not found/disabled on page ${pageNum}. End of dataset.`);
        break;
      }
      await waitForNextBatch(intercept.version, prevVer, pageNum);
    }

    const batch = intercept.latestBatch() || [];
    const newRecords = batch.filter(r => !capturedIds.has(recordId(r)));

    if (newRecords.length === 0) {
      consecutiveNoNew++;
      console.log(`  Page ${pageNum}: 0 new records.`);
      if (consecutiveNoNew >= 2) { console.log('  2 consecutive empty pages — stopping.'); break; }
    } else {
      consecutiveNoNew = 0;
      for (const r of newRecords) capturedIds.add(recordId(r));
      checkpoint.records.push(...newRecords);
      checkpoint.capturedIds = Array.from(capturedIds);
      checkpoint.lastPage = pageNum;
      saveCheckpoint(checkpointFilePath, checkpoint);

      const pct = total ? ` (${Math.round((checkpoint.records.length / total) * 100)}%)` : '';
      console.log(`  Page ${pageNum}: +${newRecords.length} → total: ${checkpoint.records.length}${pct}`);
    }

    if (isEndOfRecords(intercept.latestJson())) {
      console.log(`\n  API signals endOfRecords=true on page ${pageNum}. Done.`);
      break;
    }

    if (batch.length < 25) {
      console.log(`\n  Last page: only ${batch.length} rows on page ${pageNum}.`);
      break;
    }
  }

  if (checkpoint.records.length === 0) {
    console.log('\n  No records captured for this phase. Excel not written.');
  } else {
    await writeExcel(checkpoint.records, outputExcelPath, phase.label);
    if (fs.existsSync(checkpointFilePath)) fs.unlinkSync(checkpointFilePath);
  }
}

// ============================================================
// MAIN
// ============================================================
(async () => {
  // Build the phase list: procurement first, then sales year by year.
  const phases = [
    { ...CONFIG.procurement },
    ...CONFIG.salesYears.map(y => ({ ...y, portalUrl: CONFIG.salesPortalUrl })),
  ];

  banner([
    'PIBO DATA DOWNLOAD — CPCB EPR PORTAL',
    '',
    `Phases to run (${phases.length} total):`,
    ...phases.map((p, i) => `  ${i + 1}. ${p.label}  [${p.dateHint}]`),
  ]);

  const hasState = fs.existsSync(storageStatePath);
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext(hasState ? { storageState: storageStatePath } : {});
  const page    = await context.newPage();

  // ---- shape-based response interception (works for material AND sales) ----
  let latestBatch  = null;
  let latestJson   = null;
  let batchVersion = 0;

  page.on('response', async (response) => {
    if (response.request().method() !== 'POST') return;
    let json;
    try { json = await response.json(); } catch (_) { return; }
    const rows = extractBatch(json);
    if (rows === null) return;
    latestBatch  = rows;
    latestJson   = json;
    batchVersion++;
    const endpoint = response.url().split('/').pop().split('?')[0];
    console.log(`  [intercept] ${rows.length} records from "${endpoint}" | endOfRecords: ${isEndOfRecords(json)}`);
  });

  const intercept = {
    version:     () => batchVersion,
    latestBatch: () => latestBatch,
    latestJson:  () => latestJson,
    phaseBaseline: 0,
  };

  // ---- session check ----
  await page.goto(phases[0].portalUrl, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(2000);

  if (!hasState || (await isLoginPage(page))) {
    banner([
      'LOGIN REQUIRED — ACTION NEEDED',
      '',
      '1. Log in to the PIBO account in the browser window.',
      '2. Press ENTER here once you are logged in.',
    ]);
    await ask('  Press ENTER after logging in... ');
    await context.storageState({ path: storageStatePath });
    console.log(`  Session saved to ${CONFIG.storageState}`);
  }

  // ---- run all phases in order ----
  for (let i = 0; i < phases.length; i++) {
    intercept.phaseBaseline = batchVersion;   // ignore batches from previous phase
    latestBatch = null;
    latestJson  = null;
    await runPhase(page, intercept, phases[i], i + 1, phases.length);
    // Keep session fresh between phases
    await context.storageState({ path: storageStatePath });
  }

  await browser.close();

  banner([
    'ALL PHASES COMPLETE',
    '',
    ...phases.map(p => `  ${fs.existsSync(path.join(__dirname, p.outputExcel)) ? '✔' : '✘'} ${p.outputExcel}`),
  ]);

})().catch(err => { console.error('Fatal error:', err); process.exit(1); });
