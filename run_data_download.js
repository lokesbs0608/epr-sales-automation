/**
 * run_data_download.js
 * Extracts paginated sales records from the CPCB EPR portal into Excel.
 * API path: json.data.tableData.bodyContent
 */

const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

function waitForEnter() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question('', () => { rl.close(); resolve(); });
  });
}

const CONFIG_PATH = path.join(__dirname, 'config_data_download.json');
const CONFIG = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));

const storageStatePath   = path.join(__dirname, CONFIG.storageState);
const checkpointFilePath = path.join(__dirname, CONFIG.checkpointFile);
const outputExcelPath    = path.join(__dirname, CONFIG.outputExcel);
const maxPages           = CONFIG.max_pages && CONFIG.max_pages > 0 ? CONFIG.max_pages : 500;

// ============================================================
// CHECKPOINT
// ============================================================
function loadCheckpoint() {
  if (fs.existsSync(checkpointFilePath)) {
    try { return JSON.parse(fs.readFileSync(checkpointFilePath, 'utf-8')); }
    catch (e) { console.error('Checkpoint corrupted, starting fresh:', e.message); }
  }
  return { records: [], capturedSalesIds: [], lastPage: 0 };
}

function saveCheckpoint(cp) {
  fs.writeFileSync(checkpointFilePath, JSON.stringify(cp, null, 2));
}

// ============================================================
// WAIT FOR NEW BATCH — infinite wait with heartbeat every 15s
// prevVer: the batchVersion value before clicking Next
// Returns when batchVersion changes (i.e. new API response arrived)
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
// EXTRACT — hardcoded to confirmed API shape
//   json.data.tableData.bodyContent  → array of sale records
//   json.data.total_no               → total count
//   json.data.endOfRecords           → boolean last-page flag
// ============================================================
function extractBatch(json) {
  try {
    const rows = json.data.tableData.bodyContent;
    if (Array.isArray(rows)) return rows;
  } catch (_) {}
  console.warn('  [intercept] WARNING: expected path json.data.tableData.bodyContent not found');
  console.warn('  [intercept] Top-level keys:', Object.keys(json));
  if (json.data) console.warn('  [intercept] json.data keys:', Object.keys(json.data));
  return [];
}

function extractTotal(json) {
  try { return json.data.total_no || null; } catch (_) { return null; }
}

function isEndOfRecords(json) {
  try { return json.data.endOfRecords === true; } catch (_) { return false; }
}

// ============================================================
// STATUS LABEL HELPERS
// is_potential_generated: 0 = invoice uploaded (green tick)
//                         1 = potential generated (amber info)
// ============================================================
function invoiceFileStatus(rec) {
  // The invoice_file URL being present means the file was uploaded
  return rec.invoice_file ? 'Uploaded' : 'Not Uploaded';
}

function potentialGenerationStatus(rec) {
  if (rec.is_potential_generated === 1) return 'Generated';
  if (rec.is_potential_generated === 0) return 'Pending';
  return '';
}

// ============================================================
// EXCEL WRITER
// ============================================================
async function writeExcel(records, outputFile) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sales Data');

  sheet.columns = [
    { header: 'Sr.No',                          key: 'srNo',                                    width: 8  },
    { header: 'Sales ID',                        key: 'sales_id',                                width: 14 },
    { header: 'Seller GST',                      key: 'sellerGst',                               width: 22 },
    { header: 'Invoice No',                      key: 'invoice',                                 width: 22 },
    { header: 'Date of Sale',                    key: 'date',                                    width: 14 },
    { header: 'Amount (Rs.)',                    key: 'amount',                                  width: 16 },
    { header: 'Register Type',                   key: 'register_type',                           width: 14 },
    { header: 'Name of the Entity',              key: 'company_name',                            width: 45 },
    { header: 'Address',                         key: 'address',                                 width: 40 },
    { header: 'District',                        key: 'district',                                width: 16 },
    { header: 'State',                           key: 'state',                                   width: 16 },
    { header: 'Total Qty Product Sold (MT)',     key: 'total_qty_of_product_sold',               width: 22 },
    { header: 'Total Potential Generated (MT)',  key: 'total_potential_generated',               width: 24 },
    { header: 'Invoice File URL',                key: 'invoice_file_url',                        width: 30 },
    { header: 'Invoice File Status',             key: 'invoice_file_status',                     width: 20 },
    { header: 'Potential Generation Status',     key: 'potential_generation_status',             width: 24 },
    { header: 'Category',                        key: 'category',                                width: 14 },
    { header: 'Process Code',                    key: 'processCode',                             width: 14 },
    { header: 'Plastic Type',                    key: 'plasticType',                             width: 16 },
    { header: 'Product Type',                    key: 'productType',                             width: 16 },
    { header: '% Recycled Plastic',              key: 'recycled_plastic_percentage_in_product',  width: 18 },
    { header: 'Qty Product Sold (MT)',           key: 'qty_product_sold',                        width: 20 },
    { header: 'Potential Generated (MT)',        key: 'potential_generated',                     width: 22 },
  ];

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
    const products = (rec.product_sold && rec.product_sold.bodyContent) || [];
    const base = {
      sales_id:                   rec.sales_id,
      sellerGst:                  rec.sellerGst,
      invoice:                    rec.invoice,
      date:                       rec.date,
      amount:                     rec.amount,
      register_type:              rec.register_type,
      company_name:               rec.company_name,
      address:                    rec.address,
      district:                   rec.district,
      state:                      rec.state,
      total_qty_of_product_sold:  rec.total_qty_of_product_sold,
      total_potential_generated:  rec.total_potential_generated,
      invoice_file_url:           rec.invoice_file || '',
      invoice_file_status:        invoiceFileStatus(rec),
      potential_generation_status: potentialGenerationStatus(rec),
    };

    if (products.length === 0) {
      sheet.addRow({ srNo: srNo++, ...base });
    } else {
      for (const p of products) {
        sheet.addRow({
          srNo: srNo++, ...base,
          category:                               p.category,
          processCode:                            p.processCode,
          plasticType:                            p.plasticType,
          productType:                            p.productType,
          recycled_plastic_percentage_in_product: p.recycled_plastic_percentage_in_product,
          qty_product_sold:                       p.qty_product_sold,
          potential_generated:                    p.potential_generated,
        });
      }
    }
  }

  await workbook.xlsx.writeFile(outputFile);
  console.log(`\nExcel written: ${outputFile}`);
  console.log(`  Sales records : ${records.length}`);
  console.log(`  Total rows    : ${srNo - 1}`);
}

// ============================================================
// SESSION HELPERS  (matches run_registered_1.js pattern)
// ============================================================
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
      await logoutDirect.click().catch(() => {});
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
        await btn.click().catch(() => {});
        await page.waitForTimeout(500);
        const logout = page.locator('text=/logout/i').first();
        if (await logout.count()) {
          await logout.click().catch(() => {});
          await page.waitForTimeout(1000);
          return true;
        }
      }
    }
  } catch (_) {}
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
// MAIN
// ============================================================
(async () => {
  if (!fs.existsSync(storageStatePath)) {
    console.error(`storageState not found: ${storageStatePath}`);
    process.exit(1);
  }

  const checkpoint = loadCheckpoint();
  const capturedSalesIds = new Set(checkpoint.capturedSalesIds);
  console.log(`Resuming: ${checkpoint.records.length} records already captured (last page: ${checkpoint.lastPage}).`);

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: storageStatePath });
  const page    = await context.newPage();

  // ---- intercept listSalesReceipt ----
  let latestBatch  = null;
  let latestJson   = null;
  let batchVersion = 0;

  page.on('response', async (response) => {
    if (!response.url().includes('listSalesReceipt')) return;
    if (response.request().method() !== 'POST') return;
    try {
      const json = await response.json();
      const rows = extractBatch(json);
      latestBatch  = rows;
      latestJson   = json;
      batchVersion++;
      console.log(`  [intercept] ${rows.length} records | endOfRecords: ${isEndOfRecords(json)}`);
    } catch (e) {
      console.error('Failed to parse listSalesReceipt response:', e.message);
    }
  });

  await page.goto(CONFIG.portalUrl, { waitUntil: 'domcontentloaded' });

  // ---- session check: if not logged in, ask user to login then save state ----
  if (!fs.existsSync(storageStatePath) || (await isLoginPage(page))) {
    await attemptLogout(page);
    console.log('');
    console.log('╔══════════════════════════════════════════════════════╗');
    console.log('║          SESSION EXPIRED — ACTION REQUIRED           ║');
    console.log('║                                                      ║');
    console.log('║  1. Log in manually in the browser window.           ║');
    console.log('║  2. Press ENTER here once you are logged in.         ║');
    console.log('╚══════════════════════════════════════════════════════╝');
    console.log('');
    await waitForEnter();
    await context.storageState({ path: storageStatePath });
    console.log(`Session saved to ${storageStatePath}`);
    // reload to the portal URL now that session is saved
    await page.goto(CONFIG.portalUrl, { waitUntil: 'domcontentloaded' });
  }

  // ---- ask user to search and load first page ----
  console.log('');
  console.log('╔══════════════════════════════════════════════════════╗');
  console.log('║        MANUAL STEPS — do these in the browser        ║');
  console.log('║                                                      ║');
  console.log('║  1. Navigate to the Sales Details page.              ║');
  console.log('║  2. Enter From Date and To Date.                     ║');
  console.log('║  3. Click Fetch / Search.                            ║');
  console.log('║  4. Wait until the first 25 rows are visible.        ║');
  console.log('║  5. Come back here and press ENTER.                  ║');
  console.log('╚══════════════════════════════════════════════════════╝');
  console.log('');
  console.log('Press ENTER when the first page of results is showing...');
  await waitForEnter();
  // Save updated session state (cookies/tokens may have refreshed during use)
  await context.storageState({ path: storageStatePath });
  console.log(`Session saved to ${storageStatePath}`);

  // Wait for first batch (infinite — heartbeat every 15s)
  if (!latestBatch || latestBatch.length === 0) {
    console.log('Waiting for page 1 API response...');
    await waitForNextBatch(() => batchVersion, -1, 1);
  }

  const total = extractTotal(latestJson);
  if (total) {
    console.log(`\nTotal records: ${total} (~${Math.ceil(total / 25)} pages)\n`);
  }

  // ---- fast-forward past already-captured pages ----
  if (checkpoint.lastPage > 0) {
    console.log(`Fast-forwarding past ${checkpoint.lastPage} already-captured page(s)...`);
    for (let i = 0; i < checkpoint.lastPage; i++) {
      const prevVer = batchVersion;
      const ok = await clickNextPage(page);
      if (!ok) { console.error('Could not click Next during fast-forward.'); await browser.close(); process.exit(1); }
      await waitForNextBatch(() => batchVersion, prevVer, i + 2);
    }
    console.log(`Fast-forward complete. Resuming from page ${checkpoint.lastPage + 1}.\n`);
  }

  // ---- main pagination loop ----
  let consecutiveNoNew = 0;
  const startPage = checkpoint.lastPage > 0 ? checkpoint.lastPage + 1 : 1;

  for (let pageNum = startPage; pageNum <= maxPages; pageNum++) {

    if (pageNum !== startPage) {
      const prevVer = batchVersion;
      const ok = await clickNextPage(page);
      if (!ok) {
        console.log(`\nNext button not found/disabled on page ${pageNum}. End of dataset.`);
        break;
      }
      // Infinite wait — no timeout, heartbeat every 15s
      await waitForNextBatch(() => batchVersion, prevVer, pageNum);
    }

    const batch = latestBatch || [];
    const newRecords = batch.filter(r => {
      const id = String(r.sales_id || r.invoice || r.id || '');
      return id && !capturedSalesIds.has(id);
    });

    if (newRecords.length === 0) {
      consecutiveNoNew++;
      console.log(`Page ${pageNum}: 0 new records.`);
      if (consecutiveNoNew >= 2) { console.log('2 consecutive empty pages — stopping.'); break; }
    } else {
      consecutiveNoNew = 0;
      for (const r of newRecords) {
        capturedSalesIds.add(String(r.sales_id || r.invoice || r.id || ''));
      }
      checkpoint.records.push(...newRecords);
      checkpoint.capturedSalesIds = Array.from(capturedSalesIds);
      checkpoint.lastPage = pageNum;
      saveCheckpoint(checkpoint);

      const pct = total ? ` (${Math.round((checkpoint.records.length / total) * 100)}%)` : '';
      console.log(`Page ${pageNum}: +${newRecords.length} → total: ${checkpoint.records.length}${pct}`);
    }

    if (isEndOfRecords(latestJson)) {
      console.log(`\nAPI signals endOfRecords=true on page ${pageNum}. Done.`);
      break;
    }

    if (batch.length < 25) {
      console.log(`\nLast page: only ${batch.length} rows on page ${pageNum}.`);
      break;
    }
  }

  await browser.close();

  if (checkpoint.records.length === 0) {
    console.log('\nNo records captured. Excel not written.');
  } else {
    await writeExcel(checkpoint.records, outputExcelPath);
    if (fs.existsSync(checkpointFilePath)) fs.unlinkSync(checkpointFilePath);
    console.log('Checkpoint deleted (run complete).');
  }

  console.log('\nDone.');

})().catch(err => { console.error('Fatal error:', err); process.exit(1); });