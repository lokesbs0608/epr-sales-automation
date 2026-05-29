"use strict";

// ---------------------------------------------------------------------------
// orchestrate.js
// Dynamic chunking + parallel runner that sits ON TOP of the existing workers
// (run_registered.js / run_unregistered.js / pibo_data.js). The workers are
// unchanged — they still take `--config X.json` and process one sheet.
//
// What this does:
//   1. Reads ONE master config that lists jobs (e.g. PP + PET).
//   2. For each job: counts data rows in the sheet.
//        - rows <= chunkSize  -> 1 chunk (whole sheet)
//        - rows >  chunkSize  -> split into chunk files of <chunkSize> rows
//   3. Generates a worker config PER CHUNK from the job/base settings
//      (overriding only inputExcel / outputExcel / sheetName).
//   4. Logs in ONCE (headed) and saves the shared storageState so every
//      spawned worker reuses the SAME login (no stdin prompts in children).
//   5. Spawns workers as child processes, capped at maxWindowsPerJob PER JOB.
//      Different jobs (PP, PET) run at the same time -> their pools overlap.
//   6. Merges each job's chunk outputs + logs back into combined files.
//
// Usage:
//   node orchestrate.js --config orchestrate_polyplex.json
//   node orchestrate.js --config X.json --chunk-size 500 --max-windows 4
//   node orchestrate.js --config X.json --fresh        (rebuild chunk files)
//   node orchestrate.js --config X.json --no-merge      (skip final merge)
//   node orchestrate.js --config X.json --skip-login    (trust existing state)
// ---------------------------------------------------------------------------

const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");
const { spawn } = require("child_process");

// ---------- CLI ----------
function argVal(flag, def = null) {
    const i = process.argv.indexOf(flag);
    return i !== -1 && process.argv[i + 1] ? process.argv[i + 1] : def;
}
function argFlag(flag) {
    return process.argv.includes(flag);
}

// --config can be repeated. Each value is a NORMAL worker config (the same
// config_registered.json / config_unregistered.json you already use) — there is
// NO special "master" file. The orchestrator reads each config, auto-detects
// the sheet, counts rows, and chunks dynamically.
function argValsMulti(flag) {
    const out = [];
    for (let i = 0; i < process.argv.length - 1; i++) {
        if (process.argv[i] === flag) out.push(process.argv[i + 1]);
    }
    return out;
}

const CONFIG_PATHS = argValsMulti("--config");
const CLI_CHUNK_SIZE = argVal("--chunk-size");
const CLI_MAX_WINDOWS = argVal("--max-windows");
const CLI_WORKER = argVal("--worker"); // override worker script for all configs
const CLI_STORAGE = argVal("--storage"); // override shared storage-state file
const FRESH = argFlag("--fresh");
const NO_MERGE = argFlag("--no-merge");
const SKIP_LOGIN = argFlag("--skip-login");
const PREPARE_ONLY = argFlag("--prepare-only"); // chunk + generate configs, then stop (no browser)

const URL = "https://eprplastic.cpcb.gov.in/#/epr/details/sales";
const CHUNK_DIR = path.resolve(__dirname, "chunks");

// Pick the worker script for a config. Priority: --worker flag > config.worker
// field > inferred from filename (registered / unregistered / pibo) > default.
function resolveWorker(cfg, cfgPath) {
    if (CLI_WORKER) return CLI_WORKER;
    if (cfg.worker) return cfg.worker;
    const name = path.basename(cfgPath).toLowerCase();
    if (name.includes("unregist")) return "run_unregistered.js";
    if (name.includes("pibo")) return "pibo_data.js";
    return "run_registered.js";
}

// ---------- helpers ----------
function log(msg) {
    const ts = new Date().toISOString();
    console.log(`[orchestrate ${ts}] ${msg}`);
}

function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") {
        if (v instanceof Date) return v.toISOString();
        if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("").trim();
        if (v.text !== undefined && v.text !== null) return String(v.text).trim();
        if (v.result !== undefined && v.result !== null) return String(v.result).trim();
    }
    return String(v).trim();
}

function loadJson(p) {
    return JSON.parse(fs.readFileSync(p, "utf8"));
}

// Build the run plan dynamically from the given worker config(s). No special
// master file — each --config is your normal worker config. A single config can
// also contain an OPTIONAL "jobs" array (advanced) to run several sheets from
// one file; otherwise the config itself IS the job.
function buildPlan() {
    if (!CONFIG_PATHS.length) {
        throw new Error("Pass at least one --config <worker-config>.json (you can repeat --config).");
    }
    const chunkSize = Number(CLI_CHUNK_SIZE || 500);
    const maxWindowsPerJob = Number(CLI_MAX_WINDOWS || 2);

    const jobs = [];
    let storageState = CLI_STORAGE || null;

    for (const cp of CONFIG_PATHS) {
        const cfgPath = path.resolve(__dirname, cp);
        if (!fs.existsSync(cfgPath)) throw new Error(`Config not found: ${cfgPath}`);
        const cfg = loadJson(cfgPath);

        // The shared login = first config's storageState (unless overridden).
        if (!storageState) storageState = cfg.storageState || "storageState_registered.json";

        const worker = resolveWorker(cfg, cfgPath);
        const baseName = path.basename(cfgPath).replace(/\.config\.json$|\.json$/i, "");

        if (Array.isArray(cfg.jobs) && cfg.jobs.length) {
            // Advanced: one config lists multiple sheets/files.
            cfg.jobs.forEach((j, i) => {
                const settings = clean({ ...cfg, ...j });
                delete settings.jobs;
                jobs.push({
                    name: j.name || `${baseName}_${i + 1}`,
                    worker: j.worker || worker,
                    settings,
                });
            });
        } else {
            const settings = clean({ ...cfg });
            delete settings.jobs;
            delete settings.worker;
            jobs.push({ name: cfg.name || baseName, worker, settings });
        }
    }

    return { url: URL, chunkSize, maxWindowsPerJob, storageState, jobs };
}

function clean(o) {
    const r = {};
    for (const k of Object.keys(o)) if (o[k] !== undefined && o[k] !== null) r[k] = o[k];
    return r;
}

function isRowEmpty(row, lastCol) {
    for (let c = 1; c <= lastCol; c++) {
        if (cellText(row.getCell(c).value) !== "") return false;
    }
    return true;
}

// Returns array of non-empty data row numbers (2..end).
function collectDataRows(ws) {
    const lastCol = ws.columnCount;
    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        if (!isRowEmpty(row, lastCol)) rows.push(r);
    }
    return rows;
}

// Writes a chunk workbook containing the header row + the given source rows,
// preserving cell value types (dates/numbers/strings). Skips rewrite when the
// file already exists (resume-friendly) unless --fresh.
async function writeChunk(srcWs, sheetName, rowNumbers, chunkPath) {
    if (fs.existsSync(chunkPath) && !FRESH) {
        return { reused: true };
    }
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);
    ws.getRow(1).values = srcWs.getRow(1).values;
    let dst = 2;
    for (const r of rowNumbers) {
        ws.getRow(dst++).values = srcWs.getRow(r).values;
    }
    await wb.xlsx.writeFile(chunkPath);
    return { reused: false };
}

// ---------- per-job preparation: count + chunk + generate configs ----------
async function prepareJob(plan, job) {
    const settings = { ...job.settings };
    const inputPath = path.resolve(__dirname, settings.inputExcel || "");
    if (!settings.inputExcel || !fs.existsSync(inputPath)) {
        throw new Error(`[${job.name}] inputExcel not found: ${inputPath}`);
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(inputPath);

    // Auto-detect sheet when not specified (dynamic): use the first worksheet.
    let ws;
    if (settings.sheetName) {
        ws = wb.getWorksheet(settings.sheetName);
        if (!ws) throw new Error(`[${job.name}] sheet not found: ${settings.sheetName}`);
    } else {
        ws = wb.worksheets[0];
        settings.sheetName = ws.name;
        log(`[${job.name}] sheetName not set -> auto-detected "${ws.name}"`);
    }

    // Default outputExcel when not specified (dynamic).
    if (!settings.outputExcel) {
        const base = path.basename(inputPath, path.extname(inputPath));
        settings.outputExcel = `${base}_output.xlsx`;
    }

    const dataRows = collectDataRows(ws);
    const total = dataRows.length;
    const chunkSize = plan.chunkSize;
    const worker = job.worker;

    // Split row numbers into groups of chunkSize.
    const groups = [];
    if (total <= chunkSize) {
        groups.push(dataRows); // single chunk = whole sheet
    } else {
        for (let i = 0; i < dataRows.length; i += chunkSize) {
            groups.push(dataRows.slice(i, i + chunkSize));
        }
    }

    log(`[${job.name}] ${total} data rows -> ${groups.length} chunk(s) of <=${chunkSize} (worker: ${worker})`);

    const chunks = [];
    for (let i = 0; i < groups.length; i++) {
        const idx = i + 1;
        const tag = `${job.name}_chunk${idx}`;
        const chunkInput = path.join(CHUNK_DIR, `${tag}.xlsx`);
        const chunkOutput = path.join(CHUNK_DIR, `${tag}_output.xlsx`);
        const chunkConfig = path.join(CHUNK_DIR, `${tag}.config.json`);

        const res = await writeChunk(ws, settings.sheetName, groups[i], chunkInput);

        // Generated worker config — inherits everything, overrides file paths.
        const cfg = {
            ...settings,
            inputExcel: path.relative(__dirname, chunkInput),
            outputExcel: path.relative(__dirname, chunkOutput),
            sheetName: settings.sheetName,
            storageState: plan.storageState,
        };
        fs.writeFileSync(chunkConfig, JSON.stringify(cfg, null, 2));

        chunks.push({ idx, tag, worker, config: chunkConfig, input: chunkInput, output: chunkOutput, rows: groups[i].length });
        log(`[${job.name}] chunk ${idx}: ${groups[i].length} rows ${res.reused ? "(reused existing file)" : "(created)"}`);
    }

    return { job, settings, worker, chunks, total };
}

// ---------- login once ----------
async function ensureLogin(plan) {
    const storagePath = path.resolve(__dirname, plan.storageState);
    if (SKIP_LOGIN) {
        if (!fs.existsSync(storagePath)) {
            throw new Error(`--skip-login set but storage state missing: ${storagePath}`);
        }
        log(`Skipping login; trusting existing ${plan.storageState}`);
        return;
    }
    const { chromium } = require("playwright");
    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext(
        fs.existsSync(storagePath) ? { storageState: storagePath } : {}
    );
    const page = await context.newPage();
    await page.goto(plan.url, { waitUntil: "domcontentloaded" });

    const isLogin = async () => {
        if (await page.locator('input[type="password"]').first().count()) return true;
        if (await page.locator('button:has-text("Login"), button:has-text("Sign In")').first().count()) return true;
        return false;
    };

    if (!fs.existsSync(storagePath) || (await isLogin())) {
        log("Login required. Log in manually in the opened window, then press ENTER here...");
        await new Promise((res) => process.stdin.once("data", () => res()));
        await context.storageState({ path: storagePath });
        log(`Saved shared session -> ${plan.storageState}`);
    } else {
        // refresh/extend the saved session
        await context.storageState({ path: storagePath });
        log(`Reusing valid session -> ${plan.storageState}`);
    }
    await browser.close();
}

// ---------- spawn a single chunk worker ----------
function runWorker(chunk, jobName) {
    return new Promise((resolve) => {
        const runlog = path.join(CHUNK_DIR, `${chunk.tag}.runlog.txt`);
        const out = fs.createWriteStream(runlog, { flags: "a" });
        const prefix = `[${jobName}#${chunk.idx}]`;
        log(`${prefix} start (${chunk.rows} rows) -> ${path.basename(chunk.config)}`);

        const child = spawn("node", [path.resolve(__dirname, chunk.worker), "--config", chunk.config], {
            cwd: __dirname,
            stdio: ["ignore", "pipe", "pipe"], // no stdin -> children never block on login prompt
        });

        const pipe = (stream) => {
            let buf = "";
            stream.on("data", (d) => {
                out.write(d);
                buf += d.toString();
                const lines = buf.split("\n");
                buf = lines.pop();
                for (const ln of lines) if (ln.trim()) console.log(`${prefix} ${ln}`);
            });
        };
        pipe(child.stdout);
        pipe(child.stderr);

        child.on("close", (code) => {
            out.end();
            log(`${prefix} finished (exit ${code})`);
            resolve({ chunk, code });
        });
    });
}

// ---------- per-job pool (cap concurrency) ----------
async function runJobPool(prepared, maxWindows) {
    const { job, chunks } = prepared;
    const queue = [...chunks];
    const results = [];
    const workers = new Array(Math.min(maxWindows, queue.length)).fill(0).map(async () => {
        while (queue.length) {
            const chunk = queue.shift();
            results.push(await runWorker(chunk, job.name));
        }
    });
    await Promise.all(workers);
    return results;
}

// Find the column number of the first header that matches one of `names`.
function findCol(ws, names) {
    const want = names.map((n) => n.trim().replace(/\s+/g, " ").toLowerCase());
    let found = null;
    ws.getRow(1).eachCell((c, n) => {
        const k = cellText(c.value).trim().replace(/\s+/g, " ").toLowerCase();
        if (found === null && want.includes(k)) found = n;
    });
    return found;
}

// Unique key for a row: prefer the business invoice number; fall back to a
// hash of all cell values so genuinely identical rows still dedup.
function rowKey(row, lastCol, keyCol) {
    if (keyCol) {
        const k = cellText(row.getCell(keyCol).value);
        if (k) return `K:${k.toUpperCase()}`;
    }
    const parts = [];
    for (let c = 1; c <= lastCol; c++) parts.push(cellText(row.getCell(c).value));
    return `H:${parts.join("")}`;
}

// ---------- merge a job's chunk outputs + logs (with DEDUP) ----------
// Produces ONE final sheet per job with NO duplicate rows. Dedup key = the
// invoice number column when present, else the full row content.
async function mergeJob(prepared) {
    const { job, settings, chunks } = prepared;
    const baseName = `${job.name}_combined`;
    const mergedXlsx = path.resolve(__dirname, `${baseName}_output.xlsx`);
    const mergedLog = path.resolve(__dirname, `${baseName}_log.csv`);

    const outWb = new ExcelJS.Workbook();
    const outWs = outWb.addWorksheet(settings.sheetName);
    let headerWritten = false;
    let dst = 2;
    let kept = 0, dupes = 0, missingChunks = 0;
    const seen = new Set();
    const invoiceNames = ["E-Invoice Number*", "E-Invoice Number", "INVOICE No", "INVOICE NO", "Invoice Number"];

    for (const ch of chunks) {
        if (!fs.existsSync(ch.output)) {
            log(`[${job.name}] merge: MISSING chunk output ${path.basename(ch.output)} (chunk not completed)`);
            missingChunks++;
            continue;
        }
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(ch.output);
        const ws = wb.getWorksheet(settings.sheetName) || wb.worksheets[0];
        if (!ws) continue;
        if (!headerWritten) {
            outWs.getRow(1).values = ws.getRow(1).values;
            headerWritten = true;
        }
        const lastCol = ws.columnCount;
        const keyCol = findCol(ws, invoiceNames);
        for (let r = 2; r <= ws.rowCount; r++) {
            const row = ws.getRow(r);
            if (isRowEmpty(row, lastCol)) continue;
            const key = rowKey(row, lastCol, keyCol);
            if (seen.has(key)) { dupes++; continue; } // skip duplicate
            seen.add(key);
            outWs.getRow(dst++).values = row.values;
            kept++;
        }
    }
    if (headerWritten) {
        await outWb.xlsx.writeFile(mergedXlsx);
        log(`[${job.name}] MERGED -> ${path.basename(mergedXlsx)} | unique rows=${kept} | duplicates removed=${dupes}` +
            (missingChunks ? ` | ⚠️ ${missingChunks} chunk(s) missing` : ""));
    } else {
        log(`[${job.name}] merge: no chunk outputs found — nothing merged.`);
    }

    // Merge per-chunk log CSVs (dedup identical log lines too).
    const logLines = [];
    const seenLog = new Set();
    let logHeader = null;
    for (const ch of chunks) {
        const baseOut = path.basename(ch.output, path.extname(ch.output));
        const chLog = path.resolve(__dirname, `${baseOut}_log.csv`);
        if (!fs.existsSync(chLog)) continue;
        const lines = fs.readFileSync(chLog, "utf8").split(/\r?\n/).filter((l) => l.length);
        if (!lines.length) continue;
        if (!logHeader) logHeader = lines[0];
        for (const ln of lines.slice(1)) {
            if (seenLog.has(ln)) continue;
            seenLog.add(ln);
            logLines.push(ln);
        }
    }
    if (logHeader) {
        fs.writeFileSync(mergedLog, [logHeader, ...logLines].join("\n") + "\n");
        log(`[${job.name}] merged ${logLines.length} log rows -> ${path.basename(mergedLog)}`);
    }

    prepared.mergeStats = { kept, dupes, missingChunks };
}

// ---------- report (terminal + saved file) ----------
function parseCsv(text) {
    const rows = [];
    let f = [], cur = "", q = false;
    for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (q) {
            if (c === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
            else cur += c;
        } else if (c === '"') q = true;
        else if (c === ",") { f.push(cur); cur = ""; }
        else if (c === "\n") { f.push(cur); rows.push(f); f = []; cur = ""; }
        else if (c !== "\r") cur += c;
    }
    if (cur.length || f.length) { f.push(cur); rows.push(f); }
    return rows;
}
function classify(status) {
    const s = (status || "").toLowerCase();
    if (s.includes("fill") || s.includes("success")) return "success";
    if (s.includes("skip")) return "skipped";
    if (s.includes("fail") || s.includes("error") || s.includes("duplicate")) return "failed";
    return "other";
}
function topReasons(obj, n = 10) {
    return Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, n);
}
function statsForJob(prepared) {
    const logPath = path.resolve(__dirname, `${prepared.job.name}_combined_log.csv`);
    const ms = prepared.mergeStats || {};
    const s = {
        job: prepared.job.name, worker: prepared.job.worker, chunks: prepared.chunks.length,
        sourceRows: prepared.total, success: 0, failed: 0, skipped: 0, other: 0, totalLogged: 0,
        uniqueMerged: ms.kept ?? null, duplicatesRemoved: ms.dupes ?? null, missingChunks: ms.missingChunks ?? 0,
        failReasons: {}, skipReasons: {},
    };
    if (!fs.existsSync(logPath)) return s;
    const rows = parseCsv(fs.readFileSync(logPath, "utf8")).filter((r) => r.length > 1);
    if (!rows.length) return s;
    const si = rows[0].indexOf("status");
    const mi = rows[0].indexOf("message");
    for (const r of rows.slice(1)) {
        const status = si >= 0 ? r[si] : "";
        const msg = (mi >= 0 ? r[mi] : "") || status;
        s.totalLogged++;
        const cls = classify(status);
        s[cls]++;
        if (cls === "failed") s.failReasons[msg] = (s.failReasons[msg] || 0) + 1;
        if (cls === "skipped") s.skipReasons[msg] = (s.skipReasons[msg] || 0) + 1;
    }
    return s;
}
function buildReport(preparedList, chunkResults, startedAt) {
    const finishedAt = new Date();
    const L = [];
    const P = (s = "") => L.push(s);
    P("=".repeat(66));
    P("  EPR AUTOMATION — RUN REPORT");
    P("=".repeat(66));
    P(`Started : ${startedAt.toISOString()}`);
    P(`Finished: ${finishedAt.toISOString()}  (${Math.round((finishedAt - startedAt) / 1000)}s)`);
    P(`Configs : ${CONFIG_PATHS.join(", ")}`);
    P("");
    let g = { source: 0, success: 0, failed: 0, skipped: 0, other: 0, dupes: 0, chunkFail: 0 };
    const perJob = [];
    for (const p of preparedList) {
        const s = statsForJob(p);
        perJob.push(s);
        g.source += s.sourceRows; g.success += s.success; g.failed += s.failed;
        g.skipped += s.skipped; g.other += s.other; g.dupes += s.duplicatesRemoved || 0;
        const cr = chunkResults.find((c) => c.jobName === p.job.name);
        const chunkFail = cr ? cr.results.filter((r) => r.code !== 0).length : 0;
        g.chunkFail += chunkFail;
        P("-".repeat(66));
        P(`JOB: ${s.job}   (worker: ${s.worker})`);
        P("-".repeat(66));
        P(`  Source rows        : ${s.sourceRows}`);
        P(`  Chunks             : ${s.chunks}` + (chunkFail ? `   (${chunkFail} chunk process(es) exited non-zero!)` : "") + (s.missingChunks ? `   (${s.missingChunks} chunk output missing!)` : ""));
        P(`  ✅ Success         : ${s.success}`);
        P(`  ⏭️  Skipped         : ${s.skipped}`);
        P(`  ❌ Failed          : ${s.failed}`);
        if (s.other) P(`  ❓ Other           : ${s.other}`);
        P(`  FINAL unique rows  : ${s.uniqueMerged}   (duplicates removed: ${s.duplicatesRemoved})`);
        const acc = s.success + s.skipped + s.failed + s.other;
        if (acc !== s.sourceRows) P(`  ⚠️  processed ${acc}/${s.sourceRows} — ${s.sourceRows - acc} row(s) pending (re-run to resume).`);
        if (Object.keys(s.skipReasons).length) { P(`  Skip reasons:`); for (const [r, n] of topReasons(s.skipReasons)) P(`     ${n} x  ${r}`); }
        if (Object.keys(s.failReasons).length) { P(`  Failure reasons:`); for (const [r, n] of topReasons(s.failReasons)) P(`     ${n} x  ${r}`); }
        P(`  Final sheet : ${s.job}_combined_output.xlsx`);
        P(`  Log         : ${s.job}_combined_log.csv`);
        P("");
    }
    P("=".repeat(66));
    P("  GRAND TOTAL");
    P("=".repeat(66));
    P(`  Source rows        : ${g.source}`);
    P(`  ✅ Success         : ${g.success}`);
    P(`  ⏭️  Skipped         : ${g.skipped}`);
    P(`  ❌ Failed          : ${g.failed}`);
    if (g.other) P(`  ❓ Other           : ${g.other}`);
    P(`  Duplicates removed : ${g.dupes}`);
    P(`  Chunk proc failures: ${g.chunkFail}`);
    P("=".repeat(66));
    const text = L.join("\n");
    console.log("\n" + text + "\n");
    const stamp = finishedAt.toISOString().replace(/[:.]/g, "-");
    fs.writeFileSync(path.resolve(__dirname, `report_${stamp}.txt`), text + "\n");
    fs.writeFileSync(path.resolve(__dirname, "report_latest.json"), JSON.stringify({
        startedAt: startedAt.toISOString(), finishedAt: finishedAt.toISOString(),
        durationSec: Math.round((finishedAt - startedAt) / 1000), configs: CONFIG_PATHS,
        grandTotal: g, jobs: perJob,
    }, null, 2));
    log(`Report saved -> report_${stamp}.txt and report_latest.json`);
    return { failed: g.failed, chunkFail: g.chunkFail };
}

// ---------- main ----------
(async () => {
    const startedAt = new Date();
    const plan = buildPlan();
    if (!fs.existsSync(CHUNK_DIR)) fs.mkdirSync(CHUNK_DIR, { recursive: true });

    log(`Configs: ${CONFIG_PATHS.join(", ")} | jobs=${plan.jobs.map((j) => `${j.name}(${j.worker})`).join(", ")} | chunkSize=${plan.chunkSize} | maxWindows/job=${plan.maxWindowsPerJob} | login=${plan.storageState}`);

    // 1) Prepare every job (count + chunk + configs) before any browser opens.
    const prepared = [];
    for (const job of plan.jobs) {
        prepared.push(await prepareJob(plan, job));
    }
    const totalChunks = prepared.reduce((a, p) => a + p.chunks.length, 0);
    log(`Prepared ${totalChunks} chunk(s) across ${prepared.length} job(s).`);

    if (PREPARE_ONLY) {
        log("--prepare-only: chunk files + configs generated; stopping before browser/login.");
        process.exit(0);
    }

    // 2) Login once (shared by all workers).
    await ensureLogin(plan);

    // 3) Run all job pools in parallel. Each job is capped at maxWindowsPerJob,
    //    so concurrent windows total = jobs * maxWindowsPerJob.
    log(`Launching workers (up to ${plan.maxWindowsPerJob} per job, ${prepared.length} jobs in parallel)...`);
    const allResults = await Promise.all(
        prepared.map((p) => runJobPool(p, plan.maxWindowsPerJob))
    );

    // 4) Chunk-process exit codes.
    const chunkResults = prepared.map((p, i) => ({ jobName: p.job.name, results: allResults[i] }));
    let ok = 0, fail = 0;
    allResults.flat().forEach((r) => (r.code === 0 ? ok++ : fail++));
    log(`All chunk processes done. exited-ok=${ok} exited-nonzero=${fail}`);

    // 5) Merge each job into ONE final deduped sheet.
    if (!NO_MERGE) {
        for (const p of prepared) await mergeJob(p);
    }

    // 6) Detailed report (terminal + saved files).
    const { failed, chunkFail } = buildReport(prepared, chunkResults, startedAt);

    log("Orchestration complete.");
    process.exit(failed > 0 || chunkFail > 0 ? 1 : 0);
})().catch((e) => {
    console.error(`[orchestrate] FATAL: ${e?.stack || e}`);
    process.exit(1);
});
