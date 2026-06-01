# EPR Sales Automation

Playwright automation that logs into the CPCB EPR portal and fills the **Sales**
form from rows in an Excel sheet, in parallel, and records the generated **EPR
Invoice Number** back into the sheet.

Key features:
- **Sanitization** — dates, numbers and state/district spelling cleaned per row.
- **Parallel orchestration** — splits work across multiple browser windows on a
  single shared login.
- **`--pending` auto-chunking** — sanitize, keep only rows that still need an EPR,
  then auto-split those into 3–4 balanced chunks (the recommended way to run).
- **Resumable & safe** — re-running only does rows without a valid EPR; each
  window writes its own file (no concurrent-write corruption).
- **Reports** — one clean sheet per product with the EPR number + a plain-English
  status (why a row did / didn't get an EPR).

## Setup

```bash
npm install
npx playwright install chromium      # one-time; re-run if Playwright version changes
```

> If you see *"browserType.launch: Executable doesn't exist…"*, just re-run
> `npx playwright install chromium`. Browser binaries live in
> `~/.cache/ms-playwright/` (outside the repo), so `npm install` doesn't fetch them.

---

## Quick start (the normal workflow)

```bash
# 1. Run PP, auto-chunked into 4 parallel windows (only rows still needing an EPR)
node orchestrate.js --config config_pp_pending.json --pending --max-windows 4

#    -> opens ONE login window. Log in, wait for the dashboard, press ENTER.
#    -> 4 windows then fill the pending rows in parallel.

# 2. Build the final per-product report sheets (EPR + Status columns)
node build_final_sheets.js

# 3. Re-run step 1 to pick up any rows that failed/timed out (resume is automatic)
```

That's it for a normal run. Details below.

---

## The pieces

| File | Role |
|---|---|
| `run_registered.js` | Worker — fills the Registered sales form for one sheet/config |
| `run_unregistered.js` | Worker — Unregistered sales form |
| `pibo_data.js` | Worker — PIBO entity entries (different form/columns) |
| `sanitize.js` | Shared per-row cleaning (dates/numbers/state/district + skip rules) |
| `orchestrate.js` | Splits a sheet into chunks and runs N workers in parallel |
| `build_final_sheets.js` | Builds `PP_FINAL_REPORT.xlsx` / `PET_FINAL_REPORT.xlsx` |
| `merge_chunks.js` | Merges chunk outputs into `FINAL_PP.xlsx` / `FINAL_PET.xlsx` (deduped) |
| `count_progress.js` | Prints live per-chunk + total progress while a run is going |
| `sim_worker.js` | Browser-free simulation worker (test the pipeline, no portal) |

---

## 1. Config format

A worker config is plain JSON:

```json
{
  "inputExcel": "PP_FINAL_REPORT.xlsx",
  "sheetName": "PP BAAZPUR",
  "outputExcel": "PP_FINAL_REPORT.xlsx",
  "plasticType": "PP",
  "storageState": "storageState_registered.json",
  "max_rows": 0
}
```

- `inputExcel` / `outputExcel` — point both at the **same report file** so EPRs
  are written back into it and resume works.
- `sheetName` — optional; auto-detected (first sheet) if omitted.
- `plasticType` — `PP` or `PET` (selects the correct CAT-II category row).
- `storageState` — the saved login session file (shared across windows).
- `max_rows: 0` — process all rows.

Provided configs: `config_pp_pending.json`, `config_pet_pending.json`,
`config_registered.json`, `config_unregistered.json`, `config_pibo.json`.

---

## 2. `--pending` auto-chunking (recommended)

This is the smart mode. For each config it:

1. **Sanitizes** every row (see §4) at entry time.
2. **Keeps only rows that still need an EPR** — skips rows that already have a
   valid (all-digit) EPR, and **excludes** rows missing a mandatory field or with
   quantity 0 (logged so you know why).
3. **Auto-splits** the remaining pending rows into balanced chunks:
   - fewer than 200 pending → **3 windows**
   - 200 or more → **4 windows**
   - override with `--windows N` or `--max-windows N`.
4. Runs all chunks **in parallel**.

```bash
# PP only, force 4 windows
node orchestrate.js --config config_pp_pending.json --pending --max-windows 4

# PET only, let it auto-pick 3/4 windows
node orchestrate.js --config config_pet_pending.json --pending

# PP + PET together (each auto-chunked)
node orchestrate.js --config config_pp_pending.json --config config_pet_pending.json --pending
```

Because it filters to pending rows every run, **`--pending` IS the resume** —
just run the same command again to retry anything that failed.

---

## 3. Login

- **First run / expired session:** run **without** `--skip-login`. The
  orchestrator opens ONE window → you log in → wait for the dashboard → press
  ENTER. It saves the session; all worker windows reuse it (they never prompt).
- **Session still valid (~2 h):** add `--skip-login` to skip the login window.
- The login token lasts ~2 hours. A long run may expire partway — just re-run
  (without `--skip-login`) to log in again and resume.

---

## 4. Sanitization (`sanitize.js`)

Runs automatically inside every worker, per row. Sheet-agnostic — only acts on
columns that exist.

| What | Behavior |
|---|---|
| **Dates** | Any format → `YYYY-MM-DD`. Handles `01.04.2025`, `13/06/2025`, Excel serials (`45901`), `1-Apr-2025`. |
| **Amounts** (Principal, GST charges) | Float noise removed, rounded to 2 dp (`397670.02999999997 → 397670.03`). |
| **Quantity** | Filled **exactly** as in the sheet — never rounded. |
| **State / District** | Spell-corrected against `state_district_mapping.xlsx - State-District Mapping.csv` (`UTTRAKHAND → UTTARAKHAND`, `chenai → Chennai`). |
| **Empty / zero rows** | Mandatory column empty or `0` → row skipped with a logged reason. |

**EPR read-back** is hardened for parallel runs: it only accepts an all-digit
value from the EPR copy field, ignores it if it's already seen (stale/cross-window),
and re-reads to confirm it's stable — preventing the garbage/duplicate EPRs seen
in early runs.

---

## 5. Reports (`build_final_sheets.js`)

Builds one clean sheet per product from the **original** file, matching each
invoice to its generated EPR:

```bash
node build_final_sheets.js
```

Outputs **`PP_FINAL_REPORT.xlsx`** and **`PET_FINAL_REPORT.xlsx`**, each with the
original 16 data columns plus:

| Column | Content |
|---|---|
| **17 — EPR Invoice Number** | the generated number (blank if not done) |
| **18 — Status** | `Filled`, or a plain-English reason it didn't get one |

Example statuses (filter column 18 to triage):
- `Filled`
- `PENDING - not attempted (chunk did not run)` → just resume
- `Failed: EPR not returned after submit (retry)` → resume
- `Failed: a field click timed out (page slow) - retry` → resume
- `Failed: submit disabled - a required field was missing/invalid` → fix the row's data
- `Failed: duplicate EPR captured (re-fetch real EPR)` → verify the real EPR on the portal

---

## 6. Live progress while a run is going

```bash
node count_progress.js     # per-chunk valid/bad counts + PP/PET totals + % done
```

Read-only — safe to run anytime, won't disturb the running windows.

---

## 7. Merging chunk outputs (`merge_chunks.js`)

Combines all chunk outputs into deduped `FINAL_PP.xlsx` / `FINAL_PET.xlsx`
(dedup key = E-Invoice Number). Safe to run mid-run (falls back to a chunk's
`.bak` if it's being written). For the human-facing deliverable prefer
`build_final_sheets.js` (§5); use this for a raw merged dump.

```bash
node merge_chunks.js
```

---

## Orchestrator options

| Flag | Default | Meaning |
|---|---|---|
| `--config FILE` | — | Worker config (repeatable for multiple products) |
| `--pending` | off | Sanitize + run ONLY rows needing an EPR, auto-split 3/4 windows |
| `--windows N` | auto | In `--pending` mode, force N windows |
| `--max-windows N` | `2` | Windows per job (also overrides `--pending` auto count) |
| `--chunk-size N` | `500` | Rows per chunk (non-pending mode only) |
| `--skip-login` | off | Trust the saved session (no login window) |
| `--worker FILE` | auto | Force worker script (else inferred from config name) |
| `--storage FILE` | first config's | Shared login session file |
| `--fresh` | off | Rebuild chunk files from source (do NOT use when resuming) |
| `--no-merge` | off | Skip the final merge step |
| `--prepare-only` | off | Build chunk files + configs, then stop (no browser) |
| `--max-chunks N` | off | Smoke test: only run the first N chunks per job |

---

## How concurrent Excel writes are kept safe

- The original input Excel is **read once**, never written by a worker.
- Each worker owns **its own** chunk output file — the only writer.
- Every save is **atomic** (`.tmp` → `rename`, keeping a `.bak`).
- Merge/report builders run **single-threaded after** all workers exit.

N parallel windows = N separate files = zero write contention.

---

## Window tiling (optional)

When launched by the orchestrator, each worker window opens in its own quadrant
of a 2×2 grid (via `WIN_SLOT`) so all windows are visible without tab-switching.
Cosmetic only; on Wayland the compositor may not honor exact positions.

---

## Simulation (no portal)

```bash
node orchestrate.js --config config_pp_pending.json --pending \
  --worker sim_worker.js --skip-login --fresh
```

Runs the full pipeline (chunk → parallel → sanitize → save → merge) but fakes
the submit — no browser, no login. Good for validating data/flow.

---

## Files NOT committed to git

Local only (see `.gitignore`):

- `*.xlsx`, `*.xls`, `*.csv` — data files
- `storageState*.json` — login sessions (auth tokens)
- `chunks/`, `report_*.txt`, `report_latest.json` — generated artifacts
