# EPR Sales Automation

Playwright automation that logs into the CPCB EPR portal and fills the **Sales**
form from rows in an Excel sheet. Includes:

- **Sanitization** — dates, numbers, and state/district spelling are cleaned per row.
- **Orchestrator** — dynamic chunking + parallel runs for large sheets, with a
  single shared login and one final de-duplicated output sheet per job.

## Setup

```bash
npm install
npx playwright install chromium
```

---

## 1. The worker scripts (fill one sheet)

| Script | Use for |
|---|---|
| `run_registered.js` | Registered sales entries |
| `run_unregistered.js` | Unregistered sales entries |
| `pibo_data.js` | PIBO entity entries (different form/columns) |

Each takes a config and processes one sheet:

```bash
node run_registered.js --config config_registered.json
```

### Config format

```json
{
  "inputExcel": "POLYPLEX BAAZPUR Automation READY (PP).xlsx",
  "sheetName": "PP BAAZPUR",
  "outputExcel": "PP_output.xlsx",
  "plasticType": "PP",
  "storageState": "storageState_registered.json",
  "max_rows": 0
}
```

- `sheetName` / `outputExcel` are optional — the orchestrator auto-detects the
  first sheet and defaults the output name if omitted.
- `max_rows: 0` means "all rows".
- `storageState` is the saved login session (created on first run).

The worker writes a `Status` column back into the sheet (`Filled` /
`Failed: <reason>` / `Skipped: <reason>`) and is **resumable** — re-running skips
rows already Filled or that already have an EPR Invoice Number.

---

## 2. Sanitization (`sanitize.js`)

Runs automatically inside every worker, per row, before submitting. It is
**sheet-agnostic** — it only acts on columns that exist in the current sheet.

| What | Behavior |
|---|---|
| **Sales / invoice dates** | Any format → `YYYY-MM-DD` (the portal's date input). Handles `01.04.2025`, `13/06/2025`, Excel serials like `45901`, `1-Apr-2025`. |
| **Amounts** (Principal, GST charges) | Float noise removed, rounded to 2 decimals (`397670.02999999997 → 397670.03`). |
| **Quantity** | Filled **exactly** as in the sheet — never rounded. |
| **State / District** | Spelling auto-corrected against `state_district_mapping.xlsx - State-District Mapping.csv` (`UTTRAKHAND → UTTARAKHAND`, `chenai → Chennai`). |
| **Empty / zero rows** | If a mandatory column is empty or `0`, the row is **skipped** with the reason logged. |

---

## 3. The orchestrator (`orchestrate.js`) — large sheets & parallel runs

Sits on top of the workers. You point it at your **normal worker config(s)** —
there is no special "master" file. It then:

1. Reads each config, opens the Excel, **counts data rows**.
2. **≤ 500 rows** → runs one window (no chunking).
   **> 500 rows** → splits the sheet into 500-row **chunk files** and generates a
   worker config for each.
3. **Logs in once** (headed browser) and saves the session; every chunk worker
   reuses that same login — no repeated prompts.
4. Runs chunks in **parallel pools** (default 2 windows per job; configurable).
   Multiple jobs (e.g. PP + PET) run at the same time.
5. **Merges** each job's chunk outputs into **one final de-duplicated sheet**,
   plus a combined log.
6. Prints a **report** to the terminal and saves `report_<timestamp>.txt` +
   `report_latest.json`.

### Commands

```bash
# One sheet (chunks automatically if > 500 rows)
node orchestrate.js --config config_registered.json

# PP + PET together — just pass both normal configs
node orchestrate.js --config config_pp.json --config config_pet.json

# Unregistered / PIBO (worker auto-detected from the config filename)
node orchestrate.js --config config_unregistered.json
node orchestrate.js --config config_pibo.json
```

### Options (all optional)

| Flag | Default | Meaning |
|---|---|---|
| `--chunk-size N` | `500` | Rows per chunk file. |
| `--max-windows N` | `2` | Concurrent browser windows **per job**. |
| `--worker FILE` | auto | Force the worker script for all configs (`run_registered.js` / `run_unregistered.js` / `pibo_data.js` / `sim_worker.js`). |
| `--storage FILE` | first config's | Shared login session file for all workers. |
| `--fresh` | off | Rebuild chunk files from scratch (otherwise reused, for resume). |
| `--no-merge` | off | Skip the final merge step. |
| `--skip-login` | off | Trust the existing saved session (no login window). |
| `--prepare-only` | off | Only build chunk files + configs, then stop (no browser). |

### Example

```bash
# Big PP + PET run, 4 windows each, after logging in once:
node orchestrate.js --config config_pp.json --config config_pet.json --max-windows 4
```

> ⚠️ Start with 2 windows per job. The portal's tolerance for many simultaneous
> submissions on one login is unknown — scale up only after a small test.

### Outputs per job

| File | Contents |
|---|---|
| `<job>_combined_output.xlsx` | **Final sheet — all chunks merged, duplicates removed.** |
| `<job>_combined_log.csv` | Combined per-row log (Filled / Skipped / Failed + reasons). |
| `report_<timestamp>.txt` | Human-readable run summary. |
| `report_latest.json` | Machine-readable run summary. |
| `chunks/` | Intermediate chunk files + per-chunk configs/logs (gitignored). |

De-dup key is the invoice number column when present, otherwise the full row
content, so the final sheet never contains duplicate entries.

---

## How concurrent Excel writes are kept safe

Parallel writes to the **same** file never happen by design:

- The original input Excel is **read once** and never written to.
- Each worker owns **its own** chunk output file — it's the only writer.
- Every save is **atomic** (`.tmp` then `rename`, keeping a `.bak`), so a crash
  mid-write never corrupts the file.
- The **merge** runs single-threaded **after** all workers exit.

So N parallel windows = N separate files = zero write contention.

---

## Simulation (test without the portal)

`sim_worker.js` runs the entire pipeline (chunking, parallel pools, sanitize,
per-row atomic save, logging, merge) but **fakes the submit** — no browser, no
login. Use it to validate data/flow on real sheets:

```bash
node orchestrate.js --config config_pp.json --config config_pet.json \
  --worker sim_worker.js --max-windows 2 --skip-login --fresh
```

---

## Files NOT committed to git

These stay **local only** (see `.gitignore`):

- `*.xlsx`, `*.xls`, `*.csv` — data files
- `storageState*.json` — login sessions (auth tokens)
- `chunks/`, `report_*.txt`, `report_latest.json` — generated artifacts
