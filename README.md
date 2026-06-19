# EPR Sales Automation

## Configuration
Default config file is `config.json`. You can also pass a per-user config:
```
node run_export.js --config config_export.json
node run_registered.js --config config_registered.json
node run_unregistered.js --config config_unregistered.json
```

Example config:
```json
{
  "inputExcel": "Toppan PWP Sale For Automation.xlsx",
  "sheetName": "Active Batch",
  "outputExcel": "Toppan PWP Sale For Automation_output.xlsx",
  "max_rows": 0,
  "storageState": "storageState_export.json",
  "plasticType": "PP",
  "invoicePdfDir": "C:\\\\Users\\\\Lenovo\\\\Downloads\\\\All invoices"
}
```

Notes:
- `max_rows`: set to a positive number to limit how many rows are processed. Use `0` or remove to process all rows.
- `storageState`: login session file for each user (export/registered/unregistered).
- `plasticType`: CAT-II plastic type (e.g., `PP`, `PET`).
- `invoicePdfDir`: root folder for PDFs used by `run_upload.js`.

## Log file names (match the output file name)
All log files are created using the output file base name so they are easy to track:

- `<output>_log.csv` (all processed rows: filled + failed)
- `<output>_filled.csv` (only filled rows)
- `<output>_upload_log.csv` (file upload log)
- `<output>_upload_filled.csv` (file upload rows)

## Notes
- Logs append across runs.

## Run Commands
Export:
```
node run_export.js --config config_export.json
```

Registered:
```
node run_registered.js --config config_registered.json
```

Unregistered:
```
node run_unregistered.js --config config_unregistered.json
```

Upload:
```
node run_upload.js --config config_upload.json
```

PIBO Data Entry (fill form + generate EPR Invoice Number):
```
node pibo_data.js --config config_pibo.json
```

PIBO Invoice Upload (upload PDF for each EPR entry):
```
node pibo_upload.js --config config_pibo_upload.json
```

Delete EPR Sales Entries:
```
node run_delete.js --config config_delete.json
```
Reads the same Excel format as the upload, searches each row by `EPR Invoice Number`,
clicks the delete (trash) icon in the table's Action column, confirms in the
"Confirm Deletion" modal, and writes `Deleted` to the status column.

- Writes the outcome to a `Delete Status` column if present, otherwise falls back to `Status`.
- Rows already marked deleted are skipped.
- Add `--dry-run` (or `"dryRun": true` in the config) to search rows and open the
  confirmation modal without actually deleting anything — useful for verifying selectors:
  ```
  node run_delete.js --config config_delete.json --dry-run
  ```
- Delete activity is logged to `<output>_delete_log.csv`.
