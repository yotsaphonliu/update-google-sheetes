# update-google-sheets

Push values into Google Sheets using labels stored in an Excel workbook.

## Requirements
- Go 1.24+
- Google Sheets API enabled and credentials exposed via `GOOGLE_APPLICATION_CREDENTIALS`
- Excel workbook (copied to `cfg/Schedule.xlsx`)

## Setup
1. Run `go run ./cmd/configset`.
   - Enter the spreadsheet ID, optional sheet filter, and lookup value (defaults populate from the previous run).
   - Choose whether to keep the existing workbook or pick a new one via Finder. Selected files must be `.xls` or `.xlsx`; the file is copied to `cfg/Schedule.xlsx`.
2. The answers are saved in `cfg/config.yaml`.

## Updating the sheet
```
GOOGLE_APPLICATION_CREDENTIALS=/path/key.json go run .
```
The tool reads `cfg/config.yaml`, finds every matching cell inside `cfg/Schedule.xlsx`, confirms the target range already has data, and writes the lookup value. Logs show the ranges updated and row/cell counts.

## Authentication helper (optional)
User credentials often need a quota project. You can run:
```
./scripts/gcloud_login.sh
./scripts/run_with_gcloud.sh
```
Skip these if you rely on a service-account JSON file.

## Tweaking the config later
Re-run `go run ./cmd/configset` whenever you need to change the spreadsheet ID, sheet filter, lookup text, or workbook. Defaults are pre-filled with the current values.

## Notes
- `cfg/config.yaml` and `cfg/Schedule.xlsx` are the only inputs the updater uses. Delete the YAML to rerun the wizard from scratch.
- Finder selections only accept `.xls`/`.xlsx` files.
- The updater refuses to overwrite empty ranges—seed the sheet manually the first time.
