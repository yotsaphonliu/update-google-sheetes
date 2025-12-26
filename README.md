# update-google-sheets

Push one lookup value into many Google Sheets cells using an Excel workbook as the map.

## Main pieces
- `cfg/Schedule.xlsx`(template) ---> Google Sheet (target)

## Requirements
- Go 1.24+
- Google Sheets API enabled and credentials referenced via `GOOGLE_APPLICATION_CREDENTIALS`
- An Excel workbook you want to use as the stencil (copied to `cfg/Schedule.xlsx`)

## Configure the run
1. `go run ./cmd/configset`
   - Provide the **Google spreadsheet ID** (the part after `/d/` in the URL).
   - Optionally enter a **sheet filter** to restrict matching to a single tab inside the workbook.
   - Enter the **lookup value** (the text the updater searches for inside the workbook).
   - Decide whether to keep the existing workbook or pick a new `.xls`/`.xlsx` file; the chosen file is copied into `cfg/Schedule.xlsx`.
2. Answers land in `cfg/config.yaml`. Re-run the wizard any time you want to change the spreadsheet, lookup text, or workbook.

## Update flow
1. Double-check the Google Sheet already contains placeholder data in every target cell. The updater refuses to overwrite blank ranges.
2. The tool loads `cfg/config.yaml`, scans `cfg/Schedule.xlsx` for the lookup value, fetches the matching ranges from the Google Sheet, and writes the lookup value into any cells that currently contain something else. Logs list every range touched plus total rows/cells.

## Optional auth helpers
Run `make gcloud-all` to run both steps in one shot.
Run `make gcloud-login` to perform the scoped ADC login through `gcloud`.
Run `make run-with-gcloud` to launch the updater using that session.

## Handy notes
- `cfg/config.yaml` + `cfg/Schedule.xlsx` are the only inputs. Delete the YAML if you want to start from a clean slate.
- Finder selections only accept `.xls`/`.xlsx` files.
- Empty Google Sheet ranges are skipped; seed them manually once so the updater can detect the pre-existing data.
