# update-google-sheets

Automates writing cell values to a Google Sheet from the command line. You can target a range directly (e.g. `Sheet1!B5`) or point the tool at an Excel configuration file and let it determine the cell for you in real time.

## Requirements
- Go 1.24 or newer
- Google Cloud project with the Google Sheets API enabled
- Credentials exposed via Application Default Credentials (`GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json` works well)
- Optional: an Excel workbook (like `Schedule.xlsx`) that maps friendly labels to the actual cell coordinates

## Build
```
go build ./...
```
This produces an `update-google-sheets` binary in the repo directory; you can also run it with `go run .` while iterating.

## Basic usage
Run the tool and follow the prompts—no flags required:
```
GOOGLE_APPLICATION_CREDENTIALS=/path/key.json go run .
```
The wizard will ask for the spreadsheet ID, how to pick the destination range (manual entry or Excel lookup), how to supply the JSON matrix (paste now or point at a file), and a few behavioural choices. When the write succeeds it prints the updated ranges and row/cell counts.

### Supplying values
The CLI expects a JSON array of arrays. Choose one of the following inputs (priority order):
1. `-values` flag containing the JSON literal
2. `-values-file` pointing at a JSON file
3. Standard input (e.g. `cat data.json | go run . …`)

## Excel-driven lookup
Every run now relies on the Excel workbook to determine where to write. The workflow is identical to the previous description:
1. Every worksheet in `Schedule.xlsx` is scanned and every cell whose trimmed text equals `โอเลี้ยง` is collected.
2. Each sheet name + cell coordinate becomes part of a single Google Sheets batch update request, so duplicated labels all get updated together.
3. Before overwriting, the tool (by default) checks that the destination cell currently contains something in Google Sheets. Answer "n" when prompted if you want to allow writing to blank cells.

This is handy when you maintain schedules locally but push definitive values into a central Google Sheet. Any Unicode text—including Thai labels like `โอเลี้ยง`—is supported as long as it matches exactly.

## Interactive choices
During the wizard you will be asked to:
- Enter the spreadsheet ID.
- Provide the Excel workbook path (defaults to `Schedule.xlsx` in the repo) and optional sheet filter, plus the lookup value to locate inside the workbook.
- The tool automatically uses Sheets’ `USER_ENTERED` behavior and writes by rows.
- Updates are only applied when the Google Sheet already contains data at the target range; if the range is empty the run exits.

## Helper script for gcloud ADC
`run_with_gcloud.sh` automates the recommended authentication flow when you rely on user credentials. It performs `gcloud auth application-default login --scopes="https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/spreadsheets"`, optionally sets a quota project when `GCP_QUOTA_PROJECT` is exported, and then executes `go run .` with the arguments you pass after `--` (the wizard still handles every in-app choice).

Example:
```
./run_with_gcloud.sh -- \
  -spreadsheet 1Abc... \
  -config-xlsx Schedule.xlsx \
  -lookup-value 'โอเลี้ยง' \
  -values '[["โอเลี้ยง"]]'
```
If you are using a service-account JSON you can skip the script and rely on `GOOGLE_APPLICATION_CREDENTIALS` directly.

## Tips
- When multiple cells share the same lookup text, every match in the Excel file is updated.
- Run with `-require-non-empty=false` when seeding a blank sheet for the first time.
- Use `RAW` input mode if you want to avoid Google Sheets evaluating formulas.

## Troubleshooting
- **auth errors**: ensure the service account has edit access to the sheet and `GOOGLE_APPLICATION_CREDENTIALS` points to the JSON key.
- **lookup failures**: run with `-range` manually to rule out Excel parsing issues, or verify the text has no leading/trailing spaces (the tool trims whitespace).
- **empty-range precondition**: either populate the cell manually first or pass `-require-non-empty=false`.
