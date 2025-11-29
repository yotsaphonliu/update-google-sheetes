# update-google-sheets

Automates writing cell values to a Google Sheet from the command line. You can target a range directly (e.g. `Sheet1!B5`) or point the tool at an Excel configuration file and let it determine the cell for you in real time.

## Requirements
- Go 1.24 or newer
- Google Cloud project with the Google Sheets API enabled
- Credentials exposed via Application Default Credentials (`GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json` works well)
- Optional: an Excel workbook (like `Schedule.xlsx`) that maps friendly labels to the actual cell coordinates

```bash
export GOOGLE_APPLICATION_CREDENTIALS="$HOME/.config/gcloud/application_default_credentials.json"
```

## Build
```
go build ./...
```
This produces an `update-google-sheets` binary in the repo directory; you can also run it with `go run .` while iterating.

## Basic usage
Populate `cfg/config.yaml` (see below) and run the app:
```
GOOGLE_APPLICATION_CREDENTIALS=/path/key.json go run .
```
When the config file exists the CLI runs without prompts. If it is missing you will be guided through the same questions interactively (spreadsheet ID, Excel workbook + lookup value). Either way, the tool prints the updated ranges and row/cell counts once the write succeeds.

## Configuration file
`cfg/config.yaml` controls non-interactive runs:
```yaml
spreadsheet_id: "your-sheet-id"
config_xlsx: "cfg/Schedule.xlsx"   # optional; defaults to cfg/Schedule.xlsx
config_sheet: "Sheet1"          # optional sheet filter
lookup_value: "โอเลี้ยง"        # required
```
The CLI writes the lookup value into every matching cell in Google Sheets. Delete the file if you prefer to answer the prompts each time.

### Updating the config via CLI
Run the helper command to rewrite the YAML and copy a workbook into place:
```
go run ./cmd/configset \
  -spreadsheet 1BqpRmUO6fjIUWA840gbgVvQnkht3Xv28lE4b8J3hp_U \
  -sheet "Live IMURA Dec 25" \
  -lookup "โอเลี้ยง" \
  -workbook-src /path/to/Schedule.xlsx \
  -workbook-dest cfg/Schedule.xlsx
```
The command overwrites `cfg/Schedule.xlsx` with the provided source file and saves the YAML at `cfg/config.yaml`.

## Excel-driven lookup
Whether values came from the config file or the wizard, the workflow is identical:
1. Every worksheet in `cfg/Schedule.xlsx` is scanned and every cell whose trimmed text equals `โอเลี้ยง` is collected.
2. Each sheet name + cell coordinate becomes part of a single Google Sheets batch update request, so duplicated labels all get updated together.
3. Before overwriting, the tool (by default) checks that the destination cell currently contains something in Google Sheets. Answer "n" when prompted if you want to allow writing to blank cells.

This is handy when you maintain schedules locally but push definitive values into a central Google Sheet. Any Unicode text—including Thai labels like `โอเลี้ยง`—is supported as long as it matches exactly.

## Interactive choices
If `cfg/config.yaml` is absent you will be prompted to:
- Enter the spreadsheet ID.
- Provide the Excel workbook path (defaults to `cfg/Schedule.xlsx`) and optional sheet filter, plus the lookup value to locate inside the workbook.
- Confirm that the target range must already contain data (this guard is always enforced).

## Helper scripts for gcloud ADC
Use `scripts/gcloud_login.sh` to establish Application Default Credentials with the proper scopes and optional quota project:
```
export GCP_QUOTA_PROJECT=my-gcp-project    # optional but recommended for user creds
./scripts/gcloud_login.sh
```
Once authenticated, launch the updater (config file or wizard) via `run_with_gcloud.sh`:
```
./scripts/run_with_gcloud.sh
```
If you are using a service-account JSON you can skip both scripts and rely on `GOOGLE_APPLICATION_CREDENTIALS` directly.

## Tips
- When multiple cells share the same lookup text, every match in the Excel file is updated.
- Keep `Schedule.xlsx` in sync with the Google Sheet so lookups remain accurate.
- Delete or adjust `cfg/config.yaml` when switching to a different spreadsheet.

## Troubleshooting
- **auth errors**: ensure the service account has edit access to the sheet and `GOOGLE_APPLICATION_CREDENTIALS` points to the JSON key.
- **lookup failures**: verify the lookup text has no leading/trailing spaces (the tool trims whitespace) and that the workbook path is correct.
- **empty-range precondition**: seed the Google Sheet manually before running; empty ranges are skipped by design.
