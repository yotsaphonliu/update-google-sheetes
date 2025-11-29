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
Provide the spreadsheet ID from the Google Sheets URL, the range in A1 notation, and the JSON matrix you want written:
```
GOOGLE_APPLICATION_CREDENTIALS=/path/key.json \
  go run . \
  -spreadsheet 1AbcDeFgHij1234567890 \
  -range "Sheet1!A2:C3" \
  -value-input USER_ENTERED \
  -values '[["Name","Score"],["Ada",98],["Ben",91]]'
```
The program echoes how many cells were modified and prints the values returned by the Sheets API so you can confirm the write.

### Supplying values
The CLI expects a JSON array of arrays. Choose one of the following inputs (priority order):
1. `-values` flag containing the JSON literal
2. `-values-file` pointing at a JSON file
3. Standard input (e.g. `cat data.json | go run . …`)

## Excel-driven lookup
If an Excel file already maps labels to their spreadsheet positions, let the tool resolve the range automatically:
```
GOOGLE_APPLICATION_CREDENTIALS=/path/key.json \
  go run . \
  -spreadsheet 1Abc... \
  -config-xlsx Schedule.xlsx \
  -lookup-value 'โอเลี้ยง' \
  -values '[["โอเลี้ยง"]]'
```
Workflow:
1. Every worksheet in `Schedule.xlsx` is scanned until the first cell whose trimmed text equals `โอเลี้ยง` is found.
2. That sheet name + cell coordinate becomes the A1 range for the Sheets API call.
3. Before overwriting, the tool (by default) checks that the destination cell currently contains something in Google Sheets. Set `-require-non-empty=false` if you want to allow writing to blank cells.

This is handy when you maintain schedules locally but push definitive values into a central Google Sheet. Any Unicode text—including Thai labels like `โอเลี้ยง`—is supported as long as it matches exactly.

## Flags
- `-spreadsheet` (required): Spreadsheet ID from the sheet URL.
- `-range`: A1 notation for the target range. Skip when using `-config-xlsx` + `-lookup-value`.
- `-config-xlsx`: Path to an Excel workbook that contains the lookup text.
- `-lookup-value`: Exact cell text to search for within the Excel config.
- `-require-non-empty`: Guard that stops the update if the Google Sheet destination is blank (default `true`).
- `-value-input`: `RAW` or `USER_ENTERED` (default `USER_ENTERED`).
- `-dimension`: `ROWS` or `COLUMNS` for how the data should be applied (default `ROWS`).
- `-values`: Inline JSON matrix.
- `-values-file`: Path to a JSON matrix file.

## Helper script for gcloud ADC
`run_with_gcloud.sh` automates the recommended authentication flow when you rely on user credentials. It performs `gcloud auth application-default login --scopes="https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/spreadsheets"`, optionally sets a quota project when `GCP_QUOTA_PROJECT` is exported, and then executes `go run .` with the arguments you pass after `--`.

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
- When multiple cells share the same lookup text, only the first match in the Excel file is used. Consider expanding the lookup logic if you need disambiguation.
- Run with `-require-non-empty=false` when seeding a blank sheet for the first time.
- Use `RAW` input mode if you want to avoid Google Sheets evaluating formulas.

## Troubleshooting
- **auth errors**: ensure the service account has edit access to the sheet and `GOOGLE_APPLICATION_CREDENTIALS` points to the JSON key.
- **lookup failures**: run with `-range` manually to rule out Excel parsing issues, or verify the text has no leading/trailing spaces (the tool trims whitespace).
- **empty-range precondition**: either populate the cell manually first or pass `-require-non-empty=false`.
