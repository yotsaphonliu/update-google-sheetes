package main

import (
	"context"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// cliOptions describes all user controllable switches.
type cliOptions struct {
	spreadsheetID   string
	targetRange     string
	valueInputMode  string
	majorDimension  string
	inlineValues    string
	valuesFile      string
	configExcel     string
	lookupValue     string
	requireNonEmpty bool
}

func main() {
	opts := parseFlags()
	if err := run(context.Background(), opts); err != nil {
		exitErr("%v", err)
	}
}

// run drives the overall Sheet update flow so main() stays focused on exit handling.
func run(ctx context.Context, opts cliOptions) error {
	values, err := loadMatrix(opts)
	if err != nil {
		return fmt.Errorf("load values: %w", err)
	}

	if len(values) == 0 {
		return errors.New("no values decoded; provide at least one row")
	}

	svc, err := newSheetsService(ctx)
	if err != nil {
		return fmt.Errorf("initialise Sheets service: %w", err)
	}

	if opts.requireNonEmpty {
		if err := ensureRangeHasValue(ctx, svc, opts.spreadsheetID, opts.targetRange); err != nil {
			return fmt.Errorf("precondition failed: %w", err)
		}
	}

	resp, err := updateValues(ctx, svc, opts, values)
	if err != nil {
		return err
	}

	fmt.Printf("Updated %d cells across %d rows in %q.\n", resp.UpdatedCells, resp.UpdatedRows, opts.targetRange)
	if resp.UpdatedData != nil && len(resp.UpdatedData.Values) > 0 {
		encoded, _ := json.Marshal(resp.UpdatedData.Values)
		fmt.Printf("Values now stored in range: %s\n", string(encoded))
	}

	return nil
}

func newSheetsService(ctx context.Context) (*sheets.Service, error) {
	return sheets.NewService(ctx, option.WithScopes(sheets.SpreadsheetsScope))
}

// updateValues performs the actual API call and returns the Sheets response for logging.
func updateValues(ctx context.Context, svc *sheets.Service, opts cliOptions, values [][]interface{}) (*sheets.UpdateValuesResponse, error) {
	payload := &sheets.ValueRange{
		MajorDimension: opts.majorDimension,
		Range:          opts.targetRange,
		Values:         values,
	}

	call := svc.Spreadsheets.
		Values.
		Update(opts.spreadsheetID, opts.targetRange, payload).
		ValueInputOption(opts.valueInputMode).
		IncludeValuesInResponse(true)

	resp, err := call.Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("update call failed: %w", err)
	}

	return resp, nil
}

func parseFlags() cliOptions {
	var opts cliOptions

	flag.StringVar(&opts.spreadsheetID, "spreadsheet", "", "ID of the Google Sheet to update")
	flag.StringVar(&opts.targetRange, "range", "", "Target range in A1 notation (e.g. Sheet1!A2:C4)")
	flag.StringVar(&opts.valueInputMode, "value-input", "USER_ENTERED", "Value input option: RAW or USER_ENTERED")
	flag.StringVar(&opts.majorDimension, "dimension", "ROWS", "Major dimension to use when writing (ROWS or COLUMNS)")
	flag.StringVar(&opts.inlineValues, "values", "", "JSON encoded 2D array of values (overrides stdin)")
	flag.StringVar(&opts.valuesFile, "values-file", "", "Path to a JSON file with the 2D values array")
	flag.StringVar(&opts.configExcel, "config-xlsx", "", "Path to an Excel config file used to derive the range")
	flag.StringVar(&opts.lookupValue, "lookup-value", "", "Exact cell value to search for inside the Excel config")
	flag.BoolVar(&opts.requireNonEmpty, "require-non-empty", true, "Only update when the current Google Sheet range already contains data")

	flag.Usage = func() {
		fmt.Fprintf(flag.CommandLine.Output(), "Usage: %s [flags] < optional JSON from stdin >\n", os.Args[0])
		fmt.Fprintln(flag.CommandLine.Output(), "\nFlags:")
		flag.PrintDefaults()
		fmt.Fprintln(flag.CommandLine.Output(), "\nProvide the values as a JSON array of arrays. Example: [[\"Name\",\"Score\"],[\"Mia\",42]]")
		fmt.Fprintln(flag.CommandLine.Output(), "The program relies on Application Default Credentials, e.g. set GOOGLE_APPLICATION_CREDENTIALS to a service account key file.")
	}

	flag.Parse()

	if err := opts.validateAndPopulate(); err != nil {
		exitErr("%v", err)
	}

	return opts
}

// validateAndPopulate makes sure mandatory flags are present and derives dependent values.
func (opts *cliOptions) validateAndPopulate() error {
	if opts.spreadsheetID == "" {
		return errors.New("-spreadsheet is required")
	}

	if err := opts.fillRangeIfMissing(); err != nil {
		return err
	}

	opts.valueInputMode = strings.ToUpper(opts.valueInputMode)
	if opts.valueInputMode != "RAW" && opts.valueInputMode != "USER_ENTERED" {
		return errors.New("-value-input must be RAW or USER_ENTERED")
	}

	opts.majorDimension = strings.ToUpper(opts.majorDimension)
	if opts.majorDimension != "ROWS" && opts.majorDimension != "COLUMNS" {
		return errors.New("-dimension must be ROWS or COLUMNS")
	}

	if opts.inlineValues != "" && opts.valuesFile != "" {
		return errors.New("specify -values or -values-file, not both")
	}

	return nil
}

func (opts *cliOptions) fillRangeIfMissing() error {
	if opts.targetRange != "" {
		return nil
	}
	if opts.configExcel == "" || opts.lookupValue == "" {
		return errors.New("provide -range or both -config-xlsx and -lookup-value")
	}
	rng, err := deriveRangeFromExcel(opts.configExcel, opts.lookupValue)
	if err != nil {
		return fmt.Errorf("derive range from config: %w", err)
	}
	fmt.Fprintf(os.Stderr, "Located %q in %s -> using range %s\n", opts.lookupValue, opts.configExcel, rng)
	opts.targetRange = rng
	return nil
}

// ensureRangeHasValue guards destructive updates by confirming we are not overwriting emptiness.
func ensureRangeHasValue(ctx context.Context, svc *sheets.Service, sheetID, rng string) error {
	resp, err := svc.Spreadsheets.Values.Get(sheetID, rng).Context(ctx).Do()
	if err != nil {
		return fmt.Errorf("fetch current value: %w", err)
	}
	if !hasAnyValue(resp.Values) {
		return fmt.Errorf("range %s is empty", rng)
	}
	return nil
}

func hasAnyValue(values [][]interface{}) bool {
	for _, row := range values {
		for _, cell := range row {
			if fmt.Sprint(cell) != "" {
				return true
			}
		}
	}
	return false
}

// deriveRangeFromExcel finds the lookup value and returns an A1 range pointing at it.
func deriveRangeFromExcel(path, lookup string) (string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return "", fmt.Errorf("open config workbook: %w", err)
	}
	defer func() {
		_ = f.Close()
	}()

	want := normalizeCellText(lookup)
	for _, sheet := range f.GetSheetList() {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return "", fmt.Errorf("read sheet %s: %w", sheet, err)
		}
		for rIdx, row := range rows {
			for cIdx, cell := range row {
				if normalizeCellText(cell) == want {
					cellRef, err := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
					if err != nil {
						return "", fmt.Errorf("build cell name: %w", err)
					}
					return formatSheetRange(sheet, cellRef), nil
				}
			}
		}
	}

	return "", fmt.Errorf("value %q not found in %s", lookup, path)
}

func formatSheetRange(sheet, cell string) string {
	if strings.ContainsAny(sheet, " !'") {
		escaped := strings.ReplaceAll(sheet, "'", "''")
		return fmt.Sprintf("'%s'!%s", escaped, cell)
	}
	return fmt.Sprintf("%s!%s", sheet, cell)
}

func normalizeCellText(s string) string {
	return strings.TrimSpace(s)
}

// loadMatrix fetches the JSON-encoded 2D array and returns it as Go values.
func loadMatrix(opts cliOptions) ([][]interface{}, error) {
	raw, err := readRawData(opts)
	if err != nil {
		return nil, err
	}

	if len(strings.TrimSpace(string(raw))) == 0 {
		return nil, errors.New("no data provided")
	}

	var matrix [][]interface{}
	if err := json.Unmarshal(raw, &matrix); err != nil {
		return nil, fmt.Errorf("decode JSON values: %w", err)
	}

	return matrix, nil
}

// readRawData abstracts over inline flag, file, or stdin input.
func readRawData(opts cliOptions) ([]byte, error) {
	switch {
	case opts.inlineValues != "":
		return []byte(opts.inlineValues), nil
	case opts.valuesFile != "":
		return os.ReadFile(opts.valuesFile)
	default:
		stat, err := os.Stdin.Stat()
		if err != nil {
			return nil, fmt.Errorf("stat stdin: %w", err)
		}
		if stat.Mode()&os.ModeCharDevice != 0 {
			return nil, errors.New("no stdin input; use -values or -values-file")
		}
		data, err := io.ReadAll(os.Stdin)
		if err != nil {
			return nil, fmt.Errorf("read stdin: %w", err)
		}
		return data, nil
	}
}

func exitErr(msg string, args ...interface{}) {
	fmt.Fprintf(os.Stderr, msg+"\n", args...)
	os.Exit(1)
}
