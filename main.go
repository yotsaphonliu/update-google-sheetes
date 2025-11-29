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
	"time"

	"github.com/xuri/excelize/v2"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

// cliOptions describes all user controllable switches.
type cliOptions struct {
	spreadsheetID   string
	targetRange     string
	targetRanges    []string
	valueInputMode  string
	majorDimension  string
	inlineValues    string
	valuesFile      string
	configExcel     string
	configSheet     string
	lookupValue     string
	requireNonEmpty bool
}

type updateSummary struct {
	Ranges        []string
	TotalCells    int64
	TotalRows     int64
	SkippedReason string
}

func main() {
	opts := parseFlags()
	logger, err := newLogger()
	if err != nil {
		exitErr("initialise logger: %v", err)
	}
	defer func() {
		_ = logger.Sync()
	}()

	summary, err := run(context.Background(), opts)
	if err != nil {
		logger.Error("update failed", zap.Error(err))
		exitErr("%v", err)
	}

	if summary.SkippedReason != "" {
		logger.Info("no updates performed", zap.String("reason", summary.SkippedReason))
		return
	}

	logger.Info(
		"update complete",
		zap.Strings("ranges", summary.Ranges),
		zap.Int64("rows", summary.TotalRows),
		zap.Int64("cells", summary.TotalCells),
	)
}

func newLogger() (*zap.Logger, error) {
	loc, err := time.LoadLocation("Asia/Bangkok")
	if err != nil {
		return nil, fmt.Errorf("load timezone: %w", err)
	}
	cfg := zap.NewProductionConfig()
	cfg.Encoding = "console"
	cfg.EncoderConfig = zap.NewProductionEncoderConfig()
	cfg.EncoderConfig.EncodeLevel = zapcore.CapitalColorLevelEncoder
	cfg.EncoderConfig.EncodeTime = func(t time.Time, enc zapcore.PrimitiveArrayEncoder) {
		enc.AppendString(t.In(loc).Format(time.RFC3339))
	}
	return cfg.Build()
}

// run drives the overall Sheet update flow so main() stays focused on exit handling.
func run(ctx context.Context, opts cliOptions) (updateSummary, error) {
	var summary updateSummary
	values, err := loadMatrix(opts)
	if err != nil {
		return summary, fmt.Errorf("load values: %w", err)
	}

	if len(values) == 0 {
		return summary, errors.New("no values decoded; provide at least one row")
	}

	svc, err := newSheetsService(ctx)
	if err != nil {
		return summary, fmt.Errorf("initialise Sheets service: %w", err)
	}

	if len(opts.targetRanges) == 0 && opts.targetRange != "" {
		// Backwards compatibility when validation only populated targetRange.
		opts.targetRanges = []string{opts.targetRange}
	}
	if len(opts.targetRanges) == 0 {
		return summary, errors.New("no target ranges resolved")
	}

	var (
		payloads        []*sheets.ValueRange
		requestedRanges []string
	)
	for _, rng := range opts.targetRanges {
		rangeValues := values
		if opts.requireNonEmpty {
			existing, err := fetchRangeValues(ctx, svc, opts.spreadsheetID, rng)
			if err != nil {
				return summary, fmt.Errorf("precondition failed for %s: %w", rng, err)
			}
			merged, hasWrites := preserveExistingValues(existing, values)
			if !hasWrites {
				continue
			}
			rangeValues = merged
		}

		payloads = append(payloads, &sheets.ValueRange{
			MajorDimension: opts.majorDimension,
			Range:          rng,
			Values:         rangeValues,
		})
		requestedRanges = append(requestedRanges, rng)
	}

	if len(payloads) == 0 {
		if opts.requireNonEmpty {
			summary.SkippedReason = "require-non-empty=true and destination cells already populated"
			return summary, nil
		}
		return summary, errors.New("no target ranges resolved")
	}

	resp, err := batchUpdateValues(ctx, svc, opts, payloads)
	if err != nil {
		return summary, err
	}

	summary.TotalCells = resp.TotalUpdatedCells
	summary.TotalRows = resp.TotalUpdatedRows
	for idx, updateResp := range resp.Responses {
		rng := updateResp.UpdatedRange
		if rng == "" && idx < len(requestedRanges) {
			rng = requestedRanges[idx]
		}
		if rng == "" {
			rng = fmt.Sprintf("response_%d", idx)
		}
		summary.Ranges = append(summary.Ranges, rng)
	}
	if len(summary.Ranges) == 0 {
		summary.Ranges = requestedRanges
	}

	return summary, nil
}

func newSheetsService(ctx context.Context) (*sheets.Service, error) {
	return sheets.NewService(ctx, option.WithScopes(sheets.SpreadsheetsScope))
}

// batchUpdateValues performs a single Sheets API batch update for all payloads.
func batchUpdateValues(ctx context.Context, svc *sheets.Service, opts cliOptions, data []*sheets.ValueRange) (*sheets.BatchUpdateValuesResponse, error) {
	req := &sheets.BatchUpdateValuesRequest{
		ValueInputOption:        opts.valueInputMode,
		IncludeValuesInResponse: true,
		Data:                    data,
	}

	resp, err := svc.Spreadsheets.
		Values.
		BatchUpdate(opts.spreadsheetID, req).
		Context(ctx).
		Do()
	if err != nil {
		return nil, fmt.Errorf("batch update failed: %w", err)
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
	flag.StringVar(&opts.configSheet, "config-sheet", "", "Limit Excel lookup to this sheet name")
	flag.StringVar(&opts.lookupValue, "lookup-value", "", "Exact cell value to search for inside the Excel config (all matches are updated)")
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

	if err := opts.populateRanges(); err != nil {
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

func (opts *cliOptions) populateRanges() error {
	switch {
	case opts.targetRange != "":
		opts.targetRanges = []string{opts.targetRange}
		return nil
	case opts.configExcel == "" || opts.lookupValue == "":
		return errors.New("provide -range or both -config-xlsx and -lookup-value")
	default:
		rngs, err := deriveRangesFromExcel(opts.configExcel, opts.configSheet, opts.lookupValue)
		if err != nil {
			return fmt.Errorf("derive range from config: %w", err)
		}
		if len(rngs) == 0 {
			return fmt.Errorf("value %q not found in %s", opts.lookupValue, opts.configExcel)
		}
		opts.targetRanges = rngs
		return nil
	}
}

// fetchRangeValues reads the target range so we can decide whether to write over each cell.
func fetchRangeValues(ctx context.Context, svc *sheets.Service, sheetID, rng string) ([][]interface{}, error) {
	resp, err := svc.Spreadsheets.Values.Get(sheetID, rng).Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("fetch current value: %w", err)
	}
	return resp.Values, nil
}

// preserveExistingValues keeps non-empty current cells intact by substituting them back into the payload.
func preserveExistingValues(existing, incoming [][]interface{}) ([][]interface{}, bool) {
	merged := make([][]interface{}, len(incoming))
	var hasWrites bool
	for r, row := range incoming {
		mergedRow := make([]interface{}, len(row))
		for c, val := range row {
			if cellHasValue(existing, r, c) {
				mergedRow[c] = existing[r][c]
				continue
			}
			mergedRow[c] = val
			if fmt.Sprint(val) != "" {
				hasWrites = true
			}
		}
		merged[r] = mergedRow
	}
	return merged, hasWrites
}

func cellHasValue(values [][]interface{}, row, col int) bool {
	if row >= len(values) {
		return false
	}
	if col >= len(values[row]) {
		return false
	}
	return fmt.Sprint(values[row][col]) != ""
}

// deriveRangesFromExcel finds every cell that matches the lookup value and returns their A1 ranges.
// When sheetFilter is set, only that worksheet is scanned.
func deriveRangesFromExcel(path, sheetFilter, lookup string) ([]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open config workbook: %w", err)
	}
	defer func() {
		_ = f.Close()
	}()

	want := normalizeCellText(lookup)
	var matches []string
	sheets := f.GetSheetList()
	if sheetFilter != "" {
		found := false
		for _, s := range sheets {
			if s == sheetFilter {
				found = true
				break
			}
		}
		if !found {
			return nil, fmt.Errorf("sheet %q not found in %s", sheetFilter, path)
		}
		sheets = []string{sheetFilter}
	}
	for _, sheet := range sheets {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return nil, fmt.Errorf("read sheet %s: %w", sheet, err)
		}
		for rIdx, row := range rows {
			for cIdx, cell := range row {
				if normalizeCellText(cell) == want {
					cellRef, err := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
					if err != nil {
						return nil, fmt.Errorf("build cell name: %w", err)
					}
					matches = append(matches, formatSheetRange(sheet, cellRef))
				}
			}
		}
	}

	if len(matches) == 0 {
		return nil, fmt.Errorf("value %q not found in %s", lookup, path)
	}
	return matches, nil
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
