package sheets

import (
	"context"
	"errors"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"

	"update-google-sheets/src/config"
)

// Summary describes the outcome of an update run.
type Summary struct {
	Ranges         []string
	TotalCells     int64
	TotalRows      int64
	SkippedReason  string
	TemplateSheets []string
	TargetSheets   []string
}

// Update synchronises lookup-derived cells with the given spreadsheet.
func Update(ctx context.Context, cfg config.Config) (Summary, error) {
	var summary Summary

	values := [][]interface{}{{cfg.LookupValue}}

	svc, err := sheets.NewService(ctx, option.WithScopes(sheets.SpreadsheetsScope))
	if err != nil {
		return summary, fmt.Errorf("initialise Sheets service: %w", err)
	}

	ranges, templateSheets, err := loadCachedRanges()
	if err != nil {
		return summary, err
	}
	summary.TemplateSheets = templateSheets
	summary.TargetSheets = uniqueSheetNames(ranges)

	payloads, err := buildPayloads(ctx, svc, cfg.SpreadsheetID, ranges, values)
	if err != nil {
		return summary, err
	}
	if len(payloads) == 0 {
		summary.SkippedReason = "all target cells already contain data"
		return summary, nil
	}

	resp, err := batchUpdate(ctx, svc, cfg.SpreadsheetID, payloads)
	if err != nil {
		return summary, err
	}

	summary.TotalCells = resp.TotalUpdatedCells
	summary.TotalRows = resp.TotalUpdatedRows
	for _, p := range payloads {
		summary.Ranges = append(summary.Ranges, p.Range)
	}

	return summary, nil
}

func buildPayloads(ctx context.Context, svc *sheets.Service, sheetID string, ranges []string, desired [][]interface{}) ([]*sheets.ValueRange, error) {
	var payloads []*sheets.ValueRange
	existingValues, err := fetchRangeValuesBatch(ctx, svc, sheetID, ranges)
	if err != nil {
		return nil, err
	}
	for _, rng := range ranges {
		existing := existingValues[rng]
		merged, needsUpdate := mergeValues(existing, desired)
		if !needsUpdate {
			continue
		}
		payloads = append(payloads, &sheets.ValueRange{
			MajorDimension: "ROWS",
			Range:          rng,
			Values:         merged,
		})
	}
	return payloads, nil
}

func batchUpdate(ctx context.Context, svc *sheets.Service, sheetID string, data []*sheets.ValueRange) (*sheets.BatchUpdateValuesResponse, error) {
	req := &sheets.BatchUpdateValuesRequest{
		ValueInputOption:        "USER_ENTERED",
		IncludeValuesInResponse: true,
		Data:                    data,
	}
	resp, err := svc.Spreadsheets.Values.BatchUpdate(sheetID, req).Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("batch update failed: %w", err)
	}
	return resp, nil
}

func fetchRangeValuesBatch(ctx context.Context, svc *sheets.Service, sheetID string, ranges []string) (map[string][][]interface{}, error) {
	if len(ranges) == 0 {
		return map[string][][]interface{}{}, nil
	}
	call := svc.Spreadsheets.Values.BatchGet(sheetID)
	call = call.Ranges(ranges...)
	resp, err := call.Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("batch fetch current values: %w", err)
	}
	values := make(map[string][][]interface{}, len(ranges))
	for _, vr := range resp.ValueRanges {
		values[vr.Range] = vr.Values
	}
	for _, rng := range ranges {
		if _, ok := values[rng]; !ok {
			values[rng] = nil
		}
	}
	return values, nil
}

func mergeValues(existing, desired [][]interface{}) ([][]interface{}, bool) {
	merged := make([][]interface{}, len(desired))
	var wrote bool
	for r, row := range desired {
		mergedRow := make([]interface{}, len(row))
		for c, val := range row {
			if cellHasValue(existing, r, c) {
				mergedRow[c] = existing[r][c]
				continue
			}
			mergedRow[c] = val
			if strings.TrimSpace(fmt.Sprint(val)) != "" {
				wrote = true
			}
		}
		merged[r] = mergedRow
	}
	return merged, wrote
}

func cellHasValue(values [][]interface{}, row, col int) bool {
	if row >= len(values) {
		return false
	}
	if col >= len(values[row]) {
		return false
	}
	return strings.TrimSpace(fmt.Sprint(values[row][col])) != ""
}

func loadCachedRanges() ([]string, []string, error) {
	ranges, err := ReadRangeCache(DefaultRangeCache)
	if err != nil {
		if errors.Is(err, ErrRangeCacheNotFound) {
			return nil, nil, fmt.Errorf("range cache missing; run go run ./cmd/lookupscan to refresh it: %w", err)
		}
		return nil, nil, err
	}
	sheetNames := uniqueSheetNames(ranges)
	return ranges, sheetNames, nil
}

func ScanWorkbook(path, sheetFilter, lookup string) ([]string, []string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, fmt.Errorf("open config workbook: %w", err)
	}
	defer func() { _ = f.Close() }()

	want := strings.TrimSpace(lookup)
	sheetsList := filterSheets(f.GetSheetList(), sheetFilter)
	if sheetFilter != "" && len(sheetsList) == 0 {
		return nil, nil, fmt.Errorf("sheet %q not found in %s", sheetFilter, path)
	}

	var matches []string
	for _, sheet := range sheetsList {
		rows, err := f.Rows(sheet)
		if err != nil {
			return nil, nil, fmt.Errorf("iterate sheet %s: %w", sheet, err)
		}
		rowIdx := 0
		for rows.Next() {
			rowIdx++
			row, err := rows.Columns()
			if err != nil {
				_ = rows.Close()
				return nil, nil, fmt.Errorf("read row %d in %s: %w", rowIdx, sheet, err)
			}
			for cIdx, cell := range row {
				if strings.TrimSpace(cell) != want {
					continue
				}
				cellName, err := excelize.CoordinatesToCellName(cIdx+1, rowIdx)
				if err != nil {
					_ = rows.Close()
					return nil, nil, fmt.Errorf("build cell name: %w", err)
				}
				matches = append(matches, formatRange(sheet, cellName))
			}
		}
		if err := rows.Close(); err != nil {
			return nil, nil, fmt.Errorf("close sheet reader %s: %w", sheet, err)
		}
	}

	if len(matches) == 0 {
		return nil, nil, fmt.Errorf("value %q not found in %s", lookup, path)
	}
	return matches, sheetsList, nil
}

func filterSheets(all []string, filter string) []string {
	if filter == "" {
		return all
	}
	for _, s := range all {
		if s == filter {
			return []string{filter}
		}
	}
	return nil
}

func formatRange(sheet, cell string) string {
	if strings.ContainsAny(sheet, " !'") {
		return fmt.Sprintf("'%s'!%s", strings.ReplaceAll(sheet, "'", "''"), cell)
	}
	return fmt.Sprintf("%s!%s", sheet, cell)
}

func uniqueSheetNames(ranges []string) []string {
	seen := make(map[string]struct{})
	var names []string
	for _, rng := range ranges {
		name := sheetNameFromRange(rng)
		if name == "" {
			continue
		}
		if _, ok := seen[name]; ok {
			continue
		}
		seen[name] = struct{}{}
		names = append(names, name)
	}
	return names
}

func sheetNameFromRange(rng string) string {
	idx := strings.Index(rng, "!")
	if idx == -1 {
		return ""
	}
	sheet := rng[:idx]
	if strings.HasPrefix(sheet, "'") && strings.HasSuffix(sheet, "'") && len(sheet) >= 2 {
		sheet = sheet[1 : len(sheet)-1]
		sheet = strings.ReplaceAll(sheet, "''", "'")
	}
	return sheet
}
