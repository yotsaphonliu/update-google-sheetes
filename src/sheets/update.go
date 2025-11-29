package sheets

import (
	"context"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"

	"update-google-sheets/src/config"
)

// Summary describes the outcome of an update run.
type Summary struct {
	Ranges        []string
	TotalCells    int64
	TotalRows     int64
	SkippedReason string
}

// Update synchronises lookup-derived cells with the given spreadsheet.
func Update(ctx context.Context, cfg config.Config) (Summary, error) {
	var summary Summary

	values := [][]interface{}{{cfg.LookupValue}}

	svc, err := sheets.NewService(ctx, option.WithScopes(sheets.SpreadsheetsScope))
	if err != nil {
		return summary, fmt.Errorf("initialise Sheets service: %w", err)
	}

	ranges, err := deriveRangesFromExcel(config.DefaultWorkbook, cfg.SheetFilter, cfg.LookupValue)
	if err != nil {
		return summary, err
	}

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
	for _, rng := range ranges {
		existing, err := fetchRangeValues(ctx, svc, sheetID, rng)
		if err != nil {
			return nil, fmt.Errorf("precondition failed for %s: %w", rng, err)
		}
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

func fetchRangeValues(ctx context.Context, svc *sheets.Service, sheetID, rng string) ([][]interface{}, error) {
	resp, err := svc.Spreadsheets.Values.Get(sheetID, rng).Context(ctx).Do()
	if err != nil {
		return nil, fmt.Errorf("fetch current value: %w", err)
	}
	return resp.Values, nil
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

func deriveRangesFromExcel(path, sheetFilter, lookup string) ([]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open config workbook: %w", err)
	}
	defer func() { _ = f.Close() }()

	want := strings.TrimSpace(lookup)
	sheetsList := filterSheets(f.GetSheetList(), sheetFilter)
	if sheetFilter != "" && len(sheetsList) == 0 {
		return nil, fmt.Errorf("sheet %q not found in %s", sheetFilter, path)
	}

	var matches []string
	for _, sheet := range sheetsList {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return nil, fmt.Errorf("read sheet %s: %w", sheet, err)
		}
		for rIdx, row := range rows {
			for cIdx, cell := range row {
				if strings.TrimSpace(cell) != want {
					continue
				}
				cellName, err := excelize.CoordinatesToCellName(cIdx+1, rIdx+1)
				if err != nil {
					return nil, fmt.Errorf("build cell name: %w", err)
				}
				matches = append(matches, formatRange(sheet, cellName))
			}
		}
	}

	if len(matches) == 0 {
		return nil, fmt.Errorf("value %q not found in %s", lookup, path)
	}
	return matches, nil
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
