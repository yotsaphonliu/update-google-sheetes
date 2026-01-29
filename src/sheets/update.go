package sheets

import (
	"context"
	"errors"
	"fmt"
	"sort"
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

	svc, err := sheets.NewService(ctx, option.WithScopes(sheets.SpreadsheetsScope))
	if err != nil {
		return summary, fmt.Errorf("initialise Sheets service: %w", err)
	}

	ranges, templateSheets, err := loadCachedRanges()
	if err != nil {
		return summary, err
	}
	summary.TemplateSheets = templateSheets
	segments, err := compressRanges(ranges)
	if err != nil {
		return summary, err
	}
	segmentRefs := make([]string, len(segments))
	for i, seg := range segments {
		segmentRefs[i] = seg.Reference
	}
	summary.TargetSheets = uniqueSheetNames(segmentRefs)

	payloads, err := buildPayloads(ctx, svc, cfg.SpreadsheetID, segments, cfg.LookupValue)
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

func buildPayloads(ctx context.Context, svc *sheets.Service, sheetID string, segments []rangeSegment, value interface{}) ([]*sheets.ValueRange, error) {
	var payloads []*sheets.ValueRange
	var rangeRefs []string
	for _, seg := range segments {
		rangeRefs = append(rangeRefs, seg.Reference)
	}
	existingValues, err := fetchRangeValuesBatch(ctx, svc, sheetID, rangeRefs)
	if err != nil {
		return nil, err
	}
	for _, seg := range segments {
		existing := existingValues[seg.Reference]
		desired := fillDesiredMatrix(value, seg.Rows, seg.Cols)
		merged, needsUpdate := mergeValues(existing, desired)
		if !needsUpdate {
			continue
		}
		payloads = append(payloads, &sheets.ValueRange{
			MajorDimension: "ROWS",
			Range:          seg.Reference,
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

type rangeSegment struct {
	Reference string
	Rows      int
	Cols      int
}

func compressRanges(ranges []string) ([]rangeSegment, error) {
	if len(ranges) == 0 {
		return nil, errors.New("no ranges provided")
	}
	type columnGroup struct {
		sheet string
		col   int
		rows  []int
	}
	groups := make(map[string]*columnGroup)
	for _, rng := range ranges {
		sheet := sheetNameFromRange(rng)
		if sheet == "" {
			return nil, fmt.Errorf("invalid range %q", rng)
		}
		cellRef, err := cellReference(rng)
		if err != nil {
			return nil, err
		}
		col, row, err := excelize.CellNameToCoordinates(cellRef)
		if err != nil {
			return nil, fmt.Errorf("parse %s: %w", rng, err)
		}
		key := fmt.Sprintf("%s|%d", sheet, col)
		grp, ok := groups[key]
		if !ok {
			grp = &columnGroup{sheet: sheet, col: col}
			groups[key] = grp
		}
		grp.rows = append(grp.rows, row)
	}
	var segments []rangeSegment
	keys := make([]string, 0, len(groups))
	for key := range groups {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, key := range keys {
		grp := groups[key]
		sort.Ints(grp.rows)
		start := grp.rows[0]
		prev := grp.rows[0]
		for i := 1; i < len(grp.rows); i++ {
			current := grp.rows[i]
			if current == prev+1 {
				prev = current
				continue
			}
			segment, err := buildSegment(grp.sheet, grp.col, start, prev)
			if err != nil {
				return nil, err
			}
			segments = append(segments, segment)
			start = current
			prev = current
		}
		segment, err := buildSegment(grp.sheet, grp.col, start, prev)
		if err != nil {
			return nil, err
		}
		segments = append(segments, segment)
	}
	return segments, nil
}

func buildSegment(sheet string, col, startRow, endRow int) (rangeSegment, error) {
	startCell, err := excelize.CoordinatesToCellName(col, startRow)
	if err != nil {
		return rangeSegment{}, fmt.Errorf("build start cell: %w", err)
	}
	ref := startCell
	rows := endRow - startRow + 1
	if rows < 1 {
		rows = 1
	}
	if rows > 1 {
		endCell, err := excelize.CoordinatesToCellName(col, endRow)
		if err != nil {
			return rangeSegment{}, fmt.Errorf("build end cell: %w", err)
		}
		ref = fmt.Sprintf("%s:%s", startCell, endCell)
	}
	return rangeSegment{
		Reference: formatRange(sheet, ref),
		Rows:      rows,
		Cols:      1,
	}, nil
}

func cellReference(rng string) (string, error) {
	idx := strings.Index(rng, "!")
	if idx == -1 {
		return "", fmt.Errorf("invalid range %q", rng)
	}
	return rng[idx+1:], nil
}

func fillDesiredMatrix(value interface{}, rows, cols int) [][]interface{} {
	matrix := make([][]interface{}, rows)
	for r := 0; r < rows; r++ {
		row := make([]interface{}, cols)
		for c := 0; c < cols; c++ {
			row[c] = value
		}
		matrix[r] = row
	}
	return matrix
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
