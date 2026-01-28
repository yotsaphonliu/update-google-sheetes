package sheets

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

const DefaultRangeCache = "cfg/lookup_ranges.csv"

var ErrRangeCacheNotFound = errors.New("range cache not found")

// WriteRangeCache persists the resolved lookup ranges to a CSV file.
func WriteRangeCache(path string, ranges []string) error {
	if len(ranges) == 0 {
		return errors.New("no ranges to cache")
	}
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return fmt.Errorf("ensure cache dir: %w", err)
	}
	f, err := os.Create(path)
	if err != nil {
		return fmt.Errorf("create cache: %w", err)
	}
	defer func() { _ = f.Close() }()
	w := csv.NewWriter(f)
	if err := w.Write([]string{"range"}); err != nil {
		return fmt.Errorf("write header: %w", err)
	}
	for _, rng := range ranges {
		if err := w.Write([]string{rng}); err != nil {
			return fmt.Errorf("write row: %w", err)
		}
	}
	w.Flush()
	if err := w.Error(); err != nil {
		return fmt.Errorf("flush cache: %w", err)
	}
	return nil
}

// ReadRangeCache loads cached lookup ranges from disk.
func ReadRangeCache(path string) ([]string, error) {
	f, err := os.Open(path)
	if err != nil {
		if errors.Is(err, os.ErrNotExist) {
			return nil, ErrRangeCacheNotFound
		}
		return nil, fmt.Errorf("open cache: %w", err)
	}
	defer func() { _ = f.Close() }()
	r := csv.NewReader(f)
	var ranges []string
	for {
		rec, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("read cache: %w", err)
		}
		if len(rec) == 0 {
			continue
		}
		value := strings.TrimSpace(rec[0])
		if value == "" {
			continue
		}
		if len(ranges) == 0 && strings.EqualFold(value, "range") {
			continue
		}
		ranges = append(ranges, value)
	}
	if len(ranges) == 0 {
		return nil, errors.New("range cache is empty")
	}
	return ranges, nil
}
