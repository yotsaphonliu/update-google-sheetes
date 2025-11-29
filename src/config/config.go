package config

import (
	"bufio"
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"gopkg.in/yaml.v3"
)

const (
	DefaultPath     = "cfg/config.yaml"
	defaultWorkbook = "cfg/Schedule.xlsx"
)

// Config captures the data needed to perform an update.
type Config struct {
	SpreadsheetID string `yaml:"spreadsheet_id"`
	Workbook      string `yaml:"config_xlsx"`
	SheetFilter   string `yaml:"config_sheet"`
	LookupValue   string `yaml:"lookup_value"`
}

// Load reads the config file or falls back to interactive prompts.
func Load(path string) (Config, error) {
	data, err := os.ReadFile(path)
	if err == nil {
		var cfg Config
		if err := yaml.Unmarshal(data, &cfg); err != nil {
			return Config{}, fmt.Errorf("parse %s: %w", path, err)
		}
		return cfg, nil
	}
	if !errors.Is(err, os.ErrNotExist) {
		return Config{}, fmt.Errorf("read %s: %w", path, err)
	}
	fmt.Printf("%s not found; switching to interactive setup.\n\n", path)
	return prompt(os.Stdin)
}

// Validate normalises defaults and checks required fields.
func (c *Config) Validate() error {
	c.SpreadsheetID = strings.TrimSpace(c.SpreadsheetID)
	c.Workbook = strings.TrimSpace(c.Workbook)
	if c.Workbook == "" {
		c.Workbook = defaultWorkbook
	}
	c.SheetFilter = strings.TrimSpace(c.SheetFilter)
	c.LookupValue = strings.TrimSpace(c.LookupValue)

	if c.SpreadsheetID == "" {
		return errors.New("spreadsheet_id is required")
	}
	if c.LookupValue == "" {
		return errors.New("lookup_value is required")
	}
	if _, err := os.Stat(c.Workbook); err != nil {
		return fmt.Errorf("access %s: %w", c.Workbook, err)
	}
	return nil
}

func prompt(input io.Reader) (Config, error) {
	r := bufio.NewReader(input)
	spreadsheetID, err := promptRequired(r, "Google Spreadsheet ID:")
	if err != nil {
		return Config{}, err
	}
	workbook, err := promptFile(r, "Path to the Excel workbook (default cfg/Schedule.xlsx):", defaultWorkbook)
	if err != nil {
		return Config{}, err
	}
	sheetFilter, err := promptLine(r, "Limit lookup to a single sheet (press Enter for all):")
	if err != nil {
		return Config{}, err
	}
	lookup, err := promptRequired(r, "Lookup value to search for:")
	if err != nil {
		return Config{}, err
	}
	fmt.Println()
	fmt.Println("Tip: store these answers in config.yaml to skip the wizard next time.")

	return Config{
		SpreadsheetID: spreadsheetID,
		Workbook:      workbook,
		SheetFilter:   strings.TrimSpace(sheetFilter),
		LookupValue:   lookup,
	}, nil
}

func promptRequired(r *bufio.Reader, question string) (string, error) {
	for {
		answer, err := promptLine(r, question)
		if err != nil {
			return "", err
		}
		answer = strings.TrimSpace(answer)
		if answer != "" {
			return answer, nil
		}
		fmt.Println("Please enter a value.")
	}
}

func promptFile(r *bufio.Reader, question, defaultPath string) (string, error) {
	for {
		answer, err := promptLine(r, question)
		if err != nil {
			return "", err
		}
		answer = strings.TrimSpace(answer)
		if answer == "" {
			answer = defaultPath
		}
		if _, statErr := os.Stat(answer); statErr == nil {
			return answer, nil
		}
		fmt.Printf("File %q is not accessible.\n", answer)
	}
}

func promptLine(r *bufio.Reader, question string) (string, error) {
	fmt.Print(question + " ")
	line, err := r.ReadString('\n')
	if err != nil && !errors.Is(err, io.EOF) {
		return "", err
	}
	return strings.TrimSpace(line), nil
}
