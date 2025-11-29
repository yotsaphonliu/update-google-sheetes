package config

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"

	survey "github.com/AlecAivazis/survey/v2"
	"gopkg.in/yaml.v3"
)

const (
	DefaultPath     = "cfg/config.yaml"
	DefaultWorkbook = "cfg/Schedule.xlsx"
)

// Config captures the data needed to perform an update.
type Config struct {
	SpreadsheetID string `yaml:"spreadsheet_id"`
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
	return prompt(), nil
}

// Validate normalises defaults and checks required fields.
func (c *Config) Validate() error {
	c.SpreadsheetID = strings.TrimSpace(c.SpreadsheetID)
	c.SheetFilter = strings.TrimSpace(c.SheetFilter)
	c.LookupValue = strings.TrimSpace(c.LookupValue)

	if c.SpreadsheetID == "" {
		return errors.New("spreadsheet_id is required")
	}
	if c.LookupValue == "" {
		return errors.New("lookup_value is required")
	}
	if _, err := os.Stat(DefaultWorkbook); err != nil {
		return fmt.Errorf("access %s: %w", DefaultWorkbook, err)
	}
	return nil
}

func prompt() Config {
	var cfg Config
	if err := survey.AskOne(&survey.Input{Message: "Google Spreadsheet ID"}, &cfg.SpreadsheetID, survey.WithValidator(survey.Required)); err != nil {
		fmt.Fprintln(os.Stderr, "input cancelled:", err)
		os.Exit(1)
	}
	if err := survey.AskOne(&survey.Input{Message: "Limit lookup to a single sheet (press Enter for all)"}, &cfg.SheetFilter); err != nil {
		fmt.Fprintln(os.Stderr, "input cancelled:", err)
		os.Exit(1)
	}
	if err := survey.AskOne(&survey.Input{Message: "Lookup value to search for"}, &cfg.LookupValue, survey.WithValidator(survey.Required)); err != nil {
		fmt.Fprintln(os.Stderr, "input cancelled:", err)
		os.Exit(1)
	}
	fmt.Println()
	fmt.Println("Tip: store these answers in config.yaml to skip the wizard next time.")
	return cfg
}

// Write saves the configuration and optionally copies a workbook into place.
func Write(cfg Config, workbookSource string) error {
	if workbookSource != "" {
		if err := copyFile(workbookSource, DefaultWorkbook); err != nil {
			return fmt.Errorf("copy workbook: %w", err)
		}
	}
	if err := cfg.Validate(); err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(DefaultPath), 0o755); err != nil {
		return fmt.Errorf("ensure config dir: %w", err)
	}
	data, err := yaml.Marshal(cfg)
	if err != nil {
		return fmt.Errorf("encode config: %w", err)
	}
	return os.WriteFile(DefaultPath, data, 0o644)
}

func copyFile(src, dest string) error {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer func() { _ = in.Close() }()
	if err := os.MkdirAll(filepath.Dir(dest), 0o755); err != nil {
		return err
	}
	out, err := os.Create(dest)
	if err != nil {
		return err
	}
	defer func() { _ = out.Close() }()
	if _, err := io.Copy(out, in); err != nil {
		return err
	}
	return out.Sync()
}
