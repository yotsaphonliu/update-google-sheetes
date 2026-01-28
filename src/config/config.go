package config

import (
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"time"

	survey "github.com/AlecAivazis/survey/v2"
	"github.com/spf13/viper"
	"gopkg.in/yaml.v3"
)

const (
	DefaultPath     = "cfg/config.yaml"
	DefaultWorkbook = "cfg/Schedule.xlsx"
)

// Config captures the data needed to perform an update.
type Config struct {
	SpreadsheetID string `yaml:"spreadsheet_id" mapstructure:"spreadsheet_id"`
	SheetFilter   string `yaml:"config_sheet" mapstructure:"config_sheet"`
	LookupValue   string `yaml:"lookup_value" mapstructure:"lookup_value"`
	ScheduleTime  string `yaml:"schedule_time" mapstructure:"schedule_time"`
}

// Load reads the config file or falls back to interactive prompts.
func Load(path string) (Config, error) {
	v := viper.New()
	v.SetConfigFile(path)
	v.SetConfigType("yaml")
	if err := v.ReadInConfig(); err != nil {
		var notFound viper.ConfigFileNotFoundError
		if errors.As(err, &notFound) {
			fmt.Printf("%s not found; switching to interactive setup.\n\n", path)
			return prompt(), nil
		}
		return Config{}, fmt.Errorf("read %s: %w", path, err)
	}
	var cfg Config
	if err := v.Unmarshal(&cfg); err != nil {
		return Config{}, fmt.Errorf("parse %s: %w", path, err)
	}
	return cfg, nil
}

// Validate normalises defaults and checks required fields.
func (c *Config) Validate() error {
	c.SpreadsheetID = strings.TrimSpace(c.SpreadsheetID)
	c.SheetFilter = strings.TrimSpace(c.SheetFilter)
	c.LookupValue = strings.TrimSpace(c.LookupValue)
	c.ScheduleTime = strings.TrimSpace(c.ScheduleTime)

	if c.SpreadsheetID == "" {
		return errors.New("spreadsheet_id is required")
	}
	if c.LookupValue == "" {
		return errors.New("lookup_value is required")
	}
	if _, err := os.Stat(DefaultWorkbook); err != nil {
		return fmt.Errorf("access %s: %w", DefaultWorkbook, err)
	}
	if c.ScheduleTime != "" {
		if _, err := c.ScheduleTimeOffset(); err != nil {
			return err
		}
	}
	return nil
}

// ScheduleTimeOffset returns the configured number of seconds since midnight.
func (c Config) ScheduleTimeOffset() (time.Duration, error) {
	if c.ScheduleTime == "" {
		return 0, errors.New("schedule_time is empty")
	}
	if d, err := parseClock(c.ScheduleTime, "15:04:05"); err == nil {
		return d, nil
	}
	return parseClock(c.ScheduleTime, "15:04")
}

func parseClock(value, layout string) (time.Duration, error) {
	parsed, err := time.Parse(layout, value)
	if err != nil {
		return 0, fmt.Errorf("schedule_time must use HH:MM or HH:MM:SS (24h) format")
	}
	return time.Duration(parsed.Hour())*time.Hour + time.Duration(parsed.Minute())*time.Minute + time.Duration(parsed.Second())*time.Second, nil
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
	if err := survey.AskOne(&survey.Input{Message: "Run after Bangkok time (HH:MM, Enter to skip)"}, &cfg.ScheduleTime); err != nil {
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
