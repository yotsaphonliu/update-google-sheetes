package main

import (
	"errors"
	"flag"
	"fmt"
	"log"
	"os"
	"os/exec"
	"strings"

	survey "github.com/AlecAivazis/survey/v2"

	"update-google-sheets/src/config"
)

func main() {
	existing, _ := config.Load(config.DefaultPath)
	nonInteractive := flag.Bool("non-interactive", false, "Use flags instead of prompts")
	spreadsheet := flag.String("spreadsheet", existing.SpreadsheetID, "Spreadsheet ID")
	sheetFilter := flag.String("sheet", existing.SheetFilter, "Sheet name filter")
	lookup := flag.String("lookup", existing.LookupValue, "Lookup value")
	workbookSrc := flag.String("workbook-src", "", "Path to workbook to copy into cfg (blank keeps existing)")
	flag.Parse()

	if *nonInteractive {
		cfg := config.Config{
			SpreadsheetID: strings.TrimSpace(*spreadsheet),
			SheetFilter:   strings.TrimSpace(*sheetFilter),
			LookupValue:   strings.TrimSpace(*lookup),
		}
		if cfg.SpreadsheetID == "" || cfg.LookupValue == "" {
			log.Fatal("provide -spreadsheet and -lookup")
		}
		copySrc := strings.TrimSpace(*workbookSrc)
		if err := writeConfig(cfg, copySrc); err != nil {
			log.Fatal(err)
		}
		log.Println("Configuration updated at", config.DefaultPath)
		return
	}

	runInteractive()
}

func runInteractive() {
	existing, _ := config.Load(config.DefaultPath)

	prompt := &survey.Input{Message: "Google Spreadsheet ID", Default: existing.SpreadsheetID}
	var spreadsheetID string
	if err := survey.AskOne(prompt, &spreadsheetID); err != nil {
		log.Fatal(err)
	}

	prompt = &survey.Input{Message: "Sheet filter", Default: existing.SheetFilter}
	var sheetFilter string
	if err := survey.AskOne(prompt, &sheetFilter); err != nil {
		log.Fatal(err)
	}

	prompt = &survey.Input{Message: "Lookup value", Default: existing.LookupValue}
	var lookupValue string
	if err := survey.AskOne(prompt, &lookupValue, survey.WithValidator(survey.Required)); err != nil {
		log.Fatal(err)
	}

	workbookSrc, err := chooseWorkbookInteractive()
	if err != nil {
		log.Fatal(err)
	}

	cfg := config.Config{
		SpreadsheetID: strings.TrimSpace(spreadsheetID),
		SheetFilter:   strings.TrimSpace(sheetFilter),
		LookupValue:   strings.TrimSpace(lookupValue),
	}

	if err := writeConfig(cfg, workbookSrc); err != nil {
		log.Fatal(err)
	}
	log.Println("Configuration updated at", config.DefaultPath)
}

func chooseWorkbookInteractive() (string, error) {
	options := []string{"Keep existing cfg/Schedule.xlsx", "Choose new template via Finder"}
	var selection string
	prompt := &survey.Select{
		Message: "Select Excel template",
		Options: options,
		Default: options[0],
	}
	if err := survey.AskOne(prompt, &selection); err != nil {
		return "", err
	}
	if selection == options[1] {
		selected, err := chooseFileWithFinder()
		if err != nil {
			return "", err
		}
		return selected, nil
	}
	return "", nil
}

func writeConfig(cfg config.Config, workbookSrc string) error {
	if err := config.Write(cfg, workbookSrc); err != nil {
		return fmt.Errorf("write config: %w", err)
	}
	return nil
}

func destExists(path string) error {
	_, err := os.Stat(path)
	return err
}

func chooseFileWithFinder() (string, error) {
	cmd := exec.Command("osascript", "-e", `POSIX path of (choose file with prompt "Select the workbook to copy into cfg/Schedule.xlsx")`)
	out, err := cmd.Output()
	if err != nil {
		return "", err
	}
	path := strings.TrimSpace(string(out))
	if path == "" {
		return "", errors.New("no file selected")
	}
	if !isExcelFile(path) {
		return "", fmt.Errorf("%s is not an .xls or .xlsx file", path)
	}
	return path, nil
}

func isExcelFile(path string) bool {
	lower := strings.ToLower(path)
	return strings.HasSuffix(lower, ".xls") || strings.HasSuffix(lower, ".xlsx")
}
