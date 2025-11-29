package main

import (
	"bufio"
	"errors"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"strings"

	"update-google-sheets/src/config"
)

func main() {
	existing, _ := config.Load(config.DefaultPath)
	defaultWorkbook := existing.Workbook
	if defaultWorkbook == "" {
		defaultWorkbook = config.DefaultWorkbook
	}
	nonInteractive := flag.Bool("non-interactive", false, "Use existing flags instead of prompts")
	spreadsheet := flag.String("spreadsheet", existing.SpreadsheetID, "Spreadsheet ID")
	sheetFilter := flag.String("sheet", existing.SheetFilter, "Sheet name filter")
	lookup := flag.String("lookup", existing.LookupValue, "Lookup value")
	workbookSrc := flag.String("workbook-src", defaultWorkbook, "Path to workbook to copy into place (blank to skip)")
	workbookDest := flag.String("workbook-dest", defaultWorkbook, "Destination workbook path")
	flag.Parse()

	if *nonInteractive {
		cfg := config.Config{
			SpreadsheetID: strings.TrimSpace(*spreadsheet),
			SheetFilter:   strings.TrimSpace(*sheetFilter),
			LookupValue:   strings.TrimSpace(*lookup),
			Workbook:      strings.TrimSpace(*workbookDest),
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
	r := bufio.NewReader(os.Stdin)
	fmt.Printf("Loaded defaults from %s. Press Enter to reuse existing values.\n\n", config.DefaultPath)

	spreadsheetID := promptWithDefault(r, "Google Spreadsheet ID", existing.SpreadsheetID)
	sheetFilter := promptWithDefault(r, "Sheet filter", existing.SheetFilter)
	lookupValue := promptWithDefault(r, "Lookup value to search for", existing.LookupValue)
	if lookupValue == "" {
		lookupValue = existing.LookupValue
	}
	defaultSrc := existing.Workbook
	if defaultSrc == "" {
		defaultSrc = config.DefaultWorkbook
	}
	workbookSrc := promptWithDefault(r, "Path to the Excel workbook to copy (press Enter to skip copying):", defaultSrc)
	workbookSrc = strings.TrimSpace(workbookSrc)
	if workbookSrc == "" {
		workbookSrc = ""
	}

	dest := existing.Workbook
	if dest == "" {
		dest = config.DefaultWorkbook
	}
	dest = promptWithDefault(r, "Destination workbook path:", dest)
	if infoErr := destExists(dest); infoErr == nil {
		if !promptYesNo(r, fmt.Sprintf("Destination %s exists. Overwrite?", dest), true) {
			dest = promptRequired(r, "Enter alternate destination path inside the repo:")
		}
	}

	cfg := config.Config{
		SpreadsheetID: spreadsheetID,
		SheetFilter:   sheetFilter,
		LookupValue:   lookupValue,
		Workbook:      dest,
	}

	if err := writeConfig(cfg, workbookSrc); err != nil {
		log.Fatal(err)
	}
	log.Println("Configuration updated at", config.DefaultPath)
}

func writeConfig(cfg config.Config, workbookSrc string) error {
	if err := config.Write(cfg, workbookSrc); err != nil {
		return fmt.Errorf("write config: %w", err)
	}
	return nil
}

func promptWithDefault(r *bufio.Reader, question, def string) string {
	trimmed := strings.TrimSpace(def)
	if trimmed == "" {
		fmt.Printf("%s: ", question)
	} else {
		fmt.Printf("%s [%s]: ", question, trimmed)
	}
	line, err := readLine(r)
	if err != nil {
		log.Fatal(err)
	}
	if strings.TrimSpace(line) == "" {
		return trimmed
	}
	return strings.TrimSpace(line)
}

func promptRequired(r *bufio.Reader, question string) string {
	for {
		fmt.Print(question + " ")
		line, err := readLine(r)
		if err != nil {
			log.Fatal(err)
		}
		if strings.TrimSpace(line) != "" {
			return strings.TrimSpace(line)
		}
		fmt.Println("Please enter a value.")
	}
}

func promptYesNo(r *bufio.Reader, question string, def bool) bool {
	defLabel := "y"
	if !def {
		defLabel = "n"
	}
	for {
		fmt.Printf("%s [y/n, default %s]: ", question, defLabel)
		line, err := readLine(r)
		if err != nil {
			log.Fatal(err)
		}
		line = strings.TrimSpace(strings.ToLower(line))
		if line == "" {
			return def
		}
		if line == "y" || line == "yes" {
			return true
		}
		if line == "n" || line == "no" {
			return false
		}
		fmt.Println("Please answer y or n.")
	}
}

func readLine(r *bufio.Reader) (string, error) {
	line, err := r.ReadString('\n')
	if err != nil && !errors.Is(err, io.EOF) {
		return "", err
	}
	return line, nil
}

func destExists(path string) error {
	_, err := os.Stat(path)
	return err
}
