package main

import (
	"log"

	"update-google-sheets/src/config"
	sheetops "update-google-sheets/src/sheets"
)

func main() {
	cfg, err := config.Load(config.DefaultPath)
	if err != nil {
		log.Fatal(err)
	}
	if err := cfg.Validate(); err != nil {
		log.Fatal(err)
	}

	ranges, templateSheets, err := sheetops.ScanWorkbook(config.DefaultWorkbook, cfg.SheetFilter, cfg.LookupValue)
	if err != nil {
		log.Fatal(err)
	}

	if err := sheetops.WriteRangeCache(sheetops.DefaultRangeCache, ranges); err != nil {
		log.Fatal(err)
	}

	log.Printf("cached %d lookup cells across %d sheets at %s\n", len(ranges), len(templateSheets), sheetops.DefaultRangeCache)
}
