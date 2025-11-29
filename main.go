package main

import (
	"context"
	"fmt"
	"os"

	"go.uber.org/zap"

	"update-google-sheets/src/config"
	"update-google-sheets/src/logger"
	sheetops "update-google-sheets/src/sheets"
)

func main() {
	cfg, err := config.Load(config.DefaultPath)
	if err != nil {
		exitErr("%v", err)
	}

	if err := cfg.Validate(); err != nil {
		exitErr("%v", err)
	}

	log, err := logger.New()
	if err != nil {
		exitErr("initialise logger: %v", err)
	}
	defer func() { _ = log.Sync() }()
	log.Info(
		"using configuration",
		zap.String("spreadsheet_id", cfg.SpreadsheetID),
		zap.String("workbook", cfg.Workbook),
		zap.String("sheet_filter", cfg.SheetFilter),
		zap.String("lookup_value", cfg.LookupValue),
	)

	summary, err := sheetops.Update(context.Background(), cfg)
	if err != nil {
		log.Error("update failed", zap.Error(err))
		exitErr("%v", err)
	}

	if summary.SkippedReason != "" {
		log.Info("no updates performed", zap.String("reason", summary.SkippedReason))
		return
	}

	log.Info(
		"update complete",
		zap.Strings("ranges", summary.Ranges),
		zap.Int64("rows", summary.TotalRows),
		zap.Int64("cells", summary.TotalCells),
	)
}

func exitErr(msg string, args ...interface{}) {
	fmt.Fprintf(os.Stderr, msg+"\n", args...)
	os.Exit(1)
}
