package exstruct

import (
	"path/filepath"

	"github.com/ukaji3/exstruct-go/pkg/exstruct/models"
	"github.com/ukaji3/exstruct-go/pkg/exstruct/parser"
	"github.com/xuri/excelize/v2"
)

// Extract extracts structured data from an Excel file.
func Extract(path string, opts Options) (*models.WorkbookData, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	bookName := filepath.Base(path)
	sheets := make(map[string]models.SheetData)

	// Get sheet names
	sheetList := f.GetSheetList()

	// Extract cells for all sheets
	for _, sheetName := range sheetList {
		includeLinks := opts.ShouldIncludeLinks()
		rows, err := parser.ExtractCells(f, sheetName, includeLinks)
		if err != nil {
			// Log warning and continue with empty rows
			rows = nil
		}

		// Detect tables
		tables, err := parser.DetectTables(f, sheetName, parser.DefaultTableParams())
		if err != nil {
			tables = nil
		}

		sheets[sheetName] = models.SheetData{
			Rows:            rows,
			TableCandidates: tables,
		}
	}

	// Extract shapes (requires direct OOXML parsing)
	if opts.Mode != ModeLight {
		shapeData, err := parser.ExtractShapes(path, string(opts.Mode))
		if err == nil {
			for sheetName, shapes := range shapeData {
				if sheet, ok := sheets[sheetName]; ok {
					sheet.Shapes = shapes
					sheets[sheetName] = sheet
				}
			}
		}
	}

	// Extract charts (requires direct OOXML parsing)
	if opts.Mode != ModeLight {
		chartData, err := parser.ExtractCharts(path, string(opts.Mode))
		if err == nil {
			for sheetName, charts := range chartData {
				if sheet, ok := sheets[sheetName]; ok {
					sheet.Charts = charts
					sheets[sheetName] = sheet
				}
			}
		}
	}

	// Extract print areas
	if opts.ShouldIncludePrintAreas() {
		printAreas, err := parser.ExtractPrintAreas(f)
		if err == nil {
			for sheetName, areas := range printAreas {
				if sheet, ok := sheets[sheetName]; ok {
					sheet.PrintAreas = areas
					sheets[sheetName] = sheet
				}
			}
		}
	}

	return &models.WorkbookData{
		BookName: bookName,
		Sheets:   sheets,
	}, nil
}
