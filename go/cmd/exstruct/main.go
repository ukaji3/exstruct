// Package main provides the CLI entry point for exstruct-go.
package main

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
	"github.com/ukaji3/exstruct-go/pkg/exstruct"
	"github.com/ukaji3/exstruct-go/pkg/exstruct/models"
	"github.com/ukaji3/exstruct-go/pkg/exstruct/output"
)

var (
	outputPath    string
	pretty        bool
	mode          string
	sheetsDir     string
	printAreasDir string
)

func main() {
	rootCmd := &cobra.Command{
		Use:   "exstruct [input.xlsx]",
		Short: "Extract structured data from Excel files",
		Long: `exstruct-go extracts structured data (cells, shapes, charts, tables) 
from Excel files and outputs JSON.`,
		Args: cobra.ExactArgs(1),
		RunE: run,
	}

	rootCmd.Flags().StringVarP(&outputPath, "output", "o", "", "Output file path (default: stdout)")
	rootCmd.Flags().BoolVar(&pretty, "pretty", false, "Pretty-print JSON output")
	rootCmd.Flags().StringVar(&mode, "mode", "standard", "Extraction mode: light, standard, verbose")
	rootCmd.Flags().StringVar(&sheetsDir, "sheets-dir", "", "Directory for per-sheet output files")
	rootCmd.Flags().StringVar(&printAreasDir, "print-areas-dir", "", "Directory for per-print-area output files")

	if err := rootCmd.Execute(); err != nil {
		os.Exit(1)
	}
}

func run(cmd *cobra.Command, args []string) error {
	inputPath := args[0]

	// Validate input file exists
	if _, err := os.Stat(inputPath); os.IsNotExist(err) {
		return fmt.Errorf("file not found: %s", inputPath)
	}

	// Parse mode
	var extractMode exstruct.Mode
	switch mode {
	case "light":
		extractMode = exstruct.ModeLight
	case "standard":
		extractMode = exstruct.ModeStandard
	case "verbose":
		extractMode = exstruct.ModeVerbose
	default:
		return fmt.Errorf("invalid mode: %s (must be light, standard, or verbose)", mode)
	}

	opts := exstruct.Options{
		Mode: extractMode,
	}

	// Extract data
	wb, err := exstruct.Extract(inputPath, opts)
	if err != nil {
		return fmt.Errorf("extraction failed: %w", err)
	}

	// Serialize to JSON
	jsonData, err := output.ToJSON(wb, pretty)
	if err != nil {
		return fmt.Errorf("serialization failed: %w", err)
	}

	// Write output
	if outputPath != "" {
		if err := os.WriteFile(outputPath, jsonData, 0644); err != nil {
			return fmt.Errorf("failed to write output: %w", err)
		}
	} else if sheetsDir == "" && printAreasDir == "" {
		fmt.Println(string(jsonData))
	}

	// Write per-sheet files
	if sheetsDir != "" {
		if err := writeSheetFiles(wb, sheetsDir); err != nil {
			return fmt.Errorf("failed to write sheet files: %w", err)
		}
	}

	// Write per-print-area files
	if printAreasDir != "" {
		if err := writePrintAreaFiles(wb, printAreasDir); err != nil {
			return fmt.Errorf("failed to write print area files: %w", err)
		}
	}

	return nil
}

func writeSheetFiles(wb *models.WorkbookData, dir string) error {
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}

	for sheetName, sheet := range wb.Sheets {
		jsonData, err := output.SheetToJSON(&sheet, pretty)
		if err != nil {
			return err
		}

		filename := filepath.Join(dir, sheetName+".json")
		if err := os.WriteFile(filename, jsonData, 0644); err != nil {
			return err
		}
	}

	return nil
}

func writePrintAreaFiles(wb *models.WorkbookData, dir string) error {
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}

	for sheetName, sheet := range wb.Sheets {
		for i, area := range sheet.PrintAreas {
			view := createPrintAreaView(wb.BookName, sheetName, sheet, area)
			jsonData, err := output.PrintAreaViewToJSON(&view, pretty)
			if err != nil {
				return err
			}

			filename := filepath.Join(dir, fmt.Sprintf("%s_area%d.json", sheetName, i+1))
			if err := os.WriteFile(filename, jsonData, 0644); err != nil {
				return err
			}
		}
	}

	return nil
}

func createPrintAreaView(bookName, sheetName string, sheet models.SheetData, area models.PrintArea) models.PrintAreaView {
	view := models.PrintAreaView{
		BookName:  bookName,
		SheetName: sheetName,
		Area:      area,
	}

	// Filter rows within area
	for _, row := range sheet.Rows {
		if row.R >= area.R1 && row.R <= area.R2 {
			view.Rows = append(view.Rows, row)
		}
	}

	// Filter shapes overlapping area
	for _, shape := range sheet.Shapes {
		if shapeOverlapsArea(shape, area) {
			view.Shapes = append(view.Shapes, shape)
		}
	}

	// Filter charts overlapping area
	for _, chart := range sheet.Charts {
		if chartOverlapsArea(chart, area) {
			view.Charts = append(view.Charts, chart)
		}
	}

	// Filter table candidates intersecting area
	// (simplified: include all for now)
	view.TableCandidates = sheet.TableCandidates

	return view
}

func shapeOverlapsArea(shape models.Shape, area models.PrintArea) bool {
	// Simplified check: just check if shape position is within area bounds
	// A more accurate check would consider shape dimensions
	return true // Include all shapes for now
}

func chartOverlapsArea(chart models.Chart, area models.PrintArea) bool {
	// Simplified check: include all charts for now
	return true
}
