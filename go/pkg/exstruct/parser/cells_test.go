package parser

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestExtractCells(t *testing.T) {
	// Create a temporary Excel file for testing
	f := excelize.NewFile()
	defer f.Close()

	sheetName := "Sheet1"
	// Set some test data
	f.SetCellValue(sheetName, "A1", "Header1")
	f.SetCellValue(sheetName, "B1", "Header2")
	f.SetCellValue(sheetName, "A2", 100)
	f.SetCellValue(sheetName, "B2", 200.5)
	f.SetCellValue(sheetName, "A3", "Text")

	// Save to temp file
	tmpDir := t.TempDir()
	tmpFile := filepath.Join(tmpDir, "test.xlsx")
	if err := f.SaveAs(tmpFile); err != nil {
		t.Fatalf("Failed to save test file: %v", err)
	}

	// Open and extract
	f2, err := excelize.OpenFile(tmpFile)
	if err != nil {
		t.Fatalf("Failed to open test file: %v", err)
	}
	defer f2.Close()

	rows, err := ExtractCells(f2, sheetName, false)
	if err != nil {
		t.Fatalf("ExtractCells failed: %v", err)
	}

	// Verify results
	if len(rows) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(rows))
	}

	// Check first row
	if rows[0].R != 1 {
		t.Errorf("Expected row 1, got %d", rows[0].R)
	}
	if rows[0].C["1"] != "Header1" {
		t.Errorf("Expected 'Header1', got %v", rows[0].C["1"])
	}

	// Check numeric values
	if rows[1].C["1"] != int64(100) {
		t.Errorf("Expected int64(100), got %v (type: %T)", rows[1].C["1"], rows[1].C["1"])
	}
	if rows[1].C["2"] != 200.5 {
		t.Errorf("Expected 200.5, got %v", rows[1].C["2"])
	}

	// Cleanup
	os.Remove(tmpFile)
}

func TestParseValue(t *testing.T) {
	tests := []struct {
		input    string
		expected interface{}
	}{
		{"123", int64(123)},
		{"123.45", 123.45},
		{"-100", int64(-100)},
		{"hello", "hello"},
		{"", ""},
	}

	for _, tt := range tests {
		result := parseValue(tt.input)
		if result != tt.expected {
			t.Errorf("parseValue(%q) = %v (type: %T), expected %v (type: %T)",
				tt.input, result, result, tt.expected, tt.expected)
		}
	}
}
