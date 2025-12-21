package models

// WorkbookData represents workbook-level container with per-sheet data.
type WorkbookData struct {
	// BookName is the workbook file name (no path).
	BookName string `json:"book_name"`
	// Sheets maps sheet name to SheetData.
	Sheets map[string]SheetData `json:"sheets"`
}
