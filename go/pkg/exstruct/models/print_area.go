package models

// PrintArea represents cell coordinate bounds for a print area.
type PrintArea struct {
	// R1 is the start row (1-based).
	R1 int `json:"r1"`
	// C1 is the start column (1-based).
	C1 int `json:"c1"`
	// R2 is the end row (1-based, inclusive).
	R2 int `json:"r2"`
	// C2 is the end column (1-based, inclusive).
	C2 int `json:"c2"`
}

// PrintAreaView represents a slice of a sheet restricted to a print area.
type PrintAreaView struct {
	// BookName is the workbook name owning the area.
	BookName string `json:"book_name"`
	// SheetName is the sheet name owning the area.
	SheetName string `json:"sheet_name"`
	// Area is the print area bounds.
	Area PrintArea `json:"area"`
	// Rows contains rows within the area bounds.
	Rows []CellRow `json:"rows,omitempty"`
	// Shapes contains shapes overlapping the area.
	Shapes []Shape `json:"shapes,omitempty"`
	// Charts contains charts overlapping the area.
	Charts []Chart `json:"charts,omitempty"`
	// TableCandidates contains table candidates intersecting the area.
	TableCandidates []string `json:"table_candidates,omitempty"`
}
