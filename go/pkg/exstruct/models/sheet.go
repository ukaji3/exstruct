package models

// SheetData represents structured data for a single sheet.
type SheetData struct {
	// Rows contains extracted rows with cell values and links.
	Rows []CellRow `json:"rows,omitempty"`
	// Shapes contains shapes detected on the sheet.
	Shapes []Shape `json:"shapes,omitempty"`
	// Charts contains charts detected on the sheet.
	Charts []Chart `json:"charts,omitempty"`
	// TableCandidates contains cell ranges likely representing tables.
	TableCandidates []string `json:"table_candidates,omitempty"`
	// PrintAreas contains user-defined print areas.
	PrintAreas []PrintArea `json:"print_areas,omitempty"`
}
