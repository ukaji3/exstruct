// Package models defines data structures for Excel extraction.
package models

// CellRow represents a single row of cells with optional hyperlinks.
type CellRow struct {
	// R is the row index (1-based).
	R int `json:"r"`
	// C maps column index (string) to cell value.
	C map[string]interface{} `json:"c"`
	// Links maps column index to hyperlink URL (optional).
	Links map[string]string `json:"links,omitempty"`
}
