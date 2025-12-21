package parser

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// TableDetectionParams holds parameters for table detection.
type TableDetectionParams struct {
	DensityMin      float64
	CoverageMin     float64
	MinNonemptyCells int
}

// DefaultTableParams returns default table detection parameters.
func DefaultTableParams() TableDetectionParams {
	return TableDetectionParams{
		DensityMin:      0.04,
		CoverageMin:     0.2,
		MinNonemptyCells: 3,
	}
}

// DetectTables detects table-like regions in a sheet.
// Returns a list of cell ranges (e.g., "A1:D10") that likely represent tables.
func DetectTables(f *excelize.File, sheetName string, params TableDetectionParams) ([]string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	if len(rows) == 0 {
		return nil, nil
	}

	// Find the bounding box of non-empty cells
	minRow, maxRow, minCol, maxCol := findDataBounds(rows)
	if minRow < 0 {
		return nil, nil
	}

	// Calculate density
	totalCells := (maxRow - minRow + 1) * (maxCol - minCol + 1)
	nonEmptyCells := countNonEmptyCells(rows, minRow, maxRow, minCol, maxCol)

	if nonEmptyCells < params.MinNonemptyCells {
		return nil, nil
	}

	density := float64(nonEmptyCells) / float64(totalCells)
	if density < params.DensityMin {
		return nil, nil
	}

	// Convert to Excel range notation
	startCell, _ := excelize.CoordinatesToCellName(minCol+1, minRow+1)
	endCell, _ := excelize.CoordinatesToCellName(maxCol+1, maxRow+1)
	rangeStr := fmt.Sprintf("%s:%s", startCell, endCell)

	return []string{rangeStr}, nil
}

// findDataBounds finds the bounding box of non-empty cells.
func findDataBounds(rows [][]string) (minRow, maxRow, minCol, maxCol int) {
	minRow, maxRow = -1, -1
	minCol, maxCol = -1, -1

	for rowIdx, row := range rows {
		for colIdx, cell := range row {
			if cell != "" {
				if minRow < 0 || rowIdx < minRow {
					minRow = rowIdx
				}
				if maxRow < 0 || rowIdx > maxRow {
					maxRow = rowIdx
				}
				if minCol < 0 || colIdx < minCol {
					minCol = colIdx
				}
				if maxCol < 0 || colIdx > maxCol {
					maxCol = colIdx
				}
			}
		}
	}

	return
}

// countNonEmptyCells counts non-empty cells within bounds.
func countNonEmptyCells(rows [][]string, minRow, maxRow, minCol, maxCol int) int {
	count := 0
	for rowIdx := minRow; rowIdx <= maxRow && rowIdx < len(rows); rowIdx++ {
		row := rows[rowIdx]
		for colIdx := minCol; colIdx <= maxCol && colIdx < len(row); colIdx++ {
			if row[colIdx] != "" {
				count++
			}
		}
	}
	return count
}
