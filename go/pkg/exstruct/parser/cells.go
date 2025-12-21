package parser

import (
	"strconv"

	"github.com/ukaji3/exstruct-go/pkg/exstruct/models"
	"github.com/xuri/excelize/v2"
)

// ExtractCells extracts cell data from a sheet.
// It returns a slice of CellRow containing non-empty rows.
func ExtractCells(f *excelize.File, sheetName string, includeLinks bool) ([]models.CellRow, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	var result []models.CellRow
	for rowIdx, row := range rows {
		rowNum := rowIdx + 1 // 1-based row index
		cellMap := make(map[string]interface{})
		linkMap := make(map[string]string)
		hasData := false

		for colIdx, cellValue := range row {
			if cellValue == "" {
				continue
			}
			hasData = true
			colStr := strconv.Itoa(colIdx + 1) // 1-based column index as string

			// Try to parse as number
			cellMap[colStr] = parseValue(cellValue)

			// Extract hyperlink if requested
			if includeLinks {
				cellName, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)
				hasLink, target, err := f.GetCellHyperLink(sheetName, cellName)
				if err == nil && hasLink && target != "" {
					linkMap[colStr] = target
				}
			}
		}

		if hasData {
			cellRow := models.CellRow{
				R: rowNum,
				C: cellMap,
			}
			if includeLinks && len(linkMap) > 0 {
				cellRow.Links = linkMap
			}
			result = append(result, cellRow)
		}
	}

	return result, nil
}

// parseValue attempts to parse a string value as a number.
// Returns int64 for integers, float64 for decimals, or the original string.
func parseValue(s string) interface{} {
	// Try integer first
	if i, err := strconv.ParseInt(s, 10, 64); err == nil {
		return i
	}
	// Try float
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		return f
	}
	// Return as string
	return s
}
