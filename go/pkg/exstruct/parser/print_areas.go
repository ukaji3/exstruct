package parser

import (
	"strings"

	"github.com/ukaji3/exstruct-go/pkg/exstruct/models"
	"github.com/xuri/excelize/v2"
)

// ExtractPrintAreas extracts print areas from a workbook.
// Returns a map of sheet name to list of print areas.
func ExtractPrintAreas(f *excelize.File) (map[string][]models.PrintArea, error) {
	result := make(map[string][]models.PrintArea)

	// Get all defined names
	definedNames := f.GetDefinedName()

	for _, dn := range definedNames {
		// Look for _xlnm.Print_Area defined name
		if strings.EqualFold(dn.Name, "_xlnm.Print_Area") {
			// Parse the reference to get sheet name and range
			sheetName, areas := parsePrintAreaReference(dn.RefersTo)
			if sheetName != "" && len(areas) > 0 {
				result[sheetName] = append(result[sheetName], areas...)
			}
		}
	}

	return result, nil
}

// parsePrintAreaReference parses a print area reference string.
// Format: 'SheetName'!$A$1:$D$10 or SheetName!$A$1:$D$10
func parsePrintAreaReference(ref string) (string, []models.PrintArea) {
	var areas []models.PrintArea

	// Split by comma for multiple print areas
	parts := strings.Split(ref, ",")

	var sheetName string
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}

		// Split by ! to separate sheet name and range
		if idx := strings.LastIndex(part, "!"); idx >= 0 {
			sheet := part[:idx]
			rangeStr := part[idx+1:]

			// Remove quotes from sheet name
			sheet = strings.Trim(sheet, "'")
			if sheetName == "" {
				sheetName = sheet
			}

			// Parse the range
			if area := parseRangeToArea(rangeStr); area != nil {
				areas = append(areas, *area)
			}
		}
	}

	return sheetName, areas
}

// parseRangeToArea parses a range string like $A$1:$D$10 to PrintArea.
func parseRangeToArea(rangeStr string) *models.PrintArea {
	// Remove $ signs
	rangeStr = strings.ReplaceAll(rangeStr, "$", "")

	// Split by :
	parts := strings.Split(rangeStr, ":")
	if len(parts) != 2 {
		return nil
	}

	startCol, startRow, err := excelize.CellNameToCoordinates(parts[0])
	if err != nil {
		return nil
	}

	endCol, endRow, err := excelize.CellNameToCoordinates(parts[1])
	if err != nil {
		return nil
	}

	return &models.PrintArea{
		R1: startRow,
		C1: startCol,
		R2: endRow,
		C2: endCol,
	}
}
