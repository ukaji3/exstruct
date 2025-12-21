package parser

import (
	"archive/zip"
	"encoding/xml"
	"io"
	"math"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/ukaji3/exstruct-go/pkg/exstruct/models"
)

// XML namespaces used in DrawingML
const (
	nsXDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
	nsA   = "http://schemas.openxmlformats.org/drawingml/2006/main"
	nsR   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

// PresetGeomMap maps OOXML preset geometry names to human-readable type labels.
var PresetGeomMap = map[string]string{
	"flowChartProcess":           "AutoShape-FlowchartProcess",
	"flowChartDecision":          "AutoShape-FlowchartDecision",
	"flowChartTerminator":        "AutoShape-FlowchartTerminator",
	"flowChartData":              "AutoShape-FlowchartData",
	"flowChartDocument":          "AutoShape-FlowchartDocument",
	"flowChartMultidocument":     "AutoShape-FlowchartMultidocument",
	"flowChartPredefinedProcess": "AutoShape-FlowchartPredefinedProcess",
	"flowChartInternalStorage":   "AutoShape-FlowchartInternalStorage",
	"flowChartPreparation":       "AutoShape-FlowchartPreparation",
	"flowChartManualInput":       "AutoShape-FlowchartManualInput",
	"flowChartManualOperation":   "AutoShape-FlowchartManualOperation",
	"flowChartConnector":         "AutoShape-FlowchartConnector",
	"flowChartOffpageConnector":  "AutoShape-FlowchartOffpageConnector",
	"rect":                       "AutoShape-Rectangle",
	"roundRect":                  "AutoShape-RoundedRectangle",
	"ellipse":                    "AutoShape-Oval",
	"diamond":                    "AutoShape-Diamond",
	"triangle":                   "AutoShape-IsoscelesTriangle",
	"straightConnector1":         "Line",
	"bentConnector2":             "AutoShape-Connector",
	"bentConnector3":             "AutoShape-Connector",
	"bentConnector4":             "AutoShape-Connector",
	"bentConnector5":             "AutoShape-Connector",
	"curvedConnector2":           "AutoShape-Connector",
	"curvedConnector3":           "AutoShape-Connector",
	"curvedConnector4":           "AutoShape-Connector",
	"curvedConnector5":           "AutoShape-Connector",
	"line":                       "Line",
	"textBox":                    "TextBox",
}

// ArrowHeadMap maps OOXML arrow head types to Excel COM style numbers.
var ArrowHeadMap = map[string]int{
	"none":     1,
	"triangle": 2,
	"stealth":  3,
	"diamond":  4,
	"oval":     5,
	"arrow":    2,
}

// shapeParseResult holds intermediate parsing results.
type shapeParseResult struct {
	shape       models.Shape
	excelID     string
	isConnector bool
	startCxnID  string
	endCxnID    string
}

// ExtractShapes extracts shapes from an xlsx file.
func ExtractShapes(xlsxPath string, mode string) (map[string][]models.Shape, error) {
	if mode == "light" {
		return make(map[string][]models.Shape), nil
	}

	r, err := zip.OpenReader(xlsxPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	// Get sheet to drawing mapping
	sheetDrawingMap, err := getSheetDrawingMap(&r.Reader)
	if err != nil {
		return nil, err
	}

	result := make(map[string][]models.Shape)
	for sheetName, drawingPath := range sheetDrawingMap {
		shapes, err := parseDrawingFile(&r.Reader, drawingPath, mode)
		if err != nil {
			// Log warning and continue
			result[sheetName] = []models.Shape{}
			continue
		}
		result[sheetName] = shapes
	}

	return result, nil
}

// getSheetDrawingMap returns a mapping of sheet names to their drawing XML paths.
func getSheetDrawingMap(r *zip.Reader) (map[string]string, error) {
	result := make(map[string]string)

	// Read workbook.xml to get sheet names and rIds
	workbookXML, err := readZipFile(r, "xl/workbook.xml")
	if err != nil {
		return result, nil
	}

	sheetsInfo := parseWorkbookSheets(workbookXML)
	if len(sheetsInfo) == 0 {
		return result, nil
	}

	// Read workbook.xml.rels to map rId to sheet file
	wbRelsXML, err := readZipFile(r, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return result, nil
	}

	sheetFiles := parseWorkbookRels(wbRelsXML, sheetsInfo)

	// For each sheet, find its drawing relationship
	for sheetName, sheetPath := range sheetFiles {
		relsPath := strings.Replace(sheetPath, "worksheets/", "worksheets/_rels/", 1)
		relsPath = strings.Replace(relsPath, ".xml", ".xml.rels", 1)

		sheetRelsXML, err := readZipFile(r, relsPath)
		if err != nil {
			continue
		}

		drawingPath := findDrawingRelationship(sheetRelsXML)
		if drawingPath != "" {
			result[sheetName] = resolveRelativePath(drawingPath, "xl/drawings")
		}
	}

	return result, nil
}

// parseDrawingFile parses a drawing XML file and extracts shapes.
func parseDrawingFile(r *zip.Reader, drawingPath string, mode string) ([]models.Shape, error) {
	drawingXML, err := readZipFile(r, drawingPath)
	if err != nil {
		return nil, err
	}

	parseResults := parseDrawingXML(drawingXML, mode)
	assignShapeIDs(parseResults)

	shapes := make([]models.Shape, len(parseResults))
	for i, pr := range parseResults {
		shapes[i] = pr.shape
	}

	return shapes, nil
}

// parseDrawingXML parses drawing XML content and returns shape parse results.
func parseDrawingXML(data []byte, mode string) []shapeParseResult {
	var results []shapeParseResult

	decoder := xml.NewDecoder(strings.NewReader(string(data)))
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			break
		}

		if se, ok := token.(xml.StartElement); ok {
			switch se.Name.Local {
			case "twoCellAnchor", "oneCellAnchor", "absoluteAnchor":
				anchorResults := parseAnchor(decoder, se, mode)
				results = append(results, anchorResults...)
			}
		}
	}

	return results
}

// parseAnchor parses an anchor element and its child shapes.
func parseAnchor(decoder *xml.Decoder, start xml.StartElement, mode string) []shapeParseResult {
	var results []shapeParseResult
	depth := 1

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "sp":
				if pr := parseShapeElement(decoder, t, mode, false); pr != nil {
					results = append(results, *pr)
				}
				depth--
			case "cxnSp":
				if pr := parseShapeElement(decoder, t, mode, true); pr != nil {
					results = append(results, *pr)
				}
				depth--
			case "grpSp":
				grpResults := parseGroupShape(decoder, t, mode)
				results = append(results, grpResults...)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return results
}

// parseShapeElement parses a single shape element.
func parseShapeElement(decoder *xml.Decoder, start xml.StartElement, mode string, isCxnSp bool) *shapeParseResult {
	var text string
	var left, top, width, height int
	var excelID, shapeName string
	var prst string
	var rotation *float64
	var beginArrowStyle, endArrowStyle *int
	var startCxnID, endCxnID string

	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "cNvPr":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "id":
						excelID = attr.Value
					case "name":
						shapeName = attr.Value
					}
				}
			case "xfrm":
				l, tp, w, h, rot := parseXfrm(decoder, t)
				left, top, width, height = l, tp, w, h
				rotation = rot
				depth--
			case "prstGeom":
				for _, attr := range t.Attr {
					if attr.Name.Local == "prst" {
						prst = attr.Value
					}
				}
			case "t":
				if txt, err := readElementText(decoder); err == nil {
					text += txt
				}
				depth--
			case "ln":
				begin, end := parseLineArrows(decoder, t)
				beginArrowStyle, endArrowStyle = begin, end
				depth--
			case "cNvCxnSpPr":
				start, end := parseConnectorEndpoints(decoder, t)
				startCxnID, endCxnID = start, end
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	text = strings.TrimSpace(text)

	// Determine type label
	typeLabel := "Unknown"
	if prst != "" {
		if label, ok := PresetGeomMap[prst]; ok {
			typeLabel = label
		} else {
			typeLabel = "AutoShape-" + prst
		}
	} else if shapeName != "" {
		typeLabel = shapeName
	}

	// Check if connector
	isConnector := isCxnSp || isConnectorShape(prst, typeLabel)

	// Apply mode filtering
	if !shouldIncludeShape(text, typeLabel, isConnector, mode) {
		return nil
	}

	shape := models.Shape{
		Text: text,
		L:    left,
		T:    top,
		Type: typeLabel,
	}

	if mode == "verbose" {
		w := width
		h := height
		shape.W = &w
		shape.H = &h
	}

	if rotation != nil {
		shape.Rotation = rotation
	}

	if isConnector {
		direction := computeDirection(width, height)
		if direction != "" {
			shape.Direction = direction
		}
		shape.BeginArrowStyle = beginArrowStyle
		shape.EndArrowStyle = endArrowStyle
	}

	return &shapeParseResult{
		shape:       shape,
		excelID:     excelID,
		isConnector: isConnector,
		startCxnID:  startCxnID,
		endCxnID:    endCxnID,
	}
}

// parseGroupShape parses a group shape element recursively.
func parseGroupShape(decoder *xml.Decoder, start xml.StartElement, mode string) []shapeParseResult {
	var results []shapeParseResult
	depth := 1

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "sp":
				if pr := parseShapeElement(decoder, t, mode, false); pr != nil {
					results = append(results, *pr)
				}
				depth--
			case "cxnSp":
				if pr := parseShapeElement(decoder, t, mode, true); pr != nil {
					results = append(results, *pr)
				}
				depth--
			case "grpSp":
				grpResults := parseGroupShape(decoder, t, mode)
				results = append(results, grpResults...)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return results
}

// parseXfrm parses xfrm element for position and size.
func parseXfrm(decoder *xml.Decoder, start xml.StartElement) (left, top, width, height int, rotation *float64) {
	// Check for rotation attribute
	for _, attr := range start.Attr {
		if attr.Name.Local == "rot" {
			if rotEmu, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
				rotDeg := float64(rotEmu) / 60000.0
				if math.Abs(rotDeg) >= 1e-6 {
					rotation = &rotDeg
				}
			}
		}
	}

	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "off":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "x":
						if x, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							left = EMUToPixels(x)
						}
					case "y":
						if y, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							top = EMUToPixels(y)
						}
					}
				}
			case "ext":
				for _, attr := range t.Attr {
					switch attr.Name.Local {
					case "cx":
						if cx, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							width = EMUToPixels(cx)
						}
					case "cy":
						if cy, err := strconv.ParseInt(attr.Value, 10, 64); err == nil {
							height = EMUToPixels(cy)
						}
					}
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// parseLineArrows parses line element for arrow styles.
func parseLineArrows(decoder *xml.Decoder, start xml.StartElement) (beginStyle, endStyle *int) {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "headEnd":
				for _, attr := range t.Attr {
					if attr.Name.Local == "type" {
						if style, ok := ArrowHeadMap[attr.Value]; ok {
							beginStyle = &style
						}
					}
				}
			case "tailEnd":
				for _, attr := range t.Attr {
					if attr.Name.Local == "type" {
						if style, ok := ArrowHeadMap[attr.Value]; ok {
							endStyle = &style
						}
					}
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// parseConnectorEndpoints parses connector endpoint IDs.
func parseConnectorEndpoints(decoder *xml.Decoder, start xml.StartElement) (startID, endID string) {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			switch t.Name.Local {
			case "stCxn":
				for _, attr := range t.Attr {
					if attr.Name.Local == "id" {
						startID = attr.Value
					}
				}
			case "endCxn":
				for _, attr := range t.Attr {
					if attr.Name.Local == "id" {
						endID = attr.Value
					}
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// computeDirection computes compass direction from connector dimensions.
func computeDirection(width, height int) string {
	if width == 0 && height == 0 {
		return ""
	}

	angle := math.Atan2(float64(-height), float64(width)) * 180 / math.Pi
	if angle < 0 {
		angle += 360
	}

	switch {
	case angle >= 337.5 || angle < 22.5:
		return "E"
	case angle >= 22.5 && angle < 67.5:
		return "NE"
	case angle >= 67.5 && angle < 112.5:
		return "N"
	case angle >= 112.5 && angle < 157.5:
		return "NW"
	case angle >= 157.5 && angle < 202.5:
		return "W"
	case angle >= 202.5 && angle < 247.5:
		return "SW"
	case angle >= 247.5 && angle < 292.5:
		return "S"
	default:
		return "SE"
	}
}

// isConnectorShape checks if a shape is a connector or line.
func isConnectorShape(prst, typeLabel string) bool {
	connectorKeywords := []string{"Connector", "line", "straightConnector"}
	for _, kw := range connectorKeywords {
		if strings.Contains(strings.ToLower(prst), strings.ToLower(kw)) {
			return true
		}
	}
	if strings.Contains(typeLabel, "Line") || strings.Contains(typeLabel, "Connector") {
		return true
	}
	return false
}

// shouldIncludeShape determines if a shape should be included based on mode.
func shouldIncludeShape(text, typeLabel string, isConnector bool, mode string) bool {
	if mode == "light" {
		return false
	}
	if mode == "verbose" {
		return true
	}
	// standard mode: include if text exists or is connector/arrow
	if text != "" {
		return true
	}
	if isConnector {
		return true
	}
	if strings.Contains(typeLabel, "Arrow") {
		return true
	}
	return false
}

// assignShapeIDs assigns sequential IDs to shapes and resolves connector endpoints.
func assignShapeIDs(results []shapeParseResult) {
	excelIDToNodeID := make(map[string]int)
	nodeIndex := 0

	// First pass: assign IDs to non-connector shapes
	for i := range results {
		if !results[i].isConnector && results[i].excelID != "" {
			nodeIndex++
			id := nodeIndex
			results[i].shape.ID = &id
			excelIDToNodeID[results[i].excelID] = nodeIndex
		}
	}

	// Second pass: resolve connector endpoints
	for i := range results {
		if results[i].isConnector {
			if results[i].startCxnID != "" {
				if nodeID, ok := excelIDToNodeID[results[i].startCxnID]; ok {
					results[i].shape.BeginID = &nodeID
				}
			}
			if results[i].endCxnID != "" {
				if nodeID, ok := excelIDToNodeID[results[i].endCxnID]; ok {
					results[i].shape.EndID = &nodeID
				}
			}
		}
	}
}

// Helper functions

func readZipFile(r *zip.Reader, name string) ([]byte, error) {
	for _, f := range r.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, nil
}

func readElementText(decoder *xml.Decoder) (string, error) {
	var text string
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return text, err
		}
		switch t := token.(type) {
		case xml.CharData:
			text += string(t)
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return text, nil
}

func resolveRelativePath(target, baseDir string) string {
	if strings.HasPrefix(target, "../") {
		clean := target
		for strings.HasPrefix(clean, "../") {
			clean = strings.TrimPrefix(clean, "../")
		}
		return "xl/" + clean
	}
	if strings.HasPrefix(target, "/") {
		return baseDir + target
	}
	return baseDir + "/" + target
}

func parseWorkbookSheets(data []byte) map[string]string {
	result := make(map[string]string) // rId -> sheet name
	decoder := xml.NewDecoder(strings.NewReader(string(data)))

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := token.(xml.StartElement); ok && se.Name.Local == "sheet" {
			var name, rID string
			for _, attr := range se.Attr {
				switch attr.Name.Local {
				case "name":
					name = attr.Value
				case "id":
					rID = attr.Value
				}
			}
			if name != "" && rID != "" {
				result[rID] = name
			}
		}
	}

	return result
}

func parseWorkbookRels(data []byte, sheetsInfo map[string]string) map[string]string {
	result := make(map[string]string) // sheet name -> file path
	decoder := xml.NewDecoder(strings.NewReader(string(data)))

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := token.(xml.StartElement); ok && se.Name.Local == "Relationship" {
			var rID, target string
			for _, attr := range se.Attr {
				switch attr.Name.Local {
				case "Id":
					rID = attr.Value
				case "Target":
					target = attr.Value
				}
			}
			if sheetName, ok := sheetsInfo[rID]; ok && strings.Contains(strings.ToLower(target), "worksheet") {
				result[sheetName] = resolveRelativePath(target, "xl")
			}
		}
	}

	return result
}

func findDrawingRelationship(data []byte) string {
	decoder := xml.NewDecoder(strings.NewReader(string(data)))

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		if se, ok := token.(xml.StartElement); ok && se.Name.Local == "Relationship" {
			var relType, target string
			for _, attr := range se.Attr {
				switch attr.Name.Local {
				case "Type":
					relType = attr.Value
				case "Target":
					target = attr.Value
				}
			}
			if strings.Contains(strings.ToLower(relType), "drawing") {
				return target
			}
		}
	}

	return ""
}

// GetShapeDrawingPath returns the drawing path for a sheet (exported for testing).
func GetShapeDrawingPath(xlsxPath, sheetName string) (string, error) {
	r, err := zip.OpenReader(xlsxPath)
	if err != nil {
		return "", err
	}
	defer r.Close()

	sheetDrawingMap, err := getSheetDrawingMap(&r.Reader)
	if err != nil {
		return "", err
	}

	return sheetDrawingMap[sheetName], nil
}

// GetDrawingFilePath returns the full path to a drawing file.
func GetDrawingFilePath(xlsxPath string, sheetIndex int) string {
	return filepath.Join("xl", "drawings", "drawing"+strconv.Itoa(sheetIndex+1)+".xml")
}
