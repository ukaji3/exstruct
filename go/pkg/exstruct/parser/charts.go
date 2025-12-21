package parser

import (
	"archive/zip"
	"encoding/xml"
	"strconv"
	"strings"

	"github.com/ukaji3/exstruct-go/pkg/exstruct/models"
)

// ChartTypeMap maps OOXML chart element tags to chart type names.
var ChartTypeMap = map[string]string{
	"lineChart":     "Line",
	"line3DChart":   "3DLine",
	"barChart":      "Bar",
	"bar3DChart":    "3DBar",
	"areaChart":     "Area",
	"area3DChart":   "3DArea",
	"pieChart":      "Pie",
	"pie3DChart":    "3DPie",
	"doughnutChart": "Doughnut",
	"scatterChart":  "XYScatter",
	"bubbleChart":   "Bubble",
	"radarChart":    "Radar",
	"surfaceChart":  "Surface",
	"surface3DChart": "3DSurface",
	"stockChart":    "Stock",
	"ofPieChart":    "PieOfPie",
}

// chartInfo holds chart metadata from drawing.xml.
type chartInfo struct {
	name      string
	chartPath string
	left      int
	top       int
	width     int
	height    int
}

// ExtractCharts extracts charts from an xlsx file.
func ExtractCharts(xlsxPath string, mode string) (map[string][]models.Chart, error) {
	if mode == "light" {
		return make(map[string][]models.Chart), nil
	}

	r, err := zip.OpenReader(xlsxPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	// Get sheet to chart mapping
	sheetChartMap, err := getSheetChartMap(&r.Reader)
	if err != nil {
		return nil, err
	}

	result := make(map[string][]models.Chart)
	for sheetName, chartInfos := range sheetChartMap {
		var charts []models.Chart
		for _, ci := range chartInfos {
			chart, err := parseChartFile(&r.Reader, ci, mode)
			if err != nil {
				continue
			}
			if chart != nil {
				charts = append(charts, *chart)
			}
		}
		result[sheetName] = charts
	}

	return result, nil
}

// getSheetChartMap returns a mapping of sheet names to their chart info.
func getSheetChartMap(r *zip.Reader) (map[string][]chartInfo, error) {
	result := make(map[string][]chartInfo)

	// Read workbook.xml to get sheet names and rIds
	workbookXML, err := readZipFile(r, "xl/workbook.xml")
	if err != nil || workbookXML == nil {
		return result, nil
	}

	sheetsInfo := parseWorkbookSheets(workbookXML)
	if len(sheetsInfo) == 0 {
		return result, nil
	}

	// Read workbook.xml.rels to map rId to sheet file
	wbRelsXML, err := readZipFile(r, "xl/_rels/workbook.xml.rels")
	if err != nil || wbRelsXML == nil {
		return result, nil
	}

	sheetFiles := parseWorkbookRels(wbRelsXML, sheetsInfo)

	// For each sheet, find its charts
	for sheetName, sheetPath := range sheetFiles {
		relsPath := strings.Replace(sheetPath, "worksheets/", "worksheets/_rels/", 1)
		relsPath = strings.Replace(relsPath, ".xml", ".xml.rels", 1)

		sheetRelsXML, err := readZipFile(r, relsPath)
		if err != nil || sheetRelsXML == nil {
			continue
		}

		drawingPath := findDrawingRelationship(sheetRelsXML)
		if drawingPath == "" {
			continue
		}

		drawingFullPath := resolveRelativePath(drawingPath, "xl/drawings")
		chartInfos := getChartInfosFromDrawing(r, drawingFullPath)
		if len(chartInfos) > 0 {
			result[sheetName] = chartInfos
		}
	}

	return result, nil
}

// getChartInfosFromDrawing extracts chart info from a drawing XML file.
func getChartInfosFromDrawing(r *zip.Reader, drawingPath string) []chartInfo {
	var result []chartInfo

	drawingXML, err := readZipFile(r, drawingPath)
	if err != nil || drawingXML == nil {
		return result
	}

	// Parse drawing XML to find graphicFrame elements with charts
	chartPositions := parseDrawingForCharts(drawingXML)
	if len(chartPositions) == 0 {
		return result
	}

	// Get drawing rels to resolve chart paths
	relsPath := strings.Replace(drawingPath, "drawings/", "drawings/_rels/", 1)
	relsPath = strings.Replace(relsPath, ".xml", ".xml.rels", 1)

	relsXML, err := readZipFile(r, relsPath)
	if err != nil || relsXML == nil {
		return result
	}

	// Resolve chart paths
	chartPaths := parseDrawingRels(relsXML)

	for rID, pos := range chartPositions {
		if chartPath, ok := chartPaths[rID]; ok {
			result = append(result, chartInfo{
				name:      pos.name,
				chartPath: resolveRelativePath(chartPath, "xl/charts"),
				left:      pos.left,
				top:       pos.top,
				width:     pos.width,
				height:    pos.height,
			})
		}
	}

	return result
}

// chartPosition holds position info from drawing.xml.
type chartPosition struct {
	name   string
	left   int
	top    int
	width  int
	height int
}

// parseDrawingForCharts parses drawing XML to find chart positions.
func parseDrawingForCharts(data []byte) map[string]chartPosition {
	result := make(map[string]chartPosition)
	decoder := xml.NewDecoder(strings.NewReader(string(data)))

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		if se, ok := token.(xml.StartElement); ok && se.Name.Local == "twoCellAnchor" {
			rID, pos := parseGraphicFrame(decoder)
			if rID != "" {
				result[rID] = pos
			}
		}
	}

	return result
}

// parseGraphicFrame parses a twoCellAnchor to find graphicFrame with chart.
func parseGraphicFrame(decoder *xml.Decoder) (string, chartPosition) {
	var rID string
	var pos chartPosition
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
			case "graphicFrame":
				rID, pos = parseGraphicFrameContent(decoder)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return rID, pos
}

// parseGraphicFrameContent parses graphicFrame content.
func parseGraphicFrameContent(decoder *xml.Decoder) (string, chartPosition) {
	var rID string
	var pos chartPosition
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
					if attr.Name.Local == "name" {
						pos.name = attr.Value
					}
				}
			case "xfrm":
				l, tp, w, h, _ := parseXfrm(decoder, t)
				pos.left, pos.top, pos.width, pos.height = l, tp, w, h
				depth--
			case "chart":
				for _, attr := range t.Attr {
					if attr.Name.Local == "id" {
						rID = attr.Value
					}
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	return rID, pos
}

// parseDrawingRels parses drawing rels to get chart paths.
func parseDrawingRels(data []byte) map[string]string {
	result := make(map[string]string)
	decoder := xml.NewDecoder(strings.NewReader(string(data)))

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		if se, ok := token.(xml.StartElement); ok && se.Name.Local == "Relationship" {
			var rID, target, relType string
			for _, attr := range se.Attr {
				switch attr.Name.Local {
				case "Id":
					rID = attr.Value
				case "Target":
					target = attr.Value
				case "Type":
					relType = attr.Value
				}
			}
			if strings.Contains(strings.ToLower(relType), "chart") {
				result[rID] = target
			}
		}
	}

	return result
}

// parseChartFile parses a chart XML file.
func parseChartFile(r *zip.Reader, ci chartInfo, mode string) (*models.Chart, error) {
	chartXML, err := readZipFile(r, ci.chartPath)
	if err != nil || chartXML == nil {
		return nil, err
	}

	chart := parseChartXML(chartXML, ci.name, ci.left, ci.top, ci.width, ci.height)
	if chart == nil {
		return nil, nil
	}

	// Apply mode filtering
	if mode != "verbose" {
		chart.W = nil
		chart.H = nil
	}

	return chart, nil
}

// parseChartXML parses chart XML content.
func parseChartXML(data []byte, name string, left, top, width, height int) *models.Chart {
	decoder := xml.NewDecoder(strings.NewReader(string(data)))

	var chartType string
	var title string
	var yAxisTitle string
	var yAxisRange []float64
	var series []models.ChartSeries

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		if se, ok := token.(xml.StartElement); ok {
			switch se.Name.Local {
			case "chart":
				chartType, title, yAxisTitle, yAxisRange, series = parseChartElement(decoder)
			}
		}
	}

	if chartType == "" {
		chartType = "unknown"
	}

	w := width
	h := height

	return &models.Chart{
		Name:       name,
		ChartType:  chartType,
		Title:      title,
		YAxisTitle: yAxisTitle,
		YAxisRange: yAxisRange,
		W:          &w,
		H:          &h,
		Series:     series,
		L:          left,
		T:          top,
	}
}

// parseChartElement parses c:chart element.
func parseChartElement(decoder *xml.Decoder) (chartType, title, yAxisTitle string, yAxisRange []float64, series []models.ChartSeries) {
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
			case "title":
				title = parseChartTitle(decoder)
				depth--
			case "plotArea":
				chartType, yAxisTitle, yAxisRange, series = parsePlotArea(decoder)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// parseChartTitle parses chart title element.
func parseChartTitle(decoder *xml.Decoder) string {
	var title string
	depth := 1

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "t" {
				if txt, err := readElementText(decoder); err == nil {
					title = strings.TrimSpace(txt)
				}
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return title
}

// parsePlotArea parses plot area element.
func parsePlotArea(decoder *xml.Decoder) (chartType, yAxisTitle string, yAxisRange []float64, series []models.ChartSeries) {
	depth := 1

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			// Check for chart type elements
			if ct, ok := ChartTypeMap[t.Name.Local]; ok {
				chartType = ct
				series = parseChartSeries(decoder)
				depth--
			} else if t.Name.Local == "valAx" {
				yAxisTitle, yAxisRange = parseValueAxis(decoder)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// parseChartSeries parses series elements within a chart type.
func parseChartSeries(decoder *xml.Decoder) []models.ChartSeries {
	var series []models.ChartSeries
	depth := 1

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "ser" {
				s := parseSingleSeries(decoder)
				series = append(series, s)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return series
}

// parseSingleSeries parses a single series element.
func parseSingleSeries(decoder *xml.Decoder) models.ChartSeries {
	var s models.ChartSeries
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
			case "tx":
				s.Name, s.NameRange = parseSeriesName(decoder)
				depth--
			case "cat":
				s.XRange = parseSeriesRange(decoder)
				depth--
			case "val":
				s.YRange = parseSeriesRange(decoder)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return s
}

// parseSeriesName parses series name from tx element.
func parseSeriesName(decoder *xml.Decoder) (name, nameRange string) {
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
			case "f":
				if txt, err := readElementText(decoder); err == nil {
					nameRange = strings.TrimSpace(txt)
				}
				depth--
			case "v":
				if txt, err := readElementText(decoder); err == nil {
					name = strings.TrimSpace(txt)
				}
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// parseSeriesRange parses range reference from cat or val element.
func parseSeriesRange(decoder *xml.Decoder) string {
	depth := 1

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			break
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			if t.Name.Local == "f" {
				if txt, err := readElementText(decoder); err == nil {
					return strings.TrimSpace(txt)
				}
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return ""
}

// parseValueAxis parses value axis element.
func parseValueAxis(decoder *xml.Decoder) (title string, axisRange []float64) {
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
			case "title":
				title = parseChartTitle(decoder)
				depth--
			case "scaling":
				axisRange = parseAxisScaling(decoder)
				depth--
			}
		case xml.EndElement:
			depth--
		}
	}

	return
}

// parseAxisScaling parses axis scaling element.
func parseAxisScaling(decoder *xml.Decoder) []float64 {
	var min, max *float64
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
			case "min":
				for _, attr := range t.Attr {
					if attr.Name.Local == "val" {
						if v, err := strconv.ParseFloat(attr.Value, 64); err == nil {
							min = &v
						}
					}
				}
			case "max":
				for _, attr := range t.Attr {
					if attr.Name.Local == "val" {
						if v, err := strconv.ParseFloat(attr.Value, 64); err == nil {
							max = &v
						}
					}
				}
			}
		case xml.EndElement:
			depth--
		}
	}

	if min != nil && max != nil {
		return []float64{*min, *max}
	}
	return nil
}
