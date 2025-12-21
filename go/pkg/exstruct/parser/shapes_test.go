package parser

import (
	"testing"
)

func TestComputeDirection(t *testing.T) {
	tests := []struct {
		width    int
		height   int
		expected string
	}{
		{100, 0, "E"},    // East (right)
		{0, -100, "N"},   // North (up)
		{0, 100, "S"},    // South (down)
		{-100, 0, "W"},   // West (left)
		{100, -100, "NE"}, // Northeast
		{100, 100, "SE"},  // Southeast
		{-100, 100, "SW"}, // Southwest
		{-100, -100, "NW"}, // Northwest
		{0, 0, ""},        // No direction
	}

	for _, tt := range tests {
		result := computeDirection(tt.width, tt.height)
		if result != tt.expected {
			t.Errorf("computeDirection(%d, %d) = %q, expected %q",
				tt.width, tt.height, result, tt.expected)
		}
	}
}

func TestIsConnectorShape(t *testing.T) {
	tests := []struct {
		prst      string
		typeLabel string
		expected  bool
	}{
		{"straightConnector1", "Line", true},
		{"bentConnector3", "AutoShape-Connector", true},
		{"line", "Line", true},
		{"rect", "AutoShape-Rectangle", false},
		{"flowChartProcess", "AutoShape-FlowchartProcess", false},
		{"", "Line", true},
		{"", "AutoShape-Connector", true},
	}

	for _, tt := range tests {
		result := isConnectorShape(tt.prst, tt.typeLabel)
		if result != tt.expected {
			t.Errorf("isConnectorShape(%q, %q) = %v, expected %v",
				tt.prst, tt.typeLabel, result, tt.expected)
		}
	}
}

func TestShouldIncludeShape(t *testing.T) {
	tests := []struct {
		text        string
		typeLabel   string
		isConnector bool
		mode        string
		expected    bool
	}{
		// Light mode excludes all
		{"text", "AutoShape", false, "light", false},
		{"", "Line", true, "light", false},
		// Verbose mode includes all
		{"", "AutoShape", false, "verbose", true},
		{"text", "AutoShape", false, "verbose", true},
		// Standard mode: include if text or connector
		{"text", "AutoShape", false, "standard", true},
		{"", "Line", true, "standard", true},
		{"", "AutoShape", false, "standard", false},
		{"", "AutoShape-Arrow", false, "standard", true},
	}

	for _, tt := range tests {
		result := shouldIncludeShape(tt.text, tt.typeLabel, tt.isConnector, tt.mode)
		if result != tt.expected {
			t.Errorf("shouldIncludeShape(%q, %q, %v, %q) = %v, expected %v",
				tt.text, tt.typeLabel, tt.isConnector, tt.mode, result, tt.expected)
		}
	}
}

func TestPresetGeomMap(t *testing.T) {
	// Test some key mappings
	tests := []struct {
		prst     string
		expected string
	}{
		{"flowChartProcess", "AutoShape-FlowchartProcess"},
		{"flowChartDecision", "AutoShape-FlowchartDecision"},
		{"rect", "AutoShape-Rectangle"},
		{"ellipse", "AutoShape-Oval"},
		{"straightConnector1", "Line"},
		{"textBox", "TextBox"},
	}

	for _, tt := range tests {
		result, ok := PresetGeomMap[tt.prst]
		if !ok {
			t.Errorf("PresetGeomMap[%q] not found", tt.prst)
			continue
		}
		if result != tt.expected {
			t.Errorf("PresetGeomMap[%q] = %q, expected %q",
				tt.prst, result, tt.expected)
		}
	}
}

func TestArrowHeadMap(t *testing.T) {
	tests := []struct {
		arrowType string
		expected  int
	}{
		{"none", 1},
		{"triangle", 2},
		{"stealth", 3},
		{"diamond", 4},
		{"oval", 5},
		{"arrow", 2},
	}

	for _, tt := range tests {
		result, ok := ArrowHeadMap[tt.arrowType]
		if !ok {
			t.Errorf("ArrowHeadMap[%q] not found", tt.arrowType)
			continue
		}
		if result != tt.expected {
			t.Errorf("ArrowHeadMap[%q] = %d, expected %d",
				tt.arrowType, result, tt.expected)
		}
	}
}

func TestResolveRelativePath(t *testing.T) {
	tests := []struct {
		target   string
		baseDir  string
		expected string
	}{
		{"../charts/chart1.xml", "xl/drawings", "xl/charts/chart1.xml"},
		{"/drawing1.xml", "xl/drawings", "xl/drawings/drawing1.xml"},
		{"drawing1.xml", "xl/drawings", "xl/drawings/drawing1.xml"},
	}

	for _, tt := range tests {
		result := resolveRelativePath(tt.target, tt.baseDir)
		if result != tt.expected {
			t.Errorf("resolveRelativePath(%q, %q) = %q, expected %q",
				tt.target, tt.baseDir, result, tt.expected)
		}
	}
}
