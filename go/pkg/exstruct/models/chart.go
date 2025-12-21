package models

// ChartSeries represents series metadata for a chart.
type ChartSeries struct {
	// Name is the series display name.
	Name string `json:"name"`
	// NameRange is the range reference for the series name.
	NameRange string `json:"name_range,omitempty"`
	// XRange is the range reference for X axis values.
	XRange string `json:"x_range,omitempty"`
	// YRange is the range reference for Y axis values.
	YRange string `json:"y_range,omitempty"`
}

// Chart represents chart metadata including series and layout.
type Chart struct {
	// Name is the chart name.
	Name string `json:"name"`
	// ChartType is the chart type (e.g., Column, Line).
	ChartType string `json:"chart_type"`
	// Title is the chart title.
	Title string `json:"title,omitempty"`
	// YAxisTitle is the Y-axis title.
	YAxisTitle string `json:"y_axis_title,omitempty"`
	// YAxisRange is the Y-axis range [min, max] when available.
	YAxisRange []float64 `json:"y_axis_range,omitempty"`
	// W is the chart width in pixels (nil if unknown or not verbose mode).
	W *int `json:"w,omitempty"`
	// H is the chart height in pixels (nil if unknown or not verbose mode).
	H *int `json:"h,omitempty"`
	// Series is the list of series included in the chart.
	Series []ChartSeries `json:"series"`
	// L is the left offset in pixels.
	L int `json:"l"`
	// T is the top offset in pixels.
	T int `json:"t"`
}
