package models

// Shape represents shape metadata including position, size, text, and styling.
type Shape struct {
	// ID is the sequential shape id within the sheet (if applicable).
	ID *int `json:"id,omitempty"`
	// Text is the visible text content of the shape.
	Text string `json:"text"`
	// L is the left offset in pixels.
	L int `json:"l"`
	// T is the top offset in pixels.
	T int `json:"t"`
	// W is the shape width in pixels (nil if unknown or not verbose mode).
	W *int `json:"w,omitempty"`
	// H is the shape height in pixels (nil if unknown or not verbose mode).
	H *int `json:"h,omitempty"`
	// Type is the Excel shape type name.
	Type string `json:"type,omitempty"`
	// Rotation is the rotation angle in degrees.
	Rotation *float64 `json:"rotation,omitempty"`
	// BeginArrowStyle is the arrow style enum for the start of a connector.
	BeginArrowStyle *int `json:"begin_arrow_style,omitempty"`
	// EndArrowStyle is the arrow style enum for the end of a connector.
	EndArrowStyle *int `json:"end_arrow_style,omitempty"`
	// BeginID is the shape id at the start of a connector.
	BeginID *int `json:"begin_id,omitempty"`
	// EndID is the shape id at the end of a connector.
	EndID *int `json:"end_id,omitempty"`
	// Direction is the connector direction (compass heading: N, NE, E, SE, S, SW, W, NW).
	Direction string `json:"direction,omitempty"`
}
