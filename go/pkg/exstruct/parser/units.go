// Package parser provides Excel file parsing utilities.
package parser

// EMUPerPixel is the number of EMUs (English Metric Units) per pixel at 96 DPI.
// 1 inch = 914400 EMU, 1 inch = 96 pixels at 96 DPI
// Therefore: 914400 / 96 = 9525 EMU per pixel
const EMUPerPixel = 9525

// EMUToPixels converts EMU (English Metric Units) to pixels at 96 DPI.
// Excel uses EMU for internal coordinate representation.
// 914400 EMU = 1 inch, and at 96 DPI, 1 inch = 96 pixels.
func EMUToPixels(emu int64) int {
	return int(emu / EMUPerPixel)
}
