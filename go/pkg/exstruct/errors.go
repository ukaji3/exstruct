package exstruct

import (
	"errors"
	"fmt"
)

// ErrFileNotFound indicates the input file does not exist.
var ErrFileNotFound = errors.New("file not found")

// ErrInvalidFormat indicates the input file is not a valid xlsx format.
var ErrInvalidFormat = errors.New("invalid xlsx format")

// ExtractionError represents an error during extraction.
type ExtractionError struct {
	SheetName string
	Component string // "cells", "shapes", "charts", "tables", "print_areas"
	Err       error
}

func (e *ExtractionError) Error() string {
	return fmt.Sprintf("extraction error in sheet %q (%s): %v", e.SheetName, e.Component, e.Err)
}

func (e *ExtractionError) Unwrap() error {
	return e.Err
}

// NewExtractionError creates a new ExtractionError.
func NewExtractionError(sheetName, component string, err error) *ExtractionError {
	return &ExtractionError{
		SheetName: sheetName,
		Component: component,
		Err:       err,
	}
}
