// Package exstruct provides Excel structured extraction functionality.
package exstruct

// Mode represents the extraction mode.
type Mode string

const (
	// ModeLight extracts cells and table candidates only (no shapes or charts).
	ModeLight Mode = "light"
	// ModeStandard extracts cells, shapes with text or connectors, charts, and table candidates.
	ModeStandard Mode = "standard"
	// ModeVerbose extracts all data including shape dimensions, cell hyperlinks, and chart dimensions.
	ModeVerbose Mode = "verbose"
)

// Options configures extraction behavior.
type Options struct {
	// Mode specifies the extraction mode (light, standard, verbose).
	Mode Mode
	// IncludeLinks specifies whether to include cell hyperlinks.
	// If nil, defaults to true for verbose mode, false otherwise.
	IncludeLinks *bool
	// IncludePrintAreas specifies whether to include print areas.
	// If nil, defaults to false for light mode, true otherwise.
	IncludePrintAreas *bool
}

// DefaultOptions returns default extraction options.
func DefaultOptions() Options {
	return Options{
		Mode: ModeStandard,
	}
}

// ShouldIncludeLinks returns whether to include cell hyperlinks.
func (o Options) ShouldIncludeLinks() bool {
	if o.IncludeLinks != nil {
		return *o.IncludeLinks
	}
	return o.Mode == ModeVerbose
}

// ShouldIncludePrintAreas returns whether to include print areas.
func (o Options) ShouldIncludePrintAreas() bool {
	if o.IncludePrintAreas != nil {
		return *o.IncludePrintAreas
	}
	return o.Mode != ModeLight
}
