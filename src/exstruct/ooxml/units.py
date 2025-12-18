"""Unit conversion utilities for OOXML coordinates.

EMU (English Metric Units) is the base unit in OOXML:
- 914400 EMU = 1 inch
- 1 inch = 72 points (Excel default)
- 1 point ≈ 1.333 pixels at 96 DPI
"""

# EMU per inch
EMU_PER_INCH: int = 914400

# Default DPI for Excel
DEFAULT_DPI: int = 96

# Points per inch
POINTS_PER_INCH: int = 72


def emu_to_pixels(emu: int, dpi: int = DEFAULT_DPI) -> int:
    """Convert EMU to pixels.

    Args:
        emu: Value in English Metric Units.
        dpi: Dots per inch (default 96).

    Returns:
        Value in pixels (rounded to nearest integer).
    """
    inches = emu / EMU_PER_INCH
    return round(inches * dpi)


def emu_to_points(emu: int) -> float:
    """Convert EMU to points.

    Args:
        emu: Value in English Metric Units.

    Returns:
        Value in points.
    """
    inches = emu / EMU_PER_INCH
    return inches * POINTS_PER_INCH
