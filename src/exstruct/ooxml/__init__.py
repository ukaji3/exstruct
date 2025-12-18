"""OOXML (Office Open XML) parsers for extracting shapes and charts without COM.

This module provides pure-Python parsers for reading shapes and charts
directly from xlsx files, enabling Linux/macOS support.
"""

from exstruct.ooxml.chart import get_charts_ooxml
from exstruct.ooxml.drawing import get_shapes_ooxml

__all__ = ["get_shapes_ooxml", "get_charts_ooxml"]
