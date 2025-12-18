"""DrawingML parser for extracting shapes from xlsx files.

Parses xl/drawings/drawing*.xml to extract shape information including
position, size, text, type, and connector arrow styles.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import TYPE_CHECKING, Literal
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from exstruct.models import Shape
from exstruct.ooxml.units import emu_to_pixels

if TYPE_CHECKING:
    from xml.etree.ElementTree import Element

logger = logging.getLogger(__name__)


def _resolve_relative_path(target: str, base_dir: str) -> str:
    """Resolve relative path from target.

    Args:
        target: Target path (may start with ..).
        base_dir: Base directory for non-relative paths.

    Returns:
        Resolved path within xl/ directory.
    """
    if target.startswith("../"):
        clean = target
        while clean.startswith("../"):
            clean = clean[3:]
        return f"xl/{clean}"
    if target.startswith("/"):
        return f"{base_dir}{target}"
    return f"{base_dir}/{target}"


# XML namespaces used in DrawingML
NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Mapping from OOXML preset geometry to ExStruct type labels
PRESET_GEOM_MAP: dict[str, str] = {
    "flowChartProcess": "AutoShape-FlowchartProcess",
    "flowChartDecision": "AutoShape-FlowchartDecision",
    "flowChartTerminator": "AutoShape-FlowchartTerminator",
    "flowChartData": "AutoShape-FlowchartData",
    "flowChartDocument": "AutoShape-FlowchartDocument",
    "flowChartMultidocument": "AutoShape-FlowchartMultidocument",
    "flowChartPredefinedProcess": "AutoShape-FlowchartPredefinedProcess",
    "flowChartInternalStorage": "AutoShape-FlowchartInternalStorage",
    "flowChartPreparation": "AutoShape-FlowchartPreparation",
    "flowChartManualInput": "AutoShape-FlowchartManualInput",
    "flowChartManualOperation": "AutoShape-FlowchartManualOperation",
    "flowChartConnector": "AutoShape-FlowchartConnector",
    "flowChartOffpageConnector": "AutoShape-FlowchartOffpageConnector",
    "flowChartPunchedCard": "AutoShape-FlowchartCard",
    "flowChartPunchedTape": "AutoShape-FlowchartPunchedTape",
    "flowChartSummingJunction": "AutoShape-FlowchartSummingJunction",
    "flowChartOr": "AutoShape-FlowchartOr",
    "flowChartCollate": "AutoShape-FlowchartCollate",
    "flowChartSort": "AutoShape-FlowchartSort",
    "flowChartExtract": "AutoShape-FlowchartExtract",
    "flowChartMerge": "AutoShape-FlowchartMerge",
    "flowChartStoredData": "AutoShape-FlowchartStoredData",
    "flowChartDelay": "AutoShape-FlowchartDelay",
    "flowChartMagneticDisk": "AutoShape-FlowchartMagneticDisk",
    "flowChartMagneticDrum": "AutoShape-FlowchartSequentialAccessStorage",
    "flowChartDisplay": "AutoShape-FlowchartDisplay",
    "rect": "AutoShape-Rectangle",
    "roundRect": "AutoShape-RoundedRectangle",
    "ellipse": "AutoShape-Oval",
    "diamond": "AutoShape-Diamond",
    "triangle": "AutoShape-IsoscelesTriangle",
    "rtTriangle": "AutoShape-RightTriangle",
    "parallelogram": "AutoShape-Parallelogram",
    "trapezoid": "AutoShape-Trapezoid",
    "pentagon": "AutoShape-Pentagon",
    "hexagon": "AutoShape-Hexagon",
    "heptagon": "AutoShape-Heptagon",
    "octagon": "AutoShape-Octagon",
    "star4": "AutoShape-4pointStar",
    "star5": "AutoShape-5pointStar",
    "star6": "AutoShape-6pointStar",
    "star7": "AutoShape-7pointStar",
    "star8": "AutoShape-8pointStar",
    "star10": "AutoShape-10pointStar",
    "star12": "AutoShape-12pointStar",
    "star16": "AutoShape-16pointStar",
    "star24": "AutoShape-24pointStar",
    "star32": "AutoShape-32pointStar",
    "rightArrow": "AutoShape-RightArrow",
    "leftArrow": "AutoShape-LeftArrow",
    "upArrow": "AutoShape-UpArrow",
    "downArrow": "AutoShape-DownArrow",
    "leftRightArrow": "AutoShape-LeftRightArrow",
    "upDownArrow": "AutoShape-UpDownArrow",
    "bentArrow": "AutoShape-BentArrow",
    "uturnArrow": "AutoShape-UTurnArrow",
    "curvedRightArrow": "AutoShape-CurvedRightArrow",
    "curvedLeftArrow": "AutoShape-CurvedLeftArrow",
    "curvedUpArrow": "AutoShape-CurvedUpArrow",
    "curvedDownArrow": "AutoShape-CurvedDownArrow",
    "stripedRightArrow": "AutoShape-StripedRightArrow",
    "notchedRightArrow": "AutoShape-NotchedRightArrow",
    "chevron": "AutoShape-Chevron",
    "homePlate": "AutoShape-Pentagon",
    "callout1": "AutoShape-LineCallout1",
    "callout2": "AutoShape-LineCallout2",
    "callout3": "AutoShape-LineCallout3",
    "accentCallout1": "AutoShape-LineCallout1AccentBar",
    "accentCallout2": "AutoShape-LineCallout2AccentBar",
    "accentCallout3": "AutoShape-LineCallout3AccentBar",
    "cloudCallout": "AutoShape-CloudCallout",
    "wedgeRectCallout": "AutoShape-RectangularCallout",
    "wedgeRoundRectCallout": "AutoShape-RoundedRectangularCallout",
    "wedgeEllipseCallout": "AutoShape-OvalCallout",
    "straightConnector1": "Line",
    "bentConnector2": "AutoShape-Connector",
    "bentConnector3": "AutoShape-Connector",
    "bentConnector4": "AutoShape-Connector",
    "bentConnector5": "AutoShape-Connector",
    "curvedConnector2": "AutoShape-Connector",
    "curvedConnector3": "AutoShape-Connector",
    "curvedConnector4": "AutoShape-Connector",
    "curvedConnector5": "AutoShape-Connector",
    "line": "Line",
    "textBox": "TextBox",
}

# Arrow head type mapping (OOXML -> Excel COM style number)
ARROW_HEAD_MAP: dict[str, int] = {
    "none": 1,
    "triangle": 2,
    "stealth": 3,
    "diamond": 4,
    "oval": 5,
    "arrow": 2,
}


def _get_text_from_element(elem: Element) -> str:
    """Extract all text content from a shape element.

    Args:
        elem: XML element containing text body.

    Returns:
        Concatenated text content, stripped.
    """
    texts: list[str] = []
    for t_elem in elem.findall(".//a:t", NS):
        if t_elem.text:
            texts.append(t_elem.text)
    return "".join(texts).strip()


def _get_xfrm_position(elem: Element) -> tuple[int, int, int, int] | None:
    """Extract position and size from xfrm element.

    Args:
        elem: XML element containing spPr/xfrm.

    Returns:
        Tuple of (left, top, width, height) in pixels, or None if not found.
    """
    xfrm = elem.find(".//a:xfrm", NS)
    if xfrm is None:
        return None

    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)

    if off is None or ext is None:
        return None

    try:
        x = int(off.get("x", "0"))
        y = int(off.get("y", "0"))
        cx = int(ext.get("cx", "0"))
        cy = int(ext.get("cy", "0"))
        return (emu_to_pixels(x), emu_to_pixels(y), emu_to_pixels(cx), emu_to_pixels(cy))
    except ValueError:
        return None


def _get_preset_geometry(elem: Element) -> str | None:
    """Extract preset geometry type from shape.

    Args:
        elem: XML element containing spPr.

    Returns:
        Preset geometry name or None.
    """
    prst_geom = elem.find(".//a:prstGeom", NS)
    if prst_geom is not None:
        return prst_geom.get("prst")
    return None


def _get_arrow_styles(elem: Element) -> tuple[int | None, int | None]:
    """Extract arrow head styles from connector line.

    Args:
        elem: XML element containing line properties.

    Returns:
        Tuple of (begin_arrow_style, end_arrow_style).
    """
    begin_style: int | None = None
    end_style: int | None = None

    ln = elem.find(".//a:ln", NS)
    if ln is None:
        return (None, None)

    head_end = ln.find("a:headEnd", NS)
    tail_end = ln.find("a:tailEnd", NS)

    if head_end is not None:
        head_type = head_end.get("type", "none")
        begin_style = ARROW_HEAD_MAP.get(head_type, 1)

    if tail_end is not None:
        tail_type = tail_end.get("type", "none")
        end_style = ARROW_HEAD_MAP.get(tail_type, 1)

    return (begin_style, end_style)


def _compute_direction(width: int, height: int) -> str | None:
    """Compute compass direction from connector dimensions.

    Args:
        width: Connector width in pixels.
        height: Connector height in pixels.

    Returns:
        Compass direction (N, NE, E, SE, S, SW, W, NW) or None.
    """
    import math

    if width == 0 and height == 0:
        return None

    angle = math.degrees(math.atan2(-height, width))
    if angle < 0:
        angle += 360

    # Map angle to compass direction
    if 337.5 <= angle or angle < 22.5:
        return "E"
    elif 22.5 <= angle < 67.5:
        return "NE"
    elif 67.5 <= angle < 112.5:
        return "N"
    elif 112.5 <= angle < 157.5:
        return "NW"
    elif 157.5 <= angle < 202.5:
        return "W"
    elif 202.5 <= angle < 247.5:
        return "SW"
    elif 247.5 <= angle < 292.5:
        return "S"
    else:
        return "SE"


def _get_rotation(elem: Element) -> float | None:
    """Extract rotation angle from xfrm element.

    Args:
        elem: XML element containing xfrm.

    Returns:
        Rotation in degrees or None.
    """
    xfrm = elem.find(".//a:xfrm", NS)
    if xfrm is None:
        return None

    rot_str = xfrm.get("rot")
    if rot_str is None:
        return None

    try:
        # OOXML rotation is in 1/60000 of a degree
        rot_emu = int(rot_str)
        rot_deg = rot_emu / 60000.0
        if abs(rot_deg) < 1e-6:
            return None
        return rot_deg
    except ValueError:
        return None


def _is_connector_shape(prst: str | None, elem: Element) -> bool:
    """Check if shape is a connector or line.

    Args:
        prst: Preset geometry name.
        elem: XML element.

    Returns:
        True if shape is a connector/line.
    """
    if prst is None:
        return False

    connector_keywords = ["Connector", "line", "straightConnector"]
    return any(kw.lower() in prst.lower() for kw in connector_keywords)


def _parse_shape_element(
    elem: Element, mode: str
) -> Shape | None:
    """Parse a single shape element into Shape model.

    Args:
        elem: xdr:sp or xdr:cxnSp element.
        mode: Output mode (light, standard, verbose).

    Returns:
        Shape model or None if should be skipped.
    """
    # Get shape name from cNvPr
    cnv_pr = elem.find(".//xdr:cNvPr", NS)
    shape_name = cnv_pr.get("name", "") if cnv_pr is not None else ""

    # Get position and size
    pos = _get_xfrm_position(elem)
    if pos is None:
        return None

    left, top, width, height = pos

    # Get text content
    text = _get_text_from_element(elem)

    # Get preset geometry
    prst = _get_preset_geometry(elem)

    # Determine type label
    if prst:
        type_label = PRESET_GEOM_MAP.get(prst, f"AutoShape-{prst}")
    elif shape_name:
        type_label = shape_name
    else:
        type_label = "Unknown"

    # Check if connector
    is_connector = _is_connector_shape(prst, elem)

    # Build shape object
    shape = Shape(
        text=text,
        l=left,
        t=top,
        w=width if mode == "verbose" else None,
        h=height if mode == "verbose" else None,
        type=type_label,
    )

    # Add connector-specific properties
    if is_connector:
        direction = _compute_direction(width, height)
        if direction:
            shape.direction = direction  # type: ignore[assignment]

        begin_style, end_style = _get_arrow_styles(elem)
        if begin_style is not None:
            shape.begin_arrow_style = begin_style
        if end_style is not None:
            shape.end_arrow_style = end_style

    # Add rotation if present
    rotation = _get_rotation(elem)
    if rotation is not None:
        shape.rotation = rotation

    return shape


def _parse_drawing_xml(
    drawing_xml: bytes, mode: str
) -> list[Shape]:
    """Parse a drawing XML file and extract shapes.

    Args:
        drawing_xml: Raw XML content.
        mode: Output mode.

    Returns:
        List of Shape models.
    """
    shapes: list[Shape] = []

    try:
        root = ET.fromstring(drawing_xml)
    except ET.ParseError as e:
        logger.warning("Failed to parse drawing XML: %s", e)
        return shapes

    # Find all shape elements (sp = shape, cxnSp = connector shape)
    for anchor in root.findall(".//xdr:twoCellAnchor", NS):
        # Regular shapes
        for sp in anchor.findall("xdr:sp", NS):
            shape = _parse_shape_element(sp, mode)
            if shape is not None:
                shapes.append(shape)

        # Connector shapes
        for cxn_sp in anchor.findall("xdr:cxnSp", NS):
            shape = _parse_shape_element(cxn_sp, mode)
            if shape is not None:
                shapes.append(shape)

    # Also check oneCellAnchor and absoluteAnchor
    for anchor in root.findall(".//xdr:oneCellAnchor", NS):
        for sp in anchor.findall("xdr:sp", NS):
            shape = _parse_shape_element(sp, mode)
            if shape is not None:
                shapes.append(shape)

    return shapes


def _get_sheet_drawing_map(xlsx_path: Path) -> dict[str, str]:
    """Map sheet names to their drawing XML paths.

    Args:
        xlsx_path: Path to xlsx file.

    Returns:
        Dict mapping sheet name to drawing XML path within zip.
    """
    sheet_drawing_map: dict[str, str] = {}

    with ZipFile(xlsx_path, "r") as zf:
        # Read workbook.xml to get sheet names and rIds
        try:
            workbook_xml = zf.read("xl/workbook.xml")
            wb_root = ET.fromstring(workbook_xml)
        except (KeyError, ET.ParseError):
            return sheet_drawing_map

        # Namespace for workbook
        wb_ns = {"": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        sheets_info: dict[str, str] = {}  # rId -> sheet name
        for sheet in wb_root.findall(".//sheet", wb_ns):
            name = sheet.get("name", "")
            r_id = sheet.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
                "",
            )
            if name and r_id:
                sheets_info[r_id] = name

        # Read workbook.xml.rels to map rId to sheet file
        try:
            wb_rels_xml = zf.read("xl/_rels/workbook.xml.rels")
            rels_root = ET.fromstring(wb_rels_xml)
        except (KeyError, ET.ParseError):
            return sheet_drawing_map

        rels_ns = {
            "": "http://schemas.openxmlformats.org/package/2006/relationships"
        }

        sheet_files: dict[str, str] = {}  # sheet name -> sheet file path
        for rel in rels_root.findall("Relationship", rels_ns):
            r_id = rel.get("Id", "")
            target = rel.get("Target", "")
            if r_id in sheets_info and "worksheet" in target.lower():
                sheet_files[sheets_info[r_id]] = _resolve_relative_path(target, "xl")

        # For each sheet, find its drawing relationship
        for sheet_name, sheet_path in sheet_files.items():
            rels_path = sheet_path.replace(
                "worksheets/", "worksheets/_rels/"
            ).replace(".xml", ".xml.rels")

            try:
                sheet_rels_xml = zf.read(rels_path)
                sheet_rels_root = ET.fromstring(sheet_rels_xml)
            except (KeyError, ET.ParseError):
                continue

            for rel in sheet_rels_root.findall("Relationship", rels_ns):
                rel_type = rel.get("Type", "")
                if "drawing" in rel_type.lower():
                    target = rel.get("Target", "")
                    # Resolve relative path
                    drawing_path = _resolve_relative_path(target, "xl/drawings")
                    sheet_drawing_map[sheet_name] = drawing_path
                    break

    return sheet_drawing_map


def get_shapes_ooxml(
    xlsx_path: str | Path, mode: Literal["light", "standard", "verbose"] = "standard"
) -> dict[str, list[Shape]]:
    """Extract shapes from xlsx file using OOXML parsing.

    This function provides COM-free shape extraction for Linux/macOS.

    Args:
        xlsx_path: Path to xlsx file.
        mode: Output mode (light, standard, verbose).

    Returns:
        Dict mapping sheet name to list of Shape models.
    """
    xlsx_path = Path(xlsx_path)
    result: dict[str, list[Shape]] = {}

    if not xlsx_path.exists():
        logger.warning("File not found: %s", xlsx_path)
        return result

    sheet_drawing_map = _get_sheet_drawing_map(xlsx_path)

    with ZipFile(xlsx_path, "r") as zf:
        for sheet_name, drawing_path in sheet_drawing_map.items():
            try:
                drawing_xml = zf.read(drawing_path)
                shapes = _parse_drawing_xml(drawing_xml, mode)
                result[sheet_name] = shapes
            except KeyError:
                logger.debug("Drawing not found: %s", drawing_path)
                result[sheet_name] = []

    return result
