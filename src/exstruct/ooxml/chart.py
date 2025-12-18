"""ChartML parser for extracting charts from xlsx files.

Parses xl/charts/chart*.xml to extract chart information including
type, title, series data, and axis ranges.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import TYPE_CHECKING, Literal
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from exstruct.models import Chart, ChartSeries
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


# XML namespaces used in ChartML
NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
}

# Mapping from OOXML chart element tags to chart type names
CHART_TYPE_MAP: dict[str, str] = {
    "lineChart": "Line",
    "line3DChart": "3DLine",
    "barChart": "Bar",
    "bar3DChart": "3DBar",
    "areaChart": "Area",
    "area3DChart": "3DArea",
    "pieChart": "Pie",
    "pie3DChart": "3DPie",
    "doughnutChart": "Doughnut",
    "scatterChart": "XYScatter",
    "bubbleChart": "Bubble",
    "radarChart": "Radar",
    "surfaceChart": "Surface",
    "surface3DChart": "3DSurface",
    "stockChart": "Stock",
    "ofPieChart": "PieOfPie",
}


def _get_chart_title(chart_elem: Element) -> str | None:
    """Extract chart title from chart element.

    Args:
        chart_elem: c:chart element.

    Returns:
        Chart title or None.
    """
    title_elem = chart_elem.find(".//c:title", NS)
    if title_elem is None:
        return None

    # Try rich text first
    for t_elem in title_elem.findall(".//a:t", NS):
        if t_elem.text:
            return t_elem.text.strip()

    # Try string reference
    str_ref = title_elem.find(".//c:strRef/c:strCache/c:pt/c:v", NS)
    if str_ref is not None and str_ref.text:
        return str_ref.text.strip()

    return None


def _get_chart_type(plot_area: Element) -> str:
    """Determine chart type from plot area element.

    Args:
        plot_area: c:plotArea element.

    Returns:
        Chart type name.
    """
    for tag, type_name in CHART_TYPE_MAP.items():
        if plot_area.find(f"c:{tag}", NS) is not None:
            return type_name
    return "unknown"


def _extract_series_name(ser_elem: Element) -> tuple[str, str | None]:
    """Extract series name and name range from series element.

    Args:
        ser_elem: c:ser element.

    Returns:
        Tuple of (name, name_range).
    """
    name = ""
    name_range: str | None = None

    tx = ser_elem.find("c:tx", NS)
    if tx is None:
        return (name, name_range)

    str_ref = tx.find("c:strRef", NS)
    if str_ref is not None:
        f_elem = str_ref.find("c:f", NS)
        if f_elem is not None and f_elem.text:
            name_range = f_elem.text
        v_elem = str_ref.find(".//c:v", NS)
        if v_elem is not None and v_elem.text:
            name = v_elem.text

    v_elem = tx.find("c:v", NS)
    if v_elem is not None and v_elem.text:
        name = v_elem.text

    return (name, name_range)


def _extract_range_from_ref(parent: Element | None, ref_types: list[str]) -> str | None:
    """Extract range formula from reference element.

    Args:
        parent: Parent element containing reference.
        ref_types: List of reference type tags to try.

    Returns:
        Range formula or None.
    """
    if parent is None:
        return None

    for ref_type in ref_types:
        ref = parent.find(ref_type, NS)
        if ref is not None:
            f_elem = ref.find("c:f", NS)
            if f_elem is not None and f_elem.text:
                return f_elem.text
    return None


def _get_series_data(ser_elem: Element) -> ChartSeries:
    """Extract series data from series element.

    Args:
        ser_elem: c:ser element.

    Returns:
        ChartSeries model.
    """
    name, name_range = _extract_series_name(ser_elem)
    x_range = _extract_range_from_ref(ser_elem.find("c:cat", NS), ["c:strRef", "c:numRef"])
    y_range = _extract_range_from_ref(ser_elem.find("c:val", NS), ["c:numRef"])

    return ChartSeries(
        name=name,
        name_range=name_range,
        x_range=x_range,
        y_range=y_range,
    )


def _get_axis_range(plot_area: Element, axis_type: str) -> list[float]:
    """Extract axis min/max range.

    Args:
        plot_area: c:plotArea element.
        axis_type: Axis type (valAx, catAx, dateAx).

    Returns:
        List of [min, max] or empty list.
    """
    axis = plot_area.find(f"c:{axis_type}", NS)
    if axis is None:
        return []

    scaling = axis.find("c:scaling", NS)
    if scaling is None:
        return []

    min_elem = scaling.find("c:min", NS)
    max_elem = scaling.find("c:max", NS)

    result: list[float] = []
    if min_elem is not None:
        val = min_elem.get("val")
        if val:
            try:
                result.append(float(val))
            except ValueError:
                pass

    if max_elem is not None:
        val = max_elem.get("val")
        if val:
            try:
                result.append(float(val))
            except ValueError:
                pass

    return result if len(result) == 2 else []


def _get_axis_title(plot_area: Element, axis_type: str) -> str:
    """Extract axis title.

    Args:
        plot_area: c:plotArea element.
        axis_type: Axis type (valAx, catAx, dateAx).

    Returns:
        Axis title or empty string.
    """
    axis = plot_area.find(f"c:{axis_type}", NS)
    if axis is None:
        return ""

    title = axis.find("c:title", NS)
    if title is None:
        return ""

    for t_elem in title.findall(".//a:t", NS):
        if t_elem.text:
            return t_elem.text.strip()

    return ""


def _parse_chart_xml(
    chart_xml: bytes, chart_name: str, left: int, top: int, width: int, height: int
) -> Chart | None:
    """Parse a chart XML file and extract chart data.

    Args:
        chart_xml: Raw XML content.
        chart_name: Chart name.
        left: Left position in pixels.
        top: Top position in pixels.
        width: Width in pixels.
        height: Height in pixels.

    Returns:
        Chart model or None on error.
    """
    try:
        root = ET.fromstring(chart_xml)
    except ET.ParseError as e:
        logger.warning("Failed to parse chart XML: %s", e)
        return None

    chart_elem = root.find("c:chart", NS)
    if chart_elem is None:
        return None

    plot_area = chart_elem.find("c:plotArea", NS)
    if plot_area is None:
        return None

    # Get chart type
    chart_type = _get_chart_type(plot_area)

    # Get title
    title = _get_chart_title(chart_elem)

    # Get series data
    series_list: list[ChartSeries] = []
    for chart_type_elem in plot_area:
        tag = chart_type_elem.tag.split("}")[-1] if "}" in chart_type_elem.tag else chart_type_elem.tag
        if tag in CHART_TYPE_MAP:
            for ser in chart_type_elem.findall("c:ser", NS):
                series = _get_series_data(ser)
                series_list.append(series)

    # Get Y axis info
    y_axis_title = _get_axis_title(plot_area, "valAx")
    y_axis_range = _get_axis_range(plot_area, "valAx")

    return Chart(
        name=chart_name,
        chart_type=chart_type,
        title=title,
        y_axis_title=y_axis_title,
        y_axis_range=y_axis_range,
        w=width,
        h=height,
        series=series_list,
        l=left,
        t=top,
    )


def _get_chart_positions_from_drawing(
    zf: ZipFile, drawing_path: str
) -> dict[str, tuple[str, int, int, int, int]]:
    """Extract chart positions from drawing XML.

    Args:
        zf: Open ZipFile.
        drawing_path: Path to drawing XML within zip.

    Returns:
        Dict mapping chart rId to (chart_path, left, top, width, height).
    """
    result: dict[str, tuple[str, int, int, int, int]] = {}

    try:
        drawing_xml = zf.read(drawing_path)
        root = ET.fromstring(drawing_xml)
    except (KeyError, ET.ParseError):
        return result

    # Find all graphicFrame elements (charts are embedded in these)
    for anchor in root.findall(".//xdr:twoCellAnchor", NS):
        graphic_frame = anchor.find("xdr:graphicFrame", NS)
        if graphic_frame is None:
            continue

        # Get chart reference
        chart_ref = graphic_frame.find(
            ".//c:chart",
            {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"},
        )
        if chart_ref is None:
            continue

        r_id = chart_ref.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        if not r_id:
            continue

        # Get chart name
        cnv_pr = graphic_frame.find(".//xdr:cNvPr", NS)
        chart_name = cnv_pr.get("name", f"Chart_{r_id}") if cnv_pr is not None else f"Chart_{r_id}"

        # Get position from xfrm
        xfrm = graphic_frame.find("xdr:xfrm", NS)
        if xfrm is not None:
            off = xfrm.find("a:off", NS)
            ext = xfrm.find("a:ext", NS)
            if off is not None and ext is not None:
                try:
                    x = int(off.get("x", "0"))
                    y = int(off.get("y", "0"))
                    cx = int(ext.get("cx", "0"))
                    cy = int(ext.get("cy", "0"))
                    result[r_id] = (
                        chart_name,
                        emu_to_pixels(x),
                        emu_to_pixels(y),
                        emu_to_pixels(cx),
                        emu_to_pixels(cy),
                    )
                    continue
                except ValueError:
                    pass

        # Fallback: estimate from anchor cells (simplified)
        result[r_id] = (chart_name, 0, 0, 400, 300)

    return result


def _resolve_chart_paths(
    zf: ZipFile, drawing_path: str, chart_positions: dict[str, tuple[str, int, int, int, int]]
) -> dict[str, tuple[str, str, int, int, int, int]]:
    """Resolve chart rIds to actual file paths.

    Args:
        zf: Open ZipFile.
        drawing_path: Path to drawing XML.
        chart_positions: Dict from _get_chart_positions_from_drawing.

    Returns:
        Dict mapping chart path to (name, path, left, top, width, height).
    """
    result: dict[str, tuple[str, str, int, int, int, int]] = {}

    # Get drawing rels file
    rels_path = drawing_path.replace("drawings/", "drawings/_rels/").replace(
        ".xml", ".xml.rels"
    )

    try:
        rels_xml = zf.read(rels_path)
        rels_root = ET.fromstring(rels_xml)
    except (KeyError, ET.ParseError):
        return result

    rels_ns = {"": "http://schemas.openxmlformats.org/package/2006/relationships"}

    for rel in rels_root.findall("Relationship", rels_ns):
        r_id = rel.get("Id", "")
        target = rel.get("Target", "")
        rel_type = rel.get("Type", "")

        if "chart" not in rel_type.lower():
            continue

        if r_id not in chart_positions:
            continue

        # Resolve path
        chart_path = _resolve_relative_path(target, "xl/charts")

        name, left, top, width, height = chart_positions[r_id]
        result[chart_path] = (name, chart_path, left, top, width, height)

    return result


def _read_sheets_info(zf: ZipFile) -> dict[str, str]:
    """Read sheet rId to name mapping from workbook.xml.

    Args:
        zf: Open ZipFile.

    Returns:
        Dict mapping rId to sheet name.
    """
    try:
        workbook_xml = zf.read("xl/workbook.xml")
        wb_root = ET.fromstring(workbook_xml)
    except (KeyError, ET.ParseError):
        return {}

    wb_ns = {"": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets_info: dict[str, str] = {}

    for sheet in wb_root.findall(".//sheet", wb_ns):
        name = sheet.get("name", "")
        r_id = sheet.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
            "",
        )
        if name and r_id:
            sheets_info[r_id] = name

    return sheets_info


def _read_sheet_files(zf: ZipFile, sheets_info: dict[str, str]) -> dict[str, str]:
    """Read sheet name to file path mapping from workbook.xml.rels.

    Args:
        zf: Open ZipFile.
        sheets_info: Dict mapping rId to sheet name.

    Returns:
        Dict mapping sheet name to file path.
    """
    try:
        wb_rels_xml = zf.read("xl/_rels/workbook.xml.rels")
        rels_root = ET.fromstring(wb_rels_xml)
    except (KeyError, ET.ParseError):
        return {}

    rels_ns = {"": "http://schemas.openxmlformats.org/package/2006/relationships"}
    sheet_files: dict[str, str] = {}

    for rel in rels_root.findall("Relationship", rels_ns):
        r_id = rel.get("Id", "")
        target = rel.get("Target", "")
        if r_id in sheets_info and "worksheet" in target.lower():
            sheet_files[sheets_info[r_id]] = _resolve_relative_path(target, "xl")

    return sheet_files


def _find_sheet_charts(
    zf: ZipFile, sheet_path: str
) -> list[tuple[str, str, int, int, int, int]]:
    """Find charts for a single sheet.

    Args:
        zf: Open ZipFile.
        sheet_path: Path to sheet XML.

    Returns:
        List of (name, chart_path, left, top, width, height).
    """
    rels_ns = {"": "http://schemas.openxmlformats.org/package/2006/relationships"}
    rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/").replace(
        ".xml", ".xml.rels"
    )

    try:
        sheet_rels_xml = zf.read(rels_path)
        sheet_rels_root = ET.fromstring(sheet_rels_xml)
    except (KeyError, ET.ParseError):
        return []

    for rel in sheet_rels_root.findall("Relationship", rels_ns):
        rel_type = rel.get("Type", "")
        if "drawing" not in rel_type.lower():
            continue

        target = rel.get("Target", "")
        drawing_path = _resolve_relative_path(target, "xl/drawings")

        chart_positions = _get_chart_positions_from_drawing(zf, drawing_path)
        if not chart_positions:
            continue

        chart_info = _resolve_chart_paths(zf, drawing_path, chart_positions)
        return list(chart_info.values())

    return []


def _get_sheet_chart_map(
    xlsx_path: Path,
) -> dict[str, list[tuple[str, str, int, int, int, int]]]:
    """Map sheet names to their chart info.

    Args:
        xlsx_path: Path to xlsx file.

    Returns:
        Dict mapping sheet name to list of (name, chart_path, left, top, width, height).
    """
    sheet_charts: dict[str, list[tuple[str, str, int, int, int, int]]] = {}

    with ZipFile(xlsx_path, "r") as zf:
        sheets_info = _read_sheets_info(zf)
        if not sheets_info:
            return sheet_charts

        sheet_files = _read_sheet_files(zf, sheets_info)

        for sheet_name, sheet_path in sheet_files.items():
            charts = _find_sheet_charts(zf, sheet_path)
            if charts:
                sheet_charts[sheet_name] = charts

    return sheet_charts


def get_charts_ooxml(
    xlsx_path: str | Path, mode: Literal["light", "standard", "verbose"] = "standard"
) -> dict[str, list[Chart]]:
    """Extract charts from xlsx file using OOXML parsing.

    This function provides COM-free chart extraction for Linux/macOS.

    Args:
        xlsx_path: Path to xlsx file.
        mode: Output mode (light, standard, verbose).

    Returns:
        Dict mapping sheet name to list of Chart models.
    """
    xlsx_path = Path(xlsx_path)
    result: dict[str, list[Chart]] = {}

    if not xlsx_path.exists():
        logger.warning("File not found: %s", xlsx_path)
        return result

    sheet_chart_map = _get_sheet_chart_map(xlsx_path)

    with ZipFile(xlsx_path, "r") as zf:
        for sheet_name, chart_infos in sheet_chart_map.items():
            charts: list[Chart] = []

            for name, chart_path, left, top, width, height in chart_infos:
                try:
                    chart_xml = zf.read(chart_path)
                    chart = _parse_chart_xml(
                        chart_xml, name, left, top, width, height
                    )
                    if chart is not None:
                        # Apply mode-specific filtering
                        if mode != "verbose":
                            chart = Chart(
                                name=chart.name,
                                chart_type=chart.chart_type,
                                title=chart.title,
                                y_axis_title=chart.y_axis_title,
                                y_axis_range=chart.y_axis_range,
                                w=None,
                                h=None,
                                series=chart.series,
                                l=chart.l,
                                t=chart.t,
                            )
                        charts.append(chart)
                except KeyError:
                    logger.debug("Chart not found: %s", chart_path)

            result[sheet_name] = charts

    return result
