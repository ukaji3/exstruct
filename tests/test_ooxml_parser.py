"""Tests for OOXML parser (COM-free shape/chart extraction).

These tests verify that the OOXML parser can extract shapes and charts
from xlsx files without requiring Excel COM (Windows).

Test xlsx files are generated programmatically using openpyxl + OOXML injection,
so no external sample files are needed.
"""

from pathlib import Path

import pytest

from exstruct.ooxml import get_charts_ooxml, get_shapes_ooxml
from exstruct.ooxml.units import emu_to_pixels, emu_to_points


# ---------------------------------------------------------------------------
# Fixture: generate test xlsx with shapes, connectors, and a chart
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def ooxml_test_xlsx(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Generate a test xlsx containing shapes, connectors, and a chart."""
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference

    tmp = tmp_path_factory.mktemp("ooxml")
    path = tmp / "test_ooxml.xlsx"

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Sheet1"

    # Sales data for chart
    ws.append(["月", "製品A", "製品B", "製品C"])
    ws.append(["Jan-25", 120, 80, 60])
    ws.append(["Feb-25", 135, 90, 64])
    ws.append(["Mar-25", 150, 100, 70])
    ws.append(["Apr-25", 170, 110, 72])
    ws.append(["May-25", 160, 120, 75])
    ws.append(["Jun-25", 180, 130, 80])

    # Add line chart
    chart = LineChart()
    chart.title = "売上データ"
    chart.y_axis.title = "売上"
    data = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=7)
    cats = Reference(ws, min_col=1, min_row=2, max_row=7)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "F1")

    wb.save(path)
    wb.close()

    # Inject shapes via OOXML XML manipulation
    from exstruct.edit.internal import (
        _ShapeSpec,
        _inject_shapes_into_xlsx,
        _points_to_emu,
    )

    shapes = [
        _ShapeSpec(
            sheet="Sheet1", shape_type="flowChartTerminator",
            anchor_col=0, anchor_row=10,
            width_emu=_points_to_emu(120), height_emu=_points_to_emu(40),
            text="開始", fill_color="4472C4",
        ),
        _ShapeSpec(
            sheet="Sheet1", shape_type="flowChartProcess",
            anchor_col=0, anchor_row=13,
            width_emu=_points_to_emu(120), height_emu=_points_to_emu(40),
            text="入力データ読み込み", fill_color="4472C4",
        ),
        _ShapeSpec(
            sheet="Sheet1", shape_type="flowChartDecision",
            anchor_col=0, anchor_row=16,
            width_emu=_points_to_emu(120), height_emu=_points_to_emu(60),
            text="フォーマット有効？", fill_color="FFC000",
        ),
        _ShapeSpec(
            sheet="Sheet1", shape_type="rect",
            anchor_col=3, anchor_row=16,
            width_emu=_points_to_emu(100), height_emu=_points_to_emu(40),
            text="エラー表示", fill_color="FF0000",
        ),
        _ShapeSpec(
            sheet="Sheet1", shape_type="flowChartTerminator",
            anchor_col=0, anchor_row=19,
            width_emu=_points_to_emu(120), height_emu=_points_to_emu(40),
            text="終了", fill_color="4472C4",
        ),
    ]
    _inject_shapes_into_xlsx(path, shapes)
    return path


# ---------------------------------------------------------------------------
# Unit conversion tests
# ---------------------------------------------------------------------------

class TestUnits:
    """Tests for EMU unit conversion."""

    def test_emu_to_pixels_default_dpi(self) -> None:
        assert emu_to_pixels(914400) == 96

    def test_emu_to_pixels_custom_dpi(self) -> None:
        assert emu_to_pixels(914400, dpi=72) == 72

    def test_emu_to_pixels_zero(self) -> None:
        assert emu_to_pixels(0) == 0

    def test_emu_to_points(self) -> None:
        assert emu_to_points(914400) == 72.0

    def test_emu_to_points_half_inch(self) -> None:
        assert emu_to_points(457200) == 36.0


# ---------------------------------------------------------------------------
# Shape extraction tests
# ---------------------------------------------------------------------------

class TestGetShapesOoxml:
    """Tests for OOXML shape extraction."""

    def test_returns_dict_for_valid_file(self, ooxml_test_xlsx: Path) -> None:
        result = get_shapes_ooxml(ooxml_test_xlsx)
        assert isinstance(result, dict)

    def test_returns_empty_dict_for_nonexistent_file(self, tmp_path: Path) -> None:
        result = get_shapes_ooxml(tmp_path / "nonexistent.xlsx")
        assert result == {}

    def test_shapes_have_required_fields(self, ooxml_test_xlsx: Path) -> None:
        result = get_shapes_ooxml(ooxml_test_xlsx)
        for shapes in result.values():
            for shape in shapes:
                assert hasattr(shape, "text")
                assert hasattr(shape, "l")
                assert hasattr(shape, "t")
                assert hasattr(shape, "type")
                assert shape.l >= 0
                assert shape.t >= 0

    def test_verbose_mode_includes_size(self, ooxml_test_xlsx: Path) -> None:
        result = get_shapes_ooxml(ooxml_test_xlsx, mode="verbose")
        for shapes in result.values():
            for shape in shapes:
                if shape.w is not None:
                    assert shape.w >= 0
                if shape.h is not None:
                    assert shape.h >= 0

    def test_standard_mode_excludes_size(self, ooxml_test_xlsx: Path) -> None:
        result = get_shapes_ooxml(ooxml_test_xlsx, mode="standard")
        for shapes in result.values():
            for shape in shapes:
                assert shape.w is None
                assert shape.h is None

    def test_light_mode_returns_empty(self, ooxml_test_xlsx: Path) -> None:
        result = get_shapes_ooxml(ooxml_test_xlsx, mode="light")
        assert result == {}


# ---------------------------------------------------------------------------
# Chart extraction tests
# ---------------------------------------------------------------------------

class TestGetChartsOoxml:
    """Tests for OOXML chart extraction."""

    def test_returns_dict_for_valid_file(self, ooxml_test_xlsx: Path) -> None:
        result = get_charts_ooxml(ooxml_test_xlsx)
        assert isinstance(result, dict)

    def test_returns_empty_dict_for_nonexistent_file(self, tmp_path: Path) -> None:
        result = get_charts_ooxml(tmp_path / "nonexistent.xlsx")
        assert result == {}

    def test_charts_have_required_fields(self, ooxml_test_xlsx: Path) -> None:
        result = get_charts_ooxml(ooxml_test_xlsx)
        all_charts = [c for charts in result.values() for c in charts]
        assert len(all_charts) > 0
        for chart in all_charts:
            assert chart.chart_type is not None
            assert chart.error is None

    def test_verbose_mode_includes_size(self, ooxml_test_xlsx: Path) -> None:
        result = get_charts_ooxml(ooxml_test_xlsx, mode="verbose")
        for charts in result.values():
            for chart in charts:
                assert chart.w is not None
                assert chart.h is not None
                assert chart.w >= 0
                assert chart.h >= 0

    def test_standard_mode_excludes_size(self, ooxml_test_xlsx: Path) -> None:
        result = get_charts_ooxml(ooxml_test_xlsx, mode="standard")
        for charts in result.values():
            for chart in charts:
                assert chart.w is None
                assert chart.h is None

    def test_chart_series_have_ranges(self, ooxml_test_xlsx: Path) -> None:
        result = get_charts_ooxml(ooxml_test_xlsx)
        all_charts = [c for charts in result.values() for c in charts]
        assert len(all_charts) > 0
        for chart in all_charts:
            for series in chart.series:
                assert hasattr(series, "y_range")
                assert hasattr(series, "x_range")
                assert hasattr(series, "name")


# ---------------------------------------------------------------------------
# Shapes: COM-equivalent tests
# ---------------------------------------------------------------------------

class TestShapesEquivalentToCom:
    """Tests equivalent to tests/com/test_shapes_extraction.py."""

    def test_図形の種別とテキストが抽出される(self, ooxml_test_xlsx: Path) -> None:
        shapes_by_sheet = get_shapes_ooxml(ooxml_test_xlsx)
        all_shapes = [s for shapes in shapes_by_sheet.values() for s in shapes]
        assert len(all_shapes) > 0
        shapes_with_text = [s for s in all_shapes if s.text]
        assert len(shapes_with_text) > 0
        for shape in shapes_with_text:
            assert shape.type is not None
            assert shape.l >= 0
            assert shape.t >= 0

    def test_図形の種別にAutoShapeが含まれる(self, ooxml_test_xlsx: Path) -> None:
        shapes_by_sheet = get_shapes_ooxml(ooxml_test_xlsx)
        all_types = [s.type for shapes in shapes_by_sheet.values() for s in shapes if s.type]
        autoshape_types = [t for t in all_types if "AutoShape" in t]
        assert len(autoshape_types) > 0

    def test_図形にIDが割り当てられる(self, ooxml_test_xlsx: Path) -> None:
        shapes_by_sheet = get_shapes_ooxml(ooxml_test_xlsx)
        shapes_with_id = [s for shapes in shapes_by_sheet.values() for s in shapes if s.id is not None]
        assert len(shapes_with_id) > 0
        for shape in shapes_with_id:
            assert isinstance(shape.id, int)
            assert shape.id > 0
        # IDs unique per sheet
        for shapes in shapes_by_sheet.values():
            ids = [s.id for s in shapes if s.id is not None]
            assert len(ids) == len(set(ids))

    def test_lightモードでは図形が抽出されない(self, ooxml_test_xlsx: Path) -> None:
        result = get_shapes_ooxml(ooxml_test_xlsx, mode="light")
        assert result == {}


# ---------------------------------------------------------------------------
# Charts: COM-equivalent tests
# ---------------------------------------------------------------------------

class TestChartsEquivalentToCom:
    """Tests equivalent to tests/com/test_charts_extraction.py."""

    def test_チャートの基本メタ情報が抽出される(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        all_charts = [c for charts in charts_by_sheet.values() for c in charts]
        assert len(all_charts) > 0
        for chart in all_charts:
            assert chart.chart_type is not None
            assert chart.chart_type != "unknown"
            assert chart.error is None

    def test_チャートタイトルが抽出される(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        all_charts = [c for charts in charts_by_sheet.values() for c in charts]
        titles = [c.title for c in all_charts if c.title]
        assert "売上データ" in titles

    def test_Y軸情報が抽出される(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        all_charts = [c for charts in charts_by_sheet.values() for c in charts]
        assert len(all_charts) > 0
        for chart in all_charts:
            assert isinstance(chart.y_axis_range, list)

    def test_系列情報が参照式として抽出される(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        all_charts = [c for charts in charts_by_sheet.values() for c in charts]
        chart_with_series = next((c for c in all_charts if c.series), None)
        assert chart_with_series is not None
        assert len(chart_with_series.series) > 0
        for series in chart_with_series.series:
            assert hasattr(series, "name")
            assert hasattr(series, "y_range")
            if series.y_range:
                assert "!" in series.y_range or "$" in series.y_range

    def test_verboseでチャートのサイズが取得される(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx, mode="verbose")
        all_charts = [c for charts in charts_by_sheet.values() for c in charts]
        assert len(all_charts) > 0
        for chart in all_charts:
            assert chart.w is not None
            assert chart.h is not None


# ---------------------------------------------------------------------------
# Integration tests
# ---------------------------------------------------------------------------

class TestOoxmlIntegration:
    """Integration tests using generated test file."""

    def test_shapes_extraction(self, ooxml_test_xlsx: Path) -> None:
        shapes_by_sheet = get_shapes_ooxml(ooxml_test_xlsx)
        total = sum(len(shapes) for shapes in shapes_by_sheet.values())
        assert total >= 5  # We injected 5 shapes

    def test_charts_extraction(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        total = sum(len(charts) for charts in charts_by_sheet.values())
        assert total >= 1  # We created 1 chart

    def test_has_flowchart_shapes(self, ooxml_test_xlsx: Path) -> None:
        shapes_by_sheet = get_shapes_ooxml(ooxml_test_xlsx)
        all_types = [s.type for shapes in shapes_by_sheet.values() for s in shapes if s.type]
        flowchart_types = [t for t in all_types if "Flowchart" in t]
        assert len(flowchart_types) > 0

    def test_has_line_chart(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        all_types = [c.chart_type for charts in charts_by_sheet.values() for c in charts if c.chart_type]
        line_charts = [t for t in all_types if "Line" in t]
        assert len(line_charts) > 0

    def test_shape_text_extraction(self, ooxml_test_xlsx: Path) -> None:
        shapes_by_sheet = get_shapes_ooxml(ooxml_test_xlsx)
        all_texts = [s.text for shapes in shapes_by_sheet.values() for s in shapes if s.text]
        assert "開始" in all_texts
        assert "入力データ読み込み" in all_texts
        assert "フォーマット有効？" in all_texts

    def test_chart_title_extraction(self, ooxml_test_xlsx: Path) -> None:
        charts_by_sheet = get_charts_ooxml(ooxml_test_xlsx)
        all_titles = [c.title for charts in charts_by_sheet.values() for c in charts if c.title]
        assert "売上データ" in all_titles
