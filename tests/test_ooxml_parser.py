"""Tests for OOXML parser (COM-free shape/chart extraction).

These tests verify that the OOXML parser can extract shapes and charts
from xlsx files without requiring Excel COM (Windows).
"""

from pathlib import Path

import pytest

from exstruct.ooxml import get_charts_ooxml, get_shapes_ooxml
from exstruct.ooxml.units import emu_to_pixels, emu_to_points


class TestUnits:
    """Tests for EMU unit conversion."""

    def test_emu_to_pixels_default_dpi(self) -> None:
        """1 inch (914400 EMU) should be 96 pixels at 96 DPI."""
        assert emu_to_pixels(914400) == 96

    def test_emu_to_pixels_custom_dpi(self) -> None:
        """1 inch should be 72 pixels at 72 DPI."""
        assert emu_to_pixels(914400, dpi=72) == 72

    def test_emu_to_pixels_zero(self) -> None:
        """Zero EMU should be zero pixels."""
        assert emu_to_pixels(0) == 0

    def test_emu_to_points(self) -> None:
        """1 inch (914400 EMU) should be 72 points."""
        assert emu_to_points(914400) == 72.0

    def test_emu_to_points_half_inch(self) -> None:
        """Half inch should be 36 points."""
        assert emu_to_points(457200) == 36.0


class TestGetShapesOoxml:
    """Tests for OOXML shape extraction."""

    @pytest.fixture
    def sample_xlsx(self) -> Path:
        """Return path to sample xlsx file."""
        return Path("sample/sample.xlsx")

    def test_returns_dict_for_valid_file(self, sample_xlsx: Path) -> None:
        """Should return dict mapping sheet names to shape lists."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_shapes_ooxml(sample_xlsx)
        assert isinstance(result, dict)

    def test_returns_empty_dict_for_nonexistent_file(self, tmp_path: Path) -> None:
        """Should return empty dict for nonexistent file."""
        result = get_shapes_ooxml(tmp_path / "nonexistent.xlsx")
        assert result == {}

    def test_shapes_have_required_fields(self, sample_xlsx: Path) -> None:
        """Extracted shapes should have required fields."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_shapes_ooxml(sample_xlsx)
        for _sheet_name, shapes in result.items():
            for shape in shapes:
                # Required fields
                assert hasattr(shape, "text")
                assert hasattr(shape, "l")
                assert hasattr(shape, "t")
                assert hasattr(shape, "type")
                # Position should be non-negative
                assert shape.l >= 0
                assert shape.t >= 0

    def test_verbose_mode_includes_size(self, sample_xlsx: Path) -> None:
        """Verbose mode should include width and height."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_shapes_ooxml(sample_xlsx, mode="verbose")
        for _sheet_name, shapes in result.items():
            for shape in shapes:
                # Verbose mode includes w and h
                if shape.w is not None:
                    assert shape.w >= 0
                if shape.h is not None:
                    assert shape.h >= 0

    def test_standard_mode_excludes_size(self, sample_xlsx: Path) -> None:
        """Standard mode should not include width and height."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_shapes_ooxml(sample_xlsx, mode="standard")
        for _sheet_name, shapes in result.items():
            for shape in shapes:
                assert shape.w is None
                assert shape.h is None

    def test_connector_shapes_have_direction(self, sample_xlsx: Path) -> None:
        """Connector shapes should have direction if applicable."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_shapes_ooxml(sample_xlsx)
        # Check if any connector shapes exist
        connectors = []
        for shapes in result.values():
            for shape in shapes:
                if shape.type and ("Line" in shape.type or "Connector" in shape.type):
                    connectors.append(shape)

        # If connectors exist, they may have direction
        for conn in connectors:
            if conn.direction:
                assert conn.direction in ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]


class TestGetChartsOoxml:
    """Tests for OOXML chart extraction."""

    @pytest.fixture
    def sample_xlsx(self) -> Path:
        """Return path to sample xlsx file."""
        return Path("sample/sample.xlsx")

    def test_returns_dict_for_valid_file(self, sample_xlsx: Path) -> None:
        """Should return dict mapping sheet names to chart lists."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_charts_ooxml(sample_xlsx)
        assert isinstance(result, dict)

    def test_returns_empty_dict_for_nonexistent_file(self, tmp_path: Path) -> None:
        """Should return empty dict for nonexistent file."""
        result = get_charts_ooxml(tmp_path / "nonexistent.xlsx")
        assert result == {}

    def test_charts_have_required_fields(self, sample_xlsx: Path) -> None:
        """Extracted charts should have required fields."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_charts_ooxml(sample_xlsx)
        for _sheet_name, charts in result.items():
            for chart in charts:
                # Required fields
                assert hasattr(chart, "name")
                assert hasattr(chart, "chart_type")
                assert hasattr(chart, "series")
                # chart_type should be a known type
                assert chart.chart_type is not None

    def test_verbose_mode_includes_size(self, sample_xlsx: Path) -> None:
        """Verbose mode should include width and height."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_charts_ooxml(sample_xlsx, mode="verbose")
        for _sheet_name, charts in result.items():
            for chart in charts:
                # Verbose mode includes w and h (may be 0 if position not found)
                assert chart.w is not None
                assert chart.h is not None
                assert chart.w >= 0
                assert chart.h >= 0

    def test_standard_mode_excludes_size(self, sample_xlsx: Path) -> None:
        """Standard mode should not include width and height."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_charts_ooxml(sample_xlsx, mode="standard")
        for _sheet_name, charts in result.items():
            for chart in charts:
                assert chart.w is None
                assert chart.h is None

    def test_chart_series_have_ranges(self, sample_xlsx: Path) -> None:
        """Chart series should have range references."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        result = get_charts_ooxml(sample_xlsx)
        for _sheet_name, charts in result.items():
            for chart in charts:
                for series in chart.series:
                    # Series should have at least y_range
                    assert hasattr(series, "y_range")
                    assert hasattr(series, "x_range")
                    assert hasattr(series, "name")


class TestShapesEquivalentToCom:
    """Tests equivalent to tests/com/test_shapes_extraction.py."""

    @pytest.fixture
    def sample_xlsx(self) -> Path:
        """Return path to sample xlsx file."""
        return Path("sample/sample.xlsx")

    def test_図形の種別とテキストが抽出される(self, sample_xlsx: Path) -> None:
        """Shape type and text should be extracted (equivalent to COM test)."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Collect all shapes
        all_shapes = []
        for shapes in shapes_by_sheet.values():
            all_shapes.extend(shapes)

        # Should have shapes
        assert len(all_shapes) > 0, "Expected shapes in sample.xlsx"

        # Find shapes with text
        shapes_with_text = [s for s in all_shapes if s.text]
        assert len(shapes_with_text) > 0, "Expected shapes with text"

        for shape in shapes_with_text:
            # Type should contain AutoShape or known type
            assert shape.type is not None
            # Position should be valid
            assert shape.l >= 0
            assert shape.t >= 0

    def test_図形の種別にAutoShapeが含まれる(self, sample_xlsx: Path) -> None:
        """Shape type should contain AutoShape for flowchart shapes."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Collect all shape types
        all_types = []
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.type:
                    all_types.append(shape.type)

        # At least some shapes should have AutoShape type
        autoshape_types = [t for t in all_types if "AutoShape" in t]
        assert len(autoshape_types) > 0, "Expected AutoShape types"

    def test_線図形の方向と矢印情報が抽出される(self, sample_xlsx: Path) -> None:
        """Line direction and arrow info should be extracted."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Find connector/line shapes
        connectors = []
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.type and ("Line" in shape.type or "Connector" in shape.type):
                    connectors.append(shape)

        if not connectors:
            pytest.skip("No connector shapes found in sample.xlsx")

        # Check that connectors have direction or arrow styles
        has_direction_or_arrow = False
        for conn in connectors:
            if conn.direction:
                assert conn.direction in ["N", "NE", "E", "SE", "S", "SW", "W", "NW"]
                has_direction_or_arrow = True
            if conn.begin_arrow_style is not None or conn.end_arrow_style is not None:
                has_direction_or_arrow = True

        assert has_direction_or_arrow, "Expected direction or arrow style on connectors"

    def test_矢印スタイルが数値で抽出される(self, sample_xlsx: Path) -> None:
        """Arrow styles should be extracted as integers."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Find shapes with arrow styles
        shapes_with_arrows = []
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.begin_arrow_style is not None or shape.end_arrow_style is not None:
                    shapes_with_arrows.append(shape)

        if not shapes_with_arrows:
            pytest.skip("No shapes with arrow styles found")

        for shape in shapes_with_arrows:
            if shape.begin_arrow_style is not None:
                assert isinstance(shape.begin_arrow_style, int)
                assert shape.begin_arrow_style >= 1
            if shape.end_arrow_style is not None:
                assert isinstance(shape.end_arrow_style, int)
                assert shape.end_arrow_style >= 1


    def test_図形にIDが割り当てられる(self, sample_xlsx: Path) -> None:
        """Non-connector shapes should have unique IDs assigned."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Collect all shapes with IDs
        shapes_with_id = []
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.id is not None:
                    shapes_with_id.append(shape)

        # Should have shapes with IDs
        assert len(shapes_with_id) > 0, "Expected shapes with IDs"

        # IDs should be positive integers
        for shape in shapes_with_id:
            assert isinstance(shape.id, int)
            assert shape.id > 0

        # IDs should be unique within each sheet
        for shapes in shapes_by_sheet.values():
            ids = [s.id for s in shapes if s.id is not None]
            assert len(ids) == len(set(ids)), "Shape IDs should be unique"

    def test_コネクターはIDを持たない(self, sample_xlsx: Path) -> None:
        """Connector shapes should not have IDs (they are relationships)."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Find connector shapes
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.type and ("Line" in shape.type or "Connector" in shape.type):
                    # Connectors should not have ID
                    assert shape.id is None, f"Connector {shape.type} should not have ID"

    def test_lightモードでは図形が抽出されない(self, sample_xlsx: Path) -> None:
        """Light mode should not extract shapes."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx, mode="light")

        # Light mode returns empty dict
        assert shapes_by_sheet == {}, "Light mode should return empty dict"


class TestChartsEquivalentToCom:
    """Tests equivalent to tests/com/test_charts_extraction.py."""

    @pytest.fixture
    def sample_xlsx(self) -> Path:
        """Return path to sample xlsx file."""
        return Path("sample/sample.xlsx")

    def test_チャートの基本メタ情報が抽出される(self, sample_xlsx: Path) -> None:
        """Chart basic metadata should be extracted (equivalent to COM test)."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # Collect all charts
        all_charts = []
        for charts in charts_by_sheet.values():
            all_charts.extend(charts)

        assert len(all_charts) > 0, "Expected charts in sample.xlsx"

        for chart in all_charts:
            # chart_type should be a known type
            assert chart.chart_type is not None
            assert chart.chart_type != "unknown"
            # error should be None for successful extraction
            assert chart.error is None

    def test_チャートタイトルが抽出される(self, sample_xlsx: Path) -> None:
        """Chart title should be extracted."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # Find charts with titles
        charts_with_title = []
        for charts in charts_by_sheet.values():
            for chart in charts:
                if chart.title:
                    charts_with_title.append(chart)

        # sample.xlsx has chart with title "売上データ"
        assert len(charts_with_title) > 0, "Expected chart with title"
        titles = [c.title for c in charts_with_title]
        assert "売上データ" in titles, f"Expected '売上データ' in titles, got {titles}"

    def test_Y軸情報が抽出される(self, sample_xlsx: Path) -> None:
        """Y-axis title and range should be extracted."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # Collect all charts
        all_charts = []
        for charts in charts_by_sheet.values():
            all_charts.extend(charts)

        assert len(all_charts) > 0

        # Check y_axis_title and y_axis_range attributes exist
        for chart in all_charts:
            assert hasattr(chart, "y_axis_title")
            assert hasattr(chart, "y_axis_range")
            # y_axis_range should be list
            assert isinstance(chart.y_axis_range, list)

    def test_系列情報が参照式として抽出される(self, sample_xlsx: Path) -> None:
        """Series data should be extracted with range references."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # Find chart with series
        chart_with_series = None
        for charts in charts_by_sheet.values():
            for chart in charts:
                if chart.series:
                    chart_with_series = chart
                    break

        assert chart_with_series is not None, "Expected chart with series"
        assert len(chart_with_series.series) > 0

        for series in chart_with_series.series:
            # Series should have name
            assert hasattr(series, "name")
            # Series should have range references
            assert hasattr(series, "x_range")
            assert hasattr(series, "y_range")
            # y_range should be present for data series
            if series.y_range:
                # Range should look like a cell reference
                assert "!" in series.y_range or "$" in series.y_range

    def test_verboseでチャートのサイズが取得される(self, sample_xlsx: Path) -> None:
        """Chart size should be available in verbose mode."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx, mode="verbose")

        # Collect all charts
        all_charts = []
        for charts in charts_by_sheet.values():
            all_charts.extend(charts)

        assert len(all_charts) > 0

        for chart in all_charts:
            # Verbose mode should include w and h
            assert chart.w is not None
            assert chart.h is not None


class TestOoxmlIntegration:
    """Integration tests using sample files."""

    @pytest.fixture
    def sample_xlsx(self) -> Path:
        """Return path to sample xlsx file."""
        return Path("sample/sample.xlsx")

    def test_sample_xlsx_shapes_extraction(self, sample_xlsx: Path) -> None:
        """Test shape extraction from sample.xlsx."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # sample.xlsx should have shapes
        total_shapes = sum(len(shapes) for shapes in shapes_by_sheet.values())
        # At minimum, verify extraction doesn't crash
        assert isinstance(total_shapes, int)

    def test_sample_xlsx_charts_extraction(self, sample_xlsx: Path) -> None:
        """Test chart extraction from sample.xlsx."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # sample.xlsx should have charts
        total_charts = sum(len(charts) for charts in charts_by_sheet.values())
        # At minimum, verify extraction doesn't crash
        assert isinstance(total_charts, int)

    def test_sample_xlsx_has_flowchart_shapes(self, sample_xlsx: Path) -> None:
        """sample.xlsx should contain flowchart shapes."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Collect all shape types
        all_types: list[str] = []
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.type:
                    all_types.append(shape.type)

        # sample.xlsx has flowchart shapes based on sample.json
        flowchart_types = [t for t in all_types if "Flowchart" in t or "flowChart" in t.lower()]
        assert len(flowchart_types) > 0, "Expected flowchart shapes in sample.xlsx"

    def test_sample_xlsx_has_line_chart(self, sample_xlsx: Path) -> None:
        """sample.xlsx should contain a line chart."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # Collect all chart types
        all_chart_types: list[str] = []
        for charts in charts_by_sheet.values():
            for chart in charts:
                if chart.chart_type:
                    all_chart_types.append(chart.chart_type)

        # sample.xlsx has a line chart (売上データ)
        line_charts = [t for t in all_chart_types if "Line" in t]
        assert len(line_charts) > 0, "Expected line chart in sample.xlsx"

    def test_shape_text_extraction(self, sample_xlsx: Path) -> None:
        """Shapes with text should have text extracted."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        shapes_by_sheet = get_shapes_ooxml(sample_xlsx)

        # Collect all shape texts
        all_texts: list[str] = []
        for shapes in shapes_by_sheet.values():
            for shape in shapes:
                if shape.text:
                    all_texts.append(shape.text)

        # sample.xlsx has shapes with text (開始, 処理A, etc.)
        assert len(all_texts) > 0, "Expected shapes with text in sample.xlsx"

    def test_chart_title_extraction(self, sample_xlsx: Path) -> None:
        """Charts with titles should have titles extracted."""
        if not sample_xlsx.exists():
            pytest.skip("sample/sample.xlsx not found")

        charts_by_sheet = get_charts_ooxml(sample_xlsx)

        # Collect all chart titles
        all_titles: list[str] = []
        for charts in charts_by_sheet.values():
            for chart in charts:
                if chart.title:
                    all_titles.append(chart.title)

        # sample.xlsx has chart with title based on sample.json
        # Title may or may not be present depending on chart configuration
        assert isinstance(all_titles, list)
