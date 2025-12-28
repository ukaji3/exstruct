from dataclasses import dataclass

import pytest
from tests.utils import parametrize

from exstruct.core import cells
from exstruct.core.cells import (
    _collect_table_candidates_from_values,
    _color_to_key,
    _count_nonempty_cells,
    _header_like_row,
    _resolve_cell_background,
    _resolve_fill_color_key,
    _table_signal_score,
)


@dataclass(frozen=True)
class _DummyColor:
    rgb: str | None = None
    type: str | None = None
    theme: int | None = None
    tint: float | None = None
    indexed: int | None = None
    auto: bool | None = None


@dataclass(frozen=True)
class _DummyFill:
    pattern_type: str | None
    fg_color: _DummyColor | None
    bg_color: _DummyColor | None

    @property
    def patternType(self) -> str | None:  # noqa: N802
        return self.pattern_type

    @property
    def fgColor(self) -> _DummyColor | None:  # noqa: N802
        return self.fg_color

    @property
    def bgColor(self) -> _DummyColor | None:  # noqa: N802
        return self.bg_color


@dataclass(frozen=True)
class _DummyCell:
    fill: _DummyFill | None


@parametrize(
    "cell,include_default,expected",
    [
        (_DummyCell(fill=None), False, None),
        (_DummyCell(fill=None), True, "FFFFFF"),
        (
            _DummyCell(
                fill=_DummyFill(pattern_type=None, fg_color=None, bg_color=None)
            ),
            False,
            None,
        ),
        (
            _DummyCell(
                fill=_DummyFill(pattern_type=None, fg_color=None, bg_color=None)
            ),
            True,
            "FFFFFF",
        ),
        (
            _DummyCell(
                fill=_DummyFill(
                    pattern_type="solid",
                    fg_color=_DummyColor(rgb="00AABBCC"),
                    bg_color=None,
                )
            ),
            False,
            "AABBCC",
        ),
    ],
)
def test_resolve_cell_background(
    cell: _DummyCell, include_default: bool, expected: str | None
) -> None:
    """背景色の既定値と fill 分岐を確認する。"""
    assert _resolve_cell_background(cell, include_default) == expected


def test_resolve_fill_color_key_prefers_fg_over_bg() -> None:
    """fgColor があれば bgColor より優先されることを確認する。"""
    fill = _DummyFill(
        pattern_type="solid",
        fg_color=_DummyColor(rgb="00FF0000"),
        bg_color=_DummyColor(rgb="0000FF00"),
    )
    assert _resolve_fill_color_key(fill) == "FF0000"


@parametrize(
    "color,expected",
    [
        (_DummyColor(rgb="00ABCDEF"), "ABCDEF"),
        (_DummyColor(type="theme", theme=1, tint=0.5), "theme:1:0.5"),
        (_DummyColor(type="theme", theme=None, tint=None), "theme:unknown"),
        (_DummyColor(type="indexed", indexed=64), "indexed:64"),
        (_DummyColor(type="auto", auto=True), "auto:True"),
        (_DummyColor(type="auto", auto=None), "auto"),
    ],
)
def test_color_to_key_variants(color: _DummyColor, expected: str) -> None:
    """Color 種別の分岐を確認する。"""
    assert _color_to_key(color) == expected


@parametrize(
    "row,expected",
    [
        (["A", "B"], True),
        (["1", "2"], False),
        (["A", "1"], True),
        ([None, ""], False),
        (["A"], False),
    ],
)
def test_header_like_row(row: list[object], expected: bool) -> None:
    """ヘッダ行判定の境界を確認する。"""
    assert _header_like_row(row) is expected


def test_table_signal_score_prefers_header_and_coverage() -> None:
    """ヘッダありのほうがスコアが高くなることを確認する。"""
    with_header = [["Name", "Value"], ["A", "1"], ["B", "2"]]
    without_header = [["1", "2"], ["3", "4"], ["5", "6"]]
    assert _table_signal_score(with_header) > _table_signal_score(without_header)


def test_count_nonempty_cells() -> None:
    """非空セル数のカウントを確認する。"""
    values = [["", None, "x"], ["y", " ", 0]]
    assert _count_nonempty_cells(values) == 3


def test_collect_table_candidates_empty_when_below_min(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """min_nonempty_cells 未満では空配列になることを確認する。"""
    original = cells._DETECTION_CONFIG["min_nonempty_cells"]
    cells._DETECTION_CONFIG["min_nonempty_cells"] = 5
    try:
        values = [["A", "B"], ["1", "2"]]
        results = _collect_table_candidates_from_values(
            values, base_top=1, base_left=1, col_name=lambda c: chr(64 + c)
        )
        assert results == []
    finally:
        cells._DETECTION_CONFIG["min_nonempty_cells"] = original


def test_collect_table_candidates_detects_simple_table(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """単純な表が検出されることを確認する。"""
    values = [["Name", "Value"], ["A", "1"], ["B", "2"]]
    results = _collect_table_candidates_from_values(
        values, base_top=1, base_left=1, col_name=lambda c: chr(64 + c)
    )
    assert results == ["A1:B3"]
