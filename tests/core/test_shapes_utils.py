from tests.utils import parametrize

from exstruct.core.shapes import (
    _should_include_shape,
    coord_to_cell_by_edges,
    has_arrow,
)


def test_should_include_shape_lightは常に除外() -> None:
    """light モードでは常に False になることを確認する。"""
    assert not _should_include_shape(
        text="",
        shape_type_num=None,
        shape_type_str=None,
        autoshape_type_str=None,
        shape_name=None,
        output_mode="light",
    )


@parametrize(
    "text,shape_type_num,shape_type_str,autoshape_type_str,shape_name,expected",
    [
        ("", None, "Rectangle", None, None, False),
        ("text", None, "Rectangle", None, None, True),
        ("", 3, "Line", None, None, True),
        ("", None, None, "RightArrow", None, True),
        ("", None, None, None, "Connector 1", True),
    ],
)
def test_should_include_shape_standard(
    text: str,
    shape_type_num: int | None,
    shape_type_str: str | None,
    autoshape_type_str: str | None,
    shape_name: str | None,
    expected: bool,
) -> None:
    """standard モードの関係線/テキスト条件を確認する。"""
    result = _should_include_shape(
        text=text,
        shape_type_num=shape_type_num,
        shape_type_str=shape_type_str,
        autoshape_type_str=autoshape_type_str,
        shape_name=shape_name,
        output_mode="standard",
    )
    assert result is expected


def test_should_include_shape_verboseは常に含める() -> None:
    """verbose モードでは常に True になることを確認する。"""
    assert _should_include_shape(
        text="",
        shape_type_num=None,
        shape_type_str=None,
        autoshape_type_str=None,
        shape_name=None,
        output_mode="verbose",
    )


@parametrize(
    "style,expected",
    [
        (0, False),
        (1, True),
        ("2", True),
        (None, False),
    ],
)
def test_has_arrow(
    style: object,
    expected: bool,
) -> None:
    """矢印スタイルの判定を確認する。"""
    assert has_arrow(style) is expected


def test_coord_to_cell_by_edges範囲内() -> None:
    """座標からセルが推定できることを確認する。"""
    row_edges = [0.0, 10.0, 20.0]
    col_edges = [0.0, 5.0, 10.0]
    assert coord_to_cell_by_edges(row_edges, col_edges, x=7.0, y=15.0) == "B2"
    assert coord_to_cell_by_edges(row_edges, col_edges, x=0.0, y=0.0) == "A1"


def test_coord_to_cell_by_edges範囲外() -> None:
    """範囲外の場合は None になることを確認する。"""
    row_edges = [0.0, 10.0, 20.0]
    col_edges = [0.0, 5.0, 10.0]
    assert coord_to_cell_by_edges(row_edges, col_edges, x=10.0, y=5.0) is None
    assert coord_to_cell_by_edges(row_edges, col_edges, x=-1.0, y=5.0) is None
