from dataclasses import dataclass

from exstruct.core.shapes import get_shapes_with_position
from exstruct.models import Arrow


@dataclass(frozen=True)
class _DummyLine:
    begin: int
    end: int

    @property
    def BeginArrowheadStyle(self) -> int:
        return self.begin

    @property
    def EndArrowheadStyle(self) -> int:
        return self.end


@dataclass(frozen=True)
class _DummyApi:
    shape_type: int
    autoshape_type: int | None
    line: _DummyLine | None
    rotation: float = 0.0

    @property
    def Type(self) -> int:
        return self.shape_type

    @property
    def AutoShapeType(self) -> int:
        if self.autoshape_type is None:
            raise RuntimeError("AutoShapeType unavailable")
        return self.autoshape_type

    @property
    def Line(self) -> _DummyLine:
        if self.line is None:
            raise RuntimeError("Line unavailable")
        return self.line

    @property
    def Rotation(self) -> float:
        return self.rotation


@dataclass(frozen=True)
class _DummyShape:
    name: str
    text: str
    left: float
    top: float
    width: float
    height: float
    api: _DummyApi


@dataclass(frozen=True)
class _DummySheet:
    name: str
    shapes: list[_DummyShape]


@dataclass(frozen=True)
class _DummyBook:
    sheets: list[_DummySheet]


def test_get_shapes_with_position_standard_filters_textless_non_relation() -> None:
    text_shape = _DummyShape(
        name="Rect1",
        text="Hello",
        left=10.0,
        top=20.0,
        width=100.0,
        height=50.0,
        api=_DummyApi(shape_type=1, autoshape_type=1, line=None),
    )
    line_shape = _DummyShape(
        name="Line1",
        text="",
        left=5.0,
        top=5.0,
        width=10.0,
        height=0.0,
        api=_DummyApi(shape_type=9, autoshape_type=None, line=_DummyLine(0, 3)),
    )
    empty_shape = _DummyShape(
        name="Rect2",
        text="",
        left=30.0,
        top=40.0,
        width=80.0,
        height=30.0,
        api=_DummyApi(shape_type=1, autoshape_type=1, line=None),
    )
    book = _DummyBook(
        sheets=[
            _DummySheet(name="Sheet1", shapes=[text_shape, line_shape, empty_shape])
        ]
    )

    result = get_shapes_with_position(book, mode="standard")
    shapes = result["Sheet1"]

    assert len(shapes) == 2
    assert {s.text for s in shapes} == {"Hello", ""}
    line_entries = [s for s in shapes if s.text == ""]
    assert isinstance(line_entries[0], Arrow)
    assert line_entries[0].direction == "E"
    text_entries = [s for s in shapes if s.text == "Hello"]
    assert text_entries[0].id == 1


def test_get_shapes_with_position_verbose_includes_all_and_sizes() -> None:
    text_shape = _DummyShape(
        name="Rect1",
        text="Hello",
        left=10.0,
        top=20.0,
        width=100.0,
        height=50.0,
        api=_DummyApi(shape_type=1, autoshape_type=1, line=None),
    )
    line_shape = _DummyShape(
        name="Line1",
        text="",
        left=5.0,
        top=5.0,
        width=10.0,
        height=0.0,
        api=_DummyApi(shape_type=9, autoshape_type=None, line=_DummyLine(0, 3)),
    )
    empty_shape = _DummyShape(
        name="Rect2",
        text="",
        left=30.0,
        top=40.0,
        width=80.0,
        height=30.0,
        api=_DummyApi(shape_type=1, autoshape_type=1, line=None),
    )
    book = _DummyBook(
        sheets=[
            _DummySheet(name="Sheet1", shapes=[text_shape, line_shape, empty_shape])
        ]
    )

    result = get_shapes_with_position(book, mode="verbose")
    shapes = result["Sheet1"]

    assert len(shapes) == 3
    assert all(s.w is not None and s.h is not None for s in shapes)
