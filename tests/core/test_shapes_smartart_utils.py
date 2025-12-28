from __future__ import annotations

from dataclasses import dataclass
from typing import cast

import xlwings as xw

from exstruct.core import shapes as shapes_mod


@dataclass
class _DummyTextRange:
    Text: str | None  # noqa: N815


@dataclass
class _DummyTextFrame:
    HasText: bool  # noqa: N815
    TextRange: _DummyTextRange  # noqa: N815


@dataclass
class _DummyNode:
    Level: int  # noqa: N815
    TextFrame2: _DummyTextFrame  # noqa: N815


@dataclass
class _DummyLayout:
    Name: str | None  # noqa: N815


@dataclass
class _DummySmartArt:
    AllNodes: list[_DummyNode]  # noqa: N815
    Layout: object  # noqa: N815


@dataclass(frozen=True)
class _DummyApi:
    HasSmartArt: bool  # noqa: N815
    SmartArt: _DummySmartArt | None  # noqa: N815


@dataclass(frozen=True)
class _DummyApiRaises:
    @property
    def HasSmartArt(self) -> bool:  # noqa: N802
        """
        Indicates whether the shape contains SmartArt.

        This stub implementation is not available and raises an error when accessed.

        Returns:
            `True` if the shape contains SmartArt, `False` otherwise.

        Raises:
            RuntimeError: Always raised with the message "HasSmartArt unavailable".
        """
        raise RuntimeError("HasSmartArt unavailable")


@dataclass(frozen=True)
class _DummyShape:
    api_obj: object

    @property
    def api(self) -> object:
        """
        Access the underlying API object for this shape.

        Returns:
            The wrapped API object exposing the shape's underlying properties and methods.
        """
        return self.api_obj


@dataclass(frozen=True)
class _DummyShapeRaisesApi:
    @property
    def api(self) -> object:
        """
        Return the underlying API object for this wrapper.

        Returns:
            object: The underlying API object.

        Raises:
            RuntimeError: If the API is unavailable.
        """
        raise RuntimeError("api unavailable")


def test_shape_has_smartart_true_false() -> None:
    smartart = _DummySmartArt(AllNodes=[], Layout=_DummyLayout(Name="L"))
    has = shapes_mod._shape_has_smartart(
        cast(
            xw.Shape,
            _DummyShape(api_obj=_DummyApi(HasSmartArt=True, SmartArt=smartart)),
        )
    )
    assert has is True

    has_false = shapes_mod._shape_has_smartart(
        cast(xw.Shape, _DummyShape(api_obj=_DummyApi(HasSmartArt=False, SmartArt=None)))
    )
    assert has_false is False


def test_shape_has_smartart_handles_exceptions() -> None:
    has = shapes_mod._shape_has_smartart(
        cast(xw.Shape, _DummyShape(api_obj=_DummyApiRaises()))
    )
    assert has is False

    has_api_error = shapes_mod._shape_has_smartart(
        cast(xw.Shape, _DummyShapeRaisesApi())
    )
    assert has_api_error is False


def test_get_smartart_layout_name() -> None:
    assert shapes_mod._get_smartart_layout_name(None) == "Unknown"
    smartart = _DummySmartArt(AllNodes=[], Layout=_DummyLayout(Name="Layout"))
    assert (
        shapes_mod._get_smartart_layout_name(cast(shapes_mod._SmartArtLike, smartart))
        == "Layout"
    )
    smartart_no_name = _DummySmartArt(AllNodes=[], Layout=_DummyLayout(Name=None))
    assert (
        shapes_mod._get_smartart_layout_name(
            cast(shapes_mod._SmartArtLike, smartart_no_name)
        )
        == "Unknown"
    )


def test_collect_smartart_node_info_and_tree() -> None:
    nodes = [
        _DummyNode(
            Level=1,
            TextFrame2=_DummyTextFrame(
                HasText=True, TextRange=_DummyTextRange(Text="root")
            ),
        ),
        _DummyNode(
            Level=2,
            TextFrame2=_DummyTextFrame(
                HasText=True, TextRange=_DummyTextRange(Text="child")
            ),
        ),
        _DummyNode(
            Level=1,
            TextFrame2=_DummyTextFrame(
                HasText=False, TextRange=_DummyTextRange(Text=None)
            ),
        ),
    ]
    smartart = _DummySmartArt(AllNodes=nodes, Layout=_DummyLayout(Name="L"))
    info = shapes_mod._collect_smartart_node_info(
        cast(shapes_mod._SmartArtLike, smartart)
    )
    assert info == [(1, "root"), (2, "child"), (1, "")]

    roots = shapes_mod._extract_smartart_nodes(cast(shapes_mod._SmartArtLike, smartart))
    assert len(roots) == 2
    assert roots[0].text == "root"
    assert roots[0].kids[0].text == "child"
    assert roots[1].text == ""


def test_collect_smartart_node_info_none() -> None:
    assert shapes_mod._collect_smartart_node_info(None) == []
