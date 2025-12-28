from collections.abc import Callable
from typing import TypeVar, cast

import pytest
from typing_extensions import ParamSpec

from exstruct.core.cells import _normalize_color_key, _normalize_rgb

P = ParamSpec("P")
R = TypeVar("R")


def _parametrize(
    *args: object, **kwargs: object
) -> Callable[[Callable[P, R]], Callable[P, R]]:
    return cast(
        Callable[[Callable[P, R]], Callable[P, R]],
        pytest.mark.parametrize(*args, **kwargs),
    )


@_parametrize(
    "raw,expected",
    [
        (" #aabbcc ", "AABBCC"),
        ("ffaabbcc", "AABBCC"),
        ("#123", "123"),
        ("THEME:1", "theme:1"),
        ("indexed:5", "indexed:5"),
        ("auto", "auto"),
        ("AUTO:1", "auto:1"),
        ("", ""),
        ("   ", ""),
    ],
)
def test_normalize_color_key(raw: str, expected: str) -> None:
    """色キーの正規化パターンを確認する。"""
    assert _normalize_color_key(raw) == expected


@_parametrize(
    "raw,expected",
    [
        ("0xFFAABBCC", "AABBCC"),
        ("00AABBCC", "AABBCC"),
        ("AABBCC", "AABBCC"),
        ("ffaabbcc", "AABBCC"),
        ("0xAABBCC", "AABBCC"),
    ],
)
def test_normalize_rgb(raw: str, expected: str) -> None:
    """RGB/ARGB 文字列の正規化を確認する。"""
    assert _normalize_rgb(raw) == expected
