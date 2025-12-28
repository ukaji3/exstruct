from collections.abc import Callable
from typing import TypeVar, cast

import pytest
from typing_extensions import ParamSpec

from exstruct.core.cells import _coerce_numeric_preserve_format

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
    "val,expected",
    [
        ("42", 42),
        ("+7", 7),
        ("-3", -3),
        ("0042", 42),
        ("3.14", 3.14),
        ("3.1400", 3.14),
        ("-0.50", -0.5),
        ("0.0", 0.0),
        ("1.", "1."),
        (".5", 0.5),
        ("1e3", "1e3"),
        ("text", "text"),
    ],
)
def test_coerce_numeric_preserve_format(val: str, expected: int | float | str) -> None:
    """数値文字列の変換/非変換を確認する。"""
    result = _coerce_numeric_preserve_format(val)
    assert result == expected
