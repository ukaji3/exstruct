from tests.utils import parametrize

from exstruct.core.cells import _coerce_numeric_preserve_format


@parametrize(
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
