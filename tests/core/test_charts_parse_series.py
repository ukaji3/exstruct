from tests.utils import parametrize

from exstruct.core.charts import parse_series_formula


@parametrize(
    "formula,expected",
    [
        (
            '=SERIES("Sales",Sheet1!$A$1:$A$12,Sheet1!$B$1:$B$12,1)',
            {
                "name_range": None,
                "x_range": "Sheet1!$A$1:$A$12",
                "y_range": "Sheet1!$B$1:$B$12",
                "plot_order": "1",
                "bubble_size_range": None,
                "name_literal": "Sales",
            },
        ),
        (
            "=SERIES(Sheet1!$C$1,Sheet1!$A$1:$A$3,Sheet1!$B$1:$B$3,2)",
            {
                "name_range": "Sheet1!$C$1",
                "x_range": "Sheet1!$A$1:$A$3",
                "y_range": "Sheet1!$B$1:$B$3",
                "plot_order": "2",
                "bubble_size_range": None,
                "name_literal": None,
            },
        ),
        (
            '=SERIES("売上";Sheet1!$A$1:$A$3;Sheet1!$B$1:$B$3;1)',
            {
                "name_range": None,
                "x_range": "Sheet1!$A$1:$A$3",
                "y_range": "Sheet1!$B$1:$B$3",
                "plot_order": "1",
                "bubble_size_range": None,
                "name_literal": "売上",
            },
        ),
        (
            '=SERIES("A""B",Sheet1!$A$1,Sheet1!$B$1,1)',
            {
                "name_range": None,
                "x_range": "Sheet1!$A$1",
                "y_range": "Sheet1!$B$1",
                "plot_order": "1",
                "bubble_size_range": None,
                "name_literal": 'A"B',
            },
        ),
        (
            "=SERIES(,,Sheet1!$B$1:$B$3,1)",
            {
                "name_range": None,
                "x_range": None,
                "y_range": "Sheet1!$B$1:$B$3",
                "plot_order": "1",
                "bubble_size_range": None,
                "name_literal": None,
            },
        ),
    ],
)
def test_parse_series_formula_valid(
    formula: str, expected: dict[str, str | None]
) -> None:
    """SERIES の正常系パターンを確認する。"""
    assert parse_series_formula(formula) == expected


@parametrize(
    "formula",
    [
        "",
        "=OTHER(A,B,C)",
        "=SERIES(",
    ],
)
def test_parse_series_formula_invalid(formula: str) -> None:
    """SERIES 以外や不正構文は None を返す。"""
    assert parse_series_formula(formula) is None
