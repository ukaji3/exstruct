from collections.abc import Callable, Iterable, Sequence
from typing import Literal, TypeVar, cast

import pytest
from typing_extensions import ParamSpec

P = ParamSpec("P")
R = TypeVar("R")


def parametrize(
    argnames: str | Sequence[str],
    argvalues: Iterable[object],
    *,
    indirect: bool | Sequence[str] = False,
    ids: Iterable[str | float | int | bool | None]
    | Callable[[object], object | None]
    | None = None,
    scope: Literal["session", "package", "module", "class", "function"] | None = None,
) -> Callable[[Callable[P, R]], Callable[P, R]]:
    """Type-safe wrapper around pytest.mark.parametrize.

    Args:
        argnames: Parameter names for the parametrized test.
        argvalues: Parameter values for each test case.
        indirect: Whether to treat parameters as fixtures.
        ids: Optional case IDs or an ID factory.
        scope: Optional fixture scope for parametrization.

    Returns:
        Decorator preserving the wrapped callable signature.
    """
    return cast(
        Callable[[Callable[P, R]], Callable[P, R]],
        pytest.mark.parametrize(
            argnames,
            argvalues,
            indirect=indirect,
            ids=ids,
            scope=scope,
        ),
    )
