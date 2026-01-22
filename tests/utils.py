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
    """Type-safe wrapper around pytest.mark.parametrize."""
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
