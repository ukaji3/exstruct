from collections.abc import Callable
from typing import TypeVar, cast

import pytest
from typing_extensions import ParamSpec

P = ParamSpec("P")
R = TypeVar("R")


def parametrize(
    *args: object, **kwargs: object
) -> Callable[[Callable[P, R]], Callable[P, R]]:
    """Type-safe wrapper around pytest.mark.parametrize."""
    return cast(
        Callable[[Callable[P, R]], Callable[P, R]],
        pytest.mark.parametrize(*args, **kwargs),
    )
