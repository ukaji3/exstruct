from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from exstruct.core import workbook


def test_openpyxl_workbook_close_error_is_suppressed(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    class _DummyWorkbook:
        def close(self) -> None:
            raise RuntimeError("close failed")

    dummy = _DummyWorkbook()

    def _fake_load_workbook(*_args: object, **_kwargs: object) -> _DummyWorkbook:
        return dummy

    monkeypatch.setattr(workbook, "load_workbook", _fake_load_workbook)

    with workbook.openpyxl_workbook(
        tmp_path / "book.xlsx", data_only=True, read_only=False
    ) as wb:
        assert wb is dummy


def test_xlwings_workbook_uses_existing(monkeypatch: pytest.MonkeyPatch) -> None:
    class _DummyBook:
        pass

    dummy = _DummyBook()

    def _fake_find(_path: Path) -> _DummyBook:
        return dummy

    def _boom(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("xlwings App should not be created when existing")

    monkeypatch.setattr(workbook, "_find_open_workbook", _fake_find)
    monkeypatch.setattr(workbook.xw, "App", _boom)

    with workbook.xlwings_workbook(Path("book.xlsx")) as wb:
        assert wb is dummy


def test_find_open_workbook_handles_fullname_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _BadBook:
        @property
        def fullname(self) -> str:
            raise RuntimeError("boom")

    class _DummyApp:
        books = [_BadBook()]

    monkeypatch.setattr(workbook.xw, "apps", [_DummyApp()])
    assert workbook._find_open_workbook(Path("book.xlsx")) is None


def test_find_open_workbook_handles_resolve_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _DummyPath:
        def __init__(self, value: str) -> None:
            self._value = value

        def resolve(self) -> _DummyPath:
            if self._value == "bad":
                raise RuntimeError("resolve failed")
            return self

        def __eq__(self, other: object) -> bool:
            return isinstance(other, _DummyPath) and self._value == other._value

    class _DummyBook:
        fullname = "bad"

    class _DummyApp:
        books = [_DummyBook()]

    monkeypatch.setattr(workbook, "Path", _DummyPath)
    monkeypatch.setattr(workbook.xw, "apps", [_DummyApp()])

    file_path = _DummyPath("good")
    assert workbook._find_open_workbook(file_path) is None


def test_find_open_workbook_returns_none_on_iter_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class _BadApps:
        def __iter__(self) -> Iterator[object]:
            raise RuntimeError("apps failure")

    monkeypatch.setattr(workbook.xw, "apps", _BadApps())
    assert workbook._find_open_workbook(Path("book.xlsx")) is None
