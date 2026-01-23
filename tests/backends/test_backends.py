from pathlib import Path

from _pytest.monkeypatch import MonkeyPatch
from openpyxl import Workbook

from exstruct.core.backends.com_backend import ComBackend, _parse_print_area_range
from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend
from exstruct.core.ranges import parse_range_zero_based


def test_openpyxl_backend_extract_cells_switches_link_mode(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    calls: list[str] = []

    def fake_cells(file_path: Path) -> dict[str, list[object]]:
        calls.append("cells")
        return {}

    def fake_cells_links(file_path: Path) -> dict[str, list[object]]:
        calls.append("links")
        return {}

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.extract_sheet_cells",
        fake_cells,
    )
    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.extract_sheet_cells_with_links",
        fake_cells_links,
    )

    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    backend.extract_cells(include_links=False)
    backend.extract_cells(include_links=True)

    assert calls == ["cells", "links"]


def test_openpyxl_backend_detect_tables_handles_failure(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    def fake_detect(file_path: Path, sheet_name: str) -> list[str]:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.detect_tables_openpyxl",
        fake_detect,
    )

    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    assert backend.detect_tables("Sheet1") == []


def test_openpyxl_backend_extract_colors_map_returns_none_on_failure(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    def fake_colors_map(
        file_path: Path,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.extract_sheet_colors_map",
        fake_colors_map,
    )

    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    assert (
        backend.extract_colors_map(include_default_background=False, ignore_colors=None)
        is None
    )


def test_openpyxl_backend_extract_formulas_map_returns_none_on_failure(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    def fake_formulas_map(file_path: Path) -> object:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.extract_sheet_formulas_map",
        fake_formulas_map,
    )

    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    assert backend.extract_formulas_map() is None


def test_com_backend_extract_colors_map_returns_none_on_failure(
    monkeypatch: MonkeyPatch,
) -> None:
    def fake_colors_map(
        workbook: object,
        *,
        include_default_background: bool,
        ignore_colors: set[str] | None,
    ) -> object:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.com_backend.extract_sheet_colors_map_com",
        fake_colors_map,
    )

    class DummyWorkbook:
        pass

    backend = ComBackend(DummyWorkbook())
    assert (
        backend.extract_colors_map(include_default_background=False, ignore_colors=None)
        is None
    )


def test_com_backend_extract_formulas_map_returns_none_on_failure(
    monkeypatch: MonkeyPatch,
) -> None:
    def fake_formulas_map(workbook: object) -> object:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.com_backend.extract_sheet_formulas_map_com",
        fake_formulas_map,
    )

    class DummyWorkbook:
        pass

    backend = ComBackend(DummyWorkbook())
    assert backend.extract_formulas_map() is None


def test_com_backend_extract_print_areas_handles_sheet_error(
    monkeypatch: MonkeyPatch,
) -> None:
    class _FailingPageSetup:
        @property
        def PrintArea(self) -> str:
            raise RuntimeError("boom")

    class _FailingSheetApi:
        PageSetup = _FailingPageSetup()

    class _FailingSheet:
        name = "Sheet1"
        api = _FailingSheetApi()

    class _DummyWorkbook:
        sheets = [_FailingSheet()]

    backend = ComBackend(_DummyWorkbook())
    assert backend.extract_print_areas() == {}


def test_openpyxl_backend_extract_print_areas(tmp_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["A", "B"])
    ws.append([1, 2])
    ws.print_area = "A1:B2"
    file_path = tmp_path / "print_area.xlsx"
    wb.save(file_path)
    wb.close()

    backend = OpenpyxlBackend(file_path)
    areas = backend.extract_print_areas()
    assert "Sheet1" in areas
    assert areas["Sheet1"]
    assert areas["Sheet1"][0].r1 == 1
    assert areas["Sheet1"][0].c1 == 0


def test_openpyxl_backend_extract_print_areas_returns_empty_on_error(
    monkeypatch: MonkeyPatch, tmp_path: Path
) -> None:
    def _raise(*_args: object, **_kwargs: object) -> None:
        raise RuntimeError("boom")

    monkeypatch.setattr(
        "exstruct.core.backends.openpyxl_backend.openpyxl_workbook", _raise
    )
    backend = OpenpyxlBackend(tmp_path / "book.xlsx")
    assert backend.extract_print_areas() == {}


def test_parse_range_zero_based_parses_sheet_prefix() -> None:
    bounds = parse_range_zero_based("Sheet1!A1:B2")
    assert bounds is not None
    assert bounds.r1 == 0
    assert bounds.c1 == 0
    assert bounds.r2 == 1
    assert bounds.c2 == 1


def test_com_backend_extract_print_areas_success() -> None:
    class _PageSetup:
        PrintArea = "A1:B2,INVALID"

    class _SheetApi:
        PageSetup = _PageSetup()

    class _Sheet:
        name = "Sheet1"
        api = _SheetApi()

    class _DummyWorkbook:
        sheets = [_Sheet()]

    backend = ComBackend(_DummyWorkbook())
    areas = backend.extract_print_areas()
    assert "Sheet1" in areas
    assert areas["Sheet1"][0].r1 == 1
    assert areas["Sheet1"][0].c1 == 0
    assert areas["Sheet1"][0].r2 == 2
    assert areas["Sheet1"][0].c2 == 1


def test_com_backend_parse_print_area_range_invalid() -> None:
    assert _parse_print_area_range("INVALID") is None


class _Location:
    def __init__(self, row: int | None = None, col: int | None = None) -> None:
        self.Row = row
        self.Column = col


class _BreakItem:
    def __init__(self, row: int | None = None, col: int | None = None) -> None:
        self.Location = _Location(row=row, col=col)


class _Breaks:
    def __init__(self, items: list[_BreakItem]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, index: int) -> _BreakItem:
        return self._items[index - 1]


class _RangeRows:
    def __init__(self, count: int) -> None:
        self.Count = count


class _RangeCols:
    def __init__(self, count: int) -> None:
        self.Count = count


class _Range:
    Row = 1
    Column = 1
    Rows = _RangeRows(2)
    Columns = _RangeCols(2)


class _UsedRange:
    Address = "A1:B2"


class _PageSetup:
    PrintArea = "A1:B2"


class _SheetApi:
    def __init__(self) -> None:
        self.DisplayPageBreaks = False
        self.PageSetup = _PageSetup()
        self.UsedRange = _UsedRange()
        self.HPageBreaks = _Breaks([_BreakItem(row=2)])
        self.VPageBreaks = _Breaks([_BreakItem(col=2)])

    def Range(self, _addr: str) -> _Range:
        return _Range()


class _Sheet:
    name = "Sheet1"

    def __init__(self) -> None:
        self.api = _SheetApi()


class _DummyWorkbook:
    sheets = [_Sheet()]


def test_com_backend_extract_auto_page_breaks_success() -> None:
    backend = ComBackend(_DummyWorkbook())
    areas = backend.extract_auto_page_breaks()
    assert "Sheet1" in areas
    assert areas["Sheet1"]


class _RestoreErrorSheetApi:
    def __init__(self) -> None:
        self._display = False
        self.PageSetup = _PageSetup()
        self.UsedRange = _UsedRange()
        self.HPageBreaks = _Breaks([])
        self.VPageBreaks = _Breaks([])

    @property
    def DisplayPageBreaks(self) -> bool:
        return self._display

    @DisplayPageBreaks.setter
    def DisplayPageBreaks(self, value: bool) -> None:
        if value is False:
            raise RuntimeError("restore failed")
        self._display = value

    def Range(self, _addr: str) -> _Range:
        return _Range()


class _RestoreErrorSheet:
    name = "Sheet1"

    def __init__(self) -> None:
        self.api = _RestoreErrorSheetApi()


class _RestoreErrorWorkbook:
    sheets = [_RestoreErrorSheet()]


def test_com_backend_extract_auto_page_breaks_restore_error() -> None:
    backend = ComBackend(_RestoreErrorWorkbook())
    areas = backend.extract_auto_page_breaks()
    assert "Sheet1" in areas
