from pathlib import Path

from openpyxl import Workbook

from exstruct import extract
from exstruct.core.backends.openpyxl_backend import OpenpyxlBackend


def _make_book_with_print_area(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "x"
    ws.print_area = "A1:B2"
    wb.save(path)
    wb.close()


def test_light_mode_includes_print_areas_without_com(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _make_book_with_print_area(path)

    wb_data = extract(path, mode="light")
    areas = wb_data.sheets["Sheet1"].print_areas
    assert len(areas) == 1
    area = areas[0]
    assert (area.r1, area.c1, area.r2, area.c2) == (1, 0, 2, 1)


def test_openpyxl_backend_multiple_print_areas(tmp_path: Path) -> None:
    path = tmp_path / "multi_print_area.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "x"
    ws["D3"] = "y"
    ws.print_area = "A1:B2,D3:E4"
    wb.save(path)
    wb.close()

    backend = OpenpyxlBackend(path)
    areas = backend.extract_print_areas()

    assert "Sheet1" in areas
    ranges = [(a.r1, a.c1, a.r2, a.c2) for a in areas["Sheet1"]]
    assert ranges == [(1, 0, 2, 1), (3, 3, 4, 4)]
