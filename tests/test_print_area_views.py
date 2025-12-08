import json
from pathlib import Path

from exstruct.io import save_print_area_views
from exstruct.models import CellRow, PrintArea, SheetData, WorkbookData


def _workbook_with_print_area() -> WorkbookData:
    sheet = SheetData(
        rows=[
            CellRow(r=0, c={"0": "A", "2": "skip"}, links={"2": "http://example.com"}),
            CellRow(r=1, c={"1": "B"}),
            CellRow(r=2, c={"1": "C"}),
        ],
        shapes=[],
        charts=[],
        table_candidates=["A1:B2", "C1:C1"],
        print_areas=[PrintArea(r1=0, c1=0, r2=1, c2=1)],
    )
    return WorkbookData(book_name="book.xlsx", sheets={"Sheet1": sheet})


def test_save_print_area_views_filters_rows_and_tables(tmp_path: Path) -> None:
    wb = _workbook_with_print_area()
    written = save_print_area_views(wb, tmp_path, fmt="json", pretty=True)
    assert len(written) == 1
    path = next(iter(written.values()))
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["sheet_name"] == "Sheet1"
    assert data["area"] == {"r1": 0, "c1": 0, "r2": 1, "c2": 1}
    # Only cells within the area remain; out-of-range cells and links are dropped.
    assert data["rows"] == [{"r": 0, "c": {"0": "A"}}, {"r": 1, "c": {"1": "B"}}]
    # Only table candidates fully contained in the print area remain.
    assert data["table_candidates"] == ["A1:B2"]


def test_save_print_area_views_normalizes_when_requested(tmp_path: Path) -> None:
    wb = _workbook_with_print_area()
    # Shift the print area to demonstrate normalization.
    wb.sheets["Sheet1"].print_areas[0] = PrintArea(r1=1, c1=1, r2=2, c2=2)
    written = save_print_area_views(
        wb, tmp_path, fmt="json", pretty=False, normalize=True
    )
    path = next(iter(written.values()))
    data = json.loads(path.read_text(encoding="utf-8"))
    # Rows inside the area are re-based to start at 0, columns start at 0.
    assert data["rows"] == [{"r": 0, "c": {"0": "B"}}, {"r": 1, "c": {"0": "C"}}]


def test_save_print_area_views_no_print_areas_returns_empty(tmp_path: Path) -> None:
    wb = _workbook_with_print_area()
    wb.sheets["Sheet1"].print_areas = []
    written = save_print_area_views(wb, tmp_path, fmt="json")
    assert written == {}
