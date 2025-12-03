from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import xlwings as xw

from ..models import CellRow, Shape, SheetData, WorkbookData
from .cells import detect_tables, extract_sheet_cells
from .charts import get_charts
from .shapes import get_shapes_with_position


def integrate_sheet_content(
    cell_data: Dict[str, List[CellRow]],
    shape_data: Dict[str, List[Shape]],
    workbook: xw.Book,
) -> Dict[str, SheetData]:
    result: Dict[str, SheetData] = {}
    for sheet_name, rows in cell_data.items():
        sheet_shapes = shape_data.get(sheet_name, [])
        sheet = workbook.sheets[sheet_name]

        sheet_model = SheetData(
            rows=rows,
            shapes=sheet_shapes,
            charts=get_charts(sheet),
            tables=detect_tables(sheet),
        )

        result[sheet_name] = sheet_model
    return result


def extract_workbook(file_path: Path) -> WorkbookData:
    cell_data = extract_sheet_cells(file_path)
    wb = xw.Book(file_path)
    try:
        shape_data = get_shapes_with_position(wb)
        merged = integrate_sheet_content(cell_data, shape_data, wb)
    finally:
        wb.close()

    return WorkbookData(book_name=file_path.name, sheets=merged)
