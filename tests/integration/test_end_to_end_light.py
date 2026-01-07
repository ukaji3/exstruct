from io import StringIO
import json
from pathlib import Path
from typing import cast

from openpyxl import Workbook

from exstruct import (
    DestinationOptions,
    ExStructEngine,
    FilterOptions,
    FormatOptions,
    OutputOptions,
    StructOptions,
    export_sheets,
    extract,
    serialize_workbook,
)


def _make_two_sheet_book(path: Path) -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = "v1"
    ws2 = wb.create_sheet("Sheet2")
    ws2["B2"] = "v2"
    wb.save(path)
    wb.close()


def test_end_to_end_light_extract_and_export(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _make_two_sheet_book(path)

    wb_data = extract(path, mode="light")
    assert wb_data.book_name == "book.xlsx"
    assert set(wb_data.sheets.keys()) == {"Sheet1", "Sheet2"}

    json_text = serialize_workbook(wb_data, fmt="json", pretty=False, indent=None)
    payload = json.loads(json_text)
    assert isinstance(payload, dict)
    payload_dict = cast(dict[str, object], payload)
    assert payload_dict.get("book_name") == "book.xlsx"

    out_dir = tmp_path / "sheets"
    files = export_sheets(wb_data, out_dir)
    assert set(files.keys()) == {"Sheet1", "Sheet2"}
    assert all(path.exists() for path in files.values())


def test_engine_process_writes_json_to_stream(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    _make_two_sheet_book(path)

    buffer = StringIO()
    engine = ExStructEngine(
        options=StructOptions(mode="light"),
        output=OutputOptions(
            format=FormatOptions(fmt="json", pretty=False, indent=None),
            filters=FilterOptions(
                include_print_areas=False,
                include_shapes=False,
                include_charts=False,
                include_shape_size=False,
                include_chart_size=False,
            ),
            destinations=DestinationOptions(),
        ),
    )
    engine.process(path, output_path=None, out_fmt="json", stream=buffer, mode="light")

    payload = json.loads(buffer.getvalue())
    payload_dict = cast(dict[str, object], payload)
    assert payload_dict.get("book_name") == "book.xlsx"
