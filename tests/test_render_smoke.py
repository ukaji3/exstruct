import os
from pathlib import Path

from openpyxl import Workbook
import pytest
import xlwings as xw

from exstruct import process_excel


def _excel_available() -> bool:
    try:
        app = xw.App(add_book=False, visible=False)
        app.quit()
        return True
    except Exception:
        return False


def _pypdfium_available() -> bool:
    try:
        import pypdfium2  # noqa: F401

        return True
    except Exception:
        return False


_RUN_RENDER = os.environ.get("RUN_RENDER_SMOKE") == "1"


@pytest.mark.skipif(
    (not _RUN_RENDER) or (not (_excel_available() and _pypdfium_available())),
    reason="Render smoke disabled or dependencies unavailable; set RUN_RENDER_SMOKE=1 to enable.",
)
def test_render_smoke_pdf_and_png(tmp_path: Path) -> None:
    # create a tiny workbook
    xlsx = tmp_path / "sample.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "hello"
    wb.save(xlsx)

    out_json = tmp_path / "out.json"
    process_excel(
        xlsx,
        output_path=out_json,
        out_fmt="json",
        pdf=True,
        image=True,
        dpi=72,
        mode="standard",
        pretty=True,
    )
    pdf_path = out_json.with_suffix(".pdf")
    images_dir = out_json.parent / f"{out_json.stem}_images"
    assert out_json.exists()
    assert pdf_path.exists()
    assert images_dir.exists()
    assert any(images_dir.glob("*.png"))
