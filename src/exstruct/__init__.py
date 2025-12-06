from __future__ import annotations

from pathlib import Path
from typing import Literal, Optional, TextIO

from .core.integrate import extract_workbook
from .core.cells import set_table_detection_params
from .io import save_as_json, save_as_toon, save_as_yaml, save_sheets, serialize_workbook
from .models import CellRow, Chart, ChartSeries, Shape, SheetData, WorkbookData
from .render import export_pdf, export_sheet_images

__all__ = [
    "extract",
    "export",
    "export_sheets",
    "export_pdf",
    "export_sheet_images",
    "process_excel",
    "ExtractionMode",
    "CellRow",
    "Shape",
    "ChartSeries",
    "Chart",
    "SheetData",
    "WorkbookData",
    "set_table_detection_params",
]


ExtractionMode = Literal["light", "standard", "verbose"]


def extract(file_path: str | Path, mode: ExtractionMode = "standard") -> WorkbookData:
    """Extract workbook semantic structure and return WorkbookData."""
    if mode not in ("light", "standard", "verbose"):
        raise ValueError(f"Unsupported mode: {mode}")
    return extract_workbook(Path(file_path), mode=mode)


def export(
    data: WorkbookData,
    path: str | Path,
    fmt: Optional[Literal["json", "yaml", "yml", "toon"]] = None,
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> None:
    """Export WorkbookData to supported file formats (json/yaml/toon)."""
    dest = Path(path)
    format_hint = (fmt or dest.suffix.lstrip(".") or "json").lower()
    match format_hint:
        case "json":
            save_as_json(data, dest, pretty=pretty, indent=indent)
        case "yaml" | "yml":
            save_as_yaml(data, dest)
        case "toon":
            save_as_toon(data, dest)
        case _:
            raise ValueError(f"Unsupported export format: {format_hint}")


def export_sheets(data: WorkbookData, dir_path: str | Path) -> dict[str, Path]:
    """
    Export each sheet as a JSON file (book_name + SheetData) into a directory.
    Returns a mapping of sheet name to written path.
    """
    return save_sheets(data, Path(dir_path), fmt="json")


def export_sheets_as(
    data: WorkbookData,
    dir_path: str | Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> dict[str, Path]:
    """
    Export each sheet in the given format (json/yaml/toon), including book_name and SheetData; returns sheet name → path map.
    """
    return save_sheets(data, Path(dir_path), fmt=fmt, pretty=pretty, indent=indent)


def process_excel(
    file_path: Path,
    output_path: Path | None = None,
    out_fmt: str = "json",
    image: bool = False,
    pdf: bool = False,
    dpi: int = 72,
    mode: ExtractionMode = "standard",
    pretty: bool = False,
    indent: int | None = None,
    sheets_dir: Path | None = None,
    stream: TextIO | None = None,
) -> None:
    """
    Convenience wrapper for CLI: export workbook and optionally PDF/PNG images (Excel required for rendering).
    - If output_path is None, writes the serialized workbook to stdout (or provided stream).
    - If sheets_dir is given, also writes per-sheet files into that directory.
    """
    if mode not in ("light", "standard", "verbose"):
        raise ValueError(f"Unsupported mode: {mode}")
    workbook_model = extract(file_path, mode=mode)
    text = serialize_workbook(workbook_model, fmt=out_fmt, pretty=pretty, indent=indent)
    target_stream = stream

    def _suffix_for(fmt: str) -> str:
        if fmt in ("yaml", "yml"):
            return ".yaml"
        if fmt == "toon":
            return ".toon"
        if fmt == "json":
            return ".json"
        raise ValueError(f"Unsupported export format: {fmt}")

    if output_path is not None:
        output_path.write_text(text, encoding="utf-8")
    else:
        if target_stream is None:
            import sys

            target_stream = sys.stdout
        target_stream.write(text)
        if not text.endswith("\n"):
            target_stream.write("\n")

    if sheets_dir is not None:
        save_sheets(
            workbook_model,
            sheets_dir,
            fmt=out_fmt,
            pretty=pretty,
            indent=indent,
        )

    if pdf or image:
        base_target = output_path or file_path.with_suffix(_suffix_for(out_fmt))
        pdf_path = base_target.with_suffix(".pdf")
        export_pdf(file_path, pdf_path)
        if image:
            images_dir = pdf_path.parent / f"{pdf_path.stem}_images"
            export_sheet_images(file_path, images_dir, dpi=dpi)
