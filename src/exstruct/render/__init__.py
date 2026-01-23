from __future__ import annotations

import logging
import multiprocessing as mp
import os
from pathlib import Path
import shutil
import tempfile
from types import ModuleType
from typing import Protocol, cast

import xlwings as xw

from ..errors import MissingDependencyError, RenderError

logger = logging.getLogger(__name__)


def _require_excel_app() -> xw.App:
    """Ensure Excel COM is available and return an App; otherwise raise."""
    try:
        app = xw.App(add_book=False, visible=False)
        return app
    except Exception as e:
        raise RenderError(
            "Excel (COM) is not available. Rendering (PDF/image) requires a desktop Excel installation."
        ) from e


def export_pdf(excel_path: str | Path, output_pdf: str | Path) -> list[str]:
    """Export an Excel workbook to PDF via Excel COM and return sheet names in order."""
    normalized_excel_path = Path(excel_path)
    normalized_output_pdf = Path(output_pdf)
    normalized_output_pdf.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as td:
        temp_dir = Path(td)
        temp_xlsx = temp_dir / "book.xlsx"
        temp_pdf = temp_dir / "book.pdf"

        app: xw.App | None = None
        wb: xw.Book | None = None
        try:
            app = _require_excel_app()
            app.display_alerts = False
            wb = app.books.open(str(normalized_excel_path))
            sheet_names = [s.name for s in wb.sheets]
            wb.api.SaveAs(str(temp_xlsx))
            wb.api.ExportAsFixedFormat(0, str(temp_pdf))
            shutil.copy(temp_pdf, normalized_output_pdf)
        except RenderError:
            raise
        except Exception as exc:
            raise RenderError(
                "Failed to export PDF for "
                f"'{normalized_excel_path}' to '{normalized_output_pdf}'."
            ) from exc
        finally:
            if wb is not None:
                wb.close()
            if app is not None:
                app.quit()
        if not normalized_output_pdf.exists():
            raise RenderError(f"Failed to export PDF to '{normalized_output_pdf}'.")
    return sheet_names


def _require_pdfium() -> ModuleType:
    """Ensure pypdfium2 is installed; otherwise raise with guidance."""
    try:
        import pypdfium2 as pdfium
    except ImportError as e:
        raise MissingDependencyError(
            "Image rendering requires pypdfium2. Install it via `pip install pypdfium2 pillow` or add the 'render' extra."
        ) from e
    return cast(ModuleType, pdfium)


def export_sheet_images(
    excel_path: str | Path, output_dir: str | Path, dpi: int = 144
) -> list[Path]:
    """Export each sheet as PNG (via PDF then pypdfium2 rasterization) and return paths in sheet order."""
    normalized_excel_path = Path(excel_path)
    normalized_output_dir = Path(output_dir)
    normalized_output_dir.mkdir(parents=True, exist_ok=True)
    use_subprocess = _use_render_subprocess()
    pdfium = _ensure_pdfium(use_subprocess)

    try:
        with tempfile.TemporaryDirectory() as td:
            temp_dir = Path(td)
            return _export_sheet_images_with_app(
                normalized_excel_path,
                normalized_output_dir,
                temp_dir,
                dpi,
                use_subprocess,
                pdfium,
            )
    except RenderError:
        raise
    except Exception as exc:
        raise RenderError(
            f"Failed to export sheet images to '{normalized_output_dir}'."
        ) from exc


def _sanitize_sheet_filename(name: str) -> str:
    return "".join("_" if c in '\\/:*?"<>|' else c for c in name).strip() or "sheet"


class _PageSetupProtocol(Protocol):
    """Protocol for Excel PageSetup objects exposing PrintArea."""

    PrintArea: object


class _SheetApiProtocol(Protocol):
    """Protocol for Excel sheet COM APIs used by render helpers."""

    PageSetup: _PageSetupProtocol

    def ExportAsFixedFormat(  # noqa: N802
        self, file_format: int, output_path: str, *args: object, **kwargs: object
    ) -> None: ...


def _iter_sheet_apis(wb: xw.Book) -> list[tuple[int, str, _SheetApiProtocol]]:
    """Return sheet index, name, and COM api handle in order."""
    try:
        ws_collection = getattr(getattr(wb, "api", None), "Worksheets", None)
        if ws_collection is None:
            raise AttributeError("Worksheets not available")
        count = int(ws_collection.Count)
        sheets: list[tuple[int, str, _SheetApiProtocol]] = []
        for i in range(1, count + 1):
            ws_api = cast(_SheetApiProtocol, ws_collection.Item(i))
            name = str(getattr(ws_api, "Name", f"Sheet{i}"))
            sheets.append((i - 1, name, ws_api))
        return sheets
    except Exception:
        return [
            (
                index,
                sheet.name,
                cast(_SheetApiProtocol, sheet.api),
            )
            for index, sheet in enumerate(wb.sheets)
        ]


def _build_sheet_export_plan(
    wb: xw.Book,
) -> list[tuple[str, _SheetApiProtocol, str | None]]:
    """Return export plan rows for sheets and their print areas.

    Each item is (sheet_name, sheet_api, print_area).
    """
    plan: list[tuple[str, _SheetApiProtocol, str | None]] = []
    for _, sheet_name, sheet_api in _iter_sheet_apis(wb):
        areas = _extract_print_areas(sheet_api)
        if not areas:
            plan.append((sheet_name, sheet_api, None))
            continue
        for area in areas:
            plan.append((sheet_name, sheet_api, area))
    return plan


def _extract_print_areas(sheet_api: _SheetApiProtocol) -> list[str]:
    """Return print areas for a sheet API, split into individual ranges."""
    try:
        page_setup = getattr(sheet_api, "PageSetup", None)
        if page_setup is None:
            return []
        raw = str(getattr(page_setup, "PrintArea", "") or "")
    except Exception:
        return []
    if not raw:
        return []
    return _split_csv_respecting_quotes(raw)


def _split_csv_respecting_quotes(raw: str) -> list[str]:
    """Split a CSV-like string while keeping commas inside single quotes intact."""
    parts: list[str] = []
    buf: list[str] = []
    in_quote = False
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch == "'":
            if in_quote and i + 1 < len(raw) and raw[i + 1] == "'":
                buf.append("''")
                i += 2
                continue
            in_quote = not in_quote
            buf.append(ch)
            i += 1
            continue
        if ch == "," and not in_quote:
            parts.append("".join(buf).strip())
            buf = []
            i += 1
            continue
        buf.append(ch)
        i += 1
    if buf:
        parts.append("".join(buf).strip())
    return [p for p in parts if p]


def _rename_pages_for_print_area(
    paths: list[Path],
    output_dir: Path,
    base_index: int,
    safe_name: str,
) -> list[Path]:
    """Rename multi-page outputs to unique prefixes for print areas."""
    renamed: list[Path] = []
    for path in paths:
        page_index = _page_index_from_suffix(path.stem)
        new_index = base_index + page_index
        new_path = output_dir / f"{new_index + 1:02d}_{safe_name}.png"
        if path != new_path:
            path.replace(new_path)
        renamed.append(new_path)
    return renamed


def _page_index_from_suffix(stem: str) -> int:
    """Extract zero-based page index from a _pNN suffix when present."""
    if "_p" not in stem:
        return 0
    base, suffix = stem.rsplit("_p", 1)
    _ = base
    if suffix.isdigit():
        page_number = int(suffix)
        if page_number <= 0:
            return 0
        return page_number - 1
    return 0


def _export_sheet_pdf(
    sheet_api: _SheetApiProtocol,
    pdf_path: Path,
    *,
    ignore_print_areas: bool,
    print_area: str | None = None,
) -> None:
    """Export a sheet to PDF via Excel COM.

    Args:
        sheet_api: Target worksheet COM api.
        pdf_path: Output PDF path.
        ignore_print_areas: Whether to ignore print areas.
        print_area: Optional print area string to apply for this export.
    """
    original_print_area: object | None = None
    page_setup = None
    if print_area is not None:
        try:
            page_setup = getattr(sheet_api, "PageSetup", None)
            if page_setup is not None:
                original_print_area = getattr(page_setup, "PrintArea", None)
                page_setup.PrintArea = print_area
        except Exception:
            page_setup = None
    try:
        sheet_api.ExportAsFixedFormat(
            0, str(pdf_path), IgnorePrintAreas=ignore_print_areas
        )
    except TypeError:
        sheet_api.ExportAsFixedFormat(0, str(pdf_path))
    finally:
        if page_setup is not None and print_area is not None:
            try:
                page_setup.PrintArea = original_print_area
            except Exception as exc:
                logger.debug("Failed to restore PrintArea. (%r)", exc)


def _ensure_pdfium(use_subprocess: bool) -> ModuleType | None:
    """Return pdfium module when needed, or None for subprocess rendering."""
    if use_subprocess:
        _require_pdfium()
        return None
    return _require_pdfium()


def _export_sheet_images_with_app(
    excel_path: Path,
    output_dir: Path,
    temp_dir: Path,
    dpi: int,
    use_subprocess: bool,
    pdfium: ModuleType | None,
) -> list[Path]:
    """Export sheet images using Excel COM and PDF rendering."""
    written: list[Path] = []
    app: xw.App | None = None
    wb: xw.Book | None = None
    try:
        app = _require_excel_app()
        wb = app.books.open(str(excel_path))
        output_index = 0
        for sheet_name, sheet_api, print_area in _build_sheet_export_plan(wb):
            sheet_pdf = temp_dir / f"sheet_{output_index + 1:02d}.pdf"
            safe_name = _sanitize_sheet_filename(sheet_name)
            _export_sheet_pdf(
                sheet_api,
                sheet_pdf,
                ignore_print_areas=False,
                print_area=print_area,
            )
            sheet_paths = _render_sheet_images(
                pdfium,
                sheet_pdf,
                output_dir,
                output_index,
                safe_name,
                dpi,
                use_subprocess,
            )
            if not sheet_paths:
                _export_sheet_pdf(
                    sheet_api,
                    sheet_pdf,
                    ignore_print_areas=True,
                    print_area=print_area,
                )
                sheet_paths = _render_sheet_images(
                    pdfium,
                    sheet_pdf,
                    output_dir,
                    output_index,
                    safe_name,
                    dpi,
                    use_subprocess,
                )
            sheet_paths = _normalize_multipage_paths(
                sheet_paths,
                output_dir,
                output_index,
                safe_name,
            )
            written.extend(sheet_paths)
            output_index += max(1, len(sheet_paths))
        return written
    finally:
        if wb is not None:
            wb.close()
        if app is not None:
            app.quit()


def _render_sheet_images(
    pdfium: ModuleType | None,
    sheet_pdf: Path,
    output_dir: Path,
    output_index: int,
    safe_name: str,
    dpi: int,
    use_subprocess: bool,
) -> list[Path]:
    """Render sheet PDF to PNGs using the configured renderer."""
    if use_subprocess:
        return _render_pdf_pages_subprocess(
            sheet_pdf,
            output_dir,
            output_index,
            safe_name,
            dpi,
        )
    if pdfium is None:
        raise RenderError("pypdfium2 is required for in-process rendering.")
    return _render_pdf_pages_in_process(
        pdfium,
        sheet_pdf,
        output_dir,
        output_index,
        safe_name,
        dpi,
    )


def _normalize_multipage_paths(
    paths: list[Path],
    output_dir: Path,
    base_index: int,
    safe_name: str,
) -> list[Path]:
    """Normalize multi-page outputs to unique prefixes when needed."""
    if len(paths) <= 1:
        return paths
    return _rename_pages_for_print_area(paths, output_dir, base_index, safe_name)


def _use_render_subprocess() -> bool:
    """Return True when PDF->PNG rendering should run in a subprocess."""
    return os.getenv("EXSTRUCT_RENDER_SUBPROCESS", "1").lower() not in {"0", "false"}


def _render_pdf_pages_in_process(
    pdfium: ModuleType,
    pdf_path: Path,
    output_dir: Path,
    sheet_index: int,
    safe_name: str,
    dpi: int,
) -> list[Path]:
    """Render PDF pages to PNGs in the current process."""
    scale = dpi / 72.0
    written: list[Path] = []
    with pdfium.PdfDocument(str(pdf_path)) as pdf:
        for page_index in range(len(pdf)):
            page = pdf[page_index]
            bitmap = page.render(scale=scale)
            pil_image = bitmap.to_pil()
            page_suffix = f"_p{page_index + 1:02d}" if page_index > 0 else ""
            img_path = (
                output_dir / f"{sheet_index + 1:02d}_{safe_name}{page_suffix}.png"
            )
            pil_image.save(img_path, format="PNG", dpi=(dpi, dpi))
            written.append(img_path)
    return written


def _render_pdf_pages_subprocess(
    pdf_path: Path,
    output_dir: Path,
    sheet_index: int,
    safe_name: str,
    dpi: int,
) -> list[Path]:
    """Render PDF pages to PNGs in a subprocess for memory isolation."""
    ctx = mp.get_context("spawn")
    queue: mp.Queue[dict[str, list[str] | str]] = ctx.Queue()
    process = ctx.Process(
        target=_render_pdf_pages_worker,
        args=(pdf_path, output_dir, sheet_index, safe_name, dpi, queue),
    )
    process.start()
    process.join()
    result = _get_subprocess_result(queue)
    if process.exitcode != 0 or "error" in result:
        message = result.get("error", "subprocess failed")
        raise RenderError(f"Failed to render PDF pages: {message}")
    paths = result.get("paths", [])
    return [Path(path) for path in paths]


def _get_subprocess_result(
    queue: mp.Queue[dict[str, list[str] | str]],
) -> dict[str, list[str] | str]:
    """Fetch the worker result from the queue with a timeout."""
    try:
        return queue.get(timeout=5)
    except Exception as exc:
        return {"error": f"subprocess did not return results ({exc})"}


def _render_pdf_pages_worker(
    pdf_path: Path,
    output_dir: Path,
    sheet_index: int,
    safe_name: str,
    dpi: int,
    queue: mp.Queue[dict[str, list[str] | str]],
) -> None:
    """Worker process to render PDF pages into PNG files."""
    try:
        import pypdfium2 as pdfium

        scale = dpi / 72.0
        output_dir.mkdir(parents=True, exist_ok=True)
        written: list[str] = []
        with pdfium.PdfDocument(str(pdf_path)) as pdf:
            for page_index in range(len(pdf)):
                page = pdf[page_index]
                bitmap = page.render(scale=scale)
                pil_image = bitmap.to_pil()
                page_suffix = f"_p{page_index + 1:02d}" if page_index > 0 else ""
                img_path = (
                    output_dir / f"{sheet_index + 1:02d}_{safe_name}{page_suffix}.png"
                )
                pil_image.save(img_path, format="PNG", dpi=(dpi, dpi))
                written.append(str(img_path))
        queue.put({"paths": written})
    except Exception as exc:
        queue.put({"error": str(exc)})


__all__ = ["export_pdf", "export_sheet_images"]
