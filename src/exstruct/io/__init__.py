from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, Literal

from ..models import WorkbookData


def dict_without_empty_values(obj: Any):
    """Recursively drop empty values from nested structures."""
    if isinstance(obj, dict):
        return {
            k: dict_without_empty_values(v)
            for k, v in obj.items()
            if v not in [None, "", [], {}]
        }
    if isinstance(obj, list):
        return [
            dict_without_empty_values(v) for v in obj if v not in [None, "", [], {}]
        ]
    if hasattr(obj, "model_dump"):
        return dict_without_empty_values(obj.model_dump(exclude_none=True))
    return obj


def save_as_json(model: WorkbookData, path: Path, *, pretty: bool = False, indent: int | None = None) -> None:
    text = serialize_workbook(model, fmt="json", pretty=pretty, indent=indent)
    path.write_text(text, encoding="utf-8")


def save_as_yaml(model: WorkbookData, path: Path) -> None:
    text = serialize_workbook(model, fmt="yaml")
    path.write_text(text, encoding="utf-8")


def save_as_toon(model: WorkbookData, path: Path) -> None:
    text = serialize_workbook(model, fmt="toon")
    path.write_text(text, encoding="utf-8")


def _sanitize_sheet_filename(name: str) -> str:
    """Make a sheet name safe for filesystem usage."""
    safe = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return safe or "sheet"


def serialize_workbook(
    model: WorkbookData,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> str:
    """
    Convert WorkbookData to string in the requested format without writing to disk.
    """
    format_hint = fmt.lower()
    if format_hint == "yml":
        format_hint = "yaml"
    filtered_dict = dict_without_empty_values(model)

    match format_hint:
        case "json":
            indent_val = 2 if pretty and indent is None else indent
            return json.dumps(filtered_dict, ensure_ascii=False, indent=indent_val)
        case "yaml":
            yaml = _require_yaml()
            return yaml.safe_dump(
                filtered_dict,
                allow_unicode=True,
                sort_keys=False,
                indent=2,
            )
        case "toon":
            toon = _require_toon()
            return toon.encode(filtered_dict)
        case _:
            raise ValueError(f"Unsupported export format: {fmt}")


def save_sheets_as_json(workbook: WorkbookData, output_dir: Path, *, pretty: bool = False, indent: int | None = None) -> Dict[str, Path]:
    """
    Save each sheet as an individual JSON file.
    Contents include book_name and the sheet's SheetData.
    Returns a map of sheet name -> written path.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    written: Dict[str, Path] = {}
    for sheet_name, sheet_data in workbook.sheets.items():
        payload = dict_without_empty_values(
            {
                "book_name": workbook.book_name,
                "sheet_name": sheet_name,
                "sheet": sheet_data,
            }
        )
        file_name = f"{_sanitize_sheet_filename(sheet_name)}.json"
        path = output_dir / file_name
        indent_val = 2 if pretty and indent is None else indent
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=indent_val), encoding="utf-8")
        written[sheet_name] = path
    return written


def save_sheets(
    workbook: WorkbookData,
    output_dir: Path,
    fmt: Literal["json", "yaml", "yml", "toon"] = "json",
    *,
    pretty: bool = False,
    indent: int | None = None,
) -> Dict[str, Path]:
    """
    Save each sheet as an individual file in the specified format (json/yaml/toon).
    Payload includes book_name and the sheet's SheetData.
    """
    format_hint = fmt.lower()
    if format_hint == "yml":
        format_hint = "yaml"
    if format_hint not in ("json", "yaml", "toon"):
        raise ValueError(f"Unsupported sheet export format: {fmt}")

    output_dir.mkdir(parents=True, exist_ok=True)
    written: Dict[str, Path] = {}
    for sheet_name, sheet_data in workbook.sheets.items():
        payload = dict_without_empty_values(
            {
                "book_name": workbook.book_name,
                "sheet_name": sheet_name,
                "sheet": sheet_data,
            }
        )
        suffix = {"json": ".json", "yaml": ".yaml", "toon": ".toon"}[format_hint]
        file_name = f"{_sanitize_sheet_filename(sheet_name)}{suffix}"
        path = output_dir / file_name
        match format_hint:
            case "json":
                indent_val = 2 if pretty and indent is None else indent
                text = json.dumps(payload, ensure_ascii=False, indent=indent_val)
            case "yaml":
                yaml = _require_yaml()
                text = yaml.safe_dump(
                    payload, allow_unicode=True, sort_keys=False, indent=2
                )
            case "toon":
                toon = _require_toon()
                text = toon.encode(payload)
            case _:
                raise ValueError(f"Unsupported sheet export format: {format_hint}")
        path.write_text(text, encoding="utf-8")
        written[sheet_name] = path
    return written


def _require_yaml():
    try:
        import yaml
    except ImportError as e:
        raise RuntimeError(
            "YAML export requires pyyaml. Install it via `pip install pyyaml` "
            "or add the 'yaml' extra."
        ) from e
    return yaml


def _require_toon():
    try:
        import toon
    except ImportError as e:
        raise RuntimeError(
            "TOON export requires python-toon. Install it via `pip install python-toon` "
            "or add the 'toon' extra."
        ) from e
    return toon


__all__ = [
    "dict_without_empty_values",
    "save_as_json",
    "save_as_yaml",
    "save_as_toon",
    "save_sheets",
    "save_sheets_as_json",
    "serialize_workbook",
]
