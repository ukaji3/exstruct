from __future__ import annotations

from pathlib import Path

import pytest

from exstruct.mcp import extract_runner
from exstruct.mcp.io import PathPolicy


def test_resolve_input_path_missing(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        extract_runner._resolve_input_path(tmp_path / "missing.xlsx", policy=None)


def test_format_suffix() -> None:
    assert extract_runner._format_suffix("yml") == ".yml"
    assert extract_runner._format_suffix("json") == ".json"


def test_normalize_output_name(tmp_path: Path) -> None:
    input_path = tmp_path / "book.xlsx"
    assert (
        extract_runner._normalize_output_name(input_path, None, ".json") == "book.json"
    )
    assert (
        extract_runner._normalize_output_name(input_path, "out", ".json") == "out.json"
    )
    assert (
        extract_runner._normalize_output_name(input_path, "out.yaml", ".json")
        == "out.yaml"
    )


def test_resolve_output_path_denies_outside_root(tmp_path: Path) -> None:
    policy = PathPolicy(root=tmp_path)
    input_path = tmp_path / "book.xlsx"
    input_path.write_text("x", encoding="utf-8")
    outside = tmp_path.parent
    with pytest.raises(ValueError):
        extract_runner._resolve_output_path(
            input_path,
            "json",
            out_dir=outside,
            out_name=None,
            policy=policy,
        )


def test_try_read_workbook_meta(tmp_path: Path) -> None:
    from openpyxl import Workbook

    path = tmp_path / "book.xlsx"
    wb = Workbook()
    wb.active.title = "Sheet1"
    wb.save(path)
    meta, warnings = extract_runner._try_read_workbook_meta(path)
    assert meta is not None
    assert meta.sheet_count == 1
    assert meta.sheet_names == ["Sheet1"]
    assert warnings == []
