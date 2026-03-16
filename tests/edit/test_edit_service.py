from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

from exstruct.cli.availability import ComAvailability
from exstruct.edit import runtime as edit_runtime
from exstruct.edit.models import (
    MakeRequest,
    OpenpyxlEngineResult,
    PatchOp,
    PatchRequest,
)
import exstruct.edit.service as edit_service


def _create_workbook(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet.title = "Sheet1"
    sheet["A1"] = "old"
    workbook.save(path)
    workbook.close()


def test_edit_service_patch_prefers_com_when_available(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)
    calls: dict[str, bool] = {}

    monkeypatch.setattr(
        edit_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    def _fake_apply_xlwings_engine(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        calls["com"] = True
        return []

    monkeypatch.setattr(
        edit_service, "apply_xlwings_engine", _fake_apply_xlwings_engine
    )
    result = edit_service.patch_workbook(
        PatchRequest(
            xlsx_path=input_path,
            ops=[PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")],
            on_conflict="rename",
            backend="auto",
        )
    )
    assert result.error is None
    assert result.engine == "com"
    assert calls["com"] is True


def test_edit_service_patch_auto_falls_back_to_openpyxl_on_com_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_path = tmp_path / "book.xlsx"
    _create_workbook(input_path)

    monkeypatch.setattr(
        edit_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=True, reason=None),
    )

    def _raise_com_error(
        input_path: Path,
        output_path: Path,
        ops: list[PatchOp],
        auto_formula: bool,
    ) -> list[object]:
        raise RuntimeError("boom")

    def _fake_apply_openpyxl_engine(
        request: PatchRequest,
        input_path: Path,
        output_path: Path,
    ) -> OpenpyxlEngineResult:
        return OpenpyxlEngineResult()

    monkeypatch.setattr(edit_service, "apply_xlwings_engine", _raise_com_error)
    monkeypatch.setattr(
        edit_service, "apply_openpyxl_engine", _fake_apply_openpyxl_engine
    )
    result = edit_service.patch_workbook(
        PatchRequest(
            xlsx_path=input_path,
            ops=[PatchOp(op="set_value", sheet="Sheet1", cell="A1", value="new")],
            on_conflict="rename",
            backend="auto",
        )
    )
    assert result.error is None
    assert result.engine == "openpyxl"
    assert any("falling back to openpyxl" in warning for warning in result.warnings)


def test_edit_service_make_applies_ops_without_mcp_policy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(
        edit_runtime,
        "get_com_availability",
        lambda: ComAvailability(available=False, reason="test"),
    )
    out_path = tmp_path / "book.xlsx"
    result = edit_service.make_workbook(
        MakeRequest(
            out_path=out_path,
            ops=[
                PatchOp(op="add_sheet", sheet="Data"),
                PatchOp(op="set_value", sheet="Data", cell="A1", value="ok"),
            ],
            backend="openpyxl",
        )
    )

    assert result.error is None
    workbook = load_workbook(out_path)
    try:
        assert workbook["Data"]["A1"].value == "ok"
    finally:
        workbook.close()
