"""Compatibility wrappers for the legacy patch service import path."""

from __future__ import annotations

from typing import Any, cast

from exstruct.edit.engine.openpyxl_engine import apply_openpyxl_engine
from exstruct.edit.engine.xlwings_engine import apply_xlwings_engine
from exstruct.edit.models import MakeRequest, PatchRequest, PatchResult
import exstruct.edit.runtime as edit_runtime
import exstruct.edit.service as edit_service
from exstruct.mcp.io import PathPolicy
import exstruct.mcp.patch.runtime as runtime


def _sync_compat_overrides() -> None:
    """Propagate legacy monkeypatch targets into the canonical edit service."""
    service_module = cast(Any, edit_service)
    runtime_module = cast(Any, edit_runtime)
    service_module.apply_openpyxl_engine = apply_openpyxl_engine
    service_module.apply_xlwings_engine = apply_xlwings_engine
    runtime_module.get_com_availability = runtime.get_com_availability
    runtime_module.PatchOpError = runtime.PatchOpError


def run_make(request: MakeRequest, *, policy: PathPolicy | None = None) -> PatchResult:
    """Compatibility wrapper for legacy make orchestration."""
    _sync_compat_overrides()
    return edit_service.make_workbook(request, policy=policy)


def run_patch(
    request: PatchRequest, *, policy: PathPolicy | None = None
) -> PatchResult:
    """Compatibility wrapper for legacy patch orchestration."""
    _sync_compat_overrides()
    return edit_service.patch_workbook(request, policy=policy)


__all__ = [
    "apply_openpyxl_engine",
    "apply_xlwings_engine",
    "run_make",
    "run_patch",
    "runtime",
]
