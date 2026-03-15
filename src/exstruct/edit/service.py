"""Public workbook editing service wrappers."""

from __future__ import annotations

from exstruct.mcp.patch.service import run_make, run_patch

from .models import MakeRequest, PatchRequest, PatchResult


def patch_workbook(request: PatchRequest) -> PatchResult:
    """Edit an existing workbook without MCP path policy enforcement."""

    return run_patch(request, policy=None)


def make_workbook(request: MakeRequest) -> PatchResult:
    """Create a new workbook and apply initial patch operations."""

    return run_make(request, policy=None)


__all__ = ["make_workbook", "patch_workbook"]
