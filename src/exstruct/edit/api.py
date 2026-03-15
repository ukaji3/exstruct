"""Public function entry points for workbook editing."""

from __future__ import annotations

from .models import MakeRequest, PatchRequest, PatchResult
from .service import make_workbook, patch_workbook

__all__ = [
    "make_workbook",
    "patch_workbook",
    "MakeRequest",
    "PatchRequest",
    "PatchResult",
]
