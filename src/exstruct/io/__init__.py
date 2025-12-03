from __future__ import annotations

import json
from pathlib import Path
from typing import Any

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


def save_as_json(model: WorkbookData, path: Path) -> None:
    filtered_dict = dict_without_empty_values(model)
    path.write_text(
        json.dumps(filtered_dict, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


__all__ = ["dict_without_empty_values", "save_as_json"]
