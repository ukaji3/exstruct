from __future__ import annotations

from collections.abc import Mapping
import json
from pathlib import Path

import pytest

from exstruct.mcp.chunk_reader import (
    ReadJsonChunkFilter,
    ReadJsonChunkRequest,
    read_json_chunk,
)


def _write_json(path: Path, data: Mapping[str, object]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def test_read_json_chunk_raw(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, max_bytes=10_000)
    result = read_json_chunk(request)
    assert json.loads(result.chunk) == data
    assert result.next_cursor is None


def test_read_json_chunk_raw_too_large(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(out_path=out, max_bytes=10)
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_with_filters(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {
            "Sheet1": {
                "rows": [
                    {"r": 1, "c": {"0": "A", "1": "B"}},
                    {"r": 2, "c": {"0": "C", "1": "D"}},
                ]
            }
        },
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(rows=(1, 1), cols=(1, 1)),
    )
    result = read_json_chunk(request)
    payload = json.loads(result.chunk)
    rows = payload["sheet"]["rows"]
    assert len(rows) == 1
    assert rows[0]["r"] == 1
    assert rows[0]["c"] == {"0": "A"}


def test_read_json_chunk_requires_sheet(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"A": {"rows": []}, "B": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        max_bytes=10_000,
        filter=ReadJsonChunkFilter(rows=(1, 1)),
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_invalid_cursor(tmp_path: Path) -> None:
    data = {"book_name": "book", "sheets": {"Sheet1": {"rows": []}}}
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        cursor="bad",
        max_bytes=10_000,
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)


def test_read_json_chunk_cursor_beyond_rows(tmp_path: Path) -> None:
    data = {
        "book_name": "book",
        "sheets": {"Sheet1": {"rows": [{"r": 1, "c": {"0": "A"}}]}},
    }
    out = tmp_path / "out.json"
    _write_json(out, data)
    request = ReadJsonChunkRequest(
        out_path=out,
        sheet="Sheet1",
        cursor="2",
        max_bytes=10_000,
    )
    with pytest.raises(ValueError):
        read_json_chunk(request)
