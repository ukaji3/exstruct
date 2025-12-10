from __future__ import annotations

"""Shared JSON-compatible type aliases used across ExStruct."""

JsonPrimitive = str | int | float | bool | None
JsonStructure = JsonPrimitive | list["JsonStructure"] | dict[str, "JsonStructure"]

__all__ = ["JsonPrimitive", "JsonStructure"]
