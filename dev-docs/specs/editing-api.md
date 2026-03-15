# Editing API Specification

This document defines the Phase 1 public editing contract exposed from
`exstruct.edit`.

## Public import path

- Primary public package: `exstruct.edit`
- Primary functions:
  - `patch_workbook(request: PatchRequest) -> PatchResult`
  - `make_workbook(request: MakeRequest) -> PatchResult`
- Primary public models:
  - `PatchOp`
  - `PatchRequest`
  - `MakeRequest`
  - `PatchResult`
  - `PatchDiffItem`
  - `PatchErrorDetail`
  - `FormulaIssue`

## Phase 1 guarantees

- Python callers can edit workbooks through `exstruct.edit` without providing
  MCP-specific `PathPolicy` restrictions.
- The patch operation vocabulary, field names, defaults, warnings, and error
  payload shapes remain aligned with the existing MCP patch contract.
- `exstruct.edit` exposes the same operation normalization and schema metadata
  used by MCP:
  - `coerce_patch_ops`
  - `resolve_top_level_sheet_for_payload`
  - `list_patch_op_schemas`
  - `get_patch_op_schema`
- Existing MCP compatibility imports remain valid:
  - `exstruct.mcp.patch_runner`
  - `exstruct.mcp.patch.normalize`
  - `exstruct.mcp.patch.specs`
  - `exstruct.mcp.op_schema`

## Host-only responsibilities

The following behaviors are not part of the Python editing API contract and
remain owned by MCP / agent hosts:

- `PathPolicy` root restrictions and deny-glob enforcement
- MCP tool input/output models and transport mapping
- artifact mirroring for MCP hosts
- server-level defaults such as `--on-conflict`
- thread offloading, timeouts, and confirmation flows

## Current implementation boundary

- Phase 1 promotes `exstruct.edit` as the canonical public import path.
- The implementation intentionally reuses the existing patch execution pipeline
  under `exstruct.mcp.patch.*` to avoid destabilizing the tested backend logic
  during the API promotion.
- Contract metadata moved to `exstruct.edit` in Phase 1:
  - patch op types
  - chart type metadata
  - patch op alias/spec metadata
  - public op schema discovery

## Explicit non-goals for Phase 1

- No top-level `from exstruct import patch_workbook` export
- No new CLI subcommands
- No op renaming (`set_value` remains the public op name)
- No change to backend selection or fallback policy
- No change to `PatchResult` shape
