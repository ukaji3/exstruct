# ADR-0007: Editing CLI as Public Operational Interface

## Status

`accepted`

## Background

Phase 1 established `exstruct.edit` as the first-class Python API for workbook
editing while preserving MCP as the host-owned integration layer. That still
left a gap for command-line and agent-oriented workflows: extraction already had
an `exstruct` CLI, but editing was still exposed mainly through MCP tools.

Phase 2 needs to answer two policy questions that are likely to recur:

- how editing commands should coexist with the legacy extraction CLI
- whether the public editing CLI should expose JSON-first operational flows
  directly over `exstruct.edit` or continue to depend on MCP-facing entrypoints

The change also touches a public CLI contract, so the compatibility and layering
decision must be recorded explicitly rather than left in implementation notes.

## Decision

- ExStruct adds a first-class editing CLI on the existing `exstruct` console
  script with these Phase 2 commands:
  - `patch`
  - `make`
  - `ops list`
  - `ops describe`
  - `validate`
- The legacy extraction entrypoint `exstruct INPUT.xlsx ...` remains valid and
  is not replaced with `exstruct extract` in Phase 2.
- `patch`, `make`, and `ops*` are thin wrappers around the public
  `exstruct.edit` contract.
- Editing commands are JSON-first:
  - `patch` and `make` serialize `PatchResult`
  - `ops list` / `ops describe` serialize patch-op schema metadata
  - `validate` serializes workbook readability results
- Phase 2 does not introduce:
  - interactive confirmation flows
  - backup / allow-root / deny-glob flags
  - a request-envelope JSON CLI format
  - a new public Python validation API

## Consequences

- Users and agents gain a stable command-line surface for workbook editing
  without routing through MCP.
- The existing extraction CLI keeps backward compatibility because editing
  dispatch is opt-in by subcommand.
- The operational CLI now aligns with the public Python API, which reduces the
  risk of CLI-only business logic drift.
- `validate` remains a CLI-only operational helper in Phase 2, so Python API
  parity for validation is still deferred.
- The `exstruct` CLI now has two invocation styles (legacy extraction and edit
  subcommands), which is slightly less uniform than a full subcommand redesign
  but materially lowers migration risk.

## Rationale

- Tests:
  - `tests/cli/test_edit_cli.py`
  - `tests/cli/test_cli.py`
  - `tests/cli/test_cli_alpha_col.py`
  - `tests/edit/test_api.py`
  - `tests/mcp/test_validate_input.py`
- Code:
  - `src/exstruct/cli/main.py`
  - `src/exstruct/cli/edit.py`
  - `src/exstruct/edit/__init__.py`
  - `src/exstruct/mcp/validate_input.py`
- Related specs:
  - `dev-docs/specs/editing-api.md`
  - `dev-docs/specs/editing-cli.md`
  - `dev-docs/architecture/overview.md`
  - `docs/cli.md`
  - `docs/api.md`
  - `README.md`
  - `README.ja.md`

## Supersedes

- None

## Superseded by

- None
