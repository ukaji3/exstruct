# Backend Constraints

Use this file when the workflow depends on backend selection or runtime
capabilities.

## Backend summary

- `openpyxl`
  - default choice for ordinary workbook edits
  - supports `--dry-run`, `--return-inverse-ops`, and
    `--preflight-formula-check`
- `com`
  - required for COM-only editing behavior such as `create_chart`
  - required for `.xls` creation/editing
- `auto`
  - follows ExStruct's existing backend-selection policy
  - can use openpyxl for a dry run and COM for the real apply

## Constraints to call out

- `create_chart` is COM-only.
- `.xls` workflows require COM and are not valid with `backend=openpyxl`.
- Requests that need host-owned path restrictions, transport mapping, or
  artifact mirroring belong to MCP, not the local CLI Skill.

## Failure behavior

- State clearly when the requested backend or capability is unsupported.
- Do not pretend that `auto` guarantees the same engine across dry run and
  apply.
- Redirect to MCP only for host-policy concerns, not as a generic substitute
  for unsupported CLI behavior.
