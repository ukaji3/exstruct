# Safe Editing

Use this file for risky edits, unfamiliar workbooks, or any request where a
wrong apply would be expensive.

## Standard flow

1. Validate the input workbook when readability is uncertain.
2. Build or inspect the patch-op JSON.
3. Run `exstruct patch --dry-run` for risky edits.
4. Inspect `PatchResult`, warnings, `patch_diff`, and `engine`.
5. Re-run without `--dry-run` only after the plan is acceptable.
6. Perform lightweight verification or deeper verification based on risk.

## Ambiguity rules

- Ask for clarification before apply when sheet names, target cells/ranges,
  output naming, or overwrite policy are unclear.
- If the user only states a goal, inspect the op schema before inventing a
  payload shape.

## Risk checks

- Treat destructive or broad edits as risky:
  - overwrite/rename conflict decisions
  - merge / unmerge operations
  - range-wide formula fills
  - style changes across large regions
  - chart creation or COM-only workflows
- Prefer `--backend openpyxl` when you want the dry run and real apply to use
  the same engine.
- Explain that `--backend auto` may use openpyxl for the dry run but COM for
  the real apply on COM-capable hosts.

## Failure behavior

- Surface schema or validation failures directly.
- Do not silently retry with a different conceptual operation.
- If the request depends on unsupported runtime features, explain the
  constraint and route to the correct backend or MCP workflow.
