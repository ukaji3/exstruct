# Feature Spec

## 2026-03-15 tasks document cleanup

### Goal

- `tasks/feature_spec.md` と `tasks/todo.md` から、完了済みの過去 task section を除去する。
- 恒久情報は `dev-docs/`、`docs/`、`AGENTS.md`、既存 ADR に残し、tasks には今回の棚卸し結果だけを残す。

### Retained references

- ADR: `dev-docs/adr/ADR-0001-extraction-mode-boundaries.md`, `dev-docs/adr/ADR-0002-rich-backend-fallback-policy.md`, `dev-docs/adr/ADR-0003-output-serialization-omission-policy.md`, `dev-docs/adr/ADR-0004-patch-backend-selection-policy.md`, `dev-docs/adr/ADR-0005-path-policy-safety-boundary.md`
- Internal specs: `dev-docs/specs/excel-extraction.md`, `dev-docs/specs/adr-index.md`, `dev-docs/specs/adr-review.md`, `dev-docs/testing/test-requirements.md`
- Agent governance: `dev-docs/agents/adr-governance.md`, `dev-docs/agents/adr-criteria.md`, `dev-docs/agents/adr-workflow.md`, `AGENTS.md`
- Public docs: `docs/api.md`, `docs/cli.md`, `docs/mcp.md`, `docs/release-notes/v0.6.1.md`, `CHANGELOG.md`

### ADR verdict

- `adr-suggester`: `not-needed`
- rationale: tasks 側に残っていたのは完了済みタスクの仕様メモと作業ログであり、既存恒久文書で説明できない policy-level decision は確認できなかった。
- `adr-reconciler` audit scope: `ADR-0001`〜`ADR-0005`, `dev-docs/specs/excel-extraction.md`, `dev-docs/testing/test-requirements.md`, `docs/api.md`, `docs/cli.md`, `docs/mcp.md`, `AGENTS.md`
- `adr-reconciler` findings: なし
- next action: `no-action`

## 2026-03-15 issue #99 phase 1 public edit API

### Goal

- Excel editing を `exstruct.edit` から利用できる first-class Python API として公開する。
- `PatchOp` / `PatchRequest` / `MakeRequest` / `PatchResult` と既存 op 契約を維持したまま、MCP 固有の path policy を public API から外す。
- MCP は互換レイヤとして維持し、tool I/O と host safety policy を担当し続ける。

### Public contract

- Primary public import path: `exstruct.edit`
- Public entry points:
  - `patch_workbook(request: PatchRequest) -> PatchResult`
  - `make_workbook(request: MakeRequest) -> PatchResult`
- Public helpers:
  - `coerce_patch_ops`
  - `resolve_top_level_sheet_for_payload`
  - `list_patch_op_schemas`
  - `get_patch_op_schema`
- Preserved Phase 1 contract:
  - existing op names remain unchanged
  - existing warning/error payload shapes remain unchanged
  - existing MCP compatibility imports remain valid

### Boundary

- `PathPolicy`, artifact mirroring, MCP tool payloads, server defaults, and thread offloading remain MCP-owned behavior.
- Phase 1 intentionally reuses the existing `exstruct.mcp.patch.*` execution pipeline under the hood to reduce backend regression risk while promoting the public API.

### Permanent references

- ADR: `dev-docs/adr/ADR-0006-public-edit-api-and-host-boundary.md`
- Internal specs:
  - `dev-docs/specs/editing-api.md`
  - `dev-docs/specs/data-model.md`
  - `dev-docs/architecture/overview.md`
- Public docs:
  - `docs/api.md`
  - `docs/mcp.md`

## 2026-03-15 pr #102 review follow-up

### Goal

- PR #102 の妥当な review 指摘だけを取り込み、Phase 1 の公開契約を変えずに不整合と正規化漏れを修正する。
- `coerce_patch_ops` / `resolve_top_level_sheet_for_payload` の JSON-string op 対応を、dict op と同じ indexed error shape で安定化する。
- ADR-0006 の status metadata を `accepted` に統一し、PR metadata warning は最小範囲で解消する。

### Accepted findings

- `coerce_patch_ops([None])` が `AttributeError` を漏らし、indexed validation error にならない。
- `resolve_top_level_sheet_for_payload` が JSON-string op を未解決のまま返し、top-level `sheet` fallback を適用できない。
- `ADR-0006` の本文と index artifacts の status が不一致。

### Chosen constraints

- public API signature と import path は変更しない。
- invalid op 型の失敗は `ValueError(build_patch_op_error_message(...))` に統一する。
- docstring warning 対応は、この PR で新規追加した `src/exstruct/edit/*.py` の不足 module docstring 補完までに限定する。
- PR 本文は `.github/pull_request_template.md` の見出し構造に合わせるが、Acceptance Criteria は issue #99 phase 1 用に書き換える。
