# ADR-0004: Patch backend 選択方針

## 状態

`accepted`

## 背景

`exstruct_patch` は、能力差と安全制約の異なる複数 backend をサポートしている。backend の選択は互換性、fallback、許可される operation に直接影響するため、call site ごとの場当たりルールではなく、明示的な判断記録が必要である。

## 決定

- `backend="auto"` は COM が使える場合に COM を優先し、許可された runtime failure に対してのみ openpyxl へ fallback する。
- `backend="com"` は明示指定であり、openpyxl へ黙って fallback しない。
- `backend="openpyxl"` は安全な pure-Python path として維持し、chart 作成や `.xls` 処理のような feature gap は明示する。
- capability restriction は backend 実行時まで遅らせず、事前に分かるものは request validation として強制する。

## 影響

- backend 固有機能を追加するときは、`auto`, `com`, `openpyxl` それぞれの振る舞いを明示する必要がある。
- fallback policy を変える場合は、正負両側の test を組で追加する必要がある。
- documentation の backend capability table は runtime validation と同期していなければならない。

## 根拠

- Tests: `tests/mcp/patch/test_service.py`, `tests/mcp/patch/test_models_internal_coverage.py`
- Code: `src/exstruct/mcp/patch/runtime.py`, `src/exstruct/mcp/patch/service.py`, `src/exstruct/mcp/server.py`
- Related specs: `docs/mcp.md`, `dev-docs/specs/patch/legacy-dependency-inventory.md`

## Supersedes

- None

## Superseded by

- None
