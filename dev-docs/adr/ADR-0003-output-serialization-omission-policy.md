# ADR-0003: 出力直列化における省略方針

## 状態

`accepted`

## 背景

ExStruct は LLM / RAG の前処理用途で使われるため、出力サイズが重要になる。backend metadata はデバッグや忠実度分析には有用だが、既定で常に含めると通常利用時の token cost が増える。

## 決定

- 直列化出力では `provenance`, `approximation_level`, `confidence` のような backend metadata を既定で省略する。
- backend metadata は public interface の明示的な opt-in flag で利用可能なまま維持する。
- この省略方針は serialization contract の一部であり、API、CLI、MCP、schema の期待値で一致していなければならない。

## 影響

- 新しい metadata field は、明確な public need がない限り既定で省略するべきである。
- public docs には opt-in 方法を記載する必要がある。
- tests は既定省略パスと明示 include パスの両方をカバーする必要がある。

## 根拠

- Tests: `tests/models/test_models_export.py`, `tests/models/test_schemas_generated.py`
- Code: `src/exstruct/io/serialize.py`, `src/exstruct/models/__init__.py`
- Related specs: `docs/api.md`, `docs/cli.md`

## Supersedes

- None

## Superseded by

- None
