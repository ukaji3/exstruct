# ADR-0001: 抽出モードの責務境界

## 状態

`accepted`

## 背景

ExStruct は複数の抽出モードを提供しており、それぞれ保証内容と必要 runtime が異なる。明示的な判断記録がないと、`light`, `libreoffice`, `standard`, `verbose` の責務が曖昧になりやすく、新しい artifact や validation rule を追加するときに境界が崩れやすい。

## 決定

- `light` は最小構成を維持し、rich runtime 依存を持ち込まない。
- `libreoffice` は `.xlsx/.xlsm` 向けの non-COM best-effort rich mode とし、PDF、画像、自動改ページ export は明示的に拒否する。
- `standard` と `verbose` は、より高忠実度な Excel ネイティブ抽出を行う COM 利用可能モードとして維持する。
- モード別 validation は単なる実装都合ではなく、プロダクト契約の一部として扱う。

## 影響

- 新機能を追加するときは、どの mode がその振る舞いを担当するかを明示する必要がある。
- validation logic は API、CLI、engine の入口で一致していなければならない。
- mode ごとの責務表は、tests の暗黙知だけに頼らず明示的に維持される。

## 根拠

- Tests: `tests/test_constraints.py`
- Code: `src/exstruct/constraints.py`, `src/exstruct/engine.py`
- Related specs: `docs/api.md`, `docs/cli.md`, `dev-docs/specs/excel-extraction.md`

## Supersedes

- None

## Superseded by

- None
