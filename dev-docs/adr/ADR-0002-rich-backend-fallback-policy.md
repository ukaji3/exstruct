# ADR-0002: Rich Backend のフォールバック方針

## 状態

`accepted`

## 背景

rich extraction は、Windows の Excel COM や non-COM 環境の LibreOffice のように、不在または不安定になりうる runtime に依存している。rich artifact が欠ける理由を隠さず、それでも抽出結果を有用に保つためには、一貫した fallback 契約が必要である。

## 決定

- runtime unavailable は例外的な product failure ではなく、通常の fallback 条件として扱う。
- fallback reason は `FallbackReason` を通して明示的に記録・ログ出力する。
- rich backend が失敗しても、ExStruct は抽出全体を落とさず、取得可能な最善の安全な結果を残す。
- COM と LibreOffice は内部実装が異なっても、上位の fallback 方針は共通に保つ。

## 影響

- 新しい backend を導入するときは、fallback の振る舞いを最初に定義しなければならない。
- error handling を変更する場合は、返却データ形状と fallback reason の両方を確認する regression test が必要になる。
- 内部 runtime hardening によって public fallback contract が黙って変わってはならない。

## 根拠

- Tests: `tests/integration/test_integrate_fallback.py`, `tests/utils/test_logging_utils.py`
- Code: `src/exstruct/core/pipeline.py`, `src/exstruct/errors.py`
- Related specs: `dev-docs/specs/excel-extraction.md`, `docs/concept.md`

## Supersedes

- None

## Superseded by

- None
