# ADR ワークフロー

この文書は、issue や PR から ADR を扱うときの標準フローを定義する。

## Phase 1 の対象

Phase 1 では次だけを標準化する。

1. ADR 要否判定
2. ADR 草案または既存 ADR 更新提案
3. ADR 文書の lint

整合性監査、索引更新、レビュー特化フローは将来フェーズで追加する。

## 標準フロー

1. issue または PR を読む
2. 関連する `docs/`, `dev-docs/specs/`, `dev-docs/adr/`, `tests/`, `src/` を読み、判定に必要な evidence triad を集める
3. `adr-suggester` で `required` / `recommended` / `not-needed` を判定する
4. `not-needed` の場合でも、判定理由と evidence triad を issue または PR に残す
5. `required` または `recommended` の場合は、`adr-drafter` で新規 ADR 草案または既存 ADR 更新提案を作る
6. 人または AI が内容をレビューする
7. `adr-linter` で形式と evidence を検査する
8. merge 時に関連 spec / docs / tests との整合を再確認する

## 読み順

ADR 系タスクでは、次の順で資料を確認する。

1. `docs/`
2. `dev-docs/specs/`
3. `dev-docs/adr/`
4. `tests/`
5. `src/`

AI 向けの判断基準が必要なときだけ、追加で次を読む。

- `dev-docs/agents/adr-governance.md`
- `dev-docs/agents/adr-criteria.md`

## skill ごとの責務

### `adr-suggester`

- 変更を設計判断として扱うべきか判定する
- verdict 前に evidence triad を集める
- 新規 ADR 候補と既存 ADR 候補を返す
- `not-needed` を含め、判定結果には evidence triad を添える
- 草案本文は生成しない

### `adr-drafter`

- 新規 ADR 草案か既存 ADR 更新提案を作る
- `背景`, `決定`, `影響`, `根拠` を埋める
- 根拠には `Tests`, `Code`, `Related specs` を含める

### `adr-linter`

- `状態`、必須セクション、evidence、`Supersedes` / `Superseded by` を検査する
- 修正文案より findings を優先する

## merge 前チェック

- ADR の結論が spec と矛盾していない
- spec に書かれた契約が tests で裏付けられている
- 既存 ADR を supersede する場合は相互参照が埋まっている
- ADR が不要な場合でも、その理由が issue または PR に残っている
- ADR が不要な場合でも、判定に使った `specs`, `src`, `tests` の根拠が追跡できる

## 将来フェーズ

- `adr-reconciler`: ADR と specs / tests / src の継続監査
- `adr-indexer`: 一覧、タグ、decision map の更新
- `adr-reviewer`: ADR 草案の設計レビュー専用観点
