# Reconstruction Utility Benchmark (RUB) Specification

## 1. 目的

RUB は「再構築された Markdown が後続タスクにどれだけ耐えるか」を機械的に測る。
Markdown 文字列の一致ではなく、構造利用性（Reconstruction Utility）を評価対象とする。

## 2. 評価対象

- 入力: 同一 Excel 文書
- 手法: pdf / image_vlm / exstruct / html / openpyxl
- 出力: 各手法で生成した Markdown
- 評価: Markdown のみを入力にした構造クエリの解答精度

## 3. 評価フロー（2段階）

### Stage A: 再構築

各手法で Markdown を生成する。

- pdf: soffice → pdf → テキスト抽出 → Markdown
- image_vlm: 画像レンダリング → VLM → Markdown
- exstruct: exstruct JSON → LLM → Markdown
- html / openpyxl: テキスト抽出 → Markdown

### Stage B: 利用（採点対象）

Stage A の Markdown だけを入力に、構造クエリを解かせる。

- 出力は JSON のみ
- JSON はスキーマ固定
- 採点は正規化後の完全一致（deterministic）

## 4. タスク設計方針

- 文字列一致に依存せず、構造理解を問う
- 手法間で不公平にならないよう、入力は Markdown のみ
- 各タスクは以下のいずれかに分類する
  - 集合問題（項目一覧）
  - グラフ問題（ノード/エッジ）
  - 階層問題（親子関係）
  - 表問題（行列の対応）
- 丸数字や装飾記号など、表記ゆれが大きい要素は避ける

## 5. 正規化（決定的）

採点前に以下の正規化を行う。

- 文字列: 前後空白削除、連続空白を 1 つに、改行は \n に統一
- 辞書: キーソート（canonicalization）
- 配列: 順序が意味を持たないタスクはソート
- 数値: 可能な範囲で数値化（例: "012" → 12）

## 6. 採点指標

### 6.1 主指標: RUS

RUS = 正解数 / 問題数

### 6.2 副指標

- Cost-normalized RUS = RUS / cost_usd
- Token-normalized RUS = RUS / input_tokens
- Stage A failure rate = Markdown 生成失敗率

## 7. データ構成

```
benchmark/
  rub/
    README.md
    BENCHMARK_SPEC.md
    manifest.json
    truth/
      *.json
    schemas/
      *.schema.json
    scoring/
      normalize.py
      score.py
    diagrams/
      rub_overview.mmd
      scoring_flow.mmd
```

## 8. manifest 仕様（案）

- id: ケースID
- type: タスク種別
- xlsx: 元ファイルパス
- question: Stage B クエリ
- truth: 正解 JSON パス
- sheet_scope: 対象シート（null なら全体）
- render: 画像レンダ設定

## 9. 再現性

- モデル名、温度、実行日時を記録
- 正規化ルールと採点コードを完全公開
- ランダム性がある工程は温度 0 固定

## 10. 公開時の注意

- 「Markdown 一致」は補助指標としてのみ扱う
- RUS（利用可能性）を主指標として説明する
- 用途が異なる手法を同一スコアで殴らない

---

この仕様は v1 とし、変更時は履歴を残す。
