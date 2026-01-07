# ExStruct — Excel 構造化抽出エンジン（OOXML 対応フォーク）

![Licence: BSD-3-Clause](https://img.shields.io/badge/license-BSD--3--Clause-blue?style=flat-square)

![ExStruct Image](docs/assets/icon.webp)

本リポジトリは [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct) のフォークで、クロスプラットフォームでの図形/チャート抽出を可能にする OOXML パーサーを追加しています。

インストール方法や基本的な使い方については、[オリジナルリポジトリ](https://github.com/harumiWeb/exstruct) を参照してください。

[English README](README.md)

## 本フォークの新機能

本フォークは純粋な Python による OOXML パーサーを追加し、**Linux や macOS** でも Excel なしで図形・チャートを抽出できるようにしています。

### 動作の仕組み

- **Windows + Excel**: COM API（xlwings 経由）を使用（全機能対応）
- **Linux / macOS**: 自動的に OOXML パーサーにフォールバック（Excel 不要）
- **Windows（Excel なし）**: OOXML パーサーを使用

### 対応機能（OOXML）

| 機能 | 対応 |
|------|------|
| 図形の位置 (l, t) | ✓ |
| 図形のサイズ (w, h) | ✓（verbose モード） |
| 図形のテキスト | ✓ |
| 図形の種別 | ✓ |
| 図形 ID の割り当て | ✓ |
| コネクターの方向 | ✓ |
| 矢印スタイル | ✓ |
| コネクター接続先 (begin_id, end_id) | ✓ |
| 回転 | ✓ |
| グループのフラット化 | ✓ |
| チャート種別 | ✓ |
| チャートタイトル | ✓ |
| Y 軸タイトル/範囲 | ✓ |
| 系列データ | ✓ |

### 制限事項（OOXML vs COM）

一部の機能は Excel の計算エンジンが必要なため、OOXML では実装できません：

- 自動計算された Y 軸範囲（Excel で「自動」設定の場合）
- タイトル/ラベルのセル参照解決
- 条件付き書式の評価
- 自動改ページの計算
- OLE / 埋め込みオブジェクト
- VBA マクロ

詳細な比較は [docs/com-vs-ooxml-implementation.md](docs/com-vs-ooxml-implementation.md) を参照してください。

## 主な特徴

- **Excel → 構造化 JSON**: セル、図形、チャート、SmartArt、テーブル候補、結合セル範囲、印刷範囲/自動改ページ範囲をシート単位で出力。
- **出力モード**: `light`（セル＋テーブル候補のみ）、`standard`（テキスト付き図形＋矢印、チャート、結合セル範囲）、`verbose`（全図形を幅高さ付きで出力、セルのハイパーリンクも出力）。
- **フォーマット**: JSON（デフォルトはコンパクト、`--pretty` で整形）、YAML、TOON（任意依存）。
- **テーブル検出のチューニング**: API でヒューリスティックを動的に変更可能。
- **安全なフォールバック**: Excel COM 不在でもプロセスは落ちず、セル＋テーブル候補に切り替え。

## インストール

```bash
pip install exstruct
```

オプション依存:

- YAML: `pip install pyyaml`
- TOON: `pip install python-toon`
- レンダリング（PDF/PNG）: Excel + `pip install pypdfium2 pillow`
- まとめて導入: `pip install exstruct[yaml,toon,render]`

## クイックスタート CLI

```bash
exstruct input.xlsx > output.json          # デフォルトは標準出力のコンパクト JSON
exstruct input.xlsx -o out.json --pretty   # 整形 JSON をファイルへ
exstruct input.xlsx --format yaml          # YAML（pyyaml が必要）
exstruct input.xlsx --sheets-dir sheets/   # シートごとに分割出力
exstruct input.xlsx --mode light           # セル＋テーブル候補のみ
```

## クイックスタート Python

```python
from pathlib import Path
from exstruct import extract, export

wb = extract("input.xlsx", mode="standard")
export(wb, Path("out.json"), pretty=False)

first_sheet = wb["Sheet1"]
for name, sheet in wb:
    print(name, len(sheet.rows))
wb.save("out.json", pretty=True)
```

**備考 (COM 非対応環境):** Excel COM が使えない場合でもセル＋`table_candidates` は返りますが、`shapes` / `charts` は空になります。

## 出力モード

- **light**: セル＋テーブル候補のみ（COM 不要）。
- **standard**: テキスト付き図形＋矢印、チャート、テーブル候補。
- **verbose**: 全図形、チャート、ハイパーリンク、`colors_map`。

## License

BSD-3-Clause. See `LICENSE` for details.

## 謝辞

本プロジェクトは [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct) のフォークです。クリーンなアーキテクチャと充実したドキュメントを備えた優れた Excel 抽出エンジンを作成されたオリジナルの開発者の方々に深く感謝いたします。

## ドキュメント

- API リファレンス (GitHub Pages): https://harumiweb.github.io/exstruct/
- JSON Schema は `schemas/` にモデルごとに配置しています。
