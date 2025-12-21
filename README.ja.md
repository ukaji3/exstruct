# ExStruct — Excel 構造化抽出エンジン（OOXML 対応フォーク）

![Licence: BSD-3-Clause](https://img.shields.io/badge/license-BSD--3--Clause-blue?style=flat-square)

![ExStruct Image](/docs/assets/icon.webp)

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

### オリジナル版からの拡張

オリジナルの ExStruct は Windows + Excel 環境に最適化されており、他のプラットフォームではセルのみの抽出にフォールバックする設計でした。本フォークはその優れた基盤の上に、クロスプラットフォームでの図形/チャート抽出を可能にする OOXML パーサーを追加しています：

| 機能 | オリジナル (COM なし) | OOXML パーサー追加後 |
|------|----------------------|---------------------|
| セル | ✓ | ✓ |
| テーブル候補 | ✓ | ✓ |
| 印刷範囲 | ✓ | ✓ |
| 図形抽出 | —（フォールバック） | ✓ |
| チャート抽出 | —（フォールバック） | ✓ |
| コネクター接続関係 | — | ✓ |
| 自動改ページ | — | —（COM 必須） |

この拡張により以下が可能になります：
- **Linux/macOS でのフローチャート抽出**（図形 + コネクターの begin_id/end_id）
- **Excel なしでのチャートデータ抽出**
- **CI/CD や Docker 環境**（ヘッドレス動作）

## License

BSD-3-Clause. See `LICENSE` for details.

## 謝辞

本プロジェクトは [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct) のフォークです。クリーンなアーキテクチャと充実したドキュメントを備えた優れた Excel 抽出エンジンを作成されたオリジナルの開発者の方々に深く感謝いたします。本フォークの OOXML パーサー拡張は、その素晴らしい基盤の上に構築されています。

## ドキュメント

- API リファレンス (GitHub Pages): https://harumiweb.github.io/exstruct/
- JSON Schema は `schemas/` にモデルごとに配置しています。モデル変更後は `python scripts/gen_json_schema.py` で再生成してください。
