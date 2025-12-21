# exstruct-go

Go 言語版 Excel 構造化抽出エンジン

## 概要

exstruct-go は Excel ワークブック（.xlsx）から構造化データを抽出し、JSON 形式で出力するライブラリおよび CLI ツールです。

[excelize](https://github.com/xuri/excelize) ライブラリを使用し、Excel なしでクロスプラットフォーム（Linux、macOS、Windows）で動作します。

## 機能

- セルデータ抽出（数値、テキスト、ハイパーリンク）
- 図形抽出（位置、サイズ、テキスト、回転、グループ展開）
- コネクター抽出（矢印スタイル、接続先 ID、方向）
- チャート抽出（タイトル、種別、系列データ、軸範囲）
- テーブル候補検出
- 印刷範囲抽出
- 出力モード（light/standard/verbose）

## インストール

```bash
go install github.com/ukaji3/exstruct-go/cmd/exstruct@latest
```

## CLI 使用方法

```bash
# 基本的な使用方法（stdout に JSON 出力）
exstruct input.xlsx

# ファイルに出力
exstruct input.xlsx -o output.json

# 整形出力
exstruct input.xlsx --pretty

# モード指定
exstruct input.xlsx --mode verbose

# シート単位出力
exstruct input.xlsx --sheets-dir sheets/

# 印刷範囲単位出力
exstruct input.xlsx --print-areas-dir areas/
```

### オプション

| フラグ | 説明 |
|--------|------|
| `-o, --output` | 出力ファイルパス（デフォルト: stdout） |
| `--pretty` | JSON を整形出力 |
| `--mode` | 抽出モード: light, standard, verbose |
| `--sheets-dir` | シート単位出力ディレクトリ |
| `--print-areas-dir` | 印刷範囲単位出力ディレクトリ |

### 出力モード

- **light**: セル + テーブル候補のみ
- **standard**: セル + 図形（テキスト付き/コネクター） + チャート + テーブル候補
- **verbose**: 全データ（図形サイズ、ハイパーリンク、チャートサイズ含む）

## ライブラリ使用方法

```go
package main

import (
    "fmt"
    "github.com/ukaji3/exstruct-go/pkg/exstruct"
    "github.com/ukaji3/exstruct-go/pkg/exstruct/output"
)

func main() {
    opts := exstruct.Options{
        Mode: exstruct.ModeStandard,
    }

    wb, err := exstruct.Extract("input.xlsx", opts)
    if err != nil {
        panic(err)
    }

    jsonData, err := output.ToJSON(wb, true)
    if err != nil {
        panic(err)
    }

    fmt.Println(string(jsonData))
}
```

## 出力形式

```json
{
  "book_name": "sample.xlsx",
  "sheets": {
    "Sheet1": {
      "rows": [
        {"r": 1, "c": {"1": "Header1", "2": "Header2"}},
        {"r": 2, "c": {"1": 100, "2": 200}}
      ],
      "shapes": [
        {
          "id": 1,
          "text": "開始",
          "l": 100,
          "t": 200,
          "type": "AutoShape-FlowchartProcess"
        }
      ],
      "charts": [...],
      "table_candidates": ["A1:D10"],
      "print_areas": [...]
    }
  }
}
```

## ライセンス

BSD-3-Clause
