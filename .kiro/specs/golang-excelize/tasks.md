# Implementation Plan

## exstruct-go: Go 言語版 Excel 構造化抽出エンジン

- [ ] 1. プロジェクト初期化とディレクトリ構造
  - [ ] 1.1 Go モジュール初期化と依存関係設定
    - `go/` ディレクトリ作成
    - `go mod init github.com/ukaji3/exstruct-go`
    - excelize/v2, cobra 依存追加
    - _Requirements: 9.1_
  - [ ] 1.2 ディレクトリ構造作成
    - `cmd/exstruct/`, `pkg/exstruct/models/`, `pkg/exstruct/parser/`, `pkg/exstruct/output/` 作成
    - _Requirements: 9.1_

- [ ] 2. データモデル実装
  - [ ] 2.1 基本モデル定義
    - `pkg/exstruct/models/cell.go`: CellRow 構造体
    - `pkg/exstruct/models/shape.go`: Shape 構造体
    - `pkg/exstruct/models/chart.go`: Chart, ChartSeries 構造体
    - `pkg/exstruct/models/print_area.go`: PrintArea 構造体
    - `pkg/exstruct/models/sheet.go`: SheetData 構造体
    - `pkg/exstruct/models/workbook.go`: WorkbookData 構造体
    - _Requirements: 7.1, 7.4_
  - [ ]* 2.2 Property test: JSON Schema Conformance
    - **Property 9: JSON Schema Conformance**
    - **Validates: Requirements 7.1**
  - [ ]* 2.3 Property test: Omit Empty Fields
    - **Property 10: Omit Empty Fields**
    - **Validates: Requirements 7.4**

- [ ] 3. ユニット変換とユーティリティ
  - [ ] 3.1 EMU → ピクセル変換実装
    - `pkg/exstruct/parser/units.go`: emu_to_pixels 関数
    - _Requirements: 2.3, 4.6_
  - [ ]* 3.2 Property test: EMU to Pixel Conversion
    - **Property 3: EMU to Pixel Conversion**
    - **Validates: Requirements 2.3, 4.6**

- [ ] 4. セル抽出実装
  - [ ] 4.1 セルデータ抽出
    - `pkg/exstruct/parser/cells.go`: ExtractCells 関数
    - excelize の GetRows, GetCellValue 使用
    - 数値/文字列型の判定
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5_
  - [ ] 4.2 ハイパーリンク抽出
    - GetCellHyperLink 使用
    - verbose モードでのみ出力
    - _Requirements: 1.6_
  - [ ]* 4.3 Property test: Cell Data Preservation
    - **Property 1: Cell Data Preservation**
    - **Validates: Requirements 1.1, 1.2**
  - [ ]* 4.4 Property test: Cell Type Preservation
    - **Property 2: Cell Type Preservation**
    - **Validates: Requirements 1.3, 1.4**

- [ ] 5. Checkpoint - セル抽出テスト
  - Ensure all tests pass, ask the user if questions arise.

- [ ] 6. 図形抽出実装
  - [ ] 6.1 DrawingML パーサー基盤
    - `pkg/exstruct/parser/shapes.go`: 基本構造
    - xl/drawings/drawing*.xml の解析
    - _Requirements: 2.1_
  - [ ] 6.2 図形位置・サイズ抽出
    - xfrm 要素から位置・サイズ取得
    - EMU → ピクセル変換適用
    - _Requirements: 2.3, 2.4_
  - [ ] 6.3 図形テキスト抽出
    - txBody/a:t 要素からテキスト取得
    - _Requirements: 2.2_
  - [ ] 6.4 プリセットジオメトリマッピング
    - PRESET_GEOM_MAP テーブル実装
    - _Requirements: 2.5_
  - [ ] 6.5 回転角度抽出
    - xfrm/@rot から角度計算（1/60000 度単位）
    - _Requirements: 2.6_
  - [ ] 6.6 グループ図形展開
    - grpSp 要素の再帰的パース
    - _Requirements: 2.7_
  - [ ]* 6.7 Property test: Shape Text Extraction
    - **Property 4: Shape Text Extraction**
    - **Validates: Requirements 2.2**
  - [ ]* 6.8 Property test: Preset Geometry Mapping
    - **Property 5: Preset Geometry Mapping**
    - **Validates: Requirements 2.5**

- [ ] 7. コネクター抽出実装
  - [ ] 7.1 コネクター識別
    - cxnSp 要素の検出
    - straightConnector, bentConnector 等の判定
    - _Requirements: 3.1_
  - [ ] 7.2 矢印スタイル抽出
    - a:ln/a:headEnd, a:tailEnd から矢印タイプ取得
    - _Requirements: 3.2_
  - [ ] 7.3 接続先 ID 抽出
    - cNvCxnSpPr/a:stCxn, a:endCxn から接続先取得
    - _Requirements: 3.3_
  - [ ] 7.4 方向計算
    - 幅・高さから角度計算、コンパス方向へマッピング
    - _Requirements: 3.4_
  - [ ] 7.5 Shape ID 割り当て
    - 非コネクター図形に連番 ID 付与
    - コネクター接続先を解決
    - _Requirements: 3.5_
  - [ ]* 7.6 Property test: Connector Direction Computation
    - **Property 6: Connector Direction Computation**
    - **Validates: Requirements 3.4**
  - [ ]* 7.7 Property test: Shape ID Assignment Order
    - **Property 7: Shape ID Assignment Order**
    - **Validates: Requirements 3.5**

- [ ] 8. Checkpoint - 図形抽出テスト
  - Ensure all tests pass, ask the user if questions arise.

- [ ] 9. チャート抽出実装
  - [ ] 9.1 ChartML パーサー基盤
    - `pkg/exstruct/parser/charts.go`: 基本構造
    - xl/charts/chart*.xml の解析
    - _Requirements: 4.1_
  - [ ] 9.2 チャートタイトル抽出
    - c:title 要素からタイトル取得
    - _Requirements: 4.2_
  - [ ] 9.3 チャートタイプマッピング
    - CHART_TYPE_MAP テーブル実装
    - _Requirements: 4.3_
  - [ ] 9.4 系列データ抽出
    - c:ser 要素から name, x_range, y_range 取得
    - _Requirements: 4.4_
  - [ ] 9.5 Y 軸情報抽出
    - c:valAx から title, min, max 取得
    - _Requirements: 4.5_
  - [ ] 9.6 チャート位置抽出
    - drawing.xml の graphicFrame から位置取得
    - _Requirements: 4.6_

- [ ] 10. テーブル候補検出実装
  - [ ] 10.1 テーブル検出アルゴリズム
    - `pkg/exstruct/parser/tables.go`: DetectTables 関数
    - 連続セル領域の検出
    - _Requirements: 5.1, 5.2_
  - [ ] 10.2 検出閾値パラメータ
    - density_min, coverage_min 等の設定
    - _Requirements: 5.3, 5.4_

- [ ] 11. 印刷範囲抽出実装
  - [ ] 11.1 印刷範囲パーサー
    - `pkg/exstruct/parser/print_areas.go`: ExtractPrintAreas 関数
    - xl/workbook.xml の definedNames から _xlnm.Print_Area 解析
    - _Requirements: 11.1, 11.2, 11.3_

- [ ] 12. 出力モード実装
  - [ ] 12.1 モードフィルタリング
    - `pkg/exstruct/options.go`: Mode 型と Options 構造体
    - light/standard/verbose のフィルタリングロジック
    - _Requirements: 6.1, 6.2, 6.3, 6.4_
  - [ ]* 12.2 Property test: Mode Filtering
    - **Property 8: Mode Filtering**
    - **Validates: Requirements 6.1, 6.2, 6.3**

- [ ] 13. JSON 出力実装
  - [ ] 13.1 JSON シリアライザ
    - `pkg/exstruct/output/json.go`: ToJSON 関数
    - pretty/compact オプション
    - omitempty タグによる空フィールド除外
    - _Requirements: 7.1, 7.2, 7.3, 7.4_
  - [ ] 13.2 UTF-8 エンコーディング
    - 非 ASCII 文字の保持
    - _Requirements: 7.5_
  - [ ]* 13.3 Property test: UTF-8 Preservation
    - **Property 11: UTF-8 Preservation**
    - **Validates: Requirements 7.5**

- [ ] 14. Checkpoint - コア機能テスト
  - Ensure all tests pass, ask the user if questions arise.

- [ ] 15. メイン抽出ロジック統合
  - [ ] 15.1 Extract 関数実装
    - `pkg/exstruct/extract.go`: Extract 関数
    - セル、図形、チャート、テーブル、印刷範囲の統合
    - _Requirements: 9.2, 9.3, 9.4_

- [ ] 16. CLI 実装
  - [ ] 16.1 基本 CLI 構造
    - `cmd/exstruct/main.go`: エントリーポイント
    - cobra によるコマンド定義
    - _Requirements: 8.1_
  - [ ] 16.2 出力オプション
    - -o/--output, --pretty フラグ
    - _Requirements: 8.2, 8.3_
  - [ ] 16.3 モードオプション
    - --mode フラグ
    - _Requirements: 8.4_
  - [ ] 16.4 エラーハンドリング
    - stderr 出力、非ゼロ終了コード
    - _Requirements: 8.5, 10.1, 10.2_
  - [ ] 16.5 シート単位出力
    - --sheets-dir フラグ
    - _Requirements: 12.1_
  - [ ] 16.6 印刷範囲単位出力
    - --print-areas-dir フラグ
    - _Requirements: 12.2, 12.3_

- [ ] 17. エラーハンドリング実装
  - [ ] 17.1 エラー型定義
    - `pkg/exstruct/errors.go`: カスタムエラー型
    - FileNotFoundError, InvalidFormatError 等
    - _Requirements: 10.1, 10.2_
  - [ ] 17.2 部分失敗処理
    - シート/要素レベルのエラーを警告として処理
    - _Requirements: 10.3, 10.4_

- [ ] 18. 統合テスト
  - [ ]* 18.1 サンプルファイルテスト
    - sample/sample.xlsx を使用した E2E テスト
    - Python 版との出力比較
  - [ ]* 18.2 CLI テスト
    - 各フラグの動作確認

- [ ] 19. Final Checkpoint - 全テスト確認
  - Ensure all tests pass, ask the user if questions arise.

- [ ] 20. ドキュメント作成
  - [ ] 20.1 README.md 作成
    - インストール方法、使用例、API リファレンス
  - [ ] 20.2 go doc コメント整備
    - 公開 API のドキュメントコメント
