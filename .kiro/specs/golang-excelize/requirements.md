# Requirements Document

## Introduction

本ドキュメントは、excelize ライブラリを使用した Go 言語版 ExStruct（exstruct-go）の要件を定義します。exstruct-go は Excel ワークブック（.xlsx）から構造化データ（セル、図形、チャート、テーブル候補）を抽出し、JSON 形式で出力するライブラリおよび CLI ツールです。

Python 版 ExStruct の OOXML パーサー機能を Go で再実装し、クロスプラットフォーム（Linux、macOS、Windows）で動作することを目標とします。

## Glossary

- **exstruct-go**: Go 言語版 Excel 構造化抽出エンジン
- **excelize**: Go 言語用 Excel ファイル操作ライブラリ（github.com/qax-os/excelize）
- **WorkbookData**: ワークブック全体の構造化データを表すモデル
- **SheetData**: 単一シートの構造化データを表すモデル
- **Shape**: 図形（オートシェイプ、コネクター、テキストボックス等）のメタデータ
- **Chart**: チャート（グラフ）のメタデータ
- **CellRow**: 行単位のセルデータ
- **Connector**: 図形間を接続する線（矢印含む）
- **EMU**: English Metric Units（Excel 内部の座標単位、914400 EMU = 1 インチ）
- **DrawingML**: Office Open XML の図形描画仕様
- **ChartML**: Office Open XML のチャート仕様

## Requirements

### Requirement 1: セルデータ抽出

**User Story:** As a developer, I want to extract cell data from Excel worksheets, so that I can process spreadsheet content programmatically.

#### Acceptance Criteria

1. WHEN a user provides an xlsx file path THEN the exstruct-go SHALL read all non-empty cells from each worksheet
2. WHEN extracting cell data THEN the exstruct-go SHALL preserve the row index (1-based) and column index (string format) for each cell
3. WHEN a cell contains a numeric value THEN the exstruct-go SHALL output the value as a number type (int or float)
4. WHEN a cell contains text THEN the exstruct-go SHALL output the value as a string type
5. WHEN a cell contains a formula THEN the exstruct-go SHALL output the calculated value (not the formula itself)
6. WHEN a cell contains a hyperlink THEN the exstruct-go SHALL extract the link URL in verbose mode

### Requirement 2: 図形抽出

**User Story:** As a developer, I want to extract shape information from Excel worksheets, so that I can analyze flowcharts and diagrams.

#### Acceptance Criteria

1. WHEN extracting shapes THEN the exstruct-go SHALL parse DrawingML (xl/drawings/drawing*.xml) to obtain shape metadata
2. WHEN a shape has text content THEN the exstruct-go SHALL extract and include the text in the output
3. WHEN extracting shape position THEN the exstruct-go SHALL convert EMU coordinates to pixel values (96 DPI)
4. WHEN extracting shape size THEN the exstruct-go SHALL include width and height in verbose mode only
5. WHEN a shape has a preset geometry THEN the exstruct-go SHALL map the OOXML preset name to a human-readable type label
6. WHEN a shape has rotation THEN the exstruct-go SHALL include the rotation angle in degrees
7. WHEN shapes are grouped THEN the exstruct-go SHALL flatten the group and extract individual shapes

### Requirement 3: コネクター抽出

**User Story:** As a developer, I want to extract connector relationships between shapes, so that I can reconstruct flowchart logic.

#### Acceptance Criteria

1. WHEN a connector shape exists THEN the exstruct-go SHALL identify it as a connector type
2. WHEN a connector has arrow heads THEN the exstruct-go SHALL extract begin_arrow_style and end_arrow_style values
3. WHEN a connector is attached to shapes THEN the exstruct-go SHALL extract begin_id and end_id referencing the connected shape IDs
4. WHEN extracting connector direction THEN the exstruct-go SHALL compute compass direction (N, NE, E, SE, S, SW, W, NW) from connector geometry
5. WHEN assigning shape IDs THEN the exstruct-go SHALL assign sequential IDs to non-connector shapes first, then resolve connector endpoints

### Requirement 4: チャート抽出

**User Story:** As a developer, I want to extract chart information from Excel worksheets, so that I can analyze data visualizations.

#### Acceptance Criteria

1. WHEN extracting charts THEN the exstruct-go SHALL parse ChartML (xl/charts/chart*.xml) to obtain chart metadata
2. WHEN a chart has a title THEN the exstruct-go SHALL extract the title text
3. WHEN extracting chart type THEN the exstruct-go SHALL map OOXML chart element tags to human-readable type names (Line, Bar, Pie, etc.)
4. WHEN a chart has series data THEN the exstruct-go SHALL extract series name, name_range, x_range, and y_range for each series
5. WHEN a chart has Y-axis configuration THEN the exstruct-go SHALL extract y_axis_title and y_axis_range (min/max) when explicitly set
6. WHEN extracting chart position THEN the exstruct-go SHALL convert EMU coordinates to pixel values

### Requirement 5: テーブル候補検出

**User Story:** As a developer, I want to identify table-like regions in worksheets, so that I can extract structured tabular data.

#### Acceptance Criteria

1. WHEN analyzing cell data THEN the exstruct-go SHALL detect contiguous rectangular regions with data
2. WHEN a region meets table criteria THEN the exstruct-go SHALL output the range as a table candidate (e.g., "B3:E9")
3. WHEN detecting tables THEN the exstruct-go SHALL consider cell density and coverage thresholds
4. WHEN multiple table candidates exist THEN the exstruct-go SHALL output all detected candidates

### Requirement 6: 出力モード

**User Story:** As a developer, I want to control the level of detail in the output, so that I can balance between completeness and file size.

#### Acceptance Criteria

1. WHEN mode is "light" THEN the exstruct-go SHALL output cells and table_candidates only (no shapes or charts)
2. WHEN mode is "standard" THEN the exstruct-go SHALL output cells, shapes with text or connectors, charts, and table_candidates
3. WHEN mode is "verbose" THEN the exstruct-go SHALL output all data including shape dimensions, cell hyperlinks, and chart dimensions
4. WHEN no mode is specified THEN the exstruct-go SHALL default to "standard" mode

### Requirement 7: JSON 出力

**User Story:** As a developer, I want to export extracted data as JSON, so that I can integrate with other tools and systems.

#### Acceptance Criteria

1. WHEN exporting data THEN the exstruct-go SHALL produce valid JSON conforming to the WorkbookData schema
2. WHEN pretty option is enabled THEN the exstruct-go SHALL format JSON with indentation for readability
3. WHEN pretty option is disabled THEN the exstruct-go SHALL produce compact JSON (default)
4. WHEN a field has no value THEN the exstruct-go SHALL omit the field from output (exclude_none behavior)
5. WHEN serializing to JSON THEN the exstruct-go SHALL use UTF-8 encoding and preserve non-ASCII characters

### Requirement 8: CLI インターフェース

**User Story:** As a user, I want to use exstruct-go from the command line, so that I can process Excel files without writing code.

#### Acceptance Criteria

1. WHEN a user runs the CLI with an xlsx file path THEN the exstruct-go SHALL output JSON to stdout by default
2. WHEN -o/--output flag is provided THEN the exstruct-go SHALL write output to the specified file
3. WHEN --pretty flag is provided THEN the exstruct-go SHALL format output with indentation
4. WHEN --mode flag is provided THEN the exstruct-go SHALL use the specified output mode (light, standard, verbose)
5. WHEN an error occurs THEN the exstruct-go SHALL print error message to stderr and exit with non-zero code

### Requirement 9: ライブラリ API

**User Story:** As a Go developer, I want to use exstruct-go as a library, so that I can integrate Excel extraction into my applications.

#### Acceptance Criteria

1. WHEN importing the library THEN the developer SHALL have access to Extract function accepting file path and options
2. WHEN calling Extract THEN the function SHALL return WorkbookData struct and error
3. WHEN configuring extraction THEN the developer SHALL be able to specify mode via Options struct
4. WHEN accessing sheet data THEN the developer SHALL be able to iterate over sheets using WorkbookData.Sheets map

### Requirement 10: エラーハンドリング

**User Story:** As a developer, I want clear error messages when extraction fails, so that I can diagnose and fix issues.

#### Acceptance Criteria

1. WHEN the input file does not exist THEN the exstruct-go SHALL return an error indicating file not found
2. WHEN the input file is not a valid xlsx THEN the exstruct-go SHALL return an error indicating invalid format
3. WHEN a specific sheet fails to parse THEN the exstruct-go SHALL log a warning and continue with other sheets
4. WHEN shape extraction fails THEN the exstruct-go SHALL log a warning and return empty shapes list for that sheet

### Requirement 11: 印刷範囲抽出

**User Story:** As a developer, I want to extract print area information from worksheets, so that I can understand the intended print layout.

#### Acceptance Criteria

1. WHEN a worksheet has defined print areas THEN the exstruct-go SHALL extract the print area bounds (r1, c1, r2, c2)
2. WHEN multiple print areas exist on a sheet THEN the exstruct-go SHALL extract all print areas as a list
3. WHEN no print areas are defined THEN the exstruct-go SHALL return an empty print_areas list
4. WHEN mode is "light" THEN the exstruct-go SHALL exclude print_areas from output by default

### Requirement 12: シート単位・印刷範囲単位出力

**User Story:** As a developer, I want to export data per-sheet or per-print-area, so that I can process large workbooks in smaller chunks.

#### Acceptance Criteria

1. WHEN --sheets-dir flag is provided THEN the exstruct-go SHALL write one JSON file per sheet to the specified directory
2. WHEN --print-areas-dir flag is provided THEN the exstruct-go SHALL write one JSON file per print area to the specified directory
3. WHEN exporting per-print-area THEN the exstruct-go SHALL include only rows, shapes, and charts that overlap the print area bounds
