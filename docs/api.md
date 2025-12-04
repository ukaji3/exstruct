# API Reference

本ページは Python API の概要を示します。すべて `exstruct` パッケージからインポート可能です。

## extract(file_path, mode="standard")

Excel ワークブックを読み取り、`WorkbookData` を返します。

- `file_path`: `str | Path` 対象 Excel。
- `mode`: `"light" | "standard" | "verbose"`
  - `light`: セル＋テーブル候補のみ（COM 不要）。
  - `standard`: テキスト付き図形＋矢印、チャート（COM があれば取得）。
  - `verbose`: 全図形（幅・高さ付き）、チャート、テーブル候補。
- COM が無い場合は自動でセル＋テーブル候補にフォールバックします。

## export(data, path, fmt=None, \*, pretty=False, indent=None)

`WorkbookData` をファイルに書き出します。

- `fmt`: `"json" | "yaml" | "yml" | "toon"`（既定は拡張子または json）
- `pretty`: `True` でインデント 2（または `indent` 指定値）で整形
- `indent`: 整形時のインデント幅（既定は 2）
- 依存: YAML は `pyyaml`, TOON は `python-toon` が必要

## export_sheets_as(data, dir_path, fmt="json", \*, pretty=False, indent=None)

各シートを個別ファイルに書き出します（`book_name` と `sheet` を含む）。

- `fmt`: `"json" | "yaml" | "yml" | "toon"`
- `pretty` / `indent`: JSON の整形オプション
- 戻り値: シート名 → 出力パスの辞書

## process_excel(file_path, output_path, out_fmt="json", image=False, pdf=False, dpi=72, mode="standard", \*, pretty=False, indent=None)

CLI 相当のヘルパー。抽出＋書き出し＋任意で PDF/PNG を生成します。

- `image`: PNG をシートごとに出力（Excel + `pypdfium2` が必要）
- `pdf`: PDF を出力（Excel が必要）
- `mode`: 出力モード（前述）
- `pretty`/`indent`: JSON 整形オプション

## set_table_detection_params(...)

テーブル検出ヒューリスティックを動的に調整します（省略した値は現行設定を維持）。

- `table_score_threshold: float | None`
- `density_min: float | None`
- `coverage_min: float | None`
- `min_nonempty_cells: int | None`
  値を上げると厳しくなり誤検知が減少、下げると検出漏れが減ります。

## render（Excel が必要）

`export_pdf(file_path, pdf_path)` / `export_sheet_images(file_path, out_dir, dpi=144)`  
Excel COM と `pypdfium2` が必要です。COM が無い環境では例外となります。

## Data Models（主なフィールド）

- `WorkbookData`: `book_name`, `sheets: Dict[str, SheetData]`
- `SheetData`: `rows`, `shapes`, `charts`, `table_candidates`
- `Shape`: `text`, `l/t/w/h`, `type`, `angle_deg`, `direction` ほか
- `Chart`: `chart_type`, `title`, `y_axis_title`, `y_axis_range`, `series`, `error`
- `CellRow`: `r`（行番号 1-based）, `c`（列番号文字列 → 値）

## Optional Dependencies

- YAML: `pip install pyyaml`
- TOON: `pip install python-toon`
- Rendering (PDF/PNG): Excel + `pip install pypdfium2`

## エラーハンドリング

- COM 不在・失敗時はセル＋テーブル候補にフォールバック（図形・チャートは空）。
- エクスポートの失敗や不正なフォーマット指定は `ValueError` / `RuntimeError` を送出。
