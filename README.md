# ExStruct — Excel Structured Extraction Engine (Fork with OOXML Support)

![Licence: BSD-3-Clause](https://img.shields.io/badge/license-BSD--3--Clause-blue?style=flat-square)

![ExStruct Image](/docs/assets/icon.webp)

This is a fork of [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct) with added OOXML parser support for cross-platform shape/chart extraction.

For installation and basic usage, please refer to the [original repository](https://github.com/harumiWeb/exstruct).

[日本版 README](README.ja.md)

## What's New in This Fork

This fork adds a pure-Python OOXML parser that enables shape and chart extraction on **Linux and macOS** without requiring Excel.

### How it works

- **Windows + Excel**: Uses COM API via xlwings (full feature support)
- **Linux / macOS**: Automatically falls back to OOXML parser (no Excel required)
- **Windows without Excel**: Also uses OOXML parser

### Supported features (OOXML)

| Feature | Support |
|---------|---------|
| Shape position (l, t) | ✓ |
| Shape size (w, h) | ✓ (verbose mode) |
| Shape text | ✓ |
| Shape type | ✓ |
| Shape ID assignment | ✓ |
| Connector direction | ✓ |
| Arrow styles | ✓ |
| Connector endpoints (begin_id, end_id) | ✓ |
| Rotation | ✓ |
| Group flattening | ✓ |
| Chart type | ✓ |
| Chart title | ✓ |
| Y-axis title/range | ✓ |
| Series data | ✓ |

### Limitations (OOXML vs COM)

Some features require Excel's calculation engine and cannot be implemented in OOXML:

- Auto-calculated Y-axis range (when set to "auto" in Excel)
- Cell reference resolution for titles/labels
- Conditional formatting evaluation
- Auto page-break calculation
- OLE/embedded objects
- VBA macros

For detailed comparison, see [docs/com-vs-ooxml-implementation.md](docs/com-vs-ooxml-implementation.md).

### Improvement over upstream (without COM)

The original ExStruct was designed with a focus on Windows + Excel environments, providing graceful fallback to cells-only extraction on other platforms. This fork extends that foundation by adding an OOXML parser for cross-platform shape/chart extraction:

| Feature | Original (no COM) | With OOXML Parser |
|---------|-------------------|-------------------|
| Cells | ✓ | ✓ |
| Table candidates | ✓ | ✓ |
| Print areas | ✓ | ✓ |
| Shape extraction | — (fallback) | ✓ |
| Chart extraction | — (fallback) | ✓ |
| Connector relationships | — | ✓ |
| Auto page-breaks | — | — (COM only) |

This extension enables:
- **Flowchart extraction** on Linux/macOS (shapes + connectors with begin_id/end_id)
- **Chart data extraction** without Excel
- **CI/CD and Docker** environments (headless operation)

## Features

- **Excel → Structured JSON**: cells, shapes, charts, table candidates, print areas/views, and auto page-break areas per sheet.
- **Output modes**: `light` (cells + table candidates + print areas; no COM, shapes/charts empty), `standard` (texted shapes + arrows, charts, print areas), `verbose` (all shapes with width/height, charts with size, print areas). Verbose also emits cell hyperlinks and `colors_map`. Size output is flag-controlled.
- **Auto page-break export (COM only)**: capture Excel-computed auto page breaks and write per-area JSON/YAML/TOON when requested (CLI option appears only when COM is available).
- **Formats**: JSON (compact by default, `--pretty` available), YAML, TOON (optional dependencies).
- **Table detection tuning**: adjust heuristics at runtime via API.
- **CLI rendering** (Excel required): optional PDF and per-sheet PNGs.
- **Graceful fallback**: if Excel COM is unavailable, extraction falls back to cells + table candidates without crashing.

## Installation

```bash
pip install exstruct
```

Optional extras:

- YAML: `pip install pyyaml`
- TOON: `pip install python-toon`
- Rendering (PDF/PNG): Excel + `pip install pypdfium2 pillow`
- All extras at once: `pip install exstruct[yaml,toon,render]`

Platform note:

- Full extraction (shapes/charts) targets Windows + Excel (COM via xlwings). On other platforms, use `mode=light` to get cells + `table_candidates`.

## Quick Start (CLI)

```bash
exstruct input.xlsx > output.json          # compact JSON to stdout (default)
exstruct input.xlsx -o out.json --pretty   # pretty JSON to a file
exstruct input.xlsx --format yaml          # YAML (needs pyyaml)
exstruct input.xlsx --format toon          # TOON (needs python-toon)
exstruct input.xlsx --sheets-dir sheets/   # split per sheet in chosen format
exstruct input.xlsx --print-areas-dir areas/  # split per print area (if any)
exstruct input.xlsx --auto-page-breaks-dir auto_areas/  # COM only; option appears when available
exstruct input.xlsx --mode light           # cells + table candidates only
exstruct input.xlsx --pdf --image          # PDF and PNGs (Excel required)
```

Auto page-break exports are available via API and CLI when Excel/COM is available; the CLI exposes `--auto-page-breaks-dir` only in COM-capable environments.

## Quick Start (Python)

```python
from pathlib import Path
from exstruct import extract, export, set_table_detection_params

# Tune table detection (optional)
set_table_detection_params(table_score_threshold=0.3, density_min=0.04)

# Extract with modes: "light", "standard", "verbose"
wb = extract("input.xlsx", mode="standard")
export(wb, Path("out.json"), pretty=False)  # compact JSON

# Model helpers: iterate, index, and serialize directly
first_sheet = wb["Sheet1"]           # __getitem__ access
for name, sheet in wb:               # __iter__ yields (name, SheetData)
    print(name, len(sheet.rows))
wb.save("out.json", pretty=True)     # WorkbookData → file (by extension)
first_sheet.save("sheet.json")       # SheetData → file (by extension)
print(first_sheet.to_yaml())         # YAML text (requires pyyaml)

# ExStructEngine: per-instance options (nested configs)
from exstruct import (
    DestinationOptions,
    ExStructEngine,
    FilterOptions,
    FormatOptions,
    OutputOptions,
    StructOptions,
    export_auto_page_breaks,
)

engine = ExStructEngine(
    options=StructOptions(mode="verbose"),  # verbose includes hyperlinks by default
    output=OutputOptions(
        format=FormatOptions(pretty=True),
        filters=FilterOptions(include_shapes=False),  # drop shapes in output
        destinations=DestinationOptions(sheets_dir=Path("out_sheets")),  # also write per-sheet files
    ),
)
wb2 = engine.extract("input.xlsx")
engine.export(wb2, Path("out_filtered.json"))  # drops shapes via filters

# Enable hyperlinks in other modes
engine_links = ExStructEngine(options=StructOptions(mode="standard", include_cell_links=True))
with_links = engine_links.extract("input.xlsx")

# Export per print area (if print areas exist)
from exstruct import export_print_areas_as
export_print_areas_as(wb, "areas", fmt="json", pretty=True)

# Auto page-break extraction/output (COM only; raises if no auto breaks exist)
engine_auto = ExStructEngine(
    output=OutputOptions(
        destinations=DestinationOptions(auto_page_breaks_dir=Path("auto_areas"))
    )
)
wb_auto = engine_auto.extract("input.xlsx")  # includes SheetData.auto_print_areas
engine_auto.export(wb_auto, Path("out_with_auto.json"))  # also writes auto_areas/*
export_auto_page_breaks(wb_auto, "auto_areas", fmt="json", pretty=True)  # manual writer
```

**Note (non-COM environments):** If Excel COM is unavailable, extraction still runs and returns cells + `table_candidates`; `shapes`/`charts` will be empty.

## Table Detection Tuning

```python
from exstruct import set_table_detection_params

set_table_detection_params(
    table_score_threshold=0.35,  # increase to be stricter
    density_min=0.05,
    coverage_min=0.2,
    min_nonempty_cells=3,
)
```

Use higher thresholds to reduce false positives; lower them if true tables are missed.

## Output Modes

- **light**: cells + table candidates (no COM needed).
- **standard**: texted shapes + arrows, charts (COM if available), table candidates. Hyperlinks are off unless `include_cell_links=True`.
- **verbose**: all shapes (with width/height), charts, table candidates, cell hyperlinks, and `colors_map`.

## Error Handling / Fallbacks

- Excel COM unavailable → falls back to cells + table candidates; shapes/charts empty.
- Shape extraction failure → logs warning, still returns cells + table candidates.
- CLI prints errors to stdout/stderr and returns non-zero on failures.

## Optional Rendering

Requires Excel and `pypdfium2`.

```bash
exstruct input.xlsx --pdf --image --dpi 144
```

Creates `<output>.pdf` and `<output>_images/` PNGs per sheet.

## Benchmark: Excel Structuring Demo

To show how well exstruct can structure Excel, we parse a workbook that combines three elements on one sheet and share an AI reasoning benchmark that uses the JSON output.

- Table (sales data)
- Line chart
- Flowchart built only with shapes

(Screenshot below is the actual sample Excel sheet)
![Sample Excel](/docs/assets/demo_sheet.png)
Sample workbook: `sample/sample.xlsx`

### 1. Input: Excel Sheet Overview

This sample Excel contains:

### ① Table (Sales Data)

| Month  | Product A | Product B | Product C |
| ------ | --------- | --------- | --------- |
| Jan-25 | 120       | 80        | 60        |
| Feb-25 | 135       | 90        | 64        |
| Mar-25 | 150       | 100       | 70        |
| Apr-25 | 170       | 110       | 72        |
| May-25 | 160       | 120       | 75        |
| Jun-25 | 180       | 130       | 80        |

### ② Chart (Line Chart)

- Title: Sales Data
- Series: Product A / Product B / Product C (six months)
- Y axis: 0–200

### ③ Flowchart built with shapes

The sheet includes this flow:

- Start / End
- Format check
- Loop (items remaining?)
- Error handling
- Yes/No decision for sending email

### 2. Output: Structured JSON produced by exstruct (excerpt)

Below is a **shortened JSON output example** from parsing this Excel workbook.

```json
{
  "book_name": "sample.xlsx",
  "sheets": {
    "Sheet1": {
      "rows": [
        {
          "r": 3,
          "c": {
            "1": "月",
            "2": "製品A",
            "3": "製品B",
            "4": "製品C"
          }
        },
        ...
      ],
      "shapes": [
        {
          "id": 1,
          "text": "開始",
          "l": 148,
          "t": 220,
          "type": "AutoShape-FlowchartProcess"
        },
        {
          "id": 2,
          "text": "入力データ読み込み",
          "l": 132,
          "t": 282,
          "type": "AutoShape-FlowchartProcess"
        },
        {
          "l": 193,
          "t": 246,
          "type": "AutoShape-Mixed",
          "begin_arrow_style": 1,
          "end_arrow_style": 2,
          "begin_id": 1,
          "end_id": 2,
          "direction": "N"
        },
        ...
      ],
      "charts": [
        {
          "name": "Chart 1",
          "chart_type": "Line",
          "title": "売上データ",
          "y_axis_range": [
            0.0,
            200.0
          ],
          "series": [
            {
              "name": "製品A",
              "name_range": "Sheet1!$C$3",
              "x_range": "Sheet1!$B$4:$B$9",
              "y_range": "Sheet1!$C$4:$C$9"
            },
            ...
          ],
          "l": 377,
          "t": 25
        }
      ],
      "table_candidates": [
        "B3:E9"
      ]
    }
  }
}
```

### 3. How AI (Copilot / LLM) interprets the JSON

````md
Below is the Markdown reconstruction of the Excel workbook. The table, chart, and flowchart are all represented.

---

## 📊 Sales Data Table

| Month      | Product A | Product B | Product C |
| ---------- | --------- | --------- | --------- |
| 2025-01-01 | 120       | 80        | 60        |
| 2025-02-01 | 135       | 90        | 64        |
| 2025-03-01 | 150       | 100       | 70        |
| 2025-04-01 | 170       | 110       | 72        |
| 2025-05-01 | 160       | 120       | 75        |
| 2025-06-01 | 180       | 130       | 80        |

---

## 📈 Sales Data (Line Chart)

- Chart title: **売上データ (Sales Data)**
- Chart type: Line
- Y-axis range: 0 to 200
- Data series:
  - Product A: 120 → 135 → 150 → 170 → 160 → 180
  - Product B: 80 → 90 → 100 → 110 → 120 → 130
  - Product C: 60 → 64 → 70 → 72 → 75 → 80

---

## 🔄 Process Flow (Mermaid Flowchart)

```mermaid
flowchart TD
    A[Start]
    B[Load input data]
    C{Is format valid?}
    D[Show error]
    E[Process one item]
    F{Items remaining?}
    G[Generate output]
    H{Send email?}
    I[Send email]
    J[Finish]

    A --> B
    B --> C
    C -->|yes| D
    C --> H
    D --> E
    E --> F
    F --> G
    G -->|yes| I
    G -->|no| J
    H --> J
    I --> J
```
````

From this we can see:

**exstruct's JSON is already in a format that AI can read and reason over directly.**

Other LLM inference samples using this library can be found in the following directory:

- [Basic Excel](sample/basic/)
- [Flowchart](sample/flowchart/)
- [Gantt Chart](sample/gantt_chart/)

### 4. Summary

This benchmark confirms exstruct can:

- **Parse tables, charts, and shapes (flowcharts) simultaneously**
- Convert the semantic structure of Excel into JSON
- Let AI/LLMs read that JSON directly and reconstruct the workbook contents

In short, **exstruct = "an engine that converts Excel into a format AI can understand."**

## Notes

- Default JSON is compact to reduce tokens; use `--pretty` or `pretty=True` when readability matters.
- Field `table_candidates` replaces `tables`; adjust downstream consumers accordingly.

## Enterprise Use

ExStruct is used primarily as a **library**, not a service.

- No official support or SLA is provided
- Long-term stability is prioritized over rapid feature growth
- Forking and internal modification are expected in enterprise use

This project is suitable for teams that:

- need transparency over black-box tools
- are comfortable maintaining internal forks if necessary

## Print Areas and Auto Page Breaks (PrintArea / PrintAreaView)

- `SheetData.print_areas` holds print areas (cell coordinates) in light/standard/verbose.
- `SheetData.auto_print_areas` holds Excel COM-computed auto page-break areas when auto page-break extraction is enabled (COM only).
- Use `export_print_areas_as(...)` or CLI `--print-areas-dir` to write one file per print area (nothing is written if none exist).
- Use CLI `--auto-page-breaks-dir` (COM only), `DestinationOptions.auto_page_breaks_dir` (preferred), or `export_auto_page_breaks(...)` to write per-auto-page-break files; the API raises `ValueError` if no auto page breaks exist.
- `PrintAreaView` includes rows and table candidates inside the area, plus shapes/charts that overlap the area (size-less shapes are treated as points). `normalize=True` rebases row/col indices to the area origin.

## License

BSD-3-Clause. See `LICENSE` for details.

## Acknowledgments

This project is a fork of [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct). We are deeply grateful to the original authors for creating such a well-designed Excel extraction engine with clean architecture and comprehensive documentation. The OOXML parser extension in this fork builds upon their excellent foundation.

## Documentation

- API Reference (GitHub Pages): https://harumiweb.github.io/exstruct/
- JSON Schemas: see `schemas/` (one file per model); regenerate via `python scripts/gen_json_schema.py`.
