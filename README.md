# ExStruct — Excel Structured Extraction Engine (Fork with OOXML Support)

![Licence: BSD-3-Clause](https://img.shields.io/badge/license-BSD--3--Clause-blue?style=flat-square)

![ExStruct Image](docs/assets/icon.webp)

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

## Features

- **Excel → Structured JSON**: cells, shapes, charts, smartart, table candidates, merged cell ranges, print areas/views, and auto page-break areas per sheet.
- **Output modes**: `light` (cells + table candidates + print areas; no COM, shapes/charts empty), `standard` (texted shapes + arrows, charts, smartart, merged cell ranges, print areas), `verbose` (all shapes with width/height, charts with size, merged cell ranges, print areas). Verbose also emits cell hyperlinks and `colors_map`. Size output is flag-controlled.
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
```

**Note (non-COM environments):** If Excel COM is unavailable, extraction still runs and returns cells + `table_candidates`; `shapes`/`charts` will be empty.

## Output Modes

- **light**: cells + table candidates (no COM needed).
- **standard**: texted shapes + arrows, charts (COM if available), table candidates. Hyperlinks are off unless `include_cell_links=True`.
- **verbose**: all shapes, charts, table_candidates, hyperlinks, and `colors_map`.

## License

BSD-3-Clause. See `LICENSE` for details.

## Acknowledgments

This project is a fork of [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct). We are deeply grateful to the original authors for creating such a well-designed Excel extraction engine with clean architecture and comprehensive documentation. The OOXML parser extension in this fork builds upon their excellent foundation.

## Documentation

- API Reference (GitHub Pages): https://harumiweb.github.io/exstruct/
- JSON Schemas: see `schemas/` (one file per model); regenerate via `python scripts/gen_json_schema.py`.
