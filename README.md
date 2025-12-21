# ExStruct — Excel Structured Extraction Engine (Fork with OOXML Support)

![Licence: BSD-3-Clause](https://img.shields.io/badge/license-BSD--3--Clause-blue?style=flat-square)

![ExStruct Image](/docs/assets/icon.webp)

This is a fork of [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct) with added OOXML parser support for cross-platform shape/chart extraction.

For installation and basic usage, please refer to the [original repository](https://github.com/harumiWeb/exstruct).

[日本版README](README.ja.md)

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

## License

BSD-3-Clause. See `LICENSE` for details.

## Acknowledgments

This project is a fork of [harumiWeb/exstruct](https://github.com/harumiWeb/exstruct). We are deeply grateful to the original authors for creating such a well-designed Excel extraction engine with clean architecture and comprehensive documentation. The OOXML parser extension in this fork builds upon their excellent foundation.

## Documentation

- API Reference (GitHub Pages): https://harumiweb.github.io/exstruct/
- JSON Schemas: see `schemas/` (one file per model); regenerate via `python scripts/gen_json_schema.py`.
