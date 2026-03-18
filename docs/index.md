# ExStruct Documentation

![ExStruct Image](assets/logo.png)

Welcome to the ExStruct docs. ExStruct has two primary capabilities:
extracting structured data from Excel workbooks and editing workbooks through a
public patch-based core.

## Choose an Interface

| Use case | Start here | What is canonical |
| --- | --- | --- |
| Embed extraction in Python code | [**API Reference**](api.md) | Extraction keeps the top-level Python API. ExStruct's editing API exists, but ordinary Python workbook editing is usually better served by `openpyxl` / `xlwings`. |
| Run local shell or agent workflows | [**CLI Guide**](cli.md) | Editing commands such as `exstruct patch` and `exstruct make` are the canonical operational interface. |
| Run sandboxed or host-managed tools | [**MCP Server**](mcp.md) | MCP is the integration / compatibility layer and owns `PathPolicy`, transport, and artifact behavior. |
| Understand the product direction | [**Concept / Why ExStruct?**](concept.md) | Extraction fidelity, layout semantics, and downstream AI use cases. |
| Check licensing details | [**Corporate Usage License Guide**](license-guide.md) | Commercial-use guidance. |

## At a Glance

- Extract structured JSON (default), YAML, and TOON from Excel workbooks.
- Edit workbooks primarily through the JSON-first editing CLI.
- Use `dry_run`, warnings, diff output, and inverse-ops support for safer edit flows.
- Keep MCP for restricted hosts that need path controls and tool transport.

## Quick Links

- Install: `pip install exstruct`
- Extraction CLI: `exstruct input.xlsx --pretty --mode standard`
- Editing CLI: `exstruct patch --input book.xlsx --ops ops.json --dry-run`
