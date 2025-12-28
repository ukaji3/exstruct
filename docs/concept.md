# Concept / Why ExStruct?

## What problem does ExStruct solve?

Real-world Excel files contain much more than cell values:

- visually-crafted tables made with borders
- shapes, callouts, grouped objects
- charts with their series, axes, labels, and data ranges
- merged cells and layout structures
- spatial relationships that carry semantic meaning

All of this is essential for understanding the actual information encoded in an Excel workbook.
However, most Python libraries (openpyxl, pandas, etc.) can only access a small portion of this data.

This leads to several common issues in automation, documentation processing, and RAG pipelines.

---

## The “Invisible Structure” Problem of Excel

### 1. Cells alone do not capture the real meaning

Typical libraries can read cells, but not the richer objects that make Excel documents highly expressive.

| Element                             | Available in typical libraries? | Available in ExStruct? |
| ----------------------------------- | ------------------------------- | ---------------------- |
| Cell values                         | ○                               | ○                      |
| Shapes (text, position, type)       | ×                               | ◎                      |
| Grouped shapes                      | ×                               | ◎                      |
| Chart series / labels / axis ranges | ×〜△                            | ◎                      |
| Heuristic table detection           | ×                               | ◎                      |

Most business documents rely on these structures to convey meaning.
Without them, AI systems miss critical context.

---

### 2. LLMs and RAG systems struggle with Excel’s “format variability”

Excel files often contain structural irregularities:

- tables made only with borders (not actual Table objects)
- inconsistent row/column spacing
- explanatory shapes placed near cells
- charts referencing off-sheet ranges
- camera-tool snapshots showing visualized cell values

When given raw cell data or text-only extraction, LLMs lose the relational and contextual meaning.

ExStruct exists to systematically expose these hidden relationships.

---

### 3. Without Excel installed, high-fidelity extraction is nearly impossible

OpenXML parsing alone cannot reliably retrieve:

- shape positions or grouping
- chart metadata and axis structures
- camera-tool references
- layout-level semantics

In practice, many enterprise environments do have Windows + Excel installed.
ExStruct embraces this environment to provide maximum information extraction while still offering a fallback mode for non-Excel environments.

---

## ExStruct Concept

> “**Convert Excel workbooks into structured, machine-readable JSON that preserves their semantic meaning.**”

ExStruct is designed specifically for extracting structural information for AI systems, automation pipelines, and document analysis.

Key Features

1. Full extraction of shapes, groups, arrows, callouts, and positions
2. Chart metadata: series, values, labels, axis titles, ranges
3. Automatic detection of “visual tables” from borders and density
4. Maximum fidelity when Excel is available; functional fallback when not
5. Optimized output modes (light / standard / verbose) for RAG usage
6. Multiple export formats: JSON, YAML, TOON

---

## Why ExStruct?

### ✔ Excel is one of the largest “semantic black boxes” in enterprise systems

- Business-critical documents are frequently Excel-based:
- inspection checklists
- QC diagrams and cause–effect charts
- SOP manuals
- analysis sheets
- reports with charts
- specification documents with annotations

These files combine **text + layout + shapes + charts**, forming a rich structure that typical parsers cannot represent.

For RAG and AI systems, this missing structure becomes a major bottleneck.

---

## What ExStruct Provides

### 1. A structured, LLM-friendly JSON representation

ExStruct outputs a unified structure containing:

- cells, rows, and sheets
- shapes, arrows, and SmartArt nodes (nested)
- chart series and metadata
- automatically detected table candidates
- layout geometry (positions, sizes)

LLMs can reason over this representation far more effectively than raw text.

---

### 2. A programmatically analyzable view of Excel documents

By converting layout and object information to JSON, ExStruct unlocks new workflows:

- load tables into pandas
- reconstruct diagrams or charts outside Excel
- build searchable document repositories
- display Excel content in web UIs
- convert Excel documents to Markdown or other formats

### 3. A ready-to-use foundation for RAG systems and document automation

Once extracted, the workflow becomes:

```md
Excel → ExStruct → JSON → Vector DB → RAG → Answer Generation
```

Previously, teams needed to build custom extraction logic for each document type.
ExStruct provides a general solution that handles both standard Excel features and complex layouts.

---

## Use Cases

### ✔ RAG for Excel manuals

LLMs can reference both text and shape-based information (flows, diagrams, callouts).

### ✔ Automated extraction of inspection/operation checklists

Visual tables become machine-readable through ExStruct’s detection logic.

### ✔ Structural extraction of QC diagrams / fishbone charts

Positions + text allow downstream tools or AI to reconstruct the logic.

### ✔ Displaying Excel files in web applications

JSON-based layout makes frontend rendering feasible without Excel.

### ✔ Automated reporting and analytics

Chart series can be re-plotted or transformed into dashboards.

---

## Summary

ExStruct is built to:

- reveal the hidden structural elements in Excel files
- expose them through a consistent JSON representation
- enable AI and automation systems to understand Excel documents qualitatively
- support both maximum-fidelity (Excel-installed) and fallback (pure Python) extraction
- facilitate RAG pipelines, document analysis, and enterprise automation

It is not just an Excel parser—
**it is a semantic extraction engine for the most commonly used business document format in the world.**
