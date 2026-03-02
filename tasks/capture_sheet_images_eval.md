# capture_sheet_images Stability Evaluation

## Purpose

Define a repeatable evaluation set and procedure for MCP `exstruct_capture_sheet_images` (Experimental) before GA.

## Runtime Baseline

- OS: Windows (Excel desktop available)
- MCP runtime defaults:
  - `EXSTRUCT_RENDER_SUBPROCESS=0`
  - `EXSTRUCT_MCP_CAPTURE_SHEET_IMAGES_TIMEOUT_SEC=180`

## Representative Workbook Set

Use the following local files:

1. `sample/basic/sample.xlsx`
2. `sample/formula/formula.xlsx`
3. `sample/flowchart/sample-shape-connector.xlsx`
4. `sample/smartart/sample_smartart.xlsx`
5. `sample/forms_with_many_merged_cells/ja_general_form/ja_form.xlsx`
6. `sample/forms_with_many_merged_cells/en_form_sf425/sample.xlsx`
7. `tests/assets/multiple_print_ranges_4sheets.xlsx`

## Evaluation Cases

For each workbook, run all applicable cases:

1. Full workbook export (`sheet`/`range` omitted)
2. Single sheet export (`sheet` only)
3. Minimal range export (`sheet` + `range=A1:A1`)
4. Sheet-qualified range export (`'Sheet Name'!A1:B2` when sheet name contains spaces)

## Procedure

1. Start MCP server with production-like env.
2. For each case, call `exstruct_capture_sheet_images` and record:
   - start/end timestamp
   - elapsed seconds
   - success/failure
   - error message (if failed)
   - generated image count
3. Repeat the whole matrix 3 times.
4. Compute:
   - total runs
   - success rate
   - p50/p95 elapsed time
   - failure type breakdown

## Pass/Fail Gate (GA)

- Success rate >= 99% across all runs.
- No opaque timeout failures; all failures include actionable messages.
- No monotonic memory growth that indicates unacceptable leak behavior during repeated runs.

## Reporting Template

- Date:
- Environment:
- Total runs:
- Success rate:
- p50/p95 elapsed:
- Failures by type:
- Notes / next actions:
