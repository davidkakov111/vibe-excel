# Architecture notes

## Data model

- **PostgreSQL** stores users and spreadsheet metadata. Cell content is a versioned **workbook JSON** document (`Spreadsheet.cells`, v2) — not a live `.xlsx` file.
- **Import/export**: `.xlsx` is converted at upload and regenerated on download via `src/lib/excel/workbook-io.ts` (ExcelJS).

## Saving (editor)

- Edits update in-memory workbook state and sync HyperFormula immediately.
- Saves are **debounced (500ms)** and always persist **`latestWorkbook`**, not a stale snapshot from when the timer was scheduled.
- If edits occur during an in-flight save, a follow-up save runs after success.
- Failed saves mark the document dirty; the next edit schedules another attempt (no infinite retry loop).

## Excel fidelity (current)

Supported reasonably well: cell values, basic formulas, solid fills, fonts, alignment, borders (subset), column widths, row heights, merged ranges, multi-sheet names.

Not supported: Excel Tables (ListObjects), native chart objects, conditional formatting, pivot tables, theme-only colors without ARGB, macros.

See `src/lib/excel/workbook-io.test.ts` for regression coverage on import/export.