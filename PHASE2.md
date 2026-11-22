# Phase 2 architecture

## Layers

| Layer | Responsibility |
| --- | --- |
| `src/lib/workbook/` | Document model (multi-sheet), migration, filter, find/replace, fill, validation |
| `src/lib/formula/` | HyperFormula wrapper, recalculation, structural ops |
| `src/lib/excel/workbook-io.ts` | Multi-sheet `.xlsx` import/export (ExcelJS) |
| `src/components/spreadsheet/` | UI: tabs, formula bar, toolbar, virtualized grid |
| API `cells` JSON column | Persists full `WorkbookDocument` (v2) |

## Formula evaluation

[HyperFormula](https://hyperformula.handsontable.com/) (GPL) provides Excel-compatible parsing, dependency tracking, circular reference detection, and cross-sheet references. Evaluation runs **in the browser**; the server stores raw inputs, formats, and sheet structure.

## Persistence

- PATCH accepts `workbook` (v2). Legacy `cells` maps are upgraded on read/write.
- Debounced save (~500ms) with in-flight request coalescing.
- `saveGeneration` prevents stale responses from overwriting newer edits.

## Implemented (Phase 2)

- Multi-sheet workbook model + tabs (add/rename/delete/duplicate/reorder)
- HyperFormula engine (SUM, IF, VLOOKUP, COUNTIF, cross-sheet, circular errors, row/col structural ops)
- Formula bar + name box + range selection + copy/paste/cut + Ctrl+A + undo/redo
- Toolbar: formatting, borders, row/col insert/delete/hide, sort, merge, freeze, zoom, gridlines, autofit
- Auto-filter (toggle, per-column contains filter, hidden rows)
- Find & replace dialog (replace all)
- Fill handle (drag to extend selection + formula ref adjustment in `fill.ts`)
- Data validation on entry; list validation uses in-cell dropdown when editing
- Charts (bar/line/pie/area via Recharts from selected range)
- ExcelJS import/export (multi-sheet, formulas, basic styles, merges, column widths)
- Workbook v2 persistence + legacy v1 migration
- Row virtualization (`@tanstack/react-virtual`); frozen rows split + sticky freeze columns/rows
- Legacy v1 grid removed (`spreadsheet-editor.tsx`, `spreadsheet-grid.tsx`)

## Partial / not Excel-complete

| Area | Status |
| --- | --- |
| Full Excel function parity | HyperFormula subset; array/dynamic functions may differ |
| Filter UI | Column prompt filter; no rich operator UI or multi-criteria |
| Find | Replace-all only; no step-through “find next” |
| Column virtualization | Rows virtualized; all visible columns still rendered |
| Conditional formatting | Not implemented |
| Excel macros, pivot, images, full validation xlsx | Unsupported |
| Engine rebuild | Workbook changes rebuild HyperFormula (large sheets = cost) |

See `README.md` for run/test instructions.