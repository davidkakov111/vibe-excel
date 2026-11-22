import { cellKey, parseCellKey } from "@/lib/spreadsheet-types";
import type { CellData, MergeRange, Worksheet } from "@/lib/workbook/types";

function shiftRows(
  map: Record<string, number>,
  atRow: number,
  delta: number,
): Record<string, number> {
  const next: Record<string, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const row = Number(key);
    next[String(row >= atRow ? row + delta : row)] = value;
  }
  return next;
}

function shiftCols(
  map: Record<string, number>,
  atCol: number,
  delta: number,
): Record<string, number> {
  const next: Record<string, number> = {};
  for (const [key, value] of Object.entries(map)) {
    const col = Number(key);
    next[String(col >= atCol ? col + delta : col)] = value;
  }
  return next;
}

function remapCells(
  cells: Record<string, CellData>,
  rowMapper: (row: number) => number | null,
  colMapper: (col: number) => number | null,
): Record<string, CellData> {
  const next: Record<string, CellData> = {};
  for (const [key, data] of Object.entries(cells)) {
    const parsed = parseCellKey(key);
    if (!parsed) continue;
    const row = rowMapper(parsed.row);
    const col = colMapper(parsed.col);
    if (row === null || col === null) continue;
    next[cellKey(row, col)] = data;
  }
  return next;
}

function shiftMerges(
  merges: MergeRange[],
  atRow: number,
  rowDelta: number,
  atCol: number,
  colDelta: number,
): MergeRange[] {
  return merges.map((merge) => ({
    startRow: merge.startRow >= atRow ? merge.startRow + rowDelta : merge.startRow,
    endRow: merge.endRow >= atRow ? merge.endRow + rowDelta : merge.endRow,
    startCol: merge.startCol >= atCol ? merge.startCol + colDelta : merge.startCol,
    endCol: merge.endCol >= atCol ? merge.endCol + colDelta : merge.endCol,
  }));
}

export function insertRows(sheet: Worksheet, atRow: number, count: number): Worksheet {
  return {
    ...sheet,
    rowCount: sheet.rowCount + count,
    cells: remapCells(
      sheet.cells,
      (row) => (row >= atRow ? row + count : row),
      (col) => col,
    ),
    rowHeights: shiftRows(sheet.rowHeights, atRow, count),
    hiddenRows: sheet.hiddenRows.map((r) => (r >= atRow ? r + count : r)),
    merges: shiftMerges(sheet.merges, atRow, count, 0, 0),
  };
}

export function deleteRows(
  sheet: Worksheet,
  atRow: number,
  count: number,
): Worksheet {
  return {
    ...sheet,
    rowCount: Math.max(1, sheet.rowCount - count),
    cells: remapCells(
      sheet.cells,
      (row) => {
        if (row >= atRow && row < atRow + count) return null;
        if (row >= atRow + count) return row - count;
        return row;
      },
      (col) => col,
    ),
    rowHeights: shiftRows(
      Object.fromEntries(
        Object.entries(sheet.rowHeights).filter(([r]) => {
          const row = Number(r);
          return row < atRow || row >= atRow + count;
        }),
      ),
      atRow + count,
      -count,
    ),
    hiddenRows: sheet.hiddenRows
      .filter((r) => r < atRow || r >= atRow + count)
      .map((r) => (r >= atRow + count ? r - count : r)),
    merges: shiftMerges(sheet.merges, atRow + count, -count, 0, 0),
  };
}

export function insertColumns(sheet: Worksheet, atCol: number, count: number): Worksheet {
  return {
    ...sheet,
    colCount: sheet.colCount + count,
    cells: remapCells(
      sheet.cells,
      (row) => row,
      (col) => (col >= atCol ? col + count : col),
    ),
    colWidths: shiftCols(sheet.colWidths, atCol, count),
    hiddenCols: sheet.hiddenCols.map((c) => (c >= atCol ? c + count : c)),
    merges: shiftMerges(sheet.merges, 0, 0, atCol, count),
  };
}

export function deleteColumns(
  sheet: Worksheet,
  atCol: number,
  count: number,
): Worksheet {
  return {
    ...sheet,
    colCount: Math.max(1, sheet.colCount - count),
    cells: remapCells(
      sheet.cells,
      (row) => row,
      (col) => {
        if (col >= atCol && col < atCol + count) return null;
        if (col >= atCol + count) return col - count;
        return col;
      },
    ),
    colWidths: shiftCols(
      Object.fromEntries(
        Object.entries(sheet.colWidths).filter(([c]) => {
          const col = Number(c);
          return col < atCol || col >= atCol + count;
        }),
      ),
      atCol + count,
      -count,
    ),
    hiddenCols: sheet.hiddenCols
      .filter((c) => c < atCol || c >= atCol + count)
      .map((c) => (c >= atCol + count ? c - count : c)),
    merges: shiftMerges(sheet.merges, 0, 0, atCol + count, -count),
  };
}