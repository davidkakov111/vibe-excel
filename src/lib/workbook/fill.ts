import type { GridSelection } from "@/lib/workbook/selection";
import type { CellData, Worksheet } from "@/lib/workbook/types";
import { getCellInput } from "@/lib/workbook/types";
import { cellKey } from "@/lib/spreadsheet-types";
import { isFormulaInput } from "@/lib/formula/engine";
import { indexToColumnLabel } from "@/lib/spreadsheet-types";

function adjustFormulaReferences(
  formula: string,
  rowDelta: number,
  colDelta: number,
): string {
  return formula.replace(
    /(\$?)([A-Z]+)(\$?)(\d+)/g,
    (_match, colAbs, colLetters, rowAbs, rowNum) => {
      let col = 0;
      for (let i = 0; i < colLetters.length; i++) {
        col = col * 26 + (colLetters.charCodeAt(i) - 64);
      }
      col -= 1;
      let row = Number(rowNum) - 1;
      if (!colAbs) col += colDelta;
      if (!rowAbs) row += rowDelta;
      if (row < 0) row = 0;
      if (col < 0) col = 0;
      const colPrefix = colAbs ? "$" : "";
      const rowPrefix = rowAbs ? "$" : "";
      return `${colPrefix}${indexToColumnLabel(col)}${rowPrefix}${row + 1}`;
    },
  );
}

function positiveMod(n: number, mod: number): number {
  return ((n % mod) + mod) % mod;
}

function copyCellForFill(
  sheet: Worksheet,
  srcRow: number,
  srcCol: number,
  destRow: number,
  destCol: number,
): CellData | undefined {
  const src = sheet.cells[cellKey(srcRow, srcCol)];
  const raw = getCellInput(sheet, srcRow, srcCol);
  const rowDelta = destRow - srcRow;
  const colDelta = destCol - srcCol;
  const input = isFormulaInput(raw)
    ? adjustFormulaReferences(raw, rowDelta, colDelta)
    : raw;
  const hasInput = input.trim() !== "";
  const hasFormat = Boolean(src?.format);
  const hasComment = Boolean(src?.comment);
  const hasValidation = Boolean(src?.validation);
  if (!hasInput && !hasFormat && !hasComment && !hasValidation) {
    return undefined;
  }
  const next: CellData = {};
  if (hasInput) next.input = input;
  if (src?.format) next.format = structuredClone(src.format);
  if (src?.comment) next.comment = structuredClone(src.comment);
  if (src?.validation) next.validation = structuredClone(src.validation);
  return next;
}

/** Tile source selection across the rectangle union of source and drag target (Excel fill). */
export function fillSelection(
  sheet: Worksheet,
  source: GridSelection,
  target: GridSelection,
): Worksheet {
  const srcStartRow = Math.min(source.row, source.endRow);
  const srcEndRow = Math.max(source.row, source.endRow);
  const srcStartCol = Math.min(source.col, source.endCol);
  const srcEndCol = Math.max(source.col, source.endCol);

  const tgtStartRow = Math.min(target.row, target.endRow);
  const tgtEndRow = Math.max(target.row, target.endRow);
  const tgtStartCol = Math.min(target.col, target.endCol);
  const tgtEndCol = Math.max(target.col, target.endCol);

  const fillStartRow = Math.min(srcStartRow, tgtStartRow);
  const fillEndRow = Math.max(srcEndRow, tgtEndRow);
  const fillStartCol = Math.min(srcStartCol, tgtStartCol);
  const fillEndCol = Math.max(srcEndCol, tgtEndCol);

  const sourceHeight = srcEndRow - srcStartRow + 1;
  const sourceWidth = srcEndCol - srcStartCol + 1;

  const nextCells = { ...sheet.cells };

  for (let row = fillStartRow; row <= fillEndRow; row++) {
    for (let col = fillStartCol; col <= fillEndCol; col++) {
      const insideSource =
        row >= srcStartRow &&
        row <= srcEndRow &&
        col >= srcStartCol &&
        col <= srcEndCol;
      if (insideSource) continue;

      const srcRow =
        srcStartRow + positiveMod(row - srcStartRow, sourceHeight);
      const srcCol =
        srcStartCol + positiveMod(col - srcStartCol, sourceWidth);
      const key = cellKey(row, col);
      const copied = copyCellForFill(sheet, srcRow, srcCol, row, col);
      if (copied) nextCells[key] = copied;
      else delete nextCells[key];
    }
  }

  return { ...sheet, cells: nextCells };
}

export function autofitColumnWidth(
  sheet: Worksheet,
  col: number,
  measure: (row: number, col: number) => string,
): number {
  let maxLen = 8;
  for (let row = 0; row < sheet.rowCount; row++) {
    const text = measure(row, col);
    maxLen = Math.max(maxLen, text.length + 2);
  }
  return Math.min(360, Math.max(48, maxLen * 8));
}