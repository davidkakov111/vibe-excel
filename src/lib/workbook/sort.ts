import { current, isDraft } from "immer";
import { cellKey } from "@/lib/spreadsheet-types";
import type { CellData, Worksheet } from "@/lib/workbook/types";

function plainWorksheet(sheet: Worksheet): Worksheet {
  return isDraft(sheet) ? current(sheet) : sheet;
}

function copyCellData(data: CellData): CellData {
  const plain = isDraft(data) ? current(data) : data;
  return structuredClone(plain);
}

function compareValues(a: string, b: string, ascending: boolean): number {
  const aNum = Number(a);
  const bNum = Number(b);
  const bothNumeric =
    a.trim() !== "" &&
    b.trim() !== "" &&
    !Number.isNaN(aNum) &&
    !Number.isNaN(bNum);
  let result: number;
  if (bothNumeric) {
    result = aNum - bNum;
  } else {
    result = a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
  }
  return ascending ? result : -result;
}

export function sortWorksheetRange(
  sheet: Worksheet,
  startRow: number,
  endRow: number,
  startCol: number,
  endCol: number,
  sortCol: number,
  ascending: boolean,
  getDisplay?: (row: number, col: number) => string,
): Worksheet {
  const base = plainWorksheet(sheet);
  if (endRow <= startRow) return base;
  const rowIndices = Array.from(
    { length: endRow - startRow + 1 },
    (_, i) => startRow + i,
  );

  const sortKeys = rowIndices.map((row) => {
    const input = base.cells[cellKey(row, sortCol)]?.input ?? "";
    const display = getDisplay?.(row, sortCol) ?? input;
    const value = input.startsWith("=") ? display : input || display;
    return { row, value };
  });

  sortKeys.sort((a, b) => compareValues(a.value, b.value, ascending));
  const order = sortKeys.map((item) => item.row);

  const sourceRows = new Map<number, Map<number, CellData>>();
  for (const row of rowIndices) {
    const cols = new Map<number, CellData>();
    for (let col = startCol; col <= endCol; col++) {
      const data = base.cells[cellKey(row, col)];
      if (data) cols.set(col, data);
    }
    sourceRows.set(row, cols);
  }

  const nextCells = { ...base.cells };
  for (let i = 0; i < order.length; i++) {
    const targetRow = startRow + i;
    const sourceRow = order[i]!;
    for (let col = startCol; col <= endCol; col++) {
      const key = cellKey(targetRow, col);
      const source = sourceRows.get(sourceRow)?.get(col);
      if (source) {
        nextCells[key] = copyCellData(source);
      } else {
        delete nextCells[key];
      }
    }
  }

  return { ...base, cells: nextCells };
}