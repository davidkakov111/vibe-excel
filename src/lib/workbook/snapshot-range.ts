import { cellKey } from "@/lib/spreadsheet-types";
import type { CellData, Worksheet } from "@/lib/workbook/types";

export function snapshotRangeCells(
  sheet: Worksheet,
  startRow: number,
  endRow: number,
  startCol: number,
  endCol: number,
): Record<string, CellData | null> {
  const cells: Record<string, CellData | null> = {};
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const key = cellKey(row, col);
      const data = sheet.cells[key];
      cells[key] = data ? structuredClone(data) : null;
    }
  }
  return cells;
}