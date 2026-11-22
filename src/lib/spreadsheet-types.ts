export type CellMap = Record<string, string>;

export interface SpreadsheetContent {
  cells: CellMap;
  rowCount: number;
  colCount: number;
}

export function cellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

export function parseCellKey(key: string): { row: number; col: number } | null {
  const [rowPart, colPart] = key.split(":");
  const row = Number(rowPart);
  const col = Number(colPart);
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0) {
    return null;
  }
  return { row, col };
}

/** 0-based column index to Excel-style label (A, B, …, AA). */
export function indexToColumnLabel(colIndex: number): string {
  let label = "";
  let n = colIndex + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

export function emptySpreadsheet(
  rowCount = 100,
  colCount = 26,
): SpreadsheetContent {
  return { cells: {}, rowCount, colCount };
}

export function normalizeSpreadsheetContent(
  data: unknown,
  fallback: SpreadsheetContent,
): SpreadsheetContent {
  if (!data || typeof data !== "object") {
    return fallback;
  }
  const raw = data as Record<string, unknown>;
  const cells =
    raw.cells && typeof raw.cells === "object" && !Array.isArray(raw.cells)
      ? (raw.cells as CellMap)
      : fallback.cells;
  const rowCount =
    typeof raw.rowCount === "number" && raw.rowCount > 0
      ? Math.floor(raw.rowCount)
      : fallback.rowCount;
  const colCount =
    typeof raw.colCount === "number" && raw.colCount > 0
      ? Math.floor(raw.colCount)
      : fallback.colCount;
  return { cells, rowCount, colCount };
}