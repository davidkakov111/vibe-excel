import { parseA1 } from "@/lib/formula/engine";
import { indexToColumnLabel } from "@/lib/spreadsheet-types";

export interface A1Range {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export function parseA1Range(ref: string): A1Range | null {
  const clean = ref.replace(/\$/g, "").trim();
  const [startRef, endRef] = clean.includes(":") ? clean.split(":") : [clean, clean];
  if (!startRef) return null;
  const start = parseA1(startRef);
  const end = parseA1(endRef ?? startRef);
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

export function cellInRange(
  row: number,
  col: number,
  range: A1Range,
): boolean {
  return (
    row >= range.startRow &&
    row <= range.endRow &&
    col >= range.startCol &&
    col <= range.endCol
  );
}

export function rangeToA1(range: A1Range): string {
  const start = `${indexToColumnLabel(range.startCol)}${range.startRow + 1}`;
  const end = `${indexToColumnLabel(range.endCol)}${range.endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}