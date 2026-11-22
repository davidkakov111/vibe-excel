import type ExcelJS from "exceljs";

/** Normalize an ExcelJS cell value to our stored input string. */
export function excelCellToInput(cell: ExcelJS.Cell): string | undefined {
  const raw = cell.value;
  if (raw === null || raw === undefined) {
    if (cell.formula) return `=${cell.formula}`;
    return undefined;
  }
  if (typeof raw === "object") {
    if ("formula" in raw && typeof raw.formula === "string") {
      return `=${raw.formula}`;
    }
    if ("richText" in raw && Array.isArray(raw.richText)) {
      return raw.richText.map((part) => part.text ?? "").join("");
    }
    if (raw instanceof Date) {
      return raw.toISOString().slice(0, 10);
    }
    if ("error" in raw) {
      return String(raw.error);
    }
  }
  if (typeof raw === "boolean") return raw ? "TRUE" : "FALSE";
  if (typeof raw === "number") return String(raw);
  return String(raw);
}

export function hasExcelCellContent(cell: ExcelJS.Cell): boolean {
  return (
    excelCellToInput(cell) !== undefined ||
    Boolean(cell.font) ||
    Boolean(cell.fill) ||
    Boolean(cell.border) ||
    Boolean(cell.alignment)
  );
}