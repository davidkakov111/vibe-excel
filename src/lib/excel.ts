import * as XLSX from "xlsx";
import {
  cellKey,
  emptySpreadsheet,
  type SpreadsheetContent,
} from "@/lib/spreadsheet-types";

function sheetToContent(sheet: XLSX.WorkSheet): SpreadsheetContent {
  const ref = sheet["!ref"];
  if (!ref) {
    return emptySpreadsheet();
  }

  const range = XLSX.utils.decode_range(ref);
  const cells: SpreadsheetContent["cells"] = {};

  for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address];
      if (!cell || cell.v === undefined || cell.v === null) {
        continue;
      }
      const value =
        typeof cell.v === "number" || typeof cell.v === "boolean"
          ? String(cell.v)
          : String(cell.v).trim();
      if (value !== "") {
        cells[cellKey(row, col)] = value;
      }
    }
  }

  const rowCount = Math.max(range.e.r + 1, 100);
  const colCount = Math.max(range.e.c + 1, 26);
  return { cells, rowCount, colCount };
}

export function parseExcelBuffer(buffer: Buffer): SpreadsheetContent {
  const workbook = XLSX.read(buffer, { type: "buffer", cellDates: false });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    return emptySpreadsheet();
  }
  return sheetToContent(workbook.Sheets[sheetName]);
}

export function spreadsheetToExcelBuffer(
  content: SpreadsheetContent,
): Buffer {
  const workbook = XLSX.utils.book_new();
  const cellRowIndices = Object.keys(content.cells).map((key) =>
    Number(key.split(":")[0]),
  );
  const cellColIndices = Object.keys(content.cells).map((key) =>
    Number(key.split(":")[1]),
  );
  const maxRow = Math.max(content.rowCount - 1, ...cellRowIndices, 0);
  const maxCol = Math.max(content.colCount - 1, ...cellColIndices, 0);

  const rows: (string | number)[][] = [];
  for (let row = 0; row <= maxRow; row++) {
    const rowValues: (string | number)[] = [];
    for (let col = 0; col <= maxCol; col++) {
      const raw = content.cells[cellKey(row, col)];
      if (raw === undefined || raw === "") {
        rowValues.push("");
        continue;
      }
      const asNumber = Number(raw);
      rowValues.push(
        raw.trim() !== "" && !Number.isNaN(asNumber) && /^-?\d+(\.\d+)?$/.test(raw.trim())
          ? asNumber
          : raw,
      );
    }
    rows.push(rowValues);
  }

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  return Buffer.from(
    XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }),
  );
}

export function defaultSpreadsheetFileName(name: string): string {
  const base = name.replace(/[^\w\s.-]/g, "").trim() || "spreadsheet";
  return base.toLowerCase().endsWith(".xlsx") ? base : `${base}.xlsx`;
}