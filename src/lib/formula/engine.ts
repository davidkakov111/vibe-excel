import {
  CellError,
  HyperFormula,
  type CellValue,
  type SimpleCellAddress,
} from "hyperformula";
import type { WorkbookDocument, Worksheet } from "@/lib/workbook/types";
import { cellKey, indexToColumnLabel } from "@/lib/spreadsheet-types";

export type FormulaErrorCode = "#DIV/0!" | "#N/A" | "#NAME?" | "#NUM!" | "#REF!" | "#VALUE!" | "#CYCLE!";

const HF_LICENSE = "gpl-v3";

export function addressToA1(row: number, col: number): string {
  return `${indexToColumnLabel(col)}${row + 1}`;
}

export function parseA1(ref: string): { row: number; col: number } | null {
  const match = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(ref.trim());
  if (!match) return null;
  const letters = match[2]!.toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  col -= 1;
  const row = Number(match[4]) - 1;
  if (row < 0 || col < 0) return null;
  return { row, col };
}

function sheetInputs(sheet: Worksheet): (string | number | boolean)[][] {
  const rows: (string | number | boolean)[][] = [];
  for (let r = 0; r < sheet.rowCount; r++) {
    const row: (string | number | boolean)[][][number] = [];
    for (let c = 0; c < sheet.colCount; c++) {
      const input = sheet.cells[cellKey(r, c)]?.input;
      row.push(input ?? "");
    }
    rows.push(row);
  }
  return rows;
}

export class FormulaEngine {
  private hf: HyperFormula;
  private sheetNameToIndex: Map<string, number> = new Map();
  private sheetIdToName: Map<string, string> = new Map();

  private constructor(hf: HyperFormula) {
    this.hf = hf;
  }

  static fromWorkbook(workbook: WorkbookDocument): FormulaEngine {
    const sheetsData: Record<string, (string | number | boolean)[][]> = {};
    for (const sheet of workbook.sheets) {
      sheetsData[sheet.name] = sheetInputs(sheet);
    }
    const hf = HyperFormula.buildFromSheets(sheetsData, {
      licenseKey: HF_LICENSE,
      useArrayArithmetic: true,
    });
    const engine = new FormulaEngine(hf);
    workbook.sheets.forEach((sheet, index) => {
      engine.sheetNameToIndex.set(sheet.name, index);
      engine.sheetIdToName.set(sheet.id, sheet.name);
    });
    return engine;
  }

  rebuild(workbook: WorkbookDocument): void {
    this.hf.destroy();
    const fresh = FormulaEngine.fromWorkbook(workbook);
    this.hf = fresh.hf;
    this.sheetNameToIndex = fresh.sheetNameToIndex;
    this.sheetIdToName = fresh.sheetIdToName;
  }

  getSheetIndex(sheetId: string): number {
    const name = this.sheetIdToName.get(sheetId);
    if (!name) return 0;
    return this.sheetNameToIndex.get(name) ?? 0;
  }

  setCellInput(sheetId: string, row: number, col: number, input: string): void {
    const sheetIndex = this.getSheetIndex(sheetId);
    const address: SimpleCellAddress = { sheet: sheetIndex, row, col };
    if (!input.trim()) {
      this.hf.setCellContents(address, null);
      return;
    }
    this.hf.setCellContents(address, [[input]]);
  }

  getComputedValue(sheetId: string, row: number, col: number): CellValue {
    const sheetIndex = this.getSheetIndex(sheetId);
    return this.hf.getCellValue({ sheet: sheetIndex, row, col });
  }

  getDisplayValue(sheetId: string, row: number, col: number): string {
    const value = this.getComputedValue(sheetId, row, col);
    return formatCellValue(value);
  }

  addSheet(sheetName: string): void {
    this.hf.addSheet(sheetName);
  }

  removeSheet(sheetIndex: number): void {
    this.hf.removeSheet(sheetIndex);
  }

  renameSheet(sheetIndex: number, newName: string): void {
    this.hf.renameSheet(sheetIndex, newName);
  }

  addRows(sheetId: string, row: number, count = 1): void {
    const sheetIndex = this.getSheetIndex(sheetId);
    this.hf.addRows(sheetIndex, [row, count]);
  }

  removeRows(sheetId: string, row: number, count = 1): void {
    const sheetIndex = this.getSheetIndex(sheetId);
    this.hf.removeRows(sheetIndex, [row, count]);
  }

  addColumns(sheetId: string, col: number, count = 1): void {
    const sheetIndex = this.getSheetIndex(sheetId);
    this.hf.addColumns(sheetIndex, [col, count]);
  }

  removeColumns(sheetId: string, col: number, count = 1): void {
    const sheetIndex = this.getSheetIndex(sheetId);
    this.hf.removeColumns(sheetIndex, [col, count]);
  }

  hasCircularReference(): boolean {
    return false;
  }

  destroy(): void {
    this.hf.destroy();
  }

  /** Sync sheet name maps after workbook structure change (no cell rebuild). */
  syncSheetMaps(workbook: WorkbookDocument): void {
    this.sheetNameToIndex.clear();
    this.sheetIdToName.clear();
    workbook.sheets.forEach((sheet, index) => {
      this.sheetNameToIndex.set(sheet.name, index);
      this.sheetIdToName.set(sheet.id, sheet.name);
    });
  }
}

export function formatCellValue(value: CellValue): string {
  if (value === null || value === undefined) return "";
  if (value instanceof CellError) {
    return value.message ?? "#ERROR";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return "#NUM!";
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return String(value);
}

export function isFormulaInput(input: string): boolean {
  return input.startsWith("=");
}