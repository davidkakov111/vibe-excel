import { cellKey } from "@/lib/spreadsheet-types";
import { FormulaEngine } from "@/lib/formula/engine";
import { sortWorksheetRange } from "@/lib/workbook/sort";
import { chartWithDefaultAnchor } from "@/lib/workbook/chart-data";
import type { CellCommentThread } from "@/lib/workbook/comment";
import type {
  CellData,
  CellFormat,
  ChartSeries,
  ConditionalFormattingSet,
  MergeRange,
  WorkbookDocument,
  Worksheet,
} from "@/lib/workbook/types";
import { getCellInput, setCellInput } from "@/lib/workbook/types";

export type SetCellOp = {
  type: "setCell";
  sheetId: string;
  row: number;
  col: number;
  input: string | null;
};

export type SetFormatOp = {
  type: "setFormat";
  sheetId: string;
  row: number;
  col: number;
  format: Partial<CellFormat> | null;
};

export type SetCommentOp = {
  type: "setComment";
  sheetId: string;
  row: number;
  col: number;
  comment: CellCommentThread | null;
};

export type SortRangeOp = {
  type: "sortRange";
  sheetId: string;
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  sortCol: number;
  ascending: boolean;
};

export type RestoreRangeOp = {
  type: "restoreRange";
  sheetId: string;
  cells: Record<string, CellData | null>;
};

export type PatchSheetOp = {
  type: "patchSheet";
  sheetId: string;
  charts?: ChartSeries[];
  merges?: MergeRange[];
  rowCount?: number;
  colCount?: number;
  rowHeights?: Record<string, number>;
  colWidths?: Record<string, number>;
  cells?: Record<string, CellData | null>;
  conditionalFormats?: ConditionalFormattingSet[];
};

export type WorkbookOp =
  | SetCellOp
  | SetFormatOp
  | SetCommentOp
  | SortRangeOp
  | RestoreRangeOp
  | PatchSheetOp;

function sheetById(doc: WorkbookDocument, sheetId: string): Worksheet | undefined {
  return doc.sheets.find((s) => s.id === sheetId);
}

function upsertCell(
  sheet: Worksheet,
  row: number,
  col: number,
  patch: (current: Worksheet["cells"][string] | undefined) => Worksheet["cells"][string] | undefined,
): Worksheet {
  const key = cellKey(row, col);
  const nextCells = { ...sheet.cells };
  const next = patch(nextCells[key]);
  if (next) nextCells[key] = next;
  else delete nextCells[key];
  return { ...sheet, cells: nextCells };
}

export function applyWorkbookOp(
  doc: WorkbookDocument,
  op: WorkbookOp,
): WorkbookDocument {
  const index = doc.sheets.findIndex((s) => s.id === op.sheetId);
  if (index < 0) return doc;
  let sheet = doc.sheets[index]!;

  switch (op.type) {
    case "setCell": {
      if (op.input === null || !String(op.input).trim()) {
        const key = cellKey(op.row, op.col);
        const existing = sheet.cells[key];
        if (!existing?.format && !existing?.comment && !existing?.validation) {
          const nextCells = { ...sheet.cells };
          delete nextCells[key];
          sheet = { ...sheet, cells: nextCells };
        } else {
          sheet = upsertCell(sheet, op.row, op.col, (current) => {
            if (!current) return current;
            const rest = { ...current };
            delete rest.input;
            return Object.keys(rest).length ? rest : undefined;
          });
        }
      } else {
        sheet = setCellInput(sheet, op.row, op.col, op.input);
      }
      break;
    }
    case "setFormat": {
      sheet = upsertCell(sheet, op.row, op.col, (current) => {
        if (op.format === null) {
          if (!current) return current;
          const rest = { ...current };
          delete rest.format;
          return Object.keys(rest).length ? rest : undefined;
        }
        const base = current ?? {};
        return { ...base, format: { ...base.format, ...op.format } };
      });
      break;
    }
    case "setComment": {
      sheet = upsertCell(sheet, op.row, op.col, (current) => {
        if (op.comment === null) {
          if (!current) return current;
          const rest = { ...current };
          delete rest.comment;
          return Object.keys(rest).length ? rest : undefined;
        }
        const base = current ?? {};
        return { ...base, comment: op.comment };
      });
      break;
    }
    case "sortRange": {
      const engine = FormulaEngine.fromWorkbook(doc);
      sheet = sortWorksheetRange(
        sheet,
        op.startRow,
        op.endRow,
        op.startCol,
        op.endCol,
        op.sortCol,
        op.ascending,
        (row, col) => {
          const cellInput = getCellInput(sheet, row, col);
          if (cellInput && !cellInput.startsWith("=")) return cellInput;
          return engine.getDisplayValue(op.sheetId, row, col) || cellInput;
        },
      );
      engine.destroy();
      break;
    }
    case "restoreRange": {
      const nextCells = { ...sheet.cells };
      for (const [key, value] of Object.entries(op.cells)) {
        if (value === null) delete nextCells[key];
        else nextCells[key] = value;
      }
      sheet = { ...sheet, cells: nextCells };
      break;
    }
    case "patchSheet": {
      if (op.charts !== undefined) {
        sheet = {
          ...sheet,
          charts: op.charts.map((c) => chartWithDefaultAnchor(c)),
        };
      }
      if (op.merges !== undefined) sheet = { ...sheet, merges: op.merges };
      if (op.rowCount !== undefined) sheet = { ...sheet, rowCount: op.rowCount };
      if (op.colCount !== undefined) sheet = { ...sheet, colCount: op.colCount };
      if (op.rowHeights) {
        sheet = {
          ...sheet,
          rowHeights: { ...sheet.rowHeights, ...op.rowHeights },
        };
      }
      if (op.colWidths) {
        sheet = {
          ...sheet,
          colWidths: { ...sheet.colWidths, ...op.colWidths },
        };
      }
      if (op.cells) {
        const nextCells = { ...sheet.cells };
        for (const [key, value] of Object.entries(op.cells)) {
          if (value === null) delete nextCells[key];
          else nextCells[key] = value;
        }
        sheet = { ...sheet, cells: nextCells };
      }
      if (op.conditionalFormats !== undefined) {
        sheet = {
          ...sheet,
          conditionalFormats: structuredClone(op.conditionalFormats),
        };
      }
      break;
    }
    default:
      break;
  }

  const sheets = [...doc.sheets];
  sheets[index] = sheet;
  return { ...doc, sheets };
}

export function applyWorkbookOps(
  doc: WorkbookDocument,
  ops: WorkbookOp[],
): WorkbookDocument {
  return ops.reduce((acc, op) => applyWorkbookOp(acc, op), doc);
}

export function cellOpFromInput(
  sheetId: string,
  row: number,
  col: number,
  input: string,
): SetCellOp {
  const trimmed = input.trim();
  return {
    type: "setCell",
    sheetId,
    row,
    col,
    input: trimmed === "" ? null : input,
  };
}

export function formatOp(
  sheetId: string,
  row: number,
  col: number,
  format: Partial<CellFormat> | null,
): SetFormatOp {
  return { type: "setFormat", sheetId, row, col, format };
}

export function commentOp(
  sheetId: string,
  row: number,
  col: number,
  comment: CellCommentThread | null,
): SetCommentOp {
  return { type: "setComment", sheetId, row, col, comment };
}

export function readCellInput(
  doc: WorkbookDocument,
  sheetId: string,
  row: number,
  col: number,
): string {
  const sheet = sheetById(doc, sheetId);
  if (!sheet) return "";
  return getCellInput(sheet, row, col);
}

export function sortRangeOp(
  sheetId: string,
  startRow: number,
  endRow: number,
  startCol: number,
  endCol: number,
  sortCol: number,
  ascending: boolean,
): SortRangeOp {
  return {
    type: "sortRange",
    sheetId,
    startRow,
    endRow,
    startCol,
    endCol,
    sortCol,
    ascending,
  };
}

export function restoreRangeOp(
  sheetId: string,
  cells: Record<string, CellData | null>,
): RestoreRangeOp {
  return { type: "restoreRange", sheetId, cells };
}

export function patchSheetOp(
  sheetId: string,
  patch: Omit<PatchSheetOp, "type" | "sheetId">,
): PatchSheetOp {
  return { type: "patchSheet", sheetId, ...patch };
}