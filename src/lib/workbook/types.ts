import { cellKey } from "@/lib/spreadsheet-types";

export const WORKBOOK_VERSION = 2 as const;

export type HorizontalAlign = "left" | "center" | "right";
export type VerticalAlign = "top" | "middle" | "bottom";
export type BorderStyle = "thin" | "medium" | "thick" | "dashed";

export interface CellBorder {
  style?: BorderStyle;
  color?: string;
}

export interface CellBorders {
  top?: CellBorder;
  right?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
}

export type NumberFormatKind =
  | "general"
  | "number"
  | "currency"
  | "percentage"
  | "date"
  | "time"
  | "accounting"
  | "text";

export interface CellFormat {
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontColor?: string;
  fillColor?: string;
  horizontalAlign?: HorizontalAlign;
  verticalAlign?: VerticalAlign;
  wrap?: boolean;
  rotation?: number;
  borders?: CellBorders;
  numberFormat?: NumberFormatKind;
  numberDecimals?: number;
}

export interface DataValidation {
  kind: "list" | "whole" | "decimal" | "date" | "textLength";
  values?: string[];
  min?: number;
  max?: number;
  allowBlank?: boolean;
  errorMessage?: string;
  inputMessage?: string;
}

export type { CellCommentMessage, CellCommentThread } from "@/lib/workbook/comment";

export interface CellData {
  /** Raw user input (formula string or literal). */
  input?: string;
  format?: CellFormat;
  validation?: DataValidation;
  comment?: import("@/lib/workbook/comment").CellCommentThread;
}

export interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface FilterColumnState {
  col: number;
  operator:
    | "equals"
    | "contains"
    | "greater"
    | "less"
    | "notEquals"
    | "notContains"
    | "oneOf";
  value: string;
  /** When operator is `oneOf`, rows must match one of these display/input values. */
  allowedValues?: string[];
}

export interface AutoFilterState {
  enabled: boolean;
  range: { startRow: number; startCol: number; endRow: number; endCol: number };
  columns: FilterColumnState[];
  hiddenRows: number[];
}

export interface ChartSeries {
  id: string;
  type: "bar" | "line" | "pie" | "area";
  title: string;
  /** Data source range (labels in first column). */
  range: { startRow: number; startCol: number; endRow: number; endCol: number };
  /** On-sheet placement (merged region reserved for the chart). */
  anchor: { startRow: number; startCol: number; endRow: number; endCol: number };
}

export interface ExcelTableColumn {
  name: string;
  totalsRowFunction?: string;
  totalsRowLabel?: string;
  filterButton?: boolean;
}

export interface ExcelTableStyle {
  theme?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

/** Structured Excel table (ListObject), stored for import/export round-trip. */
export interface ExcelTable {
  name: string;
  displayName?: string;
  ref: string;
  headerRow?: boolean;
  totalsRow?: boolean;
  columns: ExcelTableColumn[];
  style?: ExcelTableStyle;
  autoFilterRef?: string;
}

/** One conditional-formatting region (mirrors ExcelJS worksheet model). */
export interface ConditionalFormattingSet {
  ref: string;
  rules: Array<Record<string, unknown>>;
}

/** Native ExcelJS image anchor corner (column/row + EMU offsets). */
export interface WorksheetImageAnchorCorner {
  nativeCol: number;
  nativeColOff?: number;
  nativeRow: number;
  nativeRowOff?: number;
}

/** Worksheet picture from import (charts as PNG, pasted images, etc.). */
export interface WorksheetEmbeddedImage {
  extension: string;
  dataBase64: string;
  range: {
    tl: WorksheetImageAnchorCorner;
    br?: WorksheetImageAnchorCorner;
    editAs?: string;
    ext?: { width: number; height: number };
  };
}

export interface Worksheet {
  id: string;
  name: string;
  rowCount: number;
  colCount: number;
  cells: Record<string, CellData>;
  rowHeights: Record<string, number>;
  colWidths: Record<string, number>;
  hiddenRows: number[];
  hiddenCols: number[];
  frozenRows: number;
  frozenCols: number;
  merges: MergeRange[];
  autoFilter: AutoFilterState | null;
  charts: ChartSeries[];
  tables: ExcelTable[];
  conditionalFormats: ConditionalFormattingSet[];
  embeddedImages: WorksheetEmbeddedImage[];
}

export interface WorkbookViewSettings {
  showGridlines: boolean;
  zoom: number;
}

export interface WorkbookDocument {
  version: typeof WORKBOOK_VERSION;
  activeSheetId: string;
  sheets: Worksheet[];
  settings: WorkbookViewSettings;
}

export function getCellInput(sheet: Worksheet, row: number, col: number): string {
  return sheet.cells[cellKey(row, col)]?.input ?? "";
}

export function setCellInput(
  sheet: Worksheet,
  row: number,
  col: number,
  input: string,
): Worksheet {
  const key = cellKey(row, col);
  const nextCells = { ...sheet.cells };
  const trimmed = input.trim();
  if (!trimmed) {
    delete nextCells[key];
  } else {
    const existing = nextCells[key];
    nextCells[key] = existing ? { ...existing, input: input } : { input: input };
  }
  return { ...sheet, cells: nextCells };
}