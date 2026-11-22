import ExcelJS from "exceljs";
import { parseA1 } from "@/lib/formula/engine";
import { parseA1Range } from "@/lib/excel/a1-range";
import { migrateLegacyContent } from "@/lib/workbook/migrate";
import {
  createEmptyWorkbook,
  DEFAULT_COL_COUNT,
  DEFAULT_ROW_COUNT,
  normalizeWorkbookDocument,
} from "@/lib/workbook/defaults";
import { cellKey, indexToColumnLabel } from "@/lib/spreadsheet-types";
import type {
  CellData,
  ConditionalFormattingSet,
  ExcelTable,
  ExcelTableColumn,
  MergeRange,
  WorkbookDocument,
  Worksheet,
  WorksheetEmbeddedImage,
  WorksheetImageAnchorCorner,
} from "@/lib/workbook/types";
import { FormulaEngine, isFormulaInput } from "@/lib/formula/engine";
import { excelCellToInput, hasExcelCellContent } from "@/lib/excel/cell-value";
import {
  commentThreadToExcelNote,
  excelNoteToCommentThread,
} from "@/lib/excel/comment-excel";
import { chartAnchorRef, renderChartPng } from "@/lib/excel/chart-image";
import { normalizeCellComment } from "@/lib/workbook/comment";
import { applyExcelStyle, readFormat } from "@/lib/excel/style-map";
import { parseThemePaletteFromWorkbookXml } from "@/lib/excel/theme-colors";
import { chartSeriesData } from "@/lib/workbook/chart-data";

function parseMergeRef(mergeRef: string): MergeRange | null {
  const [startRef, endRef] = mergeRef.split(":");
  if (!startRef || !endRef) return null;
  const start = parseA1(startRef.replace(/\$/g, ""));
  const end = parseA1(endRef.replace(/\$/g, ""));
  if (!start || !end) return null;
  return {
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col),
  };
}

type ExcelTableModel = {
  name: string;
  displayName?: string;
  tableRef?: string;
  ref?: string;
  totalsRow?: boolean;
  headerRow?: boolean;
  columns?: ExcelTableColumn[];
  autoFilterRef?: string;
  style?: ExcelTable["style"];
};

function importTablesFromModel(tables?: ExcelTableModel[]): ExcelTable[] {
  if (!tables?.length) return [];
  return tables.map((t) => ({
    name: t.name,
    displayName: t.displayName,
    ref: String(t.tableRef ?? t.ref ?? "A1"),
    headerRow: t.headerRow ?? true,
    totalsRow: t.totalsRow ?? false,
    columns: (t.columns ?? []).map((c) => ({
      name: c.name,
      totalsRowFunction: c.totalsRowFunction,
      totalsRowLabel: c.totalsRowLabel,
      filterButton: c.filterButton,
    })),
    style: t.style,
    autoFilterRef: t.autoFilterRef,
  }));
}

function importConditionalFormats(
  raw?: ConditionalFormattingSet[],
): ConditionalFormattingSet[] {
  if (!raw?.length) return [];
  return raw.map((block) => ({
    ref: block.ref,
    rules: (block.rules ?? []).map((rule) => ({ ...rule })),
  }));
}

type ExcelWorksheetMediaItem = {
  type: string;
  imageId: number;
  range?: {
    tl: WorksheetImageAnchorCorner;
    br?: WorksheetImageAnchorCorner;
    editAs?: string;
    ext?: { width: number; height: number };
  };
};

function copyAnchorCorner(corner: WorksheetImageAnchorCorner): WorksheetImageAnchorCorner {
  return {
    nativeCol: corner.nativeCol,
    nativeColOff: corner.nativeColOff,
    nativeRow: corner.nativeRow,
    nativeRowOff: corner.nativeRowOff,
  };
}

function importEmbeddedImages(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
): WorksheetEmbeddedImage[] {
  const media = (worksheet.model.media ?? []) as unknown as ExcelWorksheetMediaItem[];
  const images: WorksheetEmbeddedImage[] = [];
  for (const item of media) {
    if (item.type !== "image" || !item.range) continue;
    const raw = workbook.getImage(item.imageId);
    if (!raw?.buffer) continue;
    const imageBuf = Buffer.from(raw.buffer as unknown as Uint8Array);
    if (!imageBuf.length) continue;
    images.push({
      extension: raw.extension ?? "png",
      dataBase64: imageBuf.toString("base64"),
      range: {
        tl: copyAnchorCorner(item.range.tl),
        br: item.range.br ? copyAnchorCorner(item.range.br) : undefined,
        editAs: item.range.editAs,
        ext: item.range.ext ? { ...item.range.ext } : undefined,
      },
    });
  }
  return images;
}

function sheetFromExcelWorksheet(
  workbook: ExcelJS.Workbook,
  worksheet: ExcelJS.Worksheet,
  name: string,
  themePalette: readonly string[],
): Worksheet {
  const sheet = createEmptyWorkbook(name).sheets[0]!;
  sheet.name = name;

  let maxRow = 0;
  let maxCol = 0;

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    maxRow = Math.max(maxRow, rowNumber);
    if (row.height) {
      sheet.rowHeights[String(rowNumber - 1)] = Math.round(row.height * (96 / 72));
    }
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      maxCol = Math.max(maxCol, colNumber);
      if (!hasExcelCellContent(cell)) {
        const formatOnly = readFormat(cell, themePalette);
        if (!formatOnly) return;
        sheet.cells[cellKey(rowNumber - 1, colNumber - 1)] = { format: formatOnly };
        return;
      }
      const data: CellData = {};
      const input = excelCellToInput(cell);
      if (input !== undefined) data.input = input;
      const format = readFormat(cell, themePalette);
      if (format) data.format = format;
      const note = cell.note;
      if (note) {
        const thread =
          typeof note === "string"
            ? excelNoteToCommentThread(note)
            : excelNoteToCommentThread(note);
        if (thread) data.comment = normalizeCellComment(thread) ?? undefined;
      }
      if (!data.input && !data.format) return;
      sheet.cells[cellKey(rowNumber - 1, colNumber - 1)] = data;
    });
  });

  for (const mergeRef of worksheet.model.merges ?? []) {
    const range = parseMergeRef(mergeRef);
    if (range) sheet.merges.push(range);
  }

  const sheetModel = worksheet.model as typeof worksheet.model & {
    tables?: ExcelTableModel[];
    conditionalFormattings?: ConditionalFormattingSet[];
  };
  sheet.tables = importTablesFromModel(sheetModel.tables);
  sheet.conditionalFormats = importConditionalFormats(
    sheetModel.conditionalFormattings,
  );
  sheet.embeddedImages = importEmbeddedImages(workbook, worksheet);

  sheet.rowCount = Math.max(maxRow, DEFAULT_ROW_COUNT);
  sheet.colCount = Math.max(maxCol, DEFAULT_COL_COUNT);

  worksheet.columns?.forEach((col, index) => {
    if (col.width) {
      sheet.colWidths[String(index)] = Math.round(col.width * 7);
    }
  });

  return sheet;
}

function tableRowsFromSheet(sheet: Worksheet, table: ExcelTable): unknown[][] {
  const range = parseA1Range(table.ref);
  if (!range) return [];
  const rows: unknown[][] = [];
  const dataStart =
    table.headerRow !== false ? range.startRow + 1 : range.startRow;
  const dataEnd = table.totalsRow ? range.endRow - 1 : range.endRow;
  for (let r = dataStart; r <= dataEnd; r++) {
    const rowVals: unknown[] = [];
    for (let c = range.startCol; c <= range.endCol; c++) {
      const input = sheet.cells[cellKey(r, c)]?.input ?? "";
      const numeric = Number(input);
      rowVals.push(
        input !== "" && !Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(input.trim())
          ? numeric
          : input,
      );
    }
    rows.push(rowVals);
  }
  return rows;
}

function exportTables(ws: ExcelJS.Worksheet, sheet: Worksheet): void {
  for (const table of sheet.tables ?? []) {
    if (!table.columns.length) continue;
    try {
      ws.addTable({
        name: table.name,
        displayName: table.displayName ?? table.name,
        ref: table.ref,
        headerRow: table.headerRow ?? true,
        totalsRow: table.totalsRow ?? false,
        style: table.style as ExcelJS.TableStyleProperties | undefined,
        columns: table.columns.map((col) => ({
          name: col.name,
          totalsRowFunction: col.totalsRowFunction as ExcelJS.TableColumnProperties["totalsRowFunction"],
          totalsRowLabel: col.totalsRowLabel,
          filterButton: col.filterButton,
        })),
        rows: tableRowsFromSheet(sheet, table),
      });
    } catch {
      // Skip invalid table definitions rather than failing export.
    }
  }
}

function exportConditionalFormats(ws: ExcelJS.Worksheet, sheet: Worksheet): void {
  for (const block of sheet.conditionalFormats ?? []) {
    if (!block.rules.length) continue;
    try {
      ws.addConditionalFormatting({
        ref: block.ref,
        rules: block.rules as unknown as ExcelJS.ConditionalFormattingRule[],
      });
    } catch {
      // Preserve export for the rest of the workbook.
    }
  }
}

function exportEmbeddedImages(
  workbook: ExcelJS.Workbook,
  ws: ExcelJS.Worksheet,
  sheet: Worksheet,
): void {
  for (const image of sheet.embeddedImages ?? []) {
    if (!image.dataBase64) continue;
    try {
      const imageId = workbook.addImage({
        buffer: Buffer.from(image.dataBase64, "base64") as unknown as ExcelJS.Buffer,
        extension: (image.extension || "png") as "jpeg" | "png" | "gif",
      });
      ws.addImage(
        imageId,
        {
          tl: image.range.tl,
          br: image.range.br,
          editAs: image.range.editAs,
          ext: image.range.ext,
        } as unknown as string,
      );
    } catch {
      // Skip invalid placements rather than failing export.
    }
  }
}

export async function parseExcelWorkbook(buffer: Buffer): Promise<WorkbookDocument> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as unknown as ExcelJS.Buffer);
  const themeXml = (workbook.model as { themes?: { theme1?: string } }).themes
    ?.theme1;
  const themePalette = parseThemePaletteFromWorkbookXml(themeXml);

  if (!workbook.worksheets.length) {
    return createEmptyWorkbook();
  }
  const sheets: Worksheet[] = workbook.worksheets.map((ws) =>
    sheetFromExcelWorksheet(workbook, ws, ws.name, themePalette),
  );
  return normalizeWorkbookDocument({
    version: 2,
    activeSheetId: sheets[0]!.id,
    sheets,
    settings: { showGridlines: true, zoom: 100 },
  });
}

export async function workbookToExcelBuffer(
  document: WorkbookDocument,
): Promise<Buffer> {
  const doc = normalizeWorkbookDocument(document);
  const engine = FormulaEngine.fromWorkbook(doc);
  const workbook = new ExcelJS.Workbook();

  for (const sheet of doc.sheets) {
    const ws = workbook.addWorksheet(sheet.name);
    for (const [key, data] of Object.entries(sheet.cells)) {
      const [row, col] = key.split(":").map(Number);
      const excelCell = ws.getCell(row + 1, col + 1);
      const input = data.input ?? "";
      if (isFormulaInput(input)) {
        excelCell.value = { formula: input.slice(1) };
      } else if (input !== "") {
        const numeric = Number(input);
        excelCell.value =
          !Number.isNaN(numeric) && /^-?\d+(\.\d+)?$/.test(input.trim())
            ? numeric
            : input;
      }
      applyExcelStyle(excelCell, data.format);
      if (data.comment) {
        const thread = normalizeCellComment(data.comment);
        if (thread) excelCell.note = commentThreadToExcelNote(thread);
      }
    }

    for (const [colKey, width] of Object.entries(sheet.colWidths)) {
      const col = ws.getColumn(Number(colKey) + 1);
      col.width = width / 7;
    }

    for (const [rowKey, height] of Object.entries(sheet.rowHeights)) {
      const row = ws.getRow(Number(rowKey) + 1);
      row.height = height * (72 / 96);
    }

    for (const merge of sheet.merges) {
      ws.mergeCells(
        merge.startRow + 1,
        merge.startCol + 1,
        merge.endRow + 1,
        merge.endCol + 1,
      );
    }

    exportTables(ws, sheet);
    exportConditionalFormats(ws, sheet);
    exportEmbeddedImages(workbook, ws, sheet);

    for (const chart of sheet.charts ?? []) {
      const points = chartSeriesData(
        sheet,
        chart,
        (sheetId, row, col) => engine.getDisplayValue(sheetId, row, col),
        sheet.id,
      );
      const png = renderChartPng(chart, points);
      if (!png) continue;
      const imageId = workbook.addImage({
        buffer: png as unknown as ExcelJS.Buffer,
        extension: "png",
      });
      try {
        ws.addImage(imageId, chartAnchorRef(chart.anchor));
      } catch {
        // skip invalid anchors
      }
    }

    for (let row = 0; row < sheet.rowCount; row++) {
      for (let col = 0; col < sheet.colCount; col++) {
        const display = engine.getDisplayValue(sheet.id, row, col);
        const input = sheet.cells[cellKey(row, col)]?.input;
        if (!input && display) {
          const excelCell = ws.getCell(row + 1, col + 1);
          excelCell.value = display;
        }
      }
    }
  }

  engine.destroy();
  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

/** @deprecated use parseExcelWorkbook */
export function legacyContentFromBuffer(buffer: Buffer): Promise<WorkbookDocument> {
  return parseExcelWorkbook(buffer);
}

export function workbookColumnLabel(col: number): string {
  return indexToColumnLabel(col);
}

export function migrateFromLegacyCells(
  cells: Record<string, string>,
  rowCount: number,
  colCount: number,
): WorkbookDocument {
  return migrateLegacyContent({ cells, rowCount, colCount });
}