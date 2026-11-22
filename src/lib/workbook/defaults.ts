import { nanoid } from "nanoid";
import { normalizeCellComment, type StoredCellComment } from "@/lib/workbook/comment";
import { chartWithDefaultAnchor } from "@/lib/workbook/chart-data";
import {
  type WorkbookDocument,
  type WorkbookViewSettings,
  type Worksheet,
  WORKBOOK_VERSION,
} from "@/lib/workbook/types";

export const DEFAULT_ROW_COUNT = 100;
export const DEFAULT_COL_COUNT = 26;
export const DEFAULT_ROW_HEIGHT = 28;
export const DEFAULT_COL_WIDTH = 112;

const defaultView: WorkbookViewSettings = {
  showGridlines: true,
  zoom: 100,
};

export function createWorksheet(name: string): Worksheet {
  return {
    id: nanoid(),
    name,
    rowCount: DEFAULT_ROW_COUNT,
    colCount: DEFAULT_COL_COUNT,
    cells: {},
    rowHeights: {},
    colWidths: {},
    hiddenRows: [],
    hiddenCols: [],
    frozenRows: 0,
    frozenCols: 0,
    merges: [],
    autoFilter: null,
    charts: [],
    tables: [],
    conditionalFormats: [],
    embeddedImages: [],
  };
}

/** Ensure optional v2 fields exist on worksheets loaded from older JSON. */
export function normalizeWorksheetFields(sheet: Worksheet): Worksheet {
  const cells = { ...sheet.cells };
  for (const [key, data] of Object.entries(cells)) {
    if (!data?.comment) continue;
    const normalized = normalizeCellComment(data.comment as StoredCellComment);
    if (normalized) cells[key] = { ...data, comment: normalized };
  }
  return {
    ...sheet,
    cells,
    tables: sheet.tables ?? [],
    conditionalFormats: sheet.conditionalFormats ?? [],
    embeddedImages: sheet.embeddedImages ?? [],
    charts: (sheet.charts ?? []).map(chartWithDefaultAnchor),
  };
}

export function normalizeWorkbookDocument(doc: WorkbookDocument): WorkbookDocument {
  return {
    ...doc,
    sheets: doc.sheets.map(normalizeWorksheetFields),
  };
}

export function createEmptyWorkbook(sheetName = "Sheet1"): WorkbookDocument {
  const sheet = createWorksheet(sheetName);
  return {
    version: WORKBOOK_VERSION,
    activeSheetId: sheet.id,
    sheets: [sheet],
    settings: { ...defaultView },
  };
}

export function duplicateWorksheet(sheet: Worksheet, name: string): Worksheet {
  return {
    ...structuredClone(sheet),
    id: nanoid(),
    name,
  };
}