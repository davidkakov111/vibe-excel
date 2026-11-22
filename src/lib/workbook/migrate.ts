import { createEmptyWorkbook, normalizeWorkbookDocument } from "@/lib/workbook/defaults";
import {
  type CellData,
  type WorkbookDocument,
  WORKBOOK_VERSION,
} from "@/lib/workbook/types";
import { type CellMap, type SpreadsheetContent } from "@/lib/spreadsheet-types";

function isV2Workbook(value: unknown): value is WorkbookDocument {
  if (!value || typeof value !== "object") return false;
  const raw = value as Record<string, unknown>;
  return (
    raw.version === WORKBOOK_VERSION &&
    Array.isArray(raw.sheets) &&
    typeof raw.activeSheetId === "string"
  );
}

export function migrateLegacyContent(
  legacy: SpreadsheetContent,
  sheetName = "Sheet1",
): WorkbookDocument {
  const workbook = createEmptyWorkbook(sheetName);
  const sheet = workbook.sheets[0]!;
  sheet.rowCount = legacy.rowCount;
  sheet.colCount = legacy.colCount;
  const cells: Record<string, CellData> = {};
  for (const [key, value] of Object.entries(legacy.cells)) {
    if (value !== "") {
      cells[key] = { input: value };
    }
  }
  sheet.cells = cells;
  return workbook;
}

export function normalizeStoredDocument(
  stored: unknown,
  legacy: SpreadsheetContent,
  sheetName?: string,
): WorkbookDocument {
  if (isV2Workbook(stored)) {
    return normalizeWorkbookDocument(stored);
  }
  if (
    stored &&
    typeof stored === "object" &&
    !Array.isArray(stored) &&
    "cells" in (stored as object)
  ) {
    const raw = stored as Record<string, unknown>;
    const cells =
      raw.cells && typeof raw.cells === "object" && !Array.isArray(raw.cells)
        ? (raw.cells as CellMap)
        : {};
    return migrateLegacyContent(
      {
        cells,
        rowCount:
          typeof raw.rowCount === "number"
            ? raw.rowCount
            : legacy.rowCount,
        colCount:
          typeof raw.colCount === "number"
            ? raw.colCount
            : legacy.colCount,
      },
      sheetName,
    );
  }
  if (stored && typeof stored === "object" && !Array.isArray(stored)) {
    const maybeMap = stored as CellMap;
    if (Object.keys(maybeMap).every((k) => k.includes(":"))) {
      return migrateLegacyContent({
        cells: maybeMap,
        rowCount: legacy.rowCount,
        colCount: legacy.colCount,
      });
    }
  }
  return migrateLegacyContent(legacy, sheetName);
}

export function workbookToStorage(workbook: WorkbookDocument): WorkbookDocument {
  return workbook;
}