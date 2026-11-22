import { normalizeStoredDocument } from "@/lib/workbook/migrate";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import type { WorkbookDocument } from "@/lib/workbook/types";
import type { CellMap } from "@/lib/spreadsheet-types";

export function cellsFromDb(value: unknown): CellMap {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }
  const raw = value as Record<string, unknown>;
  if (raw.version === 2) {
    return {};
  }
  return value as CellMap;
}

export function workbookFromDb(
  stored: unknown,
  rowCount: number,
  colCount: number,
  spreadsheetName?: string,
): WorkbookDocument {
  const legacy = {
    cells: cellsFromDb(stored),
    rowCount,
    colCount,
  };
  return normalizeStoredDocument(stored, legacy, spreadsheetName);
}

export function emptyWorkbookFallback(): WorkbookDocument {
  return createEmptyWorkbook();
}