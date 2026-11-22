import { cellKey } from "@/lib/spreadsheet-types";
import type {
  AutoFilterState,
  FilterColumnState,
  Worksheet,
} from "@/lib/workbook/types";
import { getCellInput } from "@/lib/workbook/types";

function cellMatchesFilter(
  raw: string,
  filter: FilterColumnState,
  display: string,
): boolean {
  const value = raw || display;
  const target = filter.value;
  const numValue = Number(value);
  const numTarget = Number(target);
  const bothNumeric =
    target.trim() !== "" &&
    value.trim() !== "" &&
    !Number.isNaN(numValue) &&
    !Number.isNaN(numTarget);

  switch (filter.operator) {
    case "oneOf":
      if (!filter.allowedValues?.length) return false;
      return (
        filter.allowedValues.includes(value) ||
        filter.allowedValues.includes(display)
      );
    case "equals":
      return bothNumeric ? numValue === numTarget : value === target;
    case "notEquals":
      return bothNumeric ? numValue !== numTarget : value !== target;
    case "contains":
      return value.toLowerCase().includes(target.toLowerCase());
    case "notContains":
      return !value.toLowerCase().includes(target.toLowerCase());
    case "greater":
      return bothNumeric ? numValue > numTarget : value > target;
    case "less":
      return bothNumeric ? numValue < numTarget : value < target;
    default:
      return true;
  }
}

export function computeHiddenRows(
  sheet: Worksheet,
  filter: AutoFilterState,
  getDisplay: (row: number, col: number) => string,
): number[] {
  const hidden: number[] = [];
  const { startRow, endRow, startCol, endCol } = filter.range;
  for (let row = startRow; row <= endRow; row++) {
    if (row === startRow) continue;
    let visible = true;
    for (const columnFilter of filter.columns) {
      if (columnFilter.col < startCol || columnFilter.col > endCol) continue;
      const raw = getCellInput(sheet, row, columnFilter.col);
      const display = getDisplay(row, columnFilter.col);
      if (columnFilter.operator === "oneOf") {
        if (columnFilter.allowedValues === undefined) continue;
        if (!cellMatchesFilter(raw, columnFilter, display)) {
          visible = false;
          break;
        }
        continue;
      }
      if (!columnFilter.value.trim()) continue;
      if (!cellMatchesFilter(raw, columnFilter, display)) {
        visible = false;
        break;
      }
    }
    if (!visible) hidden.push(row);
  }
  return hidden;
}

export function enableAutoFilter(
  sheet: Worksheet,
  range: AutoFilterState["range"],
): Worksheet {
  return {
    ...sheet,
    autoFilter: {
      enabled: true,
      range,
      columns: [],
      hiddenRows: [],
    },
  };
}

export function setColumnFilter(
  sheet: Worksheet,
  filter: FilterColumnState,
  getDisplay: (row: number, col: number) => string,
): Worksheet {
  if (!sheet.autoFilter?.enabled) return sheet;
  const columns = sheet.autoFilter.columns.filter((c) => c.col !== filter.col);
  const hasFilter =
    filter.operator === "oneOf"
      ? filter.allowedValues !== undefined
      : filter.value.trim().length > 0;
  if (hasFilter) {
    columns.push(filter);
  }
  const nextFilter: AutoFilterState = {
    ...sheet.autoFilter,
    columns,
    hiddenRows: [],
  };
  const hiddenRows = computeHiddenRows(sheet, nextFilter, getDisplay);
  return {
    ...sheet,
    autoFilter: { ...nextFilter, hiddenRows },
  };
}

export function clearAutoFilter(sheet: Worksheet): Worksheet {
  return {
    ...sheet,
    autoFilter: sheet.autoFilter
      ? { ...sheet.autoFilter, columns: [], hiddenRows: [] }
      : null,
  };
}

export function disableAutoFilter(sheet: Worksheet): Worksheet {
  return { ...sheet, autoFilter: null };
}

export function uniqueColumnValues(
  sheet: Worksheet,
  col: number,
  startRow: number,
  endRow: number,
): string[] {
  const values = new Set<string>();
  for (let row = startRow; row <= endRow; row++) {
    const input = sheet.cells[cellKey(row, col)]?.input ?? "";
  if (input) values.add(input);
  }
  return [...values].sort();
}