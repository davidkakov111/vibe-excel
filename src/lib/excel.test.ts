import { describe, expect, it } from "vitest";
import { parseExcelBuffer, spreadsheetToExcelBuffer } from "@/lib/excel";
import { cellKey } from "@/lib/spreadsheet-types";

describe("excel import/export", () => {
  it("preserves basic cell values across a round trip", () => {
    const source = {
      cells: {
        [cellKey(0, 0)]: "Revenue",
        [cellKey(0, 1)]: "42",
        [cellKey(1, 0)]: "Notes",
      },
      rowCount: 20,
      colCount: 10,
    };

    const workbook = spreadsheetToExcelBuffer(source);
    const imported = parseExcelBuffer(workbook);

    expect(imported.cells[cellKey(0, 0)]).toBe("Revenue");
    expect(imported.cells[cellKey(0, 1)]).toBe("42");
    expect(imported.cells[cellKey(1, 0)]).toBe("Notes");
  });
});