import { describe, expect, it } from "vitest";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { fillSelection } from "@/lib/workbook/fill";
import { setCellInput } from "@/lib/workbook/types";

describe("fillSelection", () => {
  it("extends values downward", () => {
    const wb = createEmptyWorkbook();
    let sheet = wb.sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "1");
    sheet = setCellInput(sheet, 1, 0, "2");
    sheet = fillSelection(
      sheet,
      { row: 0, col: 0, endRow: 1, endCol: 0 },
      { row: 0, col: 0, endRow: 4, endCol: 0 },
    );
    expect(sheet.cells["2:0"]?.input).toBe("1");
    expect(sheet.cells["3:0"]?.input).toBe("2");
  });

  it("fills a 2D extension with tiled source pattern", () => {
    const wb = createEmptyWorkbook();
    let sheet = wb.sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "A");
    sheet = setCellInput(sheet, 0, 1, "B");
    sheet = setCellInput(sheet, 1, 0, "C");
    sheet = setCellInput(sheet, 1, 1, "D");
    sheet = fillSelection(
      sheet,
      { row: 0, col: 0, endRow: 1, endCol: 1 },
      { row: 0, col: 0, endRow: 3, endCol: 3 },
    );
    expect(sheet.cells["2:2"]?.input).toBe("A");
    expect(sheet.cells["2:3"]?.input).toBe("B");
    expect(sheet.cells["3:2"]?.input).toBe("C");
    expect(sheet.cells["3:3"]?.input).toBe("D");
  });
});