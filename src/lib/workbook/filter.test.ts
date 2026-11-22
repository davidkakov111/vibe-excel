import { describe, expect, it } from "vitest";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { setColumnFilter, enableAutoFilter } from "@/lib/workbook/filter";
import { setCellInput } from "@/lib/workbook/types";

describe("auto filter", () => {
  it("hides rows that do not match", () => {
    const wb = createEmptyWorkbook();
    let sheet = wb.sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "Name");
    sheet = setCellInput(sheet, 1, 0, "Alice");
    sheet = setCellInput(sheet, 2, 0, "Bob");
    sheet = setCellInput(sheet, 3, 0, "Alice");
    sheet = enableAutoFilter(sheet, {
      startRow: 0,
      startCol: 0,
      endRow: 3,
      endCol: 0,
    });
    sheet = setColumnFilter(
      sheet,
      { col: 0, operator: "equals", value: "Alice" },
      () => "",
    );
    expect(sheet.autoFilter?.hiddenRows).toContain(2);
    expect(sheet.autoFilter?.hiddenRows).not.toContain(1);
  });
});