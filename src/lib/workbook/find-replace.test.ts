import { describe, expect, it } from "vitest";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { findInSheet, replaceInSheet } from "@/lib/workbook/find-replace";
import { setCellInput } from "@/lib/workbook/types";

describe("find and replace", () => {
  it("finds and replaces cell values", () => {
    const wb = createEmptyWorkbook();
    let sheet = wb.sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "Hello World");
    sheet = setCellInput(sheet, 1, 1, "hello again");
    const matches = findInSheet(sheet, "hello", false);
    expect(matches).toHaveLength(2);
    const result = replaceInSheet(sheet, "hello", "Hi", false, true);
    expect(result.count).toBe(2);
    expect(result.sheet.cells["0:0"]?.input).toBe("Hi World");
  });
});