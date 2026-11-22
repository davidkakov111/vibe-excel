import { produce } from "immer";
import { describe, expect, it } from "vitest";
import { cellKey } from "@/lib/spreadsheet-types";
import { createCommentThread } from "@/lib/workbook/comment";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { sortWorksheetRange } from "@/lib/workbook/sort";
import { setCellInput } from "@/lib/workbook/types";

describe("sortWorksheetRange", () => {
  it("reorders row values within range", () => {
    let sheet = createEmptyWorkbook().sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "3");
    sheet = setCellInput(sheet, 1, 0, "1");
    sheet = setCellInput(sheet, 2, 0, "2");
    sheet = setCellInput(sheet, 0, 1, "x");
    sheet = setCellInput(sheet, 1, 1, "y");
    sheet = setCellInput(sheet, 2, 1, "z");

    const sorted = sortWorksheetRange(
      sheet,
      0,
      2,
      0,
      sheet.colCount - 1,
      0,
      true,
    );

    expect(sorted.cells["0:0"]?.input).toBe("1");
    expect(sorted.cells["1:0"]?.input).toBe("2");
    expect(sorted.cells["2:0"]?.input).toBe("3");
    expect(sorted.cells["0:1"]?.input).toBe("y");
    expect(sorted.cells["1:1"]?.input).toBe("z");
    expect(sorted.cells["2:1"]?.input).toBe("x");
  });

  it("applies through immer produce", () => {
    const workbook = createEmptyWorkbook();
    const sheetId = workbook.sheets[0]!.id;
    workbook.sheets[0] = setCellInput(workbook.sheets[0]!, 0, 0, "30");
    workbook.sheets[0] = setCellInput(workbook.sheets[0]!, 1, 0, "10");
    workbook.sheets[0] = setCellInput(workbook.sheets[0]!, 2, 0, "20");

    const next = produce(workbook, (draft) => {
      const index = draft.sheets.findIndex((s) => s.id === sheetId);
      const sorted = sortWorksheetRange(
        draft.sheets[index]!,
        0,
        2,
        0,
        draft.sheets[index]!.colCount - 1,
        0,
        true,
      );
      draft.sheets[index]!.cells = sorted.cells;
    });

    expect(next.sheets[0]?.cells["0:0"]?.input).toBe("10");
    expect(next.sheets[0]?.cells["1:0"]?.input).toBe("20");
    expect(next.sheets[0]?.cells["2:0"]?.input).toBe("30");
  });

  it("preserves comments and formats when permuting rows", () => {
    let sheet = createEmptyWorkbook().sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "b");
    sheet = setCellInput(sheet, 1, 0, "a");
    sheet = {
      ...sheet,
      cells: {
        ...sheet.cells,
        [cellKey(0, 0)]: {
          ...sheet.cells[cellKey(0, 0)]!,
          format: { bold: true },
          comment: createCommentThread("u", "Ann", "note on b"),
        },
        [cellKey(1, 0)]: {
          ...sheet.cells[cellKey(1, 0)]!,
          format: { fillColor: "#FF0000" },
        },
      },
    };
    const sorted = sortWorksheetRange(sheet, 0, 1, 0, 3, 0, true);
    expect(sorted.cells[cellKey(0, 0)]?.input).toBe("a");
    expect(sorted.cells[cellKey(1, 0)]?.input).toBe("b");
    expect(sorted.cells[cellKey(0, 0)]?.format?.fillColor).toBe("#FF0000");
    expect(sorted.cells[cellKey(1, 0)]?.comment?.messages[0]?.text).toBe(
      "note on b",
    );
  });
});