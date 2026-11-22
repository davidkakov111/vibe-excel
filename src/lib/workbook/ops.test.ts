import { describe, expect, it } from "vitest";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { cellKey } from "@/lib/spreadsheet-types";
import {
  applyWorkbookOps,
  formatOp,
  patchSheetOp,
  sortRangeOp,
} from "@/lib/workbook/ops";
import { setCellInput } from "@/lib/workbook/types";

describe("workbook ops batch", () => {
  it("applies format to multiple cells in one pass", () => {
    const doc = createEmptyWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const ops = [
      formatOp(sheetId, 0, 0, { fillColor: "#FF0000" }),
      formatOp(sheetId, 0, 1, { fillColor: "#FF0000" }),
      formatOp(sheetId, 1, 0, { fillColor: "#FF0000" }),
    ];
    const next = applyWorkbookOps(doc, ops);
    const sheet = next.sheets[0]!;
    expect(sheet.cells[cellKey(0, 0)]?.format?.fillColor).toBe("#FF0000");
    expect(sheet.cells[cellKey(0, 1)]?.format?.fillColor).toBe("#FF0000");
    expect(sheet.cells[cellKey(1, 0)]?.format?.fillColor).toBe("#FF0000");
  });
});

describe("patchSheet op", () => {
  it("persists conditional formatting rules on the sheet", () => {
    const doc = createEmptyWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const op = patchSheetOp(sheetId, {
      conditionalFormats: [
        {
          ref: "A1:B2",
          rules: [
            {
              type: "cellIs",
              operator: "greaterThan",
              formulae: [10],
              priority: 1,
              style: { fill: { type: "pattern", pattern: "solid" } },
            },
          ],
        },
      ],
    });
    const next = applyWorkbookOps(doc, [op]);
    expect(next.sheets[0]!.conditionalFormats).toHaveLength(1);
    expect(next.sheets[0]!.conditionalFormats[0]?.ref).toBe("A1:B2");
  });

  it("updates charts and layout for collaborators", () => {
    const doc = createEmptyWorkbook();
    const sheetId = doc.sheets[0]!.id;
    const op = patchSheetOp(sheetId, {
      charts: [
        {
          id: "c1",
          type: "bar",
          title: "Revenue",
          range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
          anchor: { startRow: 5, startCol: 0, endRow: 12, endCol: 5 },
        },
      ],
      rowCount: 120,
    });
    const next = applyWorkbookOps(doc, [op]);
    expect(next.sheets[0]!.charts[0]?.title).toBe("Revenue");
    expect(next.sheets[0]!.rowCount).toBe(120);
  });
});

describe("sortRange op", () => {
  it("sorts rows via collaborative op", () => {
    let doc = createEmptyWorkbook();
    const sheetId = doc.sheets[0]!.id;
    let sheet = doc.sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "b");
    sheet = setCellInput(sheet, 1, 0, "a");
    sheet = setCellInput(sheet, 2, 0, "c");
    doc = { ...doc, sheets: [sheet] };
    const op = sortRangeOp(sheetId, 0, 2, 0, 5, 0, true);
    const next = applyWorkbookOps(doc, [op]);
    expect(next.sheets[0]!.cells[cellKey(0, 0)]?.input).toBe("a");
    expect(next.sheets[0]!.cells[cellKey(1, 0)]?.input).toBe("b");
    expect(next.sheets[0]!.cells[cellKey(2, 0)]?.input).toBe("c");
  });
});