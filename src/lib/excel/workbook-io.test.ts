import ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { cellKey } from "@/lib/spreadsheet-types";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { setCellInput } from "@/lib/workbook/types";
import { parseExcelWorkbook, workbookToExcelBuffer } from "@/lib/excel/workbook-io";

const fixtureXlsx = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "fixtures",
  "untitled-spreadsheet.xlsx",
);

function cellComparableValue(cell: ExcelJS.Cell): string | number | undefined {
  const v = cell.value;
  if (v == null) return undefined;
  if (typeof v === "object" && "formula" in v) {
    return `=${String((v as { formula: string }).formula)}`;
  }
  if (typeof v === "object" && "richText" in v) {
    const parts = (v as ExcelJS.CellRichTextValue).richText.map((p) => p.text).join("");
    return parts || undefined;
  }
  if (typeof v === "object" && "result" in v) {
    const r = (v as ExcelJS.CellFormulaValue).result;
    if (r != null && typeof r !== "object") return r;
  }
  if (typeof v === "number" || typeof v === "string" || typeof v === "boolean") {
    return v;
  }
  return String(v);
}

function collectWorksheetValues(ws: ExcelJS.Worksheet): Map<string, string | number> {
  const out = new Map<string, string | number>();
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const value = cellComparableValue(cell);
      if (value === undefined) return;
      out.set(`${rowNumber}:${colNumber}`, value);
    });
  });
  return out;
}

async function buildFilledWorkbookBuffer(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Data");
  const valueCell = ws.getCell("A1");
  valueCell.value = "Revenue";
  const filledOnly = ws.getCell("B1");
  filledOnly.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFF0000" },
  };
  ws.mergeCells("C1:D1");
  const merged = ws.getCell("C1");
  merged.value = "Merged title";
  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

describe("workbook-io", () => {
  it("imports format-only cells and merges", async () => {
    const buffer = await buildFilledWorkbookBuffer();
    const doc = await parseExcelWorkbook(buffer);
    const sheet = doc.sheets[0]!;

    expect(sheet.cells[cellKey(0, 0)]?.input).toBe("Revenue");
    expect(sheet.cells[cellKey(0, 1)]?.format?.fillColor?.toUpperCase()).toBe("#FF0000");
    expect(sheet.merges).toEqual([
      { startRow: 0, startCol: 2, endRow: 0, endCol: 3 },
    ]);
    expect(sheet.cells[cellKey(0, 2)]?.input).toBe("Merged title");
  });

  it("imports theme-only font colors", async () => {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Theme");
    const cell = ws.getCell("A1");
    cell.value = "Accent";
    cell.font = { color: { theme: 4 } };
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
    const doc = await parseExcelWorkbook(buffer);
    expect(doc.sheets[0]!.cells[cellKey(0, 0)]?.format?.fontColor?.toUpperCase()).toBe(
      "#4F81BD",
    );
  });

  it("round-trips Excel tables and conditional formatting", async () => {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Features");
    ws.getCell("A1").value = "Item";
    ws.getCell("B1").value = "Qty";
    ws.getCell("A2").value = "Widget";
    ws.getCell("B2").value = 5;
    ws.addTable({
      name: "Inventory",
      ref: "A1:B2",
      columns: [{ name: "Item" }, { name: "Qty" }],
      rows: [["Widget", 5]],
    });
    ws.addConditionalFormatting({
      ref: "B2:B20",
      rules: [
        {
          type: "cellIs",
          operator: "greaterThan",
          formulae: [3],
          priority: 1,
          style: {
            fill: {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFFFFF00" },
            },
          },
        },
      ],
    });
    const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
    const imported = await parseExcelWorkbook(buffer);
    const sheet = imported.sheets[0]!;
    expect(sheet.tables[0]?.name).toBe("Inventory");
    expect(sheet.tables[0]?.ref).toBe("A1:B2");
    expect(sheet.conditionalFormats[0]?.ref).toBe("B2:B20");
    expect(sheet.conditionalFormats[0]?.rules[0]?.type).toBe("cellIs");

    const xlsx = await workbookToExcelBuffer(imported);
    const again = await parseExcelWorkbook(xlsx);
    expect(again.sheets[0]!.tables[0]?.name).toBe("Inventory");
    expect(again.sheets[0]!.conditionalFormats.length).toBe(1);
  });

  it("exports charts as worksheet images", async () => {
    let doc = createEmptyWorkbook("ChartExport");
    let sheet = doc.sheets[0]!;
    sheet = {
      ...sheet,
      cells: {
        ...sheet.cells,
        [cellKey(0, 0)]: { input: "Category" },
        [cellKey(0, 1)]: { input: "Value" },
        [cellKey(1, 0)]: { input: "A" },
        [cellKey(1, 1)]: { input: "10" },
        [cellKey(2, 0)]: { input: "B" },
        [cellKey(2, 1)]: { input: "20" },
      },
      charts: [
        {
          id: "c1",
          type: "bar",
          title: "Sales",
          range: { startRow: 0, startCol: 0, endRow: 2, endCol: 1 },
          anchor: { startRow: 4, startCol: 0, endRow: 15, endCol: 5 },
        },
      ],
      merges: [{ startRow: 4, startCol: 0, endRow: 15, endCol: 5 }],
    };
    doc = { ...doc, sheets: [sheet] };
    const buffer = await workbookToExcelBuffer(doc);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet("ChartExport");
    const media = ws?.model.media as unknown[] | undefined;
    expect(media?.length).toBeGreaterThan(0);
  });

  it("round-trips workbook values, fill, and merges", async () => {
    let doc = createEmptyWorkbook("Roundtrip");
    let sheet = doc.sheets[0]!;
    sheet = setCellInput(sheet, 0, 0, "Total");
    sheet = {
      ...sheet,
      cells: {
        ...sheet.cells,
        [cellKey(0, 1)]: {
          format: { fillColor: "#00B050" },
        },
        [cellKey(1, 0)]: { input: "42" },
      },
      merges: [{ startRow: 0, startCol: 2, endRow: 1, endCol: 2 }],
    };
    doc = { ...doc, sheets: [sheet] };

    const xlsx = await workbookToExcelBuffer(doc);
    const imported = await parseExcelWorkbook(xlsx);
    const out = imported.sheets[0]!;

    expect(out.cells[cellKey(0, 0)]?.input).toBe("Total");
    expect(out.cells[cellKey(0, 1)]?.format?.fillColor?.toUpperCase()).toBe("#00B050");
    expect(out.cells[cellKey(1, 0)]?.input).toBe("42");
    expect(out.merges.length).toBe(1);
  });

  it("round-trips the example spreadsheet without losing values, merges, or images", async () => {
    if (!fs.existsSync(fixtureXlsx)) return;
    const originalBuf = fs.readFileSync(fixtureXlsx);

    const wbOrig = new ExcelJS.Workbook();
    await wbOrig.xlsx.load(originalBuf);
    const wsOrig = wbOrig.worksheets[0]!;

    const doc = await parseExcelWorkbook(originalBuf);
    const sheet = doc.sheets[0]!;
    expect(sheet.embeddedImages.length).toBe(wsOrig.model.media?.length ?? 0);
    expect(sheet.merges.length).toBe(wsOrig.model.merges?.length ?? 0);

    const exported = await workbookToExcelBuffer(doc);
    const wbRound = new ExcelJS.Workbook();
    await wbRound.xlsx.load(exported);
    const wsRound = wbRound.worksheets[0]!;

    expect(wsRound.model.media?.length ?? 0).toBe(wsOrig.model.media?.length ?? 0);

    const origValues = collectWorksheetValues(wsOrig);
    const roundValues = collectWorksheetValues(wsRound);
    expect(roundValues.size).toBe(origValues.size);
    for (const [key, value] of origValues) {
      expect(roundValues.get(key)).toEqual(value);
    }

    const origImageLens = (wsOrig.model.media ?? []).map((_, idx) =>
      wbOrig.getImage(idx)?.buffer?.length ?? 0,
    );
    const roundImageLens = (wsRound.model.media ?? []).map((_, idx) =>
      wbRound.getImage(idx)?.buffer?.length ?? 0,
    );
    expect(roundImageLens).toEqual(origImageLens);
  });
});