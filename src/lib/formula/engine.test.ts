import { describe, expect, it } from "vitest";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { setCellInput } from "@/lib/workbook/types";
import { FormulaEngine, parseA1 } from "@/lib/formula/engine";
import { createWorksheet } from "@/lib/workbook/defaults";
import { insertRows } from "@/lib/workbook/structure";

describe("parseA1", () => {
  it("parses references", () => {
    expect(parseA1("A1")).toEqual({ row: 0, col: 0 });
    expect(parseA1("$B$10")).toEqual({ row: 9, col: 1 });
  });
});

describe("FormulaEngine", () => {
  it("evaluates SUM and nested IF", () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0]!;
    let s = setCellInput(sheet, 0, 0, "10");
    s = setCellInput(s, 1, 0, "20");
    s = setCellInput(s, 2, 0, "=SUM(A1:A2)");
    s = setCellInput(s, 3, 0, '=IF(A3>25,"high","low")');
    wb.sheets[0] = s;

    const engine = FormulaEngine.fromWorkbook(wb);
    expect(engine.getDisplayValue(sheet.id, 2, 0)).toBe("30");
    expect(engine.getDisplayValue(sheet.id, 3, 0)).toBe("high");
    engine.destroy();
  });

  it("recalculates dependents", () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0]!;
    let s = setCellInput(sheet, 0, 0, "5");
    s = setCellInput(s, 0, 1, "=A1*2");
    wb.sheets[0] = s;
    const engine = FormulaEngine.fromWorkbook(wb);
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("10");
    engine.setCellInput(sheet.id, 0, 0, "7");
    expect(engine.getDisplayValue(sheet.id, 0, 1)).toBe("14");
    engine.destroy();
  });

  it("supports cross-sheet references", () => {
    const wb = createEmptyWorkbook();
    const sheet2 = createWorksheet("Sheet2");
    wb.sheets.push(sheet2);
    const s1 = setCellInput(wb.sheets[0]!, 0, 0, "3");
    const s2 = setCellInput(sheet2, 0, 0, "=Sheet1!A1+1");
    wb.sheets[0] = s1;
    wb.sheets[1] = s2;

    const engine = FormulaEngine.fromWorkbook(wb);
    expect(engine.getDisplayValue(sheet2.id, 0, 0)).toBe("4");
    engine.destroy();
  });

  it("detects circular references", () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0]!;
    let s = setCellInput(sheet, 0, 0, "=B1");
    s = setCellInput(s, 0, 1, "=A1");
    wb.sheets[0] = s;
    const engine = FormulaEngine.fromWorkbook(wb);
    const v = engine.getDisplayValue(sheet.id, 0, 0);
    expect(v).toMatch(/CYCLE|cycle|#/i);
    engine.destroy();
  });

  it("supports COUNTIF and VLOOKUP", () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0]!;
    let s = setCellInput(sheet, 0, 0, "apple");
    s = setCellInput(s, 1, 0, "pear");
    s = setCellInput(s, 2, 0, "apple");
    s = setCellInput(s, 0, 2, '=COUNTIF(A1:A3,"apple")');
    s = setCellInput(s, 1, 2, "1");
    s = setCellInput(s, 1, 3, "alpha");
    s = setCellInput(s, 2, 2, "2");
    s = setCellInput(s, 2, 3, "beta");
    s = setCellInput(s, 0, 4, "=VLOOKUP(2,C2:D3,2,0)");
    wb.sheets[0] = s;
    const engine = FormulaEngine.fromWorkbook(wb);
    expect(engine.getDisplayValue(sheet.id, 0, 2)).toBe("2");
    expect(engine.getDisplayValue(sheet.id, 0, 4)).toBe("beta");
    engine.destroy();
  });
});

describe("insertRows with formulas", () => {
  it("keeps cell data aligned after structural edit + rebuild", () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0]!;
    let s = setCellInput(sheet, 0, 0, "1");
    s = setCellInput(s, 1, 0, "=A1+1");
    const inserted = insertRows(s, 1, 1);
    wb.sheets[0] = inserted;

    const engine = FormulaEngine.fromWorkbook(wb);
    expect(engine.getDisplayValue(sheet.id, 2, 0)).toBe("2");
    engine.destroy();
  });
});