import { describe, expect, it } from "vitest";
import { validateCellValue } from "@/lib/workbook/validation";

describe("validateCellValue", () => {
  it("accepts list values and rejects others", () => {
    const validation = { kind: "list" as const, values: ["A", "B"], allowBlank: true };
    expect(validateCellValue("A", validation).valid).toBe(true);
    expect(validateCellValue("C", validation).valid).toBe(false);
  });

  it("validates whole numbers with bounds", () => {
    const validation = { kind: "whole" as const, min: 1, max: 10, allowBlank: false };
    expect(validateCellValue("5", validation).valid).toBe(true);
    expect(validateCellValue("11", validation).valid).toBe(false);
    expect(validateCellValue("3.5", validation).valid).toBe(false);
  });
});