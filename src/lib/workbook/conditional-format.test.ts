import { describe, expect, it } from "vitest";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { effectiveCellFormat } from "@/lib/workbook/conditional-format";

describe("conditional-format", () => {
  it("applies cellIs rule fill when value matches", () => {
    const sheet = createEmptyWorkbook("S").sheets[0]!;
    sheet.conditionalFormats = [
      {
        ref: "A1:A10",
        rules: [
          {
            type: "cellIs",
            operator: "greaterThan",
            priority: 1,
            formulae: ["10"],
            style: {
              fill: {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "FF00FF00" },
              },
            },
          },
        ],
      },
    ];
    const fmt = effectiveCellFormat(sheet, 0, 0, undefined, {
      input: "42",
      display: "42",
    });
    expect(fmt?.fillColor?.toUpperCase()).toBe("#00FF00");
  });
});