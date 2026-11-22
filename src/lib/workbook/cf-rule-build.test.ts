import { describe, expect, it } from "vitest";
import {
  buildConditionalRule,
  selectionToA1,
  upsertConditionalFormatSet,
} from "@/lib/workbook/cf-rule-build";

describe("cf-rule-build", () => {
  it("builds A1 range from selection", () => {
    expect(
      selectionToA1({ row: 0, col: 0, endRow: 1, endCol: 2 }),
    ).toBe("A1:C2");
  });

  it("upserts conditional format rules for a range", () => {
    const rule = buildConditionalRule({
      preset: "greaterThan",
      value: "10",
      fillColor: "#ff0000",
      priority: 1,
    });
    const sets = upsertConditionalFormatSet([], "A1:B10", [rule]);
    expect(sets[0]?.rules[0]).toMatchObject({ type: "cellIs", operator: "greaterThan" });
  });
});