import { describe, expect, it } from "vitest";
import {
  cellKey,
  indexToColumnLabel,
  parseCellKey,
} from "@/lib/spreadsheet-types";

describe("indexToColumnLabel", () => {
  it("maps column indices to Excel labels", () => {
    expect(indexToColumnLabel(0)).toBe("A");
    expect(indexToColumnLabel(25)).toBe("Z");
    expect(indexToColumnLabel(26)).toBe("AA");
  });
});

describe("cellKey", () => {
  it("round-trips row/col coordinates", () => {
    const key = cellKey(4, 7);
    expect(parseCellKey(key)).toEqual({ row: 4, col: 7 });
  });
});