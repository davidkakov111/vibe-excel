import { describe, expect, it } from "vitest";
import { normalizeStoredDocument } from "@/lib/workbook/migrate";
import { WORKBOOK_VERSION } from "@/lib/workbook/types";

describe("normalizeStoredDocument", () => {
  it("migrates legacy flat cells", () => {
    const doc = normalizeStoredDocument(
      { "0:0": "hello", "1:1": "42" },
      { cells: {}, rowCount: 50, colCount: 10 },
    );
    expect(doc.version).toBe(WORKBOOK_VERSION);
    expect(doc.sheets[0]?.cells["0:0"]?.input).toBe("hello");
    expect(doc.sheets[0]?.cells["1:1"]?.input).toBe("42");
  });
});