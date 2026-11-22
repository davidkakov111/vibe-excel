import { describe, expect, it } from "vitest";
import { reconcileWorkbookWithOps } from "@/lib/collaboration/op-sync";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { sortRangeOp } from "@/lib/workbook/ops";
import { setCellInput } from "@/lib/workbook/types";

describe("reconcileWorkbookWithOps", () => {
  it("applies sort on stale server snapshot instead of reverting", () => {
    const server = createEmptyWorkbook();
    const sheetId = server.sheets[0]!.id;
    server.sheets[0] = setCellInput(server.sheets[0]!, 0, 0, "b");
    server.sheets[0] = setCellInput(server.sheets[0]!, 1, 0, "a");

    const sortOp = sortRangeOp(sheetId, 0, 1, 0, 5, 0, true);
    const merged = reconcileWorkbookWithOps(server, [sortOp]);
    expect(merged.sheets[0]!.cells["0:0"]?.input).toBe("a");
    expect(merged.sheets[0]!.cells["1:0"]?.input).toBe("b");
  });
});