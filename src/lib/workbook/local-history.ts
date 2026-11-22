import type { WorkbookOp } from "@/lib/workbook/ops";

export type LocalHistoryEntry = {
  undo: WorkbookOp[];
  redo: WorkbookOp[];
};

export function pushLocalHistory(
  stack: LocalHistoryEntry[],
  entry: LocalHistoryEntry,
  max = 50,
) {
  stack.push(entry);
  if (stack.length > max) stack.shift();
}