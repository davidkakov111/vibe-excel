export interface GridSelection {
  row: number;
  col: number;
  endRow: number;
  endCol: number;
}

export function normalizeSelection(selection: GridSelection): GridSelection {
  return {
    row: Math.min(selection.row, selection.endRow),
    col: Math.min(selection.col, selection.endCol),
    endRow: Math.max(selection.row, selection.endRow),
    endCol: Math.max(selection.col, selection.endCol),
  };
}