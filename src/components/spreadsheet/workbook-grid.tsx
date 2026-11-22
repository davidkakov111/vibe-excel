"use client";

import { useVirtualizer } from "@tanstack/react-virtual";
import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
  type MutableRefObject,
  type ReactNode,
  type RefObject,
} from "react";
import type { FilterColumnState } from "@/lib/workbook/types";
import {
  remoteHighlightsAtCell,
  type RemotePresence,
} from "@/lib/collaboration/presence";
import type { GridSelection } from "@/lib/workbook/selection";
export type { GridSelection } from "@/lib/workbook/selection";
import {
  cellStyleFromFormat,
  formatDisplayValue,
} from "@/lib/workbook/format-display";
import { effectiveCellFormat } from "@/lib/workbook/conditional-format";
import {
  CLIPBOARD_MIME,
  parseClipboardPayload,
  payloadFromPlainTsv,
  plainTextFromPayload,
  serializeClipboardPayload,
  type ClipboardPayload,
} from "@/lib/workbook/clipboard";
import type { Worksheet } from "@/lib/workbook/types";
import { getCellInput } from "@/lib/workbook/types";
import { cellKey, indexToColumnLabel } from "@/lib/spreadsheet-types";
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from "@/lib/workbook/defaults";
import {
  ColumnFilterMenu,
  type ColumnFilterMenuState,
} from "@/components/spreadsheet/column-filter-menu";
import { SheetChartOverlays } from "@/components/spreadsheet/sheet-chart-overlay";
import { CellCommentPanel } from "@/components/spreadsheet/cell-comment-panel";
import {
  normalizeCellComment,
  threadAuthorLabel,
} from "@/lib/workbook/comment";

const ROW_HEADER_WIDTH = 46;

interface WorkbookGridProps {
  sheet: Worksheet;
  sheetId: string;
  zoom: number;
  showGridlines: boolean;
  getDisplay: (sheetId: string, row: number, col: number) => string;
  selection: GridSelection;
  onSelectionChange: (selection: GridSelection) => void;
  onCommitCell: (row: number, col: number, value: string) => void;
  onColResize: (col: number, width: number) => void;
  onRowResize: (row: number, height: number) => void;
  onFillEnd: (target: GridSelection) => void;
  onCut: () => void;
  onPasteCells?: (startRow: number, startCol: number, cells: ClipboardPayload["cells"]) => void;
  onApplyColumnFilter: (filter: FilterColumnState) => void;
  onSortColumn: (col: number, ascending: boolean) => void;
  onClearColumnFilter: (col: number) => void;
  readOnly?: boolean;
  remotePresence?: RemotePresence[];
  commentUserId: string;
  onCommentStart: (row: number, col: number, text: string) => void;
  onCommentReply: (row: number, col: number, text: string) => void;
  onCommentResolve: (row: number, col: number) => void;
  onCommentReopen: (row: number, col: number) => void;
  onCommentDelete: (row: number, col: number) => void;
  onEnsureCapacity?: (minRow?: number, minCol?: number) => void;
}

export type WorkbookGridHandle = {
  commitPendingEdit: () => void;
  openCommentOnSelection: () => void;
};

function isCommentShortcut(event: KeyboardEvent): boolean {
  const key = event.key.toLowerCase();
  if (key !== "m" && key !== "c") return false;
  if (!(event.ctrlKey || event.metaKey) || !event.shiftKey) return false;
  return true;
}

function mergeAt(sheet: Worksheet, row: number, col: number) {
  return sheet.merges.find(
    (m) =>
      row >= m.startRow &&
      row <= m.endRow &&
      col >= m.startCol &&
      col <= m.endCol,
  );
}

function isHiddenByFilter(sheet: Worksheet, row: number): boolean {
  return sheet.autoFilter?.hiddenRows.includes(row) ?? false;
}

export const WorkbookGrid = forwardRef<WorkbookGridHandle, WorkbookGridProps>(
  function WorkbookGrid(
    {
      sheet,
      sheetId,
      zoom,
      showGridlines,
      getDisplay,
      selection,
      onSelectionChange,
      onCommitCell,
      onColResize,
      onRowResize,
      onFillEnd,
      onCut,
      onPasteCells,
      onApplyColumnFilter,
      onSortColumn,
      onClearColumnFilter,
      readOnly = false,
      remotePresence = [],
      commentUserId,
      onCommentStart,
      onCommentReply,
      onCommentResolve,
      onCommentReopen,
      onCommentDelete,
      onEnsureCapacity,
    },
    ref,
  ) {
  const [filterMenu, setFilterMenu] = useState<ColumnFilterMenuState | null>(
    null,
  );
  const [editing, setEditing] = useState<{ row: number; col: number } | null>(null);
  const [draft, setDraft] = useState("");
  const [fillDrag, setFillDrag] = useState<GridSelection | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const selectRef = useRef<HTMLSelectElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const dragAnchor = useRef<{ row: number; col: number } | null>(null);
  const scale = zoom / 100;
  const [commentPanel, setCommentPanel] = useState<{ row: number; col: number } | null>(
    null,
  );

  const visibleRows = useMemo(
    () =>
      Array.from({ length: sheet.rowCount }, (_, row) => row).filter(
        (row) =>
          !sheet.hiddenRows.includes(row) && !isHiddenByFilter(sheet, row),
      ),
    [sheet],
  );

  const visibleCols = useMemo(
    () =>
      Array.from({ length: sheet.colCount }, (_, col) => col).filter(
        (col) => !sheet.hiddenCols.includes(col),
      ),
    [sheet.colCount, sheet.hiddenCols],
  );

  const colWidth = useCallback(
    (col: number) => (sheet.colWidths[String(col)] ?? DEFAULT_COL_WIDTH) * scale,
    [sheet.colWidths, scale],
  );
  const rowHeight = useCallback(
    (row: number) => (sheet.rowHeights[String(row)] ?? DEFAULT_ROW_HEIGHT) * scale,
    [sheet.rowHeights, scale],
  );

  const frozenRowsList = useMemo(
    () => visibleRows.filter((row) => row < sheet.frozenRows),
    [visibleRows, sheet.frozenRows],
  );
  const scrollRows = useMemo(
    () => visibleRows.filter((row) => row >= sheet.frozenRows),
    [visibleRows, sheet.frozenRows],
  );

  const frozenHeight = useMemo(
    () => frozenRowsList.reduce((sum, row) => sum + rowHeight(row), 0),
    [frozenRowsList, rowHeight],
  );

  const columnsWidth = useMemo(
    () => visibleCols.reduce((sum, col) => sum + colWidth(col), 0),
    [visibleCols, colWidth],
  );
  const rowHeaderWidth = ROW_HEADER_WIDTH * scale;
  const sheetWidth = rowHeaderWidth + columnsWidth;

  const rowVirtualizer = useVirtualizer({
    count: scrollRows.length,
    getScrollElement: () => scrollRef.current,
    estimateSize: (index) => rowHeight(scrollRows[index] ?? 0),
    overscan: 8,
  });

  useEffect(() => {
    rowVirtualizer.measure();
  }, [rowVirtualizer, scale, sheet.rowHeights, sheet.frozenRows]);

  useEffect(() => {
    const el = scrollRef.current;
    if (!el || !onEnsureCapacity) return;
    let raf = 0;
    const onScroll = () => {
      if (raf) return;
      raf = window.requestAnimationFrame(() => {
        raf = 0;
        const nearBottom =
          el.scrollHeight - el.scrollTop - el.clientHeight < 320;
        const nearRight =
          el.scrollWidth - el.scrollLeft - el.clientWidth < 280;
        const items = rowVirtualizer.getVirtualItems();
        const lastScrollRow = items.length
          ? scrollRows[items[items.length - 1]!.index]
          : scrollRows[scrollRows.length - 1];
        if (nearBottom && lastScrollRow !== undefined) {
          onEnsureCapacity(lastScrollRow + 1, undefined);
        }
        if (nearRight) {
          const maxVisibleCol = visibleCols[visibleCols.length - 1];
          if (maxVisibleCol !== undefined) {
            onEnsureCapacity(undefined, maxVisibleCol + 1);
          }
        }
      });
    };
    el.addEventListener("scroll", onScroll, { passive: true });
    return () => {
      el.removeEventListener("scroll", onScroll);
      if (raf) window.cancelAnimationFrame(raf);
    };
  }, [onEnsureCapacity, rowVirtualizer, scrollRows, visibleCols]);

  const getRowTopInFrozen = useCallback(
    (row: number) => {
      let top = 0;
      for (const r of frozenRowsList) {
        if (r >= row) break;
        top += rowHeight(r);
      }
      return top;
    },
    [frozenRowsList, rowHeight],
  );

  const getRowTopInScroll = useCallback(
    (row: number) => {
      const index = scrollRows.indexOf(row);
      if (index < 0) return 0;
      const virtual = rowVirtualizer
        .getVirtualItems()
        .find((item) => item.index === index);
      return virtual?.start ?? 0;
    },
    [rowVirtualizer, scrollRows],
  );

  const commitDraft = useCallback(
    (commit: boolean) => {
      if (!editing) return;
      if (commit) {
        onCommitCell(editing.row, editing.col, draft);
      }
      setEditing(null);
    },
    [draft, editing, onCommitCell],
  );

  const openCommentOnSelection = useCallback(() => {
    if (readOnly) return;
  const s = normalizeSelection(selection);
    if (editing) commitDraft(true);
    setCommentPanel({ row: s.row, col: s.col });
  }, [commitDraft, editing, readOnly, selection]);

  useImperativeHandle(
    ref,
    () => ({
      commitPendingEdit: () => commitDraft(true),
      openCommentOnSelection,
    }),
    [commitDraft, openCommentOnSelection],
  );

  const requestColumnSort = useCallback(
    (col: number, ascending: boolean) => {
      if (editing) commitDraft(true);
      onSortColumn(col, ascending);
    },
    [commitDraft, editing, onSortColumn],
  );

  const resolveAnchor = useCallback(
    (row: number, col: number) => {
      const merge = mergeAt(sheet, row, col);
      if (merge) {
        return { row: merge.startRow, col: merge.startCol };
      }
      return { row, col };
    },
    [sheet],
  );

  const beginEdit = useCallback(
    (row: number, col: number) => {
      if (readOnly) return;
      const anchor = resolveAnchor(row, col);
      onSelectionChange({
        row: anchor.row,
        col: anchor.col,
        endRow: anchor.row,
        endCol: anchor.col,
      });
      setDraft(getCellInput(sheet, anchor.row, anchor.col));
      setEditing({ row: anchor.row, col: anchor.col });
    },
    [onSelectionChange, readOnly, resolveAnchor, sheet],
  );

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus();
      inputRef.current?.select();
      selectRef.current?.focus();
    }
  }, [editing]);

  useEffect(() => {
    if (!fillDrag) return;
    const end = () => {
      onFillEnd(fillDrag);
      setFillDrag(null);
    };
    window.addEventListener("mouseup", end);
    return () => window.removeEventListener("mouseup", end);
  }, [fillDrag, onFillEnd]);

  const normalizeSelection = (sel: GridSelection): GridSelection => ({
    row: Math.min(sel.row, sel.endRow),
    col: Math.min(sel.col, sel.endCol),
    endRow: Math.max(sel.row, sel.endRow),
    endCol: Math.max(sel.col, sel.endCol),
  });

  const selectEntireRow = useCallback(
    (row: number, extend: boolean) => {
      if (editing) commitDraft(true);
      const s = normalizeSelection(selection);
      onSelectionChange(
        extend
          ? { row: s.row, col: 0, endRow: row, endCol: sheet.colCount - 1 }
          : { row, col: 0, endRow: row, endCol: sheet.colCount - 1 },
      );
    },
    [commitDraft, editing, onSelectionChange, selection, sheet.colCount],
  );

  const selectEntireColumn = useCallback(
    (col: number, extend: boolean) => {
      if (editing) commitDraft(true);
      const s = normalizeSelection(selection);
      onSelectionChange(
        extend
          ? { row: 0, col: s.col, endRow: sheet.rowCount - 1, endCol: col }
          : { row: 0, col, endRow: sheet.rowCount - 1, endCol: col },
      );
    },
    [commitDraft, editing, onSelectionChange, selection, sheet.rowCount],
  );

  const activeSelection = fillDrag ?? selection;

  const inSelection = useCallback(
    (row: number, col: number) => {
      const s = normalizeSelection(activeSelection);
      return row >= s.row && row <= s.endRow && col >= s.col && col <= s.endCol;
    },
    [activeSelection],
  );

  const onFillHover = useCallback(
    (row: number, col: number) => {
      if (!fillDrag) return;
      setFillDrag({
        row: Math.min(fillDrag.row, fillDrag.endRow),
        col: Math.min(fillDrag.col, fillDrag.endCol),
        endRow: row,
        endCol: col,
      });
    },
    [fillDrag],
  );

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (!readOnly && isCommentShortcut(event)) {
        event.preventDefault();
        openCommentOnSelection();
        return;
      }
      const target = event.target;
      if (target instanceof HTMLElement) {
        const tag = target.tagName;
        if (
          tag === "INPUT" ||
          tag === "TEXTAREA" ||
          tag === "SELECT" ||
          target.isContentEditable ||
          target.closest('[role="dialog"]')
        ) {
          return;
        }
      }

      if (editing) {
        if (event.key === "Enter") {
          event.preventDefault();
          commitDraft(true);
        }
        if (event.key === "Escape") {
          event.preventDefault();
          commitDraft(false);
        }
        if (event.key === "Tab") {
          event.preventDefault();
          commitDraft(true);
          const nextCol = event.shiftKey ? selection.col - 1 : selection.col + 1;
          const col = Math.max(0, Math.min(sheet.colCount - 1, nextCol));
          beginEdit(selection.row, col);
        }
        return;
      }

      const s = normalizeSelection(selection);
      let row = s.row;
      let col = s.col;
      if (event.key === "ArrowDown") row = Math.min(sheet.rowCount - 1, row + 1);
      if (event.key === "ArrowUp") row = Math.max(0, row - 1);
      if (event.key === "ArrowRight") col = Math.min(sheet.colCount - 1, col + 1);
      if (event.key === "ArrowLeft") col = Math.max(0, col - 1);
      if (event.key === "Enter") {
        event.preventDefault();
        beginEdit(s.row, s.col);
        return;
      }
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
        event.preventDefault();
        onSelectionChange({
          row: 0,
          col: 0,
          endRow: sheet.rowCount - 1,
          endCol: sheet.colCount - 1,
        });
        return;
      }
      if (event.key === "Delete" || event.key === "Backspace") {
        event.preventDefault();
        for (let r = s.row; r <= s.endRow; r++) {
          for (let c = s.col; c <= s.endCol; c++) {
            onCommitCell(r, c, "");
          }
        }
        return;
      }
      if (row !== s.row || col !== s.col) {
        event.preventDefault();
        onSelectionChange(
          event.shiftKey
            ? { row: s.row, col: s.col, endRow: row, endCol: col }
            : { row, col, endRow: row, endCol: col },
        );
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [
    beginEdit,
    commitDraft,
    editing,
    onCommitCell,
    onSelectionChange,
    selection,
    sheet,
    sheet.colCount,
    sheet.rowCount,
    openCommentOnSelection,
    readOnly,
  ]);

  const buildClipboardPayload = useCallback((): ClipboardPayload => {
    const s = normalizeSelection(selection);
    const cells: ClipboardPayload["cells"] = [];
    for (let r = s.row; r <= s.endRow; r++) {
      const row: ClipboardPayload["cells"][number] = [];
      for (let c = s.col; c <= s.endCol; c++) {
        const data = sheet.cells[cellKey(r, c)];
        row.push({
          input: getCellInput(sheet, r, c),
          format: data?.format ? structuredClone(data.format) : undefined,
        });
      }
      cells.push(row);
    }
    return { version: 1, cells };
  }, [selection, sheet]);

  const writeClipboard = useCallback(
    (clipboardData: DataTransfer | null) => {
      if (!clipboardData) return;
      const payload = buildClipboardPayload();
      clipboardData.setData("text/plain", plainTextFromPayload(payload));
      clipboardData.setData(CLIPBOARD_MIME, serializeClipboardPayload(payload));
    },
    [buildClipboardPayload],
  );

  const onCopy = useCallback(
    (event: ClipboardEvent) => {
      writeClipboard(event.clipboardData);
      event.preventDefault();
    },
    [writeClipboard],
  );

  const onPaste = useCallback(
    (event: ClipboardEvent) => {
      if (readOnly) return;
      const rich = event.clipboardData?.getData(CLIPBOARD_MIME);
      const text = event.clipboardData?.getData("text/plain");
      const payload =
        (rich && parseClipboardPayload(rich)) ||
        (text ? payloadFromPlainTsv(text) : null);
      if (!payload) return;
      event.preventDefault();
      const s = normalizeSelection(selection);
      if (onPasteCells) {
        onPasteCells(s.row, s.col, payload.cells);
        return;
      }
      payload.cells.forEach((row, rOffset) => {
        row.forEach((cell, cOffset) => {
          if (cell.input !== undefined) {
            onCommitCell(s.row + rOffset, s.col + cOffset, cell.input);
          }
        });
      });
    },
    [onCommitCell, onPasteCells, readOnly, selection],
  );

  const onCutClipboard = useCallback(
    (event: ClipboardEvent) => {
      if (readOnly) return;
      writeClipboard(event.clipboardData);
      onCut();
      event.preventDefault();
    },
    [onCut, readOnly, writeClipboard],
  );

  useEffect(() => {
    window.addEventListener("copy", onCopy);
    window.addEventListener("cut", onCutClipboard);
    window.addEventListener("paste", onPaste);
    return () => {
      window.removeEventListener("copy", onCopy);
      window.removeEventListener("cut", onCutClipboard);
      window.removeEventListener("paste", onPaste);
    };
  }, [onCopy, onCutClipboard, onPaste]);

  const headerHeight = 22 * scale;

  const chartRowTop = useCallback(
    (row: number) => {
      if (row < sheet.frozenRows) {
        return headerHeight + getRowTopInFrozen(row);
      }
      return headerHeight + frozenHeight + getRowTopInScroll(row);
    },
    [
      frozenHeight,
      getRowTopInFrozen,
      getRowTopInScroll,
      headerHeight,
      sheet.frozenRows,
    ],
  );

  const chartColOffset = useCallback(
    (col: number) => colOffsetsVisible(visibleCols, col, colWidth),
    [colWidth, visibleCols],
  );

  const activeColumnFilter = (col: number) =>
    sheet.autoFilter?.columns.find((c) => c.col === col);

  return (
    <div className="flex min-h-0 flex-1 flex-col overflow-hidden border border-[#b4b4b4] bg-[#e6e6e6] text-zinc-900">
      <ColumnFilterMenu
        sheet={sheet}
        sheetId={sheetId}
        menu={filterMenu}
        getDisplay={getDisplay}
        activeFilter={
          filterMenu ? activeColumnFilter(filterMenu.col) : undefined
        }
        onClose={() => setFilterMenu(null)}
        onSortAsc={(col) => requestColumnSort(col, true)}
        onSortDesc={(col) => requestColumnSort(col, false)}
        onApply={onApplyColumnFilter}
        onClearColumn={onClearColumnFilter}
      />
      <div
        ref={scrollRef}
        className="min-h-0 flex-1 overscroll-contain overflow-auto"
      >
        <div
          className="relative"
          style={{
            width: sheetWidth,
            minHeight:
              headerHeight + frozenHeight + rowVirtualizer.getTotalSize(),
          }}
        >
          <div
            className="sticky top-0 z-30 flex border-b border-[#b4b4b4] bg-[#f3f3f3]"
            style={{ width: sheetWidth, height: headerHeight }}
          >
            <div
              className="shrink-0 cursor-cell border-r border-[#b4b4b4] bg-[#f3f3f3]"
              style={{ width: rowHeaderWidth, height: headerHeight }}
              title="Select all"
              onMouseDown={(event) => {
                event.preventDefault();
                if (editing) commitDraft(true);
                onSelectionChange({
                  row: 0,
                  col: 0,
                  endRow: sheet.rowCount - 1,
                  endCol: sheet.colCount - 1,
                });
              }}
            />
            <div
              className="flex shrink-0"
              style={{ width: columnsWidth, height: headerHeight }}
            >
              {visibleCols.map((col) => (
                <ColHeader
                  key={col}
                  col={col}
                  width={colWidth(col)}
                  height={headerHeight}
                  filterActive={sheet.autoFilter?.columns.some(
                    (c) => c.col === col,
                  )}
                  showFilter={Boolean(sheet.autoFilter?.enabled)}
                  filterValues={
                    activeColumnFilter(col)?.operator === "oneOf"
                      ? activeColumnFilter(col)?.allowedValues
                      : undefined
                  }
                  onOpenFilter={(anchor) =>
                    setFilterMenu({
                      col,
                      top: anchor.bottom + 4,
                      left: anchor.left,
                    })
                  }
                  onResize={(w) => onColResize(col, w / scale)}
                  onSelect={(extend) => selectEntireColumn(col, extend)}
                />
              ))}
            </div>
          </div>

          {frozenHeight > 0 ? (
            <div
              className="sticky z-[25] flex shrink-0 border-b-2 border-[#217346] bg-[#f3f3f3]"
              style={{
                top: headerHeight,
                width: sheetWidth,
                height: frozenHeight,
              }}
            >
              <div
                className="relative shrink-0 border-r border-[#b4b4b4] bg-[#f3f3f3]"
                style={{ width: rowHeaderWidth, height: frozenHeight }}
              >
                {frozenRowsList.map((row) => (
                  <div
                    key={`frozen-rh-${row}`}
                    className="absolute left-0"
                    style={{
                      top: getRowTopInFrozen(row),
                      width: rowHeaderWidth,
                      height: rowHeight(row),
                    }}
                  >
                    <RowHeader
                      row={row}
                      height={rowHeight(row)}
                      onResize={(h) => onRowResize(row, h / scale)}
                      onSelect={(extend) => selectEntireRow(row, extend)}
                    />
                  </div>
                ))}
              </div>
              <div
                className="relative shrink-0"
                style={{ width: columnsWidth, height: frozenHeight }}
              >
                {frozenRowsList.flatMap((row) =>
                  renderRowCells({
                    row,
                    top: getRowTopInFrozen(row),
                    sheet,
                    sheetId,
                    visibleCols,
                    getDisplay,
                    colWidth,
                    rowHeight,
                    showGridlines,
                    editing,
                    draft,
                    setDraft,
                    inputRef,
                    selectRef,
                    inSelection,
                    selectionBounds: normalizeSelection(activeSelection),
                    fillDrag,
                    onFillHover,
                    commitDraft,
                    beginEdit,
                    onCommitCell,
                    onSelectionChange,
                    dragAnchor,
                    scale,
                    remotePresence,
                    onOpenComment: (r, c) => setCommentPanel({ row: r, col: c }),
                  }),
                )}
              </div>
            </div>
          ) : null}

          <div
            className="flex"
            style={{
              width: sheetWidth,
              height: rowVirtualizer.getTotalSize(),
            }}
          >
            <div
              className="sticky left-0 z-20 relative shrink-0 border-r border-[#b4b4b4] bg-[#f3f3f3]"
              style={{
                width: rowHeaderWidth,
                height: rowVirtualizer.getTotalSize(),
              }}
            >
              {rowVirtualizer.getVirtualItems().map((virtualRow) => {
                const row = scrollRows[virtualRow.index]!;
                return (
                  <div
                    key={row}
                    className="absolute left-0"
                    style={{
                      top: virtualRow.start,
                      width: rowHeaderWidth,
                      height: virtualRow.size,
                    }}
                  >
                    <RowHeader
                      row={row}
                      height={virtualRow.size}
                      onResize={(h) => onRowResize(row, h / scale)}
                      onSelect={(extend) => selectEntireRow(row, extend)}
                    />
                  </div>
                );
              })}
            </div>
            <div
              className="relative shrink-0"
              style={{
                width: columnsWidth,
                height: rowVirtualizer.getTotalSize(),
              }}
            >
          {rowVirtualizer.getVirtualItems().flatMap((virtualRow) => {
            const row = scrollRows[virtualRow.index]!;
            return renderRowCells({
              row,
              top: virtualRow.start,
              sheet,
              sheetId,
              visibleCols,
              getDisplay,
              colWidth,
              rowHeight,
              showGridlines,
              editing,
              draft,
              setDraft,
              inputRef,
              selectRef,
              inSelection,
              selectionBounds: normalizeSelection(activeSelection),
              fillDrag,
              onFillHover,
              commitDraft,
              beginEdit,
              onCommitCell,
              onSelectionChange,
              dragAnchor,
              scale,
              remotePresence,
              onOpenComment: (r, c) => setCommentPanel({ row: r, col: c }),
            });
          })}
              <FillHandle
                selection={selection}
                getRowTop={getRowTopInScroll}
                rowHeight={rowHeight}
                visibleCols={visibleCols}
                colWidth={colWidth}
                minRow={
                  sheet.frozenRows > 0 ? sheet.frozenRows : undefined
                }
                onFillStart={() => setFillDrag(normalizeSelection(selection))}
              />
              {sheet.frozenRows > 0 ? (
                <FillHandle
                  selection={selection}
                  getRowTop={getRowTopInFrozen}
                  rowHeight={rowHeight}
                  visibleCols={visibleCols}
                  colWidth={colWidth}
                  maxRow={sheet.frozenRows - 1}
                  onFillStart={() => setFillDrag(normalizeSelection(selection))}
                />
              ) : null}
            </div>
          </div>
          <SheetChartOverlays
            sheet={sheet}
            sheetId={sheetId}
            getDisplay={getDisplay}
            charts={sheet.charts}
            colWidth={colWidth}
            rowHeight={rowHeight}
            colOffset={chartColOffset}
            rowTop={chartRowTop}
            leftGutter={rowHeaderWidth}
          />
          {commentPanel ? (
            <div
              className="absolute z-[50]"
              style={{
                left:
                  rowHeaderWidth +
                  colOffsetsVisible(visibleCols, commentPanel.col, colWidth),
                top: chartRowTop(commentPanel.row),
              }}
            >
              <CellCommentPanel
                row={commentPanel.row}
                col={commentPanel.col}
                thread={normalizeCellComment(
                  sheet.cells[cellKey(commentPanel.row, commentPanel.col)]?.comment,
                )}
                readOnly={readOnly}
                currentUserId={commentUserId}
                onStart={(text) => {
                  onCommentStart(commentPanel.row, commentPanel.col, text);
                  setCommentPanel({ ...commentPanel });
                }}
                onReply={(text) =>
                  onCommentReply(commentPanel.row, commentPanel.col, text)
                }
                onResolve={() =>
                  onCommentResolve(commentPanel.row, commentPanel.col)
                }
                onReopen={() =>
                  onCommentReopen(commentPanel.row, commentPanel.col)
                }
                onDelete={() => {
                  onCommentDelete(commentPanel.row, commentPanel.col);
                  setCommentPanel(null);
                }}
                onClose={() => setCommentPanel(null)}
              />
            </div>
          ) : null}
        </div>
      </div>
    </div>
  );
},
);
WorkbookGrid.displayName = "WorkbookGrid";

function FillHandle({
  selection,
  getRowTop,
  rowHeight,
  visibleCols,
  colWidth,
  minRow,
  maxRow,
  onFillStart,
}: {
  selection: GridSelection;
  getRowTop: (row: number) => number;
  rowHeight: (row: number) => number;
  visibleCols: number[];
  colWidth: (col: number) => number;
  minRow?: number;
  maxRow?: number;
  onFillStart: () => void;
}) {
  const endRow = Math.max(selection.row, selection.endRow);
  const endCol = Math.max(selection.col, selection.endCol);
  if (minRow !== undefined && endRow < minRow) return null;
  if (maxRow !== undefined && endRow > maxRow) return null;
  const left =
    colOffsetsVisible(visibleCols, endCol, colWidth) + colWidth(endCol) - 6;
  const top = getRowTop(endRow) + rowHeight(endRow) - 6;
  return (
    <div
      className="absolute z-30 h-2 w-2 cursor-crosshair bg-blue-600"
      style={{ left, top }}
      onMouseDown={(event) => {
        event.preventDefault();
        event.stopPropagation();
        onFillStart();
      }}
      title="Fill handle"
    />
  );
}

function colOffsetsVisible(
  visibleCols: number[],
  col: number,
  colWidth: (c: number) => number,
): number {
  let offset = 0;
  for (const c of visibleCols) {
    if (c === col) break;
    offset += colWidth(c);
  }
  return offset;
}

interface RenderRowCellsArgs {
  row: number;
  top: number;
  sheet: Worksheet;
  sheetId: string;
  visibleCols: number[];
  getDisplay: (sheetId: string, row: number, col: number) => string;
  colWidth: (col: number) => number;
  rowHeight: (row: number) => number;
  showGridlines: boolean;
  editing: { row: number; col: number } | null;
  draft: string;
  setDraft: (value: string) => void;
  inputRef: RefObject<HTMLInputElement | null>;
  selectRef: RefObject<HTMLSelectElement | null>;
  inSelection: (row: number, col: number) => boolean;
  selectionBounds: GridSelection;
  fillDrag: GridSelection | null;
  onFillHover: (row: number, col: number) => void;
  commitDraft: (commit: boolean) => void;
  beginEdit: (row: number, col: number) => void;
  onCommitCell: (row: number, col: number, value: string) => void;
  onSelectionChange: (selection: GridSelection) => void;
  dragAnchor: MutableRefObject<{ row: number; col: number } | null>;
  scale: number;
  remotePresence: RemotePresence[];
  onOpenComment: (row: number, col: number) => void;
}

function mergePixelHeight(
  merge: NonNullable<ReturnType<typeof mergeAt>>,
  rowHeight: (row: number) => number,
): number {
  let height = 0;
  for (let row = merge.startRow; row <= merge.endRow; row++) {
    height += rowHeight(row);
  }
  return height;
}

function renderRowCells(args: RenderRowCellsArgs) {
  const {
    row,
    top,
    sheet,
    sheetId,
    visibleCols,
    getDisplay,
    colWidth,
    rowHeight,
    showGridlines,
    editing,
    draft,
    setDraft,
    inputRef,
    selectRef,
    inSelection,
    selectionBounds,
    fillDrag,
    onFillHover,
    commitDraft,
    beginEdit,
    onSelectionChange,
    dragAnchor,
    scale,
    remotePresence,
    onOpenComment,
  } = args;

  return visibleCols.map((col) => {
    const merge = mergeAt(sheet, row, col);
    if (merge && !(merge.startRow === row && merge.startCol === col)) {
      return null;
    }
    const data = sheet.cells[cellKey(row, col)];
    const input = getCellInput(sheet, row, col);
    const rawDisplay =
      input && !input.startsWith("=") ? input : getDisplay(sheetId, row, col);
    const display = formatDisplayValue(rawDisplay, data?.format, input);
    const selected = inSelection(row, col);
    const isEditing = editing?.row === row && editing?.col === col;
    const listValues =
      data?.validation?.kind === "list" ? data.validation.values ?? [] : null;
    const w = merge
      ? visibleCols
          .filter((c) => c >= merge.startCol && c <= merge.endCol)
          .reduce((sum, c) => sum + colWidth(c), 0)
      : colWidth(col);
    const h = merge
      ? mergePixelHeight(merge, rowHeight)
      : rowHeight(row);
    const displayFormat = effectiveCellFormat(sheet, row, col, data?.format, {
      input,
      display: rawDisplay,
    });
    const align = displayFormat?.horizontalAlign;
    const valign = displayFormat?.verticalAlign ?? "middle";
    const editFont: CSSProperties = {
      ...cellStyleFromFormat(displayFormat),
      lineHeight: 1.2,
      margin: 0,
      fontSize: displayFormat?.fontSize
        ? `${displayFormat.fontSize * scale}pt`
        : `${11 * scale}pt`,
    };

    const remoteHighlights = remoteHighlightsAtCell(
      row,
      col,
      sheetId,
      remotePresence,
    );

    return (
      <GridCell
        key={cellKey(row, col)}
        row={row}
        col={col}
        top={top}
        left={colOffsetsVisible(visibleCols, col, colWidth)}
        width={w}
        height={h}
        showGridlines={showGridlines}
        selected={selected}
        selectionBounds={selectionBounds}
        merged={Boolean(merge)}
        remoteHighlights={remoteHighlights}
        style={{
          ...cellStyleFromFormat(displayFormat),
          fontSize: displayFormat?.fontSize
            ? `${displayFormat.fontSize * scale}pt`
            : `${11 * scale}pt`,
        }}
        onMouseDown={() => {
          if (editing) commitDraft(true);
          dragAnchor.current = { row, col };
          onSelectionChange({ row, col, endRow: row, endCol: col });
        }}
        onMouseEnter={() => {
          if (fillDrag) {
            onFillHover(row, col);
            return;
          }
          if (dragAnchor.current) {
            onSelectionChange({
              row: dragAnchor.current.row,
              col: dragAnchor.current.col,
              endRow: row,
              endCol: col,
            });
          }
        }}
        onMouseUp={() => {
          dragAnchor.current = null;
        }}
        onDoubleClick={() => beginEdit(row, col)}
      >
        {isEditing && listValues && listValues.length > 0 ? (
          <select
            ref={selectRef}
            className="h-full w-full bg-white px-1 text-zinc-900 outline-none"
            style={editFont}
            value={draft}
            onChange={(e) => {
              setDraft(e.target.value);
              commitDraft(true);
            }}
            onBlur={() => commitDraft(true)}
          >
            {listValues.map((value) => (
              <option key={value} value={value}>
                {value}
              </option>
            ))}
          </select>
        ) : isEditing ? (
          <input
            ref={inputRef}
            className="h-full w-full bg-white px-1 text-zinc-900 outline-none"
            style={editFont}
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={() => commitDraft(true)}
            list={
              listValues && listValues.length > 0
                ? `list-${row}-${col}`
                : undefined
            }
          />
        ) : (
          <>
            {data?.comment ? (
              <button
                type="button"
                className={`absolute right-0 top-0 z-[5] h-0 w-0 border-l-[7px] border-t-[7px] border-l-transparent ${
                  normalizeCellComment(data.comment)?.status === "resolved"
                    ? "border-t-zinc-400 opacity-60"
                    : "border-t-red-500"
                }`}
                title={`${threadAuthorLabel(normalizeCellComment(data.comment)! )}: thread`}
                onMouseDown={(event) => {
                  event.stopPropagation();
                  onOpenComment(row, col);
                }}
              />
            ) : null}
            <CellDisplay
              align={align}
              valign={valign}
              wrap={Boolean(data?.format?.wrap)}
            >
              {display}
            </CellDisplay>
          </>
        )}
      </GridCell>
    );
  });
}

function ColHeader({
  col,
  width,
  height,
  onResize,
  showFilter,
  onOpenFilter,
  filterActive,
  onSelect,
}: {
  col: number;
  width: number;
  height: number;
  onResize: (width: number) => void;
  showFilter?: boolean;
  onOpenFilter?: (anchor: DOMRect) => void;
  filterActive?: boolean;
  filterValues?: string[];
  onSelect?: (extend: boolean) => void;
}) {
  return (
    <div
      className="relative flex shrink-0 cursor-pointer items-center justify-center gap-0.5 border-r border-[#b4b4b4] text-xs font-semibold text-zinc-700 select-none hover:bg-[#e8f5ec]"
      style={{ width, height }}
      onMouseDown={(event) => {
        if (event.button !== 0) return;
        event.preventDefault();
        onSelect?.(event.shiftKey);
      }}
    >
      {indexToColumnLabel(col)}
      {showFilter ? (
        <button
          type="button"
          className={`rounded px-0.5 text-sm leading-none ${
            filterActive
              ? "bg-[#217346] text-white"
              : "text-zinc-600 hover:bg-zinc-200"
          }`}
          onClick={(event) => {
            event.stopPropagation();
            onOpenFilter?.(event.currentTarget.getBoundingClientRect());
          }}
          title="Filter & sort"
        >
          ▾
        </button>
      ) : null}
      <span
        className="absolute top-0 right-0 h-full w-1 cursor-col-resize"
        onMouseDown={(event) => {
          event.preventDefault();
          const startX = event.clientX;
          const startW = width;
          const onMove = (move: MouseEvent) => {
            onResize(Math.max(40, startW + (move.clientX - startX)));
          };
          const onUp = () => {
            window.removeEventListener("mousemove", onMove);
            window.removeEventListener("mouseup", onUp);
          };
          window.addEventListener("mousemove", onMove);
          window.addEventListener("mouseup", onUp);
        }}
      />
    </div>
  );
}

function RowHeader({
  row,
  height,
  onResize,
  onSelect,
}: {
  row: number;
  height: number;
  onResize: (height: number) => void;
  onSelect?: (extend: boolean) => void;
}) {
  return (
    <div
      className="relative flex cursor-pointer items-center justify-center border-b border-zinc-200 text-xs text-zinc-600 select-none hover:bg-[#e8f5ec]"
      style={{ height, width: ROW_HEADER_WIDTH }}
      onMouseDown={(event) => {
        if (event.button !== 0) return;
        event.preventDefault();
        onSelect?.(event.shiftKey);
      }}
    >
      {row + 1}
      <span
        className="absolute bottom-0 left-0 h-1 w-full cursor-row-resize"
        onMouseDown={(event) => {
          event.preventDefault();
          const startY = event.clientY;
          const startH = height;
          const onMove = (move: MouseEvent) => {
            onResize(Math.max(18, startH + (move.clientY - startY)));
          };
          const onUp = () => {
            window.removeEventListener("mousemove", onMove);
            window.removeEventListener("mouseup", onUp);
          };
          window.addEventListener("mousemove", onMove);
          window.addEventListener("mouseup", onUp);
        }}
      />
    </div>
  );
}

function CellDisplay({
  children,
  align,
  valign,
  wrap,
}: {
  children: ReactNode;
  align?: string;
  valign?: string;
  wrap?: boolean;
}) {
  const justify =
    align === "center" ? "center" : align === "right" ? "flex-end" : "flex-start";
  const items =
    valign === "top" ? "flex-start" : valign === "bottom" ? "flex-end" : "center";
  return (
    <div
      className={`flex h-full max-h-full w-full overflow-hidden px-1 text-inherit ${
        wrap ? "whitespace-pre-wrap break-words" : "truncate"
      }`}
      style={{ justifyContent: justify, alignItems: items }}
    >
      <span className={wrap ? "" : "truncate"}>{children}</span>
    </div>
  );
}

function selectionEdgeStyle(
  row: number,
  col: number,
  bounds: GridSelection,
  selected: boolean,
  showGridlines: boolean,
): CSSProperties {
  if (!selected) return {};
  const edgeColor = "#217346";
  const gridColor = "#b0b0b0";
  const edge = 2;
  const inner = showGridlines ? 1 : 0;
  return {
    borderTop: `${row === bounds.row ? edge : inner}px solid ${
      row === bounds.row ? edgeColor : gridColor
    }`,
    borderBottom: `${row === bounds.endRow ? edge : inner}px solid ${
      row === bounds.endRow ? edgeColor : gridColor
    }`,
    borderLeft: `${col === bounds.col ? edge : inner}px solid ${
      col === bounds.col ? edgeColor : gridColor
    }`,
    borderRight: `${col === bounds.endCol ? edge : inner}px solid ${
      col === bounds.endCol ? edgeColor : gridColor
    }`,
  };
}

function GridCell({
  children,
  row,
  col,
  top,
  left,
  width,
  height,
  selected,
  selectionBounds,
  showGridlines,
  merged,
  remoteHighlights = [],
  style,
  onMouseDown,
  onMouseEnter,
  onMouseUp,
  onDoubleClick,
}: {
  children: ReactNode;
  row: number;
  col: number;
  top: number;
  left: number;
  width: number;
  height: number;
  selected: boolean;
  selectionBounds: GridSelection;
  showGridlines: boolean;
  merged?: boolean;
  remoteHighlights?: { color: string; initial: string }[];
  style: CSSProperties;
  onMouseDown: () => void;
  onMouseEnter: () => void;
  onMouseUp: () => void;
  onDoubleClick: () => void;
}) {
  const remoteRing =
    remoteHighlights.length > 0
      ? remoteHighlights
          .map((h, i) => `inset 0 0 0 ${2 + i * 2}px ${h.color}`)
          .join(", ")
      : undefined;
  const gridLine =
    !selected && showGridlines ? "inset 0 0 0 1px #b0b0b0" : undefined;
  const boxShadow = [gridLine, remoteRing].filter(Boolean).join(", ") || undefined;

  return (
    <div
      className={`absolute box-border overflow-hidden text-[11pt] leading-tight text-inherit ${
        selected ? "z-[2] bg-[#cce8cc]" : "z-[1] bg-white"
      } ${merged ? "flex" : ""}`}
      style={{
        top,
        left,
        width,
        height,
        boxShadow,
        ...selectionEdgeStyle(row, col, selectionBounds, selected, showGridlines),
        ...style,
        color: style.color ?? "#18181b",
      }}
      onMouseDown={onMouseDown}
      onMouseEnter={onMouseEnter}
      onMouseUp={onMouseUp}
      onDoubleClick={onDoubleClick}
    >
      {remoteHighlights.map((h, i) => (
        <span
          key={`${h.color}-${i}`}
          className="pointer-events-none absolute z-[4] flex h-3.5 min-w-3.5 items-center justify-center rounded-br px-0.5 text-[9px] font-bold leading-none text-white shadow-sm"
          style={{
            left: 2 + i * 12,
            top: 2,
            backgroundColor: h.color,
          }}
          title="Another editor"
        >
          {h.initial}
        </span>
      ))}
      {children}
    </div>
  );
}