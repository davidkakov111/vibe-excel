"use client";

import { current, produce } from "immer";
import { useCallback, useEffect, useRef, useState } from "react";
import { FormulaEngine } from "@/lib/formula/engine";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  duplicateWorksheet,
  createWorksheet,
} from "@/lib/workbook/defaults";
import {
  deleteColumns,
  deleteRows,
  insertColumns,
  insertRows,
} from "@/lib/workbook/structure";
import { sortWorksheetRange } from "@/lib/workbook/sort";
import {
  clearAutoFilter,
  computeHiddenRows,
  disableAutoFilter,
  enableAutoFilter,
  setColumnFilter,
} from "@/lib/workbook/filter";
import {
  CHART_ANCHOR_COLS,
  CHART_ANCHOR_ROWS,
  defaultChartAnchor,
} from "@/lib/workbook/chart-data";
import { fillSelection, autofitColumnWidth } from "@/lib/workbook/fill";
import { replaceInSheet } from "@/lib/workbook/find-replace";
import type { GridSelection } from "@/lib/workbook/selection";
import { validateCellValue, attachValidation } from "@/lib/workbook/validation";
import type {
  CellData,
  WorkbookDocument,
  MergeRange,
  CellFormat,
  DataValidation,
  FilterColumnState,
  Worksheet,
  ConditionalFormattingSet,
} from "@/lib/workbook/types";
import { getCellInput, setCellInput } from "@/lib/workbook/types";
import { cellKey } from "@/lib/spreadsheet-types";
import {
  mergeCellPaste,
  type ClipboardCell,
} from "@/lib/workbook/clipboard";
import { nanoid } from "nanoid";
import {
  applyWorkbookOp,
  applyWorkbookOps,
  cellOpFromInput,
  commentOp,
  formatOp,
  readCellInput,
  patchSheetOp,
  restoreRangeOp,
  sortRangeOp,
  type WorkbookOp,
} from "@/lib/workbook/ops";
import type { ChartSeries } from "@/lib/workbook/types";
import {
  pushLocalHistory,
  type LocalHistoryEntry,
} from "@/lib/workbook/local-history";
import { snapshotRangeCells } from "@/lib/workbook/snapshot-range";
import {
  addCommentReply,
  createCommentThread,
  normalizeCellComment,
  setCommentResolved,
  type CellCommentThread,
} from "@/lib/workbook/comment";

type SaveState = "idle" | "saving" | "saved" | "error";

const MAX_UNDO = 50;

export type WorkbookEditorOptions = {
  readOnly?: boolean;
  useGranularOps?: boolean;
  sendOps?: (ops: WorkbookOp[]) => Promise<boolean>;
  /** Server revision for optimistic concurrency on PATCH (legacy saves). */
  getRevision?: () => number;
  userId?: string;
  userName?: string;
};

export function useWorkbookEditor(
  spreadsheetId: string,
  initialWorkbook: WorkbookDocument,
  options?: WorkbookEditorOptions,
) {
  const [workbook, setWorkbook] = useState(initialWorkbook);
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const [saveMessage, setSaveMessage] = useState<string | null>(null);
  const engineRef = useRef<FormulaEngine | null>(null);
  const undoStack = useRef<WorkbookDocument[]>([]);
  const redoStack = useRef<WorkbookDocument[]>([]);
  const localUndoStack = useRef<LocalHistoryEntry[]>([]);
  const localRedoStack = useRef<LocalHistoryEntry[]>([]);
  const saveTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const saveGeneration = useRef(0);
  const inFlightGeneration = useRef<number | null>(null);
  const saveDirty = useRef(false);
  const saveInFlight = useRef(false);
  const latestWorkbook = useRef(workbook);

  useEffect(() => {
    latestWorkbook.current = workbook;
  }, [workbook]);

  const syncEngine = useCallback((doc: WorkbookDocument) => {
    if (engineRef.current) {
      engineRef.current.rebuild(doc);
    } else {
      engineRef.current = FormulaEngine.fromWorkbook(doc);
    }
  }, []);

  useEffect(() => {
    syncEngine(workbook);
    return () => engineRef.current?.destroy();
  }, [syncEngine, workbook]);

  const pushUndo = useCallback((snapshot: WorkbookDocument) => {
    if (options?.useGranularOps) return;
    undoStack.current.push(structuredClone(snapshot));
    if (undoStack.current.length > MAX_UNDO) {
      undoStack.current.shift();
    }
    redoStack.current = [];
  }, [options?.useGranularOps]);

  const recordLocalHistory = useCallback((entry: LocalHistoryEntry) => {
    if (!options?.useGranularOps) return;
    pushLocalHistory(localUndoStack.current, entry, MAX_UNDO);
    localRedoStack.current = [];
  }, [options?.useGranularOps]);

  const activeSheet = workbook.sheets.find((s) => s.id === workbook.activeSheetId)
    ?? workbook.sheets[0]!;

  const clearPendingSave = useCallback(() => {
    saveDirty.current = false;
    if (saveTimer.current) {
      clearTimeout(saveTimer.current);
      saveTimer.current = null;
    }
  }, []);

  const runSave = useCallback(async () => {
    if (options?.useGranularOps) {
      clearPendingSave();
      return;
    }
    if (saveInFlight.current) return;
    if (!saveDirty.current) return;

    saveInFlight.current = true;
    saveDirty.current = false;

    const generation = ++saveGeneration.current;
    inFlightGeneration.current = generation;
    const doc = structuredClone(latestWorkbook.current);
    const baseRevision = options?.getRevision?.();

    setSaveState("saving");
    setSaveMessage(null);
    let failed = false;
    try {
      const response = await fetch(`/api/spreadsheets/${spreadsheetId}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          workbook: doc,
          ...(baseRevision !== undefined ? { baseRevision } : {}),
        }),
      });
      if (!response.ok) {
        const payload = (await response.json().catch(() => null)) as {
          error?: string;
        } | null;
        throw new Error(payload?.error ?? "Failed to save spreadsheet");
      }
      if (inFlightGeneration.current === generation) {
        setSaveState("saved");
      }
    } catch (error) {
      failed = true;
      saveDirty.current = true;
      if (inFlightGeneration.current === generation) {
        setSaveState("error");
        setSaveMessage(
          error instanceof Error ? error.message : "Failed to save spreadsheet",
        );
      }
    } finally {
      saveInFlight.current = false;
      if (saveDirty.current && !failed) {
        void runSave();
      }
    }
  }, [clearPendingSave, options, spreadsheetId]);

  const scheduleSave = useCallback(() => {
    if (options?.useGranularOps) return;
    saveDirty.current = true;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      saveTimer.current = null;
      void runSave();
    }, 500);
  }, [options?.useGranularOps, runSave]);

  const applyWorkbook = useCallback(
    (next: WorkbookDocument, save = true) => {
      latestWorkbook.current = next;
      syncEngine(next);
      setWorkbook(next);
      if (options?.useGranularOps) {
        clearPendingSave();
        setSaveState("saved");
        return;
      }
      if (save) scheduleSave();
    },
    [clearPendingSave, options?.useGranularOps, scheduleSave, syncEngine],
  );

  const commitWorkbook = useCallback(
    (updater: (draft: WorkbookDocument) => void, recordHistory = true) => {
      if (recordHistory) {
        pushUndo(latestWorkbook.current);
      }
      const next = produce(latestWorkbook.current, updater);
      applyWorkbook(next);
    },
    [applyWorkbook, pushUndo],
  );

  const applyLocalOps = useCallback(
    (ops: WorkbookOp[], save: boolean) => {
      const next = applyWorkbookOps(latestWorkbook.current, ops);
      applyWorkbook(next, save);
    },
    [applyWorkbook],
  );

  const pushCollaborativeOps = useCallback(
    (ops: WorkbookOp[], history?: LocalHistoryEntry) => {
      if (history) recordLocalHistory(history);
      clearPendingSave();
      const next = applyWorkbookOps(latestWorkbook.current, ops);
      applyWorkbook(next, false);
      if (options?.sendOps) void options.sendOps(ops);
    },
    [applyWorkbook, clearPendingSave, options, recordLocalHistory],
  );

  const setActiveSheet = useCallback((sheetId: string) => {
    commitWorkbook((draft) => {
      draft.activeSheetId = sheetId;
    }, false);
  }, [commitWorkbook]);

  const setCell = useCallback(
    (row: number, col: number, input: string) => {
      if (options?.readOnly) return;
      const sheet = workbook.sheets.find((s) => s.id === workbook.activeSheetId);
      const validation = sheet?.cells[cellKey(row, col)]?.validation;
      const check = validateCellValue(input, validation);
      if (!check.valid) {
        window.alert(check.message ?? "Invalid value for this cell.");
        return;
      }
      if (options?.useGranularOps && options.sendOps) {
        const sheetId = latestWorkbook.current.activeSheetId;
        const previous = readCellInput(latestWorkbook.current, sheetId, row, col);
        const op = cellOpFromInput(sheetId, row, col, input);
        recordLocalHistory({
          undo: [cellOpFromInput(sheetId, row, col, previous)],
          redo: [op],
        });
        const next = applyWorkbookOp(latestWorkbook.current, op);
        applyWorkbook(next, false);
        void options.sendOps([op]);
        return;
      }
      commitWorkbook((draft) => {
        const active = draft.sheets.find((s) => s.id === draft.activeSheetId);
        if (!active) return;
        const index = draft.sheets.indexOf(active);
        draft.sheets[index] = setCellInput(active, row, col, input);
      });
    },
    [
      applyWorkbook,
      commitWorkbook,
      options,
      recordLocalHistory,
      workbook.activeSheetId,
      workbook.sheets,
    ],
  );

  const applyRemoteOps = useCallback(
    (ops: WorkbookOp[]) => {
      const next = applyWorkbookOps(latestWorkbook.current, ops);
      applyWorkbook(next, false);
    },
    [applyWorkbook],
  );

  const applyConflictWorkbook = useCallback(
    (doc: WorkbookDocument) => {
      applyWorkbook(doc, false);
    },
    [applyWorkbook],
  );

  const applyRemoteWorkbookRestore = useCallback(
    (doc: WorkbookDocument) => {
      localUndoStack.current = [];
      localRedoStack.current = [];
      undoStack.current = [];
      redoStack.current = [];
      applyWorkbook(doc, false);
    },
    [applyWorkbook],
  );

  const setConditionalFormats = useCallback(
    (sets: ConditionalFormattingSet[]) => {
      if (options?.readOnly) return;
      const sheetId = latestWorkbook.current.activeSheetId;
      const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
      if (!sheet) return;
      const op = patchSheetOp(sheetId, {
        conditionalFormats: structuredClone(sets),
      });
      if (options?.useGranularOps && options.sendOps) {
        const prev = structuredClone(sheet.conditionalFormats ?? []);
        recordLocalHistory({
          undo: [patchSheetOp(sheetId, { conditionalFormats: prev })],
          redo: [op],
        });
        const next = applyWorkbookOp(latestWorkbook.current, op);
        applyWorkbook(next, false);
        void options.sendOps([op]);
        return;
      }
      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        draft.sheets[index] = {
          ...draft.sheets[index]!,
          conditionalFormats: structuredClone(sets),
        };
      });
    },
    [applyWorkbook, commitWorkbook, options, recordLocalHistory],
  );

  const setCellComment = useCallback(
    (row: number, col: number, comment: CellCommentThread | null) => {
      if (options?.readOnly) return;
      const op = commentOp(workbook.activeSheetId, row, col, comment);
      if (options?.useGranularOps && options.sendOps) {
        const sheet = latestWorkbook.current.sheets.find(
          (s) => s.id === workbook.activeSheetId,
        );
        const previous = sheet?.cells[cellKey(row, col)]?.comment ?? null;
        recordLocalHistory({
          undo: [commentOp(workbook.activeSheetId, row, col, previous)],
          redo: [op],
        });
        const next = applyWorkbookOp(latestWorkbook.current, op);
        applyWorkbook(next, false);
        void options.sendOps([op]);
        return;
      }
      commitWorkbook((draft) => {
        const next = applyWorkbookOp(current(draft), op);
        draft.sheets = next.sheets;
        draft.activeSheetId = next.activeSheetId;
        draft.settings = next.settings;
        draft.version = next.version;
      });
    },
    [applyWorkbook, commitWorkbook, options, recordLocalHistory, workbook.activeSheetId],
  );

  const mutateCellComment = useCallback(
    (
      row: number,
      col: number,
      fn: (current: CellCommentThread | null) => CellCommentThread | null,
    ) => {
      if (options?.readOnly) return;
      const sheet = latestWorkbook.current.sheets.find(
        (s) => s.id === workbook.activeSheetId,
      );
      const key = cellKey(row, col);
      const current = normalizeCellComment(sheet?.cells[key]?.comment ?? null);
      const next = fn(current);
      setCellComment(row, col, next);
    },
    [options?.readOnly, setCellComment, workbook.activeSheetId],
  );

  const startCellComment = useCallback(
    (row: number, col: number, text: string) => {
      const uid = options?.userId ?? "local";
      const name = options?.userName ?? "You";
      if (!text.trim()) return;
      const thread = createCommentThread(uid, name, text);
      setCellComment(row, col, thread);
    },
    [options?.userId, options?.userName, setCellComment],
  );

  const replyCellComment = useCallback(
    (row: number, col: number, text: string) => {
      const uid = options?.userId ?? "local";
      const name = options?.userName ?? "You";
      mutateCellComment(row, col, (current) => {
        if (!current) return createCommentThread(uid, name, text);
        return addCommentReply(current, uid, name, text);
      });
    },
    [mutateCellComment, options?.userId, options?.userName],
  );

  const resolveCellComment = useCallback(
    (row: number, col: number) => {
      mutateCellComment(row, col, (current) =>
        current ? setCommentResolved(current, true) : null,
      );
    },
    [mutateCellComment],
  );

  const reopenCellComment = useCallback(
    (row: number, col: number) => {
      mutateCellComment(row, col, (current) =>
        current ? setCommentResolved(current, false) : null,
      );
    },
    [mutateCellComment],
  );

  const deleteCellComment = useCallback(
    (row: number, col: number) => {
      setCellComment(row, col, null);
    },
    [setCellComment],
  );

  const applyFormatPatch = useCallback(
    (cells: { row: number; col: number }[], format: Partial<CellFormat>) => {
      if (options?.readOnly || cells.length === 0) return;
      const sheetId = latestWorkbook.current.activeSheetId;
      const ops = cells.map((c) => formatOp(sheetId, c.row, c.col, format));

      if (options?.useGranularOps && options.sendOps) {
        const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
        const undoOps = cells.map(({ row, col }) => {
          const prev = sheet?.cells[cellKey(row, col)]?.format;
          return formatOp(sheetId, row, col, prev ?? null);
        });
        recordLocalHistory({ undo: undoOps, redo: ops });
        const next = applyWorkbookOps(latestWorkbook.current, ops);
        applyWorkbook(next, false);
        void options.sendOps(ops);
        return;
      }

      commitWorkbook((draft) => {
        const sheet = draft.sheets.find((s) => s.id === draft.activeSheetId);
        if (!sheet) return;
        const index = draft.sheets.indexOf(sheet);
        const nextCells = { ...sheet.cells };
        for (const { row, col } of cells) {
          const key = cellKey(row, col);
          const current = nextCells[key] ?? {};
          nextCells[key] = {
            ...current,
            format: { ...current.format, ...format },
          };
        }
        draft.sheets[index] = { ...sheet, cells: nextCells };
      });
    },
    [applyWorkbook, commitWorkbook, options, recordLocalHistory],
  );

  const setCellFormat = useCallback(
    (row: number, col: number, format: Partial<CellFormat>) => {
      applyFormatPatch([{ row, col }], format);
    },
    [applyFormatPatch],
  );

  const setFormatOnSelection = useCallback(
    (selection: GridSelection, format: Partial<CellFormat>) => {
      const startRow = Math.min(selection.row, selection.endRow);
      const endRow = Math.max(selection.row, selection.endRow);
      const startCol = Math.min(selection.col, selection.endCol);
      const endCol = Math.max(selection.col, selection.endCol);
      const cells: { row: number; col: number }[] = [];
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          cells.push({ row, col });
        }
      }
      applyFormatPatch(cells, format);
    },
    [applyFormatPatch],
  );

  const addSheet = useCallback(() => {
    commitWorkbook((draft) => {
      const sheet = createWorksheet(`Sheet${draft.sheets.length + 1}`);
      draft.sheets.push(sheet);
      draft.activeSheetId = sheet.id;
    });
  }, [commitWorkbook]);

  const duplicateActiveSheet = useCallback(() => {
    commitWorkbook((draft) => {
      const active = draft.sheets.find((s) => s.id === draft.activeSheetId);
      if (!active) return;
      const copy = duplicateWorksheet(active, `${active.name} Copy`);
      draft.sheets.push(copy);
      draft.activeSheetId = copy.id;
    });
  }, [commitWorkbook]);

  const deleteSheet = useCallback((sheetId: string) => {
    commitWorkbook((draft) => {
      if (draft.sheets.length <= 1) return;
      draft.sheets = draft.sheets.filter((s) => s.id !== sheetId);
      if (draft.activeSheetId === sheetId) {
        draft.activeSheetId = draft.sheets[0]!.id;
      }
    });
  }, [commitWorkbook]);

  const renameSheet = useCallback((sheetId: string, name: string) => {
    commitWorkbook((draft) => {
      const sheet = draft.sheets.find((s) => s.id === sheetId);
      if (!sheet) return;
      sheet.name = name.slice(0, 31) || sheet.name;
    });
  }, [commitWorkbook]);

  const reorderSheet = useCallback((sheetId: string, direction: -1 | 1) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === sheetId);
      const target = index + direction;
      if (index < 0 || target < 0 || target >= draft.sheets.length) return;
      const [item] = draft.sheets.splice(index, 1);
      draft.sheets.splice(target, 0, item!);
    });
  }, [commitWorkbook]);

  const insertSheetRows = useCallback((atRow: number, count = 1) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      draft.sheets[index] = insertRows(draft.sheets[index]!, atRow, count);
    });
  }, [commitWorkbook]);

  const deleteSheetRows = useCallback((atRow: number, count = 1) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      draft.sheets[index] = deleteRows(draft.sheets[index]!, atRow, count);
    });
  }, [commitWorkbook]);

  const insertSheetColumns = useCallback((atCol: number, count = 1) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      draft.sheets[index] = insertColumns(draft.sheets[index]!, atCol, count);
    });
  }, [commitWorkbook]);

  const deleteSheetColumns = useCallback((atCol: number, count = 1) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      draft.sheets[index] = deleteColumns(draft.sheets[index]!, atCol, count);
    });
  }, [commitWorkbook]);

  const sortByColumn = useCallback(
    (col: number, ascending: boolean, startRow: number, endRow: number) => {
      const doc = latestWorkbook.current;
      const sheet = doc.sheets.find((s) => s.id === doc.activeSheetId);
      if (!sheet) return;

      const startCol = 0;
      const endCol = sheet.colCount - 1;
      const sortOp = sortRangeOp(
        sheet.id,
        startRow,
        endRow,
        startCol,
        endCol,
        col,
        ascending,
      );

      if (options?.useGranularOps && options.sendOps) {
        const snapshot = snapshotRangeCells(
          sheet,
          startRow,
          endRow,
          startCol,
          endCol,
        );
        pushCollaborativeOps([sortOp], {
          undo: [restoreRangeOp(sheet.id, snapshot)],
          redo: [sortOp],
        });
        return;
      }

      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const active = current(draft.sheets[index]!);
        const sorted = sortWorksheetRange(
          active,
          startRow,
          endRow,
          startCol,
          endCol,
          col,
          ascending,
          (row, c) => {
            const cellInput = active.cells[cellKey(row, c)]?.input ?? "";
            if (cellInput && !cellInput.startsWith("=")) {
              return cellInput;
            }
            return engineRef.current?.getDisplayValue(active.id, row, c) ?? cellInput;
          },
        );
        draft.sheets[index]!.cells = sorted.cells;
      });
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const mergeSelection = useCallback((range: MergeRange) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      const sheet = draft.sheets[index]!;
      draft.sheets[index] = {
        ...sheet,
        merges: [...sheet.merges, range],
      };
    });
  }, [commitWorkbook]);

  const unmergeAt = useCallback((row: number, col: number) => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      const sheet = draft.sheets[index]!;
      draft.sheets[index] = {
        ...sheet,
        merges: sheet.merges.filter(
          (m) =>
            !(
              row >= m.startRow &&
              row <= m.endRow &&
              col >= m.startCol &&
              col <= m.endCol
            ),
        ),
      };
    });
  }, [commitWorkbook]);

  const setFreeze = useCallback((rows: number) => {
    commitWorkbook((draft) => {
      const sheet = draft.sheets.find((s) => s.id === draft.activeSheetId);
      if (!sheet) return;
      sheet.frozenRows = rows;
      sheet.frozenCols = 0;
    });
  }, [commitWorkbook]);

  const setZoom = useCallback((zoom: number) => {
    commitWorkbook((draft) => {
      draft.settings.zoom = Math.min(200, Math.max(50, zoom));
    }, false);
  }, [commitWorkbook]);

  const toggleGridlines = useCallback(() => {
    commitWorkbook((draft) => {
      draft.settings.showGridlines = !draft.settings.showGridlines;
    }, false);
  }, [commitWorkbook]);

  const undo = useCallback(() => {
    if (options?.useGranularOps && options.sendOps) {
      const entry = localUndoStack.current.pop();
      if (!entry) return;
      localRedoStack.current.push(entry);
      applyLocalOps(entry.undo, false);
      void options.sendOps(entry.undo);
      return;
    }
    const previous = undoStack.current.pop();
    if (!previous) return;
    redoStack.current.push(structuredClone(latestWorkbook.current));
    applyWorkbook(previous);
  }, [applyLocalOps, applyWorkbook, options]);

  const redo = useCallback(() => {
    if (options?.useGranularOps && options.sendOps) {
      const entry = localRedoStack.current.pop();
      if (!entry) return;
      localUndoStack.current.push(entry);
      applyLocalOps(entry.redo, false);
      void options.sendOps(entry.redo);
      return;
    }
    const next = redoStack.current.pop();
    if (!next) return;
    undoStack.current.push(structuredClone(latestWorkbook.current));
    applyWorkbook(next);
  }, [applyLocalOps, applyWorkbook, options]);

  const getDisplay = useCallback((sheetId: string, row: number, col: number) => {
    const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
    if (!sheet) return "";
    const input = getCellInput(sheet, row, col);
    if (input && !input.startsWith("=")) {
      return input;
    }
    return engineRef.current?.getDisplayValue(sheetId, row, col) ?? input;
  }, []);

  const getRaw = useCallback(
    (sheetId: string, row: number, col: number) => {
      const sheet = workbook.sheets.find((s) => s.id === sheetId);
      if (!sheet) return "";
      return getCellInput(sheet, row, col);
    },
    [workbook.sheets],
  );

  const addChart = useCallback(
    (
      type: "bar" | "line" | "pie" | "area",
      range: MergeRange,
      title: string,
    ) => {
      const buildChartState = (sheet: Worksheet) => {
        const dataRange = {
          startRow: range.startRow,
          startCol: range.startCol,
          endRow: range.endRow,
          endCol: range.endCol,
        };
        const chartBase = {
          id: nanoid(),
          type,
          title,
          range: dataRange,
        };
        const anchor = defaultChartAnchor(chartBase);
        const anchorMerge: MergeRange = { ...anchor };
        const hasMerge = sheet.merges.some(
          (m) =>
            m.startRow === anchorMerge.startRow &&
            m.startCol === anchorMerge.startCol,
        );
        const merges = hasMerge ? sheet.merges : [...sheet.merges, anchorMerge];
        const anchorKey = cellKey(anchor.startRow, anchor.startCol);
        const anchorCell = {
          input: title,
          format: {
            bold: true,
            horizontalAlign: "center" as const,
            verticalAlign: "middle" as const,
            fillColor: "#f8fafc",
          },
        };
        let rowCount = sheet.rowCount;
        let colCount = sheet.colCount;
        if (anchor.endRow >= rowCount - 1) rowCount = anchor.endRow + CHART_ANCHOR_ROWS;
        if (anchor.endCol >= colCount - 1) colCount = anchor.endCol + CHART_ANCHOR_COLS;
        const charts: ChartSeries[] = [
          ...sheet.charts,
          { ...chartBase, anchor },
        ];
        return {
          charts,
          merges,
          rowCount,
          colCount,
          cells: { [anchorKey]: anchorCell },
        };
      };

      if (options?.useGranularOps && options.sendOps) {
        const sheetId = latestWorkbook.current.activeSheetId;
        const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
        if (!sheet) return;
        const before = {
          charts: structuredClone(sheet.charts),
          merges: structuredClone(sheet.merges),
          rowCount: sheet.rowCount,
          colCount: sheet.colCount,
        };
        const built = buildChartState(sheet);
        const redo = patchSheetOp(sheetId, built);
        const undo = patchSheetOp(sheetId, {
          charts: before.charts,
          merges: before.merges,
          rowCount: before.rowCount,
          colCount: before.colCount,
          cells: built.cells
            ? Object.fromEntries(
                Object.keys(built.cells).map((key) => [
                  key,
                  sheet.cells[key] ? structuredClone(sheet.cells[key]!) : null,
                ]),
              )
            : undefined,
        });
        pushCollaborativeOps([redo], { undo: [undo], redo: [redo] });
        return;
      }

      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const sheet = draft.sheets[index]!;
        const built = buildChartState(sheet);
        draft.sheets[index] = {
          ...sheet,
          charts: built.charts,
          merges: built.merges,
          rowCount: built.rowCount,
          colCount: built.colCount,
          cells: { ...sheet.cells, ...built.cells },
        };
      });
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  useEffect(() => {
    return () => {
      if (saveTimer.current) {
        clearTimeout(saveTimer.current);
        saveTimer.current = null;
      }
      if (saveDirty.current) {
        void runSave();
      }
    };
  }, [runSave]);

  const setColWidth = useCallback(
    (col: number, width: number) => {
      const sheetId = latestWorkbook.current.activeSheetId;
      const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
      if (!sheet) return;
      const key = String(col);
      const previous = sheet.colWidths[key];

      if (options?.useGranularOps && options.sendOps) {
        const op = patchSheetOp(sheetId, { colWidths: { [key]: width } });
        const undo = patchSheetOp(sheetId, {
          colWidths: { [key]: previous ?? DEFAULT_COL_WIDTH },
        });
        pushCollaborativeOps([op], { undo: [undo], redo: [op] });
        return;
      }

      commitWorkbook((draft) => {
        const active = draft.sheets.find((s) => s.id === draft.activeSheetId);
        if (!active) return;
        active.colWidths[String(col)] = width;
      }, false);
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const setRowHeight = useCallback(
    (row: number, height: number) => {
      const sheetId = latestWorkbook.current.activeSheetId;
      const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
      if (!sheet) return;
      const key = String(row);
      const previous = sheet.rowHeights[key];

      if (options?.useGranularOps && options.sendOps) {
        const op = patchSheetOp(sheetId, { rowHeights: { [key]: height } });
        const undo = patchSheetOp(sheetId, {
          rowHeights: { [key]: previous ?? DEFAULT_ROW_HEIGHT },
        });
        pushCollaborativeOps([op], { undo: [undo], redo: [op] });
        return;
      }

      commitWorkbook((draft) => {
        const active = draft.sheets.find((s) => s.id === draft.activeSheetId);
        if (!active) return;
        active.rowHeights[String(row)] = height;
      }, false);
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const MAX_SHEET_ROWS = 1_048_576;
  const MAX_SHEET_COLS = 16_384;
  const GROW_ROWS = 50;
  const GROW_COLS = 8;

  const ensureSheetCapacity = useCallback(
    (minRows?: number, minCols?: number) => {
      const sheetId = latestWorkbook.current.activeSheetId;
      const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
      if (!sheet) return;

      let rowCount = sheet.rowCount;
      let colCount = sheet.colCount;
      if (minRows !== undefined && minRows > rowCount - 15) {
        rowCount = Math.min(MAX_SHEET_ROWS, rowCount + GROW_ROWS);
        if (minRows >= rowCount) {
          rowCount = Math.min(MAX_SHEET_ROWS, minRows + GROW_ROWS);
        }
      }
      if (minCols !== undefined && minCols > colCount - 4) {
        colCount = Math.min(MAX_SHEET_COLS, colCount + GROW_COLS);
        if (minCols >= colCount) {
          colCount = Math.min(MAX_SHEET_COLS, minCols + GROW_COLS);
        }
      }
      if (rowCount === sheet.rowCount && colCount === sheet.colCount) return;

      const op = patchSheetOp(sheetId, { rowCount, colCount });
      if (options?.useGranularOps && options.sendOps) {
        pushCollaborativeOps([op]);
        return;
      }
      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const s = draft.sheets[index]!;
        draft.sheets[index] = { ...s, rowCount, colCount };
      }, false);
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const fillTo = useCallback(
    (source: GridSelection, target: GridSelection) => {
      const srcStartRow = Math.min(source.row, source.endRow);
      const srcEndRow = Math.max(source.row, source.endRow);
      const srcStartCol = Math.min(source.col, source.endCol);
      const srcEndCol = Math.max(source.col, source.endCol);
      const tgtEndRow = Math.max(target.row, target.endRow);
      const tgtEndCol = Math.max(target.col, target.endCol);
      const tgtStartRow = Math.min(target.row, target.endRow);
      const tgtStartCol = Math.min(target.col, target.endCol);
      const snapStartRow = Math.min(srcStartRow, tgtStartRow);
      const snapEndRow = Math.max(srcEndRow, tgtEndRow);
      const snapStartCol = Math.min(srcStartCol, tgtStartCol);
      const snapEndCol = Math.max(srcEndCol, tgtEndCol);

      if (options?.useGranularOps && options.sendOps) {
        const sheetId = latestWorkbook.current.activeSheetId;
        const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
        if (!sheet) return;
        const before = snapshotRangeCells(
          sheet,
          snapStartRow,
          snapEndRow,
          snapStartCol,
          snapEndCol,
        );
        const filled = fillSelection(sheet, source, target);
        const after = snapshotRangeCells(
          filled,
          snapStartRow,
          snapEndRow,
          snapStartCol,
          snapEndCol,
        );
        const redo = restoreRangeOp(sheetId, after);
        const undo = restoreRangeOp(sheetId, before);
        pushCollaborativeOps([redo], { undo: [undo], redo: [redo] });
        return;
      }

      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        draft.sheets[index] = fillSelection(draft.sheets[index]!, source, target);
      });
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const cutRange = useCallback(
    (selection: GridSelection): string[][] => {
      const startRow = Math.min(selection.row, selection.endRow);
      const endRow = Math.max(selection.row, selection.endRow);
      const startCol = Math.min(selection.col, selection.endCol);
      const endCol = Math.max(selection.col, selection.endCol);
      const lines: string[][] = [];
      const sheetId = latestWorkbook.current.activeSheetId;
      const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
      if (!sheet) return lines;

      for (let row = startRow; row <= endRow; row++) {
        const rowVals: string[] = [];
        for (let col = startCol; col <= endCol; col++) {
          rowVals.push(getCellInput(sheet, row, col));
        }
        lines.push(rowVals);
      }

      if (options?.useGranularOps && options.sendOps) {
        const before = snapshotRangeCells(sheet, startRow, endRow, startCol, endCol);
        const after: Record<string, CellData | null> = {};
        for (const [key, data] of Object.entries(before)) {
          if (!data) {
            after[key] = null;
            continue;
          }
          const next = { ...data };
          delete next.input;
          after[key] = Object.keys(next).length ? next : null;
        }
        const redo = restoreRangeOp(sheetId, after);
        const undo = restoreRangeOp(sheetId, before);
        pushCollaborativeOps([redo], { undo: [undo], redo: [redo] });
        return lines;
      }

      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        let nextSheet = draft.sheets[index]!;
        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            nextSheet = setCellInput(nextSheet, row, col, "");
          }
        }
        draft.sheets[index] = nextSheet;
      });
      return lines;
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const pasteCells = useCallback(
    (startRow: number, startCol: number, cells: ClipboardCell[][]) => {
      if (options?.readOnly) return;
      const maxCols = cells.reduce((max, row) => Math.max(max, row.length), 0);
      const endRow = startRow + cells.length - 1;
      const endCol = startCol + maxCols - 1;

      if (options?.useGranularOps && options.sendOps) {
        const sheetId = latestWorkbook.current.activeSheetId;
        const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
        if (!sheet) return;
        const before = snapshotRangeCells(
          sheet,
          startRow,
          Math.max(startRow, endRow),
          startCol,
          Math.max(startCol, endCol),
        );
        const after: Record<string, CellData | null> = { ...before };
        for (let rOff = 0; rOff < cells.length; rOff++) {
          const row = cells[rOff]!;
          for (let cOff = 0; cOff < row.length; cOff++) {
            const r = startRow + rOff;
            const c = startCol + cOff;
            if (r < 0 || c < 0 || r >= sheet.rowCount || c >= sheet.colCount) continue;
            const key = cellKey(r, c);
            const merged = mergeCellPaste(sheet.cells[key], row[cOff]!);
            after[key] = merged ? structuredClone(merged) : null;
          }
        }
        const redo = restoreRangeOp(sheetId, after);
        const undo = restoreRangeOp(sheetId, before);
        pushCollaborativeOps([redo], { undo: [undo], redo: [redo] });
        return;
      }

      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const sheet = draft.sheets[index]!;
        const nextCells = { ...sheet.cells };
        for (let rOff = 0; rOff < cells.length; rOff++) {
          const row = cells[rOff]!;
          for (let cOff = 0; cOff < row.length; cOff++) {
            const r = startRow + rOff;
            const c = startCol + cOff;
            if (r < 0 || c < 0 || r >= sheet.rowCount || c >= sheet.colCount) continue;
            const key = cellKey(r, c);
            const merged = mergeCellPaste(nextCells[key], row[cOff]!);
            if (merged) nextCells[key] = merged;
            else delete nextCells[key];
          }
        }
        draft.sheets[index] = { ...sheet, cells: nextCells };
      });
    },
    [commitWorkbook, options, pushCollaborativeOps],
  );

  const clearRange = useCallback((selection: GridSelection) => {
    const startRow = Math.min(selection.row, selection.endRow);
    const endRow = Math.max(selection.row, selection.endRow);
    const startCol = Math.min(selection.col, selection.endCol);
    const endCol = Math.max(selection.col, selection.endCol);

    if (options?.useGranularOps && options.sendOps) {
      const sheetId = latestWorkbook.current.activeSheetId;
      const sheet = latestWorkbook.current.sheets.find((s) => s.id === sheetId);
      if (!sheet) return;
      const before = snapshotRangeCells(sheet, startRow, endRow, startCol, endCol);
      const after: Record<string, CellData | null> = {};
      for (const [key, data] of Object.entries(before)) {
        if (!data) {
          after[key] = null;
          continue;
        }
        const next = { ...data };
        delete next.input;
        after[key] = Object.keys(next).length ? next : null;
      }
      const redo = restoreRangeOp(sheetId, after);
      const undo = restoreRangeOp(sheetId, before);
      pushCollaborativeOps([redo], { undo: [undo], redo: [redo] });
      return;
    }

    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      let sheet = draft.sheets[index]!;
      for (let row = startRow; row <= endRow; row++) {
        for (let col = startCol; col <= endCol; col++) {
          sheet = setCellInput(sheet, row, col, "");
        }
      }
      draft.sheets[index] = sheet;
    });
  }, [commitWorkbook, options, pushCollaborativeOps]);

  const toggleFilter = useCallback(() => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      const sheet = draft.sheets[index]!;
      if (sheet.autoFilter?.enabled) {
        draft.sheets[index] = disableAutoFilter(sheet);
        return;
      }
      draft.sheets[index] = enableAutoFilter(sheet, {
        startRow: 0,
        startCol: 0,
        endRow: sheet.rowCount - 1,
        endCol: sheet.colCount - 1,
      });
    });
  }, [commitWorkbook]);

  const sortColumn = useCallback(
    (col: number, ascending: boolean) => {
      const sheet = latestWorkbook.current.sheets.find(
        (s) => s.id === latestWorkbook.current.activeSheetId,
      );
      if (!sheet) return;

      let firstData = Number.POSITIVE_INFINITY;
      let lastData = -1;
      for (let row = 0; row < sheet.rowCount; row++) {
        const input = getCellInput(sheet, row, col);
        const display =
          engineRef.current?.getDisplayValue(sheet.id, row, col) ?? input;
        const value = input.trim() || display.trim();
        if (!value) continue;
        firstData = Math.min(firstData, row);
        lastData = Math.max(lastData, row);
      }

      if (lastData < 0 || lastData <= firstData) {
        window.alert(
          "Add at least two values in this column (press Enter in the formula bar to save), then sort again.",
        );
        return;
      }

      const headerRow = sheet.autoFilter?.enabled
        ? sheet.autoFilter.range.startRow
        : firstData;
      const startRow = sheet.autoFilter?.enabled
        ? Math.max(headerRow + 1, firstData)
        : firstData;
      const endRow = sheet.autoFilter?.enabled
        ? Math.min(sheet.autoFilter.range.endRow, lastData)
        : lastData;

      if (endRow <= startRow) {
        window.alert("Need at least two data rows to sort in this column.");
        return;
      }

      sortByColumn(col, ascending, startRow, endRow);
    },
    [sortByColumn],
  );

  const clearColumnFilter = useCallback(
    (col: number) => {
      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const sheet = draft.sheets[index]!;
        if (!sheet.autoFilter) return;
        const columns = sheet.autoFilter.columns.filter((c) => c.col !== col);
        const nextFilter = { ...sheet.autoFilter, columns, hiddenRows: [] as number[] };
        const hiddenRows = computeHiddenRows(sheet, nextFilter, (row, c) =>
          engineRef.current?.getDisplayValue(sheet.id, row, c) ?? "",
        );
        draft.sheets[index] = {
          ...sheet,
          autoFilter: { ...nextFilter, hiddenRows },
        };
      });
    },
    [commitWorkbook],
  );

  const applyColumnFilter = useCallback(
    (filter: FilterColumnState) => {
      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const sheet = draft.sheets[index]!;
        draft.sheets[index] = setColumnFilter(sheet, filter, (row, col) =>
          engineRef.current?.getDisplayValue(sheet.id, row, col) ?? "",
        );
      });
    },
    [commitWorkbook],
  );

  const clearFilters = useCallback(() => {
    commitWorkbook((draft) => {
      const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
      if (index < 0) return;
      draft.sheets[index] = clearAutoFilter(draft.sheets[index]!);
    });
  }, [commitWorkbook]);

  const findAndReplace = useCallback(
    (
      query: string,
      replacement: string,
      matchCase: boolean,
      replaceAll: boolean,
    ) => {
      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const result = replaceInSheet(
          draft.sheets[index]!,
          query,
          replacement,
          matchCase,
          replaceAll,
        );
        draft.sheets[index] = result.sheet;
      });
    },
    [commitWorkbook],
  );

  const setValidationOnSelection = useCallback(
    (selection: GridSelection, validation: DataValidation) => {
      commitWorkbook((draft) => {
        const index = draft.sheets.findIndex((s) => s.id === draft.activeSheetId);
        if (index < 0) return;
        const sheet = draft.sheets[index]!;
        const nextCells = { ...sheet.cells };
        for (
          let row = Math.min(selection.row, selection.endRow);
          row <= Math.max(selection.row, selection.endRow);
          row++
        ) {
          for (
            let col = Math.min(selection.col, selection.endCol);
            col <= Math.max(selection.col, selection.endCol);
            col++
          ) {
            const key = cellKey(row, col);
            nextCells[key] = attachValidation(nextCells[key] ?? {}, validation);
          }
        }
        draft.sheets[index] = { ...sheet, cells: nextCells };
      });
    },
    [commitWorkbook],
  );

  const toggleHideRows = useCallback((rows: number[]) => {
    commitWorkbook((draft) => {
      const sheet = draft.sheets.find((s) => s.id === draft.activeSheetId);
      if (!sheet) return;
      const hidden = new Set(sheet.hiddenRows);
      for (const row of rows) {
        if (hidden.has(row)) hidden.delete(row);
        else hidden.add(row);
      }
      sheet.hiddenRows = [...hidden].sort((a, b) => a - b);
    });
  }, [commitWorkbook]);

  const unhideAllRows = useCallback(() => {
    commitWorkbook((draft) => {
      const sheet = draft.sheets.find((s) => s.id === draft.activeSheetId);
      if (!sheet) return;
      sheet.hiddenRows = [];
    });
  }, [commitWorkbook]);

  const unhideAllCols = useCallback(() => {
    commitWorkbook((draft) => {
      const sheet = draft.sheets.find((s) => s.id === draft.activeSheetId);
      if (!sheet) return;
      sheet.hiddenCols = [];
    });
  }, [commitWorkbook]);

  const toggleHideCols = useCallback((cols: number[]) => {
    commitWorkbook((draft) => {
      const sheet = draft.sheets.find((s) => s.id === draft.activeSheetId);
      if (!sheet) return;
      const hidden = new Set(sheet.hiddenCols);
      for (const col of cols) {
        if (hidden.has(col)) hidden.delete(col);
        else hidden.add(col);
      }
      sheet.hiddenCols = [...hidden].sort((a, b) => a - b);
    });
  }, [commitWorkbook]);

  const autofitColumn = useCallback(
    (col: number) => {
      const sheet = activeSheet;
      const width = autofitColumnWidth(sheet, col, (row, c) =>
        engineRef.current?.getDisplayValue(sheet.id, row, c) ?? "",
      );
      setColWidth(col, width);
    },
    [activeSheet, setColWidth],
  );

  const toggleCellFormat = useCallback(
    (selection: GridSelection, key: keyof CellFormat) => {
      const startRow = Math.min(selection.row, selection.endRow);
      const startCol = Math.min(selection.col, selection.endCol);
      const primary =
        activeSheet.cells[cellKey(startRow, startCol)]?.format ?? {};
      const current = Boolean(primary[key]);
      const nextValue = !current as CellFormat[keyof CellFormat];
      setFormatOnSelection(selection, {
        [key]: nextValue,
      } as Partial<CellFormat>);
    },
    [activeSheet.cells, setFormatOnSelection],
  );

  const sortSelection = useCallback(
    (selection: GridSelection, ascending: boolean) => {
      const sortCol = Math.min(selection.col, selection.endCol);
      const r0 = Math.min(selection.row, selection.endRow);
      const r1 = Math.max(selection.row, selection.endRow);
      if (r1 <= r0) {
        sortColumn(sortCol, ascending);
        return;
      }
      sortByColumn(sortCol, ascending, r0, r1);
    },
    [sortByColumn, sortColumn],
  );

  const applyBorders = useCallback(
    (selection: GridSelection, border: CellFormat["borders"]) => {
      setFormatOnSelection(selection, { borders: border });
    },
    [setFormatOnSelection],
  );

  return {
    workbook,
    activeSheet,
    saveState,
    saveMessage,
    setColWidth,
    setRowHeight,
    ensureSheetCapacity,
    setActiveSheet,
    setCell,
    setCellFormat,
    setFormatOnSelection,
    addSheet,
    duplicateActiveSheet,
    deleteSheet,
    renameSheet,
    reorderSheet,
    insertSheetRows,
    deleteSheetRows,
    insertSheetColumns,
    deleteSheetColumns,
    sortByColumn,
    sortColumn,
    sortSelection,
    clearColumnFilter,
    toggleCellFormat,
    mergeSelection,
    unmergeAt,
    setFreeze,
    setZoom,
    toggleGridlines,
    undo,
    redo,
    getDisplay,
    getRaw,
    applyRemoteOps,
    applyConflictWorkbook,
    applyRemoteWorkbookRestore,
    setConditionalFormats,
    setCellComment,
    startCellComment,
    replyCellComment,
    resolveCellComment,
    reopenCellComment,
    deleteCellComment,
    addChart,
    fillTo,
    cutRange,
    pasteCells,
    clearRange,
    toggleFilter,
    applyColumnFilter,
    clearFilters,
    findAndReplace,
    setValidationOnSelection,
    toggleHideRows,
    unhideAllRows,
    toggleHideCols,
    unhideAllCols,
    autofitColumn,
    applyBorders,
  };
}