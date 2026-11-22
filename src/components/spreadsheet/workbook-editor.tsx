"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { nanoid } from "nanoid";
import { useCollaboration } from "@/hooks/use-collaboration";
import type { WorkbookOp } from "@/lib/workbook/ops";
import { parseA1 } from "@/lib/formula/engine";
import { cellKey } from "@/lib/spreadsheet-types";
import type { WorkbookDocument } from "@/lib/workbook/types";
import type { GridSelection } from "@/lib/workbook/selection";
import { useWorkbookEditor } from "@/hooks/use-workbook-editor";
import { FormulaBar } from "@/components/spreadsheet/formula-bar";
import { SheetTabs } from "@/components/spreadsheet/sheet-tabs";
import { SpreadsheetToolbar } from "@/components/spreadsheet/toolbar";
import { SpreadsheetChartPanel } from "@/components/spreadsheet/spreadsheet-chart";
import { FindReplaceDialog } from "@/components/spreadsheet/find-replace-dialog";
import { ChartHelpDialog } from "@/components/spreadsheet/chart-help-dialog";
import {
  WorkbookGrid,
  type WorkbookGridHandle,
} from "@/components/spreadsheet/workbook-grid";
import { ShareDialog } from "@/components/spreadsheet/share-dialog";
import { ConditionalFormatDialog } from "@/components/spreadsheet/conditional-format-dialog";
import { VersionHistoryDialog } from "@/components/spreadsheet/version-history-dialog";

interface WorkbookEditorProps {
  spreadsheetId: string;
  initialName: string;
  initialWorkbook: WorkbookDocument;
  initialRevision?: number;
  canEdit?: boolean;
  isOwner?: boolean;
  userId?: string;
  userName?: string;
}

export function WorkbookEditor({
  spreadsheetId,
  initialName,
  initialWorkbook,
  initialRevision = 0,
  canEdit = true,
  isOwner = true,
  userId = "local",
  userName = "You",
}: WorkbookEditorProps) {
  const clientId = useMemo(() => nanoid(), []);
  const [revision, setRevision] = useState(initialRevision);
  const pushOpsRef = useRef<(ops: WorkbookOp[]) => Promise<boolean>>(async () => false);
  const revisionRef = useRef(initialRevision);
  revisionRef.current = revision;

  const editor = useWorkbookEditor(spreadsheetId, initialWorkbook, {
    readOnly: !canEdit,
    useGranularOps: canEdit,
    sendOps: (ops) => pushOpsRef.current(ops),
    getRevision: () => revisionRef.current,
    userId,
    userName,
  });
  const [selection, setSelection] = useState<GridSelection>({
    row: 0,
    col: 0,
    endRow: 0,
    endCol: 0,
  });
  const [formulaDraft, setFormulaDraft] = useState("");
  const [findOpen, setFindOpen] = useState(false);
  const [chartMessage, setChartMessage] = useState<string | null>(null);
  const [chartHelpOpen, setChartHelpOpen] = useState(false);
  const [freezeHint, setFreezeHint] = useState<string | null>(null);
  const [shareOpen, setShareOpen] = useState(false);
  const [cfOpen, setCfOpen] = useState(false);
  const [versionsOpen, setVersionsOpen] = useState(false);
  const gridRef = useRef<WorkbookGridHandle>(null);

  const collaboration = useCollaboration({
    spreadsheetId,
    clientId,
    userId,
    userName,
    enabled: true,
    revision,
    onRevision: setRevision,
    onRemoteOps: editor.applyRemoteOps,
    onReconcile: (workbook, nextRevision) => {
      editor.applyConflictWorkbook(workbook);
      setRevision(nextRevision);
    },
    onRemoteWorkbookRestore: (workbook, nextRevision) => {
      editor.applyRemoteWorkbookRestore(workbook);
      setRevision(nextRevision);
    },
    activeSheetId: editor.activeSheet.id,
    selection,
  });

  useEffect(() => {
    pushOpsRef.current = collaboration.pushOps;
  }, [collaboration.pushOps]);

  const normalized = useMemo(
    () => ({
      row: Math.min(selection.row, selection.endRow),
      col: Math.min(selection.col, selection.endCol),
      endRow: Math.max(selection.row, selection.endRow),
      endCol: Math.max(selection.col, selection.endCol),
    }),
    [selection],
  );

  const rawAtPrimary = editor.getRaw(
    editor.activeSheet.id,
    normalized.row,
    normalized.col,
  );

  const syncFormulaDraft = useCallback(
    (sel: GridSelection) => {
      const row = Math.min(sel.row, sel.endRow);
      const col = Math.min(sel.col, sel.endCol);
      setFormulaDraft(editor.getRaw(editor.activeSheet.id, row, col));
    },
    [editor],
  );

  const onSelectionChange = useCallback(
    (sel: GridSelection) => {
      setSelection(sel);
      syncFormulaDraft(sel);
    },
    [syncFormulaDraft],
  );

  const commitFormulaBar = useCallback(() => {
    editor.setCell(normalized.row, normalized.col, formulaDraft);
  }, [editor, formulaDraft, normalized.col, normalized.row]);

  const flushEditsAndSort = useCallback(
    (ascending: boolean) => {
      gridRef.current?.commitPendingEdit();
      commitFormulaBar();
      editor.sortSelection(selection, ascending);
    },
    [commitFormulaBar, editor, selection],
  );

  const applyFormatToSelection = useCallback(
    (format: Parameters<typeof editor.setFormatOnSelection>[1]) => {
      editor.setFormatOnSelection(selection, format);
    },
    [editor, selection],
  );

  const primaryFormat =
    editor.activeSheet.cells[cellKey(normalized.row, normalized.col)]?.format;

  const saveLabel =
    editor.saveState === "saving"
      ? "Saving…"
      : editor.saveState === "saved"
        ? "All changes saved"
        : editor.saveState === "error"
          ? editor.saveMessage ?? "Save failed"
          : "Ready";

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-[#e6e6e6]">
      <EditorHeader
        title={initialName}
        spreadsheetId={spreadsheetId}
        saveLabel={saveLabel}
        isError={editor.saveState === "error"}
        canShare={isOwner}
        onShare={() => setShareOpen(true)}
        readOnly={!canEdit}
        collaborators={collaboration.others.length + 1}
        onVersionHistory={() => setVersionsOpen(true)}
      />
      <ShareDialog
        spreadsheetId={spreadsheetId}
        open={shareOpen}
        onClose={() => setShareOpen(false)}
      />

      <SpreadsheetToolbar
        zoom={editor.workbook.settings.zoom}
        activeFormat={primaryFormat}
        onUndo={editor.undo}
        onRedo={editor.redo}
        onFindReplace={() => setFindOpen(true)}
        onToggleFilter={editor.toggleFilter}
        onClearFilters={editor.clearFilters}
        filterEnabled={Boolean(editor.activeSheet.autoFilter?.enabled)}
        onToggleBold={() => editor.toggleCellFormat(selection, "bold")}
        onToggleItalic={() => editor.toggleCellFormat(selection, "italic")}
        onToggleUnderline={() => editor.toggleCellFormat(selection, "underline")}
        onToggleStrike={() =>
          editor.toggleCellFormat(selection, "strikethrough")
        }
        onToggleWrap={() => {
          editor.toggleCellFormat(selection, "wrap");
          for (let r = normalized.row; r <= normalized.endRow; r++) {
            const current =
              editor.activeSheet.rowHeights[String(r)] ?? 28;
            editor.setRowHeight(r, Math.max(current, 48));
          }
        }}
        onInsertRow={() => editor.insertSheetRows(normalized.row)}
        onDeleteRow={() => editor.deleteSheetRows(normalized.row)}
        onInsertColumn={() => editor.insertSheetColumns(normalized.col)}
        onDeleteColumn={() => editor.deleteSheetColumns(normalized.col)}
        onHideRows={() => {
          const rows: number[] = [];
          for (let r = normalized.row; r <= normalized.endRow; r++) rows.push(r);
          editor.toggleHideRows(rows);
        }}
        onUnhideRows={editor.unhideAllRows}
        onHideCols={() => {
          const cols: number[] = [];
          for (let c = normalized.col; c <= normalized.endCol; c++) cols.push(c);
          editor.toggleHideCols(cols);
        }}
        onAutofitCol={() => editor.autofitColumn(normalized.col)}
        onSortAsc={() => flushEditsAndSort(true)}
        onSortDesc={() => flushEditsAndSort(false)}
        onFreeze={() => {
          if (normalized.row === 0) {
            window.alert(
              "Click a cell on the first row you want to scroll (e.g. row 3), then Freeze panes — rows above stay visible while you scroll down.",
            );
            return;
          }
          editor.setFreeze(normalized.row);
          setFreezeHint(
            `Frozen: rows 1–${normalized.row} stay visible when you scroll down.`,
          );
        }}
        onUnfreeze={() => {
          editor.setFreeze(0);
          setFreezeHint(null);
        }}
        onMerge={() =>
          editor.mergeSelection({
            startRow: normalized.row,
            startCol: normalized.col,
            endRow: normalized.endRow,
            endCol: normalized.endCol,
          })
        }
        onUnmerge={() => editor.unmergeAt(normalized.row, normalized.col)}
        onToggleGridlines={editor.toggleGridlines}
        onZoomIn={() => editor.setZoom(editor.workbook.settings.zoom + 10)}
        onZoomOut={() => editor.setZoom(editor.workbook.settings.zoom - 10)}
        onFormat={applyFormatToSelection}
        readOnly={!canEdit}
        onEditComment={() => gridRef.current?.openCommentOnSelection()}
        onConditionalFormat={() => setCfOpen(true)}
        onAllBorders={() =>
          editor.applyBorders(selection, {
            top: { style: "thin", color: "#334155" },
            right: { style: "thin", color: "#334155" },
            bottom: { style: "thin", color: "#334155" },
            left: { style: "thin", color: "#334155" },
          })
        }
        onSetValidationList={() => {
          const raw = window.prompt("Dropdown values (comma-separated)", "Yes,No");
          if (!raw) return;
          editor.setValidationOnSelection(selection, {
            kind: "list",
            values: raw.split(",").map((v) => v.trim()).filter(Boolean),
            allowBlank: true,
            errorMessage: "Pick a value from the list.",
          });
        }}
        onAddChart={(type) => {
          if (
            normalized.row === normalized.endRow &&
            normalized.col === normalized.endCol
          ) {
            setChartHelpOpen(true);
            return;
          }
          editor.addChart(
            type,
            {
              startRow: normalized.row,
              startCol: normalized.col,
              endRow: normalized.endRow,
              endCol: normalized.endCol,
            },
            `${type.charAt(0).toUpperCase()}${type.slice(1)} chart`,
          );
          setChartMessage(
            `Added ${type} chart — scroll down to the chart preview below the sheet.`,
          );
          requestAnimationFrame(() => {
            document.getElementById("chart-panel")?.scrollIntoView({
              behavior: "smooth",
              block: "nearest",
            });
          });
        }}
      />

      {freezeHint ? (
        <p className="border-b border-emerald-200 bg-emerald-50 px-3 py-1 text-sm text-emerald-900">
          {freezeHint}
        </p>
      ) : null}

      {chartMessage ? (
        <p className="border-b border-amber-200 bg-amber-50 px-3 py-1 text-sm text-amber-900">
          {chartMessage}
        </p>
      ) : null}

      <FormulaBar
        sheetId={editor.activeSheet.id}
        selection={selection}
        rawValue={formulaDraft || rawAtPrimary}
        onChange={setFormulaDraft}
        onCommit={commitFormulaBar}
        onNavigateToAddress={(address) => {
          const parts = address.split(":");
          const start = parseA1(parts[0] ?? "");
          if (!start) return false;
          const end = parts[1] ? parseA1(parts[1]) : start;
          if (!end) return false;
          onSelectionChange({
            row: start.row,
            col: start.col,
            endRow: end.row,
            endCol: end.col,
          });
          return true;
        }}
      />

      <div className="flex min-h-0 flex-1 flex-col overflow-hidden px-2 pb-2">
      <WorkbookGrid
        ref={gridRef}
        sheet={editor.activeSheet}
        sheetId={editor.activeSheet.id}
        zoom={editor.workbook.settings.zoom}
        showGridlines={editor.workbook.settings.showGridlines}
        getDisplay={editor.getDisplay}
        selection={selection}
        onSelectionChange={onSelectionChange}
        onCommitCell={(row, col, value) => {
          editor.setCell(row, col, value);
          if (row === normalized.row && col === normalized.col) {
            setFormulaDraft(value);
          }
        }}
        onColResize={editor.setColWidth}
        onRowResize={editor.setRowHeight}
        onFillEnd={(target) => editor.fillTo(selection, target)}
        onCut={() => void editor.cutRange(selection)}
        onPasteCells={(row, col, cells) => editor.pasteCells(row, col, cells)}
        onApplyColumnFilter={editor.applyColumnFilter}
        onSortColumn={(col, asc) => {
          gridRef.current?.commitPendingEdit();
          commitFormulaBar();
          editor.sortColumn(col, asc);
        }}
        onClearColumnFilter={editor.clearColumnFilter}
        remotePresence={collaboration.others.filter(
          (p) => p.sheetId === editor.activeSheet.id,
        )}
        commentUserId={userId}
        onCommentStart={(row, col, text) => editor.startCellComment(row, col, text)}
        onCommentReply={(row, col, text) => editor.replyCellComment(row, col, text)}
        onCommentResolve={(row, col) => editor.resolveCellComment(row, col)}
        onCommentReopen={(row, col) => editor.reopenCellComment(row, col)}
        onCommentDelete={(row, col) => editor.deleteCellComment(row, col)}
        onEnsureCapacity={(minRow, minCol) =>
          editor.ensureSheetCapacity(minRow, minCol)
        }
        readOnly={!canEdit}
      />
      </div>

      <div id="chart-panel" className="max-h-40 shrink-0 overflow-y-auto border-t border-[#b4b4b4] bg-white">
      <SpreadsheetChartPanel
        sheet={editor.activeSheet}
        sheetId={editor.activeSheet.id}
        getDisplay={editor.getDisplay}
      />
      </div>

      <ChartHelpDialog
        open={chartHelpOpen}
        onClose={() => setChartHelpOpen(false)}
      />

      <SheetTabs
        sheets={editor.workbook.sheets}
        activeSheetId={editor.workbook.activeSheetId}
        onSelect={editor.setActiveSheet}
        onAdd={editor.addSheet}
        onRename={editor.renameSheet}
        onDelete={editor.deleteSheet}
        onDuplicate={editor.duplicateActiveSheet}
        onReorder={editor.reorderSheet}
      />

      <FindReplaceDialog
        open={findOpen}
        onClose={() => setFindOpen(false)}
        onReplace={(query, replacement, matchCase, replaceAll) => {
          editor.findAndReplace(query, replacement, matchCase, replaceAll);
        }}
      />

      <ConditionalFormatDialog
        open={cfOpen}
        onClose={() => setCfOpen(false)}
        selection={selection}
        sets={editor.activeSheet.conditionalFormats ?? []}
        readOnly={!canEdit}
        onApply={(sets) => editor.setConditionalFormats(sets)}
      />

      <VersionHistoryDialog
        spreadsheetId={spreadsheetId}
        open={versionsOpen}
        clientId={clientId}
        baseRevision={revision}
        readOnly={!canEdit}
        onClose={() => setVersionsOpen(false)}
        onRestored={(workbook, nextRevision) => {
          editor.applyRemoteWorkbookRestore(workbook);
          setRevision(nextRevision);
        }}
      />
    </div>
  );
}

function EditorHeader({
  title,
  spreadsheetId,
  saveLabel,
  isError,
  canShare,
  onShare,
  onVersionHistory,
  readOnly,
  collaborators,
}: {
  title: string;
  spreadsheetId: string;
  saveLabel: string;
  isError: boolean;
  canShare: boolean;
  onShare: () => void;
  onVersionHistory: () => void;
  readOnly: boolean;
  collaborators: number;
}) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-2 border-b border-[#0d5c2e] bg-[#217346] px-3 py-2 text-white">
      <div className="flex items-center gap-4">
        <Link
          href="/dashboard"
          className="text-sm text-white/90 underline hover:text-white"
        >
          ← File
        </Link>
        <h1 className="text-lg font-semibold">{title}</h1>
        <p
          className={`text-xs ${isError ? "text-red-200" : "text-emerald-100"}`}
          role="status"
        >
          {readOnly ? "View only" : saveLabel}
        </p>
        <p className="text-xs text-emerald-100">
          {collaborators} online
        </p>
      </div>
      {canShare ? (
        <button
          type="button"
          onClick={onShare}
          className="rounded bg-white/15 px-3 py-1.5 text-sm font-medium hover:bg-white/25"
        >
          Share
        </button>
      ) : null}
      <button
        type="button"
        onClick={onVersionHistory}
        className="rounded bg-white/15 px-3 py-1.5 text-sm font-medium hover:bg-white/25"
      >
        Version history
      </button>
      <a
        href={`/api/spreadsheets/${spreadsheetId}/export`}
        className="rounded bg-white/15 px-3 py-1.5 text-sm font-medium hover:bg-white/25"
      >
        Download .xlsx
      </a>
      <a
        href={`/api/spreadsheets/${spreadsheetId}/export/pdf`}
        className="rounded bg-white/15 px-3 py-1.5 text-sm font-medium hover:bg-white/25"
      >
        Download PDF
      </a>
    </div>
  );
}