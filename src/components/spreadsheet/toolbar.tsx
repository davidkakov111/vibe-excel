"use client";

import type { CellFormat } from "@/lib/workbook/types";

interface ToolbarProps {
  zoom: number;
  filterEnabled: boolean;
  activeFormat: CellFormat | undefined;
  onUndo: () => void;
  onRedo: () => void;
  onInsertRow: () => void;
  onDeleteRow: () => void;
  onInsertColumn: () => void;
  onDeleteColumn: () => void;
  onSortAsc: () => void;
  onSortDesc: () => void;
  onFreeze: () => void;
  onUnfreeze: () => void;
  onMerge: () => void;
  onUnmerge: () => void;
  onToggleGridlines: () => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onFormat: (format: Partial<CellFormat>) => void;
  onToggleBold: () => void;
  onToggleItalic: () => void;
  onToggleUnderline: () => void;
  onToggleStrike: () => void;
  onToggleWrap: () => void;
  onAddChart: (type: "bar" | "line" | "pie" | "area") => void;
  onFindReplace: () => void;
  onToggleFilter: () => void;
  onClearFilters: () => void;
  onHideRows: () => void;
  onUnhideRows: () => void;
  onHideCols: () => void;
  onAutofitCol: () => void;
  onSetValidationList: () => void;
  onAllBorders: () => void;
  onEditComment?: () => void;
  onConditionalFormat?: () => void;
  readOnly?: boolean;
}

export function SpreadsheetToolbar(props: ToolbarProps) {
  const fmt = props.activeFormat ?? {};
  const fontFamily = fmt.fontFamily ?? "Calibri";
  const fontSize = fmt.fontSize ?? 11;

  return (
    <div className="excel-ribbon border-b border-[#0d5c2e] bg-[#217346] text-white shadow-sm">
      <div className="flex flex-wrap items-stretch gap-0 px-1 py-1 text-xs">
        <RibbonGroup title="Clipboard">
          <RibbonBtn onClick={props.onUndo} title="Undo">
            ↶ Undo
          </RibbonBtn>
          <RibbonBtn onClick={props.onRedo} title="Redo">
            ↷ Redo
          </RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="Font">
          <select
            className="max-w-[7rem] rounded border border-[#185c37] bg-white px-1 py-1 text-zinc-900"
            value={fontFamily}
            onChange={(e) => props.onFormat({ fontFamily: e.target.value })}
            aria-label="Font family"
          >
            <option value="Calibri">Calibri</option>
            <option value="Arial">Arial</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Courier New">Courier New</option>
          </select>
          <select
            className="w-12 rounded border border-[#185c37] bg-white px-1 py-1 text-zinc-900"
            value={fontSize}
            onChange={(e) =>
              props.onFormat({ fontSize: Number(e.target.value) })
            }
            aria-label="Font size"
          >
            {[9, 10, 11, 12, 14, 16, 18, 20, 24].map((size) => (
              <option key={size} value={size}>
                {size}
              </option>
            ))}
          </select>
          <RibbonBtn
            active={Boolean(fmt.bold)}
            onClick={props.onToggleBold}
            title="Bold"
          >
            <b>B</b>
          </RibbonBtn>
          <RibbonBtn
            active={Boolean(fmt.italic)}
            onClick={props.onToggleItalic}
            title="Italic"
          >
            <i>I</i>
          </RibbonBtn>
          <RibbonBtn
            active={Boolean(fmt.underline)}
            onClick={props.onToggleUnderline}
            title="Underline"
          >
            <u>U</u>
          </RibbonBtn>
          <RibbonBtn
            active={Boolean(fmt.strikethrough)}
            onClick={props.onToggleStrike}
            title="Strikethrough"
          >
            S̶
          </RibbonBtn>
          <RibbonBtn
            onClick={() =>
              props.onFormat({
                fillColor:
                  fmt.fillColor === "#ffff00" ? "#ffffff" : "#ffff00",
              })
            }
            title="Fill color"
          >
            🪣 Fill
          </RibbonBtn>
          <RibbonBtn
            onClick={() =>
              props.onFormat({
                fontColor:
                  fmt.fontColor === "#c00000" ? "#000000" : "#c00000",
              })
            }
            title="Font color (toggles red / black text)"
          >
            <span className="text-red-400 underline">A</span>
          </RibbonBtn>
          <RibbonBtn
            active={Boolean(fmt.wrap)}
            onClick={props.onToggleWrap}
            title="Wrap text"
          >
            Wrap
          </RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="Alignment">
          <RibbonBtn
            active={fmt.horizontalAlign === "left" || !fmt.horizontalAlign}
            onClick={() => props.onFormat({ horizontalAlign: "left" })}
          >
            Left
          </RibbonBtn>
          <RibbonBtn
            active={fmt.horizontalAlign === "center"}
            onClick={() => props.onFormat({ horizontalAlign: "center" })}
          >
            Center
          </RibbonBtn>
          <RibbonBtn
            active={fmt.horizontalAlign === "right"}
            onClick={() => props.onFormat({ horizontalAlign: "right" })}
          >
            Right
          </RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="Data">
          <RibbonBtn onClick={props.onSortAsc} title="Sort selected column A→Z">
            Sort A→Z
          </RibbonBtn>
          <RibbonBtn onClick={props.onSortDesc} title="Sort selected column Z→A">
            Sort Z→A
          </RibbonBtn>
          <RibbonBtn
            active={props.filterEnabled}
            onClick={props.onToggleFilter}
            title="Turn on filters (dropdowns on column headers). Row 1 = headers."
          >
            {props.filterEnabled ? "Filters on" : "Filter"}
          </RibbonBtn>
          <RibbonBtn onClick={props.onClearFilters}>Clear filters</RibbonBtn>
          <RibbonBtn onClick={props.onMerge} title="Merge selected cells">
            Merge
          </RibbonBtn>
          <RibbonBtn onClick={props.onUnmerge}>Unmerge</RibbonBtn>
          <RibbonBtn onClick={props.onFindReplace}>Find</RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="Insert">
          <RibbonBtn onClick={props.onInsertRow}>+ Row</RibbonBtn>
          <RibbonBtn onClick={props.onDeleteRow}>− Row</RibbonBtn>
          <RibbonBtn onClick={props.onInsertColumn}>+ Column</RibbonBtn>
          <RibbonBtn onClick={props.onDeleteColumn}>− Column</RibbonBtn>
          <RibbonBtn onClick={props.onAutofitCol}>Fit column</RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="View">
          <RibbonBtn
            onClick={props.onFreeze}
            title="Keep rows above the active cell visible while scrolling (pick a cell on the first scrolling row)"
          >
            Freeze panes
          </RibbonBtn>
          <RibbonBtn onClick={props.onUnfreeze}>Unfreeze</RibbonBtn>
          <RibbonBtn onClick={props.onToggleGridlines}>Gridlines</RibbonBtn>
          <RibbonBtn onClick={props.onZoomOut}>−</RibbonBtn>
          <span className="self-center px-1 tabular-nums">{props.zoom}%</span>
          <RibbonBtn onClick={props.onZoomIn}>+</RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="Review">
          <RibbonBtn
            onClick={() => props.onEditComment?.()}
            title="Add or edit cell comment (Ctrl+Shift+M)"
            disabled={!props.onEditComment || props.readOnly}
          >
            Comment
          </RibbonBtn>
          <RibbonBtn
            onClick={() => props.onConditionalFormat?.()}
            title="Conditional formatting for selection"
            disabled={!props.onConditionalFormat || props.readOnly}
          >
            Cond. format
          </RibbonBtn>
        </RibbonGroup>

        <RibbonGroup title="Charts">
          <p className="mb-1 hidden w-full text-[10px] text-emerald-100 lg:block">
            Select cells (labels + numbers), then pick a chart —
            preview appears below the grid.
          </p>
          <RibbonBtn
            onClick={() => props.onAddChart("bar")}
            title="Bar chart from selection"
          >
            Bar chart
          </RibbonBtn>
          <RibbonBtn onClick={() => props.onAddChart("line")}>Line</RibbonBtn>
          <RibbonBtn onClick={() => props.onAddChart("pie")}>Pie</RibbonBtn>
          <RibbonBtn onClick={() => props.onAddChart("area")}>Area</RibbonBtn>
        </RibbonGroup>
      </div>
    </div>
  );
}

function RibbonGroup({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div className="flex flex-col border-r border-[#0d5c2e] px-2 last:border-r-0">
      <div className="flex flex-wrap items-center gap-1">{children}</div>
      <span className="mt-0.5 text-center text-[10px] text-emerald-100">
        {title}
      </span>
    </div>
  );
}

function RibbonBtn({
  children,
  onClick,
  title,
  active,
  disabled,
}: {
  children: React.ReactNode;
  onClick: () => void;
  title?: string;
  active?: boolean;
  disabled?: boolean;
}) {
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      onClick={onClick}
      className={`rounded px-2 py-1 text-white ${
        disabled
          ? "cursor-not-allowed bg-[#185c37]/50"
          : active
            ? "bg-[#0d5c2e] ring-1 ring-white/40"
            : "bg-[#185c37] hover:bg-[#0d5c2e]"
      }`}
    >
      {children}
    </button>
  );
}