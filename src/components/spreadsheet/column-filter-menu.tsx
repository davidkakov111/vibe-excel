"use client";

import { useEffect, useRef } from "react";
import type { FilterColumnState, Worksheet } from "@/lib/workbook/types";
import { indexToColumnLabel } from "@/lib/spreadsheet-types";
import { getCellInput } from "@/lib/workbook/types";

export interface ColumnFilterMenuState {
  col: number;
  top: number;
  left: number;
}

interface ColumnFilterMenuProps {
  sheet: Worksheet;
  sheetId: string;
  menu: ColumnFilterMenuState | null;
  getDisplay: (sheetId: string, row: number, col: number) => string;
  activeFilter?: FilterColumnState;
  onClose: () => void;
  onSortAsc: (col: number) => void;
  onSortDesc: (col: number) => void;
  onApply: (filter: FilterColumnState) => void;
  onClearColumn: (col: number) => void;
}

export function ColumnFilterMenu({
  sheet,
  sheetId,
  menu,
  getDisplay,
  activeFilter,
  onClose,
  onSortAsc,
  onSortDesc,
  onApply,
  onClearColumn,
}: ColumnFilterMenuProps) {
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!menu) return;
    const onDoc = (event: MouseEvent) => {
      if (!panelRef.current?.contains(event.target as Node)) {
        onClose();
      }
    };
    document.addEventListener("click", onDoc);
    return () => document.removeEventListener("click", onDoc);
  }, [menu, onClose]);

  if (!menu || !sheet.autoFilter?.enabled) return null;

  const { startRow, endRow } = sheet.autoFilter.range;
  const dataStart = startRow + 1;
  const values = uniqueValuesForColumn(
    sheet,
    sheetId,
    menu.col,
    dataStart,
    endRow,
    getDisplay,
  );

  const selected =
    activeFilter?.operator === "oneOf" && activeFilter.allowedValues
      ? new Set(activeFilter.allowedValues)
      : new Set(values);

  const allChecked = selected.size === values.length;

  const toggleValue = (value: string) => {
    const next = new Set(selected);
    if (next.has(value)) next.delete(value);
    else next.add(value);
    if (next.size === 0) {
      onApply({
        col: menu.col,
        operator: "oneOf",
        value: "",
        allowedValues: [],
      });
      return;
    }
    if (next.size === values.length) {
      onClearColumn(menu.col);
      return;
    }
    onApply({
      col: menu.col,
      operator: "oneOf",
      value: "",
      allowedValues: [...next],
    });
  };

  return (
    <div
      ref={panelRef}
      className="fixed z-50 w-56 rounded border border-zinc-400 bg-white py-1 text-sm text-zinc-900 shadow-lg"
      onMouseDown={(event) => event.stopPropagation()}
      style={{ top: menu.top, left: menu.left }}
      role="dialog"
      aria-label={`Filter column ${indexToColumnLabel(menu.col)}`}
    >
      <button
        type="button"
        className="block w-full px-3 py-1.5 text-left hover:bg-zinc-100"
        onClick={() => {
          onSortAsc(menu.col);
          onClose();
        }}
      >
        Sort A → Z
      </button>
      <button
        type="button"
        className="block w-full px-3 py-1.5 text-left hover:bg-zinc-100"
        onClick={() => {
          onSortDesc(menu.col);
          onClose();
        }}
      >
        Sort Z → A
      </button>
      <div className="my-1 border-t border-zinc-200" />
      <button
        type="button"
        className="block w-full px-3 py-1.5 text-left text-red-700 hover:bg-zinc-100"
        onClick={() => {
          onClearColumn(menu.col);
          onClose();
        }}
      >
        Clear filter from column
      </button>
      <div className="my-1 border-t border-zinc-200" />
      <div className="px-3 py-1 text-xs font-semibold text-zinc-500">
        Filter by value
      </div>
      <label className="flex cursor-pointer items-center gap-2 px-3 py-0.5 hover:bg-zinc-50">
        <input
          type="checkbox"
          checked={allChecked}
          onChange={() => {
            if (allChecked) {
              onApply({
                col: menu.col,
                operator: "oneOf",
                value: "",
                allowedValues: [],
              });
            } else {
              onClearColumn(menu.col);
            }
          }}
        />
        <span>(Select all)</span>
      </label>
      <div className="max-h-48 overflow-y-auto">
        {values.map((value) => (
          <label
            key={value || "__blank__"}
            className="flex cursor-pointer items-center gap-2 px-3 py-0.5 hover:bg-zinc-50"
          >
            <input
              type="checkbox"
              checked={selected.has(value)}
              onChange={() => toggleValue(value)}
            />
            <span className="truncate">{value || "(Blanks)"}</span>
          </label>
        ))}
      </div>
    </div>
  );
}

function uniqueValuesForColumn(
  sheet: Worksheet,
  sheetId: string,
  col: number,
  startRow: number,
  endRow: number,
  getDisplay: (sheetId: string, row: number, col: number) => string,
): string[] {
  const values = new Set<string>();
  for (let row = startRow; row <= endRow; row++) {
    const raw = getCellInput(sheet, row, col);
    const display = getDisplay(sheetId, row, col);
    const value = raw.startsWith("=") ? display : raw || display;
    values.add(value);
  }
  const sorted = [...values].sort((a, b) =>
    a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
  );
  return sorted;
}