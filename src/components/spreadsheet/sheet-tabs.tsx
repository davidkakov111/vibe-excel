"use client";

import type { Worksheet } from "@/lib/workbook/types";

interface SheetTabsProps {
  sheets: Worksheet[];
  activeSheetId: string;
  onSelect: (sheetId: string) => void;
  onAdd: () => void;
  onRename: (sheetId: string, name: string) => void;
  onDelete: (sheetId: string) => void;
  onDuplicate: () => void;
  onReorder: (sheetId: string, direction: -1 | 1) => void;
}

export function SheetTabs({
  sheets,
  activeSheetId,
  onSelect,
  onAdd,
  onRename,
  onDelete,
  onDuplicate,
  onReorder,
}: SheetTabsProps) {
  return (
    <div className="flex items-end gap-1 overflow-x-auto border-t border-[#b4b4b4] bg-[#f3f3f3] px-2 py-1">
      {sheets.map((sheet) => {
        const active = sheet.id === activeSheetId;
        return (
          <div
            key={sheet.id}
            className={`group flex items-center gap-1 rounded-t-md border border-b-0 px-3 py-1 text-sm ${
              active
                ? "border-[#b4b4b4] bg-white font-medium text-zinc-900"
                : "border-transparent bg-[#e6e6e6] text-zinc-600 hover:bg-white"
            }`}
          >
            <button type="button" onClick={() => onSelect(sheet.id)} className="max-w-32 truncate">
              {sheet.name}
            </button>
            {active ? (
              <div className="flex gap-0.5 opacity-0 transition group-hover:opacity-100">
                <TabAction title="Rename" onClick={() => {
                  const name = prompt("Sheet name", sheet.name);
                  if (name) onRename(sheet.id, name);
                }}>✎</TabAction>
                <TabAction title="Move left" onClick={() => onReorder(sheet.id, -1)}>◀</TabAction>
                <TabAction title="Move right" onClick={() => onReorder(sheet.id, 1)}>▶</TabAction>
                <TabAction title="Duplicate" onClick={onDuplicate}>⧉</TabAction>
                {sheets.length > 1 ? (
                  <TabAction title="Delete" onClick={() => onDelete(sheet.id)}>×</TabAction>
                ) : null}
              </div>
            ) : null}
          </div>
        );
      })}
      <button
        type="button"
        onClick={onAdd}
        className="rounded-md px-2 py-1 text-lg leading-none text-zinc-600 hover:bg-white"
        title="Add worksheet"
      >
        +
      </button>
    </div>
  );
}

function TabAction({
  title,
  onClick,
  children,
}: {
  title: string;
  onClick: () => void;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      title={title}
      onClick={onClick}
      className="rounded px-1 text-xs hover:bg-zinc-200"
    >
      {children}
    </button>
  );
}