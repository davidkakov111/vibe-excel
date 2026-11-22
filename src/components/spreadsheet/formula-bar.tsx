"use client";

import { addressToA1 } from "@/lib/formula/engine";

interface FormulaBarProps {
  sheetId: string;
  selection: { row: number; col: number; endRow: number; endCol: number };
  rawValue: string;
  onChange: (value: string) => void;
  onCommit: () => void;
  onNavigateToAddress: (address: string) => boolean;
}

export function FormulaBar({
  selection,
  rawValue,
  onChange,
  onCommit,
  onNavigateToAddress,
}: FormulaBarProps) {
  const label =
    selection.row === selection.endRow && selection.col === selection.endCol
      ? addressToA1(selection.row, selection.col)
      : `${addressToA1(selection.row, selection.col)}:${addressToA1(selection.endRow, selection.endCol)}`;

  return (
    <div className="flex items-center gap-2 border-b border-[#b4b4b4] bg-white px-2 py-1">
      <input
        className="w-28 rounded border border-[#b4b4b4] bg-[#f3f3f3] px-2 py-1 text-xs font-mono text-zinc-900"
        defaultValue={label}
        key={label}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            const ok = onNavigateToAddress(event.currentTarget.value);
            if (!ok) event.preventDefault();
          }
        }}
        aria-label="Name box"
      />
      <span className="text-sm italic text-zinc-500">fx</span>
      <input
        className="min-w-0 flex-1 rounded border border-[#b4b4b4] bg-white px-2 py-1 text-sm font-mono text-zinc-900"
        value={rawValue}
        onChange={(e) => onChange(e.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter") {
            event.preventDefault();
            onCommit();
          }
        }}
        aria-label="Formula bar"
      />
    </div>
  );
}