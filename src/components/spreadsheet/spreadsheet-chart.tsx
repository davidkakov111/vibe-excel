"use client";

import type { Worksheet } from "@/lib/workbook/types";

interface SpreadsheetChartProps {
  sheet: Worksheet;
  getDisplay: (sheetId: string, row: number, col: number) => string;
  sheetId: string;
}

export function SpreadsheetChartPanel({
  sheet,
}: SpreadsheetChartProps) {
  if (!sheet.charts.length) {
    return (
      <div className="mx-2 mb-1 rounded border border-dashed border-[#b4b4b4] bg-white px-3 py-2 text-xs text-zinc-600">
        Charts appear on the sheet in a reserved block below your data. Select
        labels + numbers, then pick Bar / Line / Pie / Area.
      </div>
    );
  }

  return (
    <div className="mx-2 mb-1 px-1 py-1 text-xs text-zinc-600">
      {sheet.charts.length} chart{sheet.charts.length === 1 ? "" : "s"} on this
      sheet (scroll to the chart area). Included in .xlsx download.
    </div>
  );
}