"use client";

import type { ChartSeries, Worksheet } from "@/lib/workbook/types";
import { chartSeriesData } from "@/lib/workbook/chart-data";
import { ChartView } from "@/components/spreadsheet/chart-view";

export function SheetChartOverlays({
  sheet,
  sheetId,
  getDisplay,
  charts,
  colWidth,
  rowHeight,
  colOffset,
  rowTop,
  leftGutter,
}: {
  sheet: Worksheet;
  sheetId: string;
  getDisplay: (sheetId: string, row: number, col: number) => string;
  charts: ChartSeries[];
  colWidth: (col: number) => number;
  rowHeight: (row: number) => number;
  colOffset: (col: number) => number;
  rowTop: (row: number) => number;
  leftGutter: number;
}) {
  if (!charts.length) return null;

  return (
    <>
      {charts.map((chart) => {
        const data = chartSeriesData(sheet, chart, getDisplay, sheetId);
        if (!data.length) return null;
        const { anchor } = chart;
        let width = 0;
        for (let c = anchor.startCol; c <= anchor.endCol; c++) width += colWidth(c);
        let height = 0;
        for (let r = anchor.startRow; r <= anchor.endRow; r++) height += rowHeight(r);

        const top = rowTop(anchor.startRow);
        const left = leftGutter + colOffset(anchor.startCol);

        return (
          <div
            key={chart.id}
            className="pointer-events-none absolute z-[6] overflow-hidden rounded-sm border-2 border-[#217346] bg-white shadow-md"
            style={{
              top,
              left,
              width,
              height,
            }}
            title={chart.title}
          >
            <ChartView chart={chart} data={data} title={chart.title} />
          </div>
        );
      })}
    </>
  );
}