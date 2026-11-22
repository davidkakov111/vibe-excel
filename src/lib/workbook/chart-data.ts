import type { ChartSeries, Worksheet } from "@/lib/workbook/types";

export type ChartPoint = { name: string; value: number };

export function chartSeriesData(
  sheet: Worksheet,
  chart: ChartSeries,
  getDisplay: (sheetId: string, row: number, col: number) => string,
  sheetId: string,
): ChartPoint[] {
  const points: ChartPoint[] = [];
  const { startRow, endRow, startCol, endCol } = chart.range;
  const valueCol = endCol > startCol ? startCol + 1 : startCol;
  for (let row = startRow + 1; row <= endRow; row++) {
    const name = getDisplay(sheetId, row, startCol) || `Row ${row + 1}`;
    const raw = getDisplay(sheetId, row, valueCol);
    const value = Number(raw);
    if (!Number.isNaN(value)) {
      points.push({ name, value });
    }
  }
  return points;
}

export const CHART_ANCHOR_ROWS = 12;
export const CHART_ANCHOR_COLS = 6;

export function defaultChartAnchor(
  chart: Pick<ChartSeries, "range">,
): ChartSeries["anchor"] {
  const startRow = chart.range.endRow + 2;
  const startCol = chart.range.startCol;
  return {
    startRow,
    startCol,
    endRow: startRow + CHART_ANCHOR_ROWS - 1,
    endCol: startCol + CHART_ANCHOR_COLS - 1,
  };
}

export function chartWithDefaultAnchor(chart: ChartSeries): ChartSeries {
  if (chart.anchor) return chart;
  return { ...chart, anchor: defaultChartAnchor(chart) };
}