import type { CSSProperties } from "react";
import type { CellFormat } from "@/lib/workbook/types";

export function formatDisplayValue(
  rawDisplay: string,
  format?: CellFormat,
  rawInput?: string,
): string {
  if (!format?.numberFormat || format.numberFormat === "general") {
    return rawDisplay;
  }
  const num = Number(rawDisplay);
  if (Number.isNaN(num)) return rawDisplay;

  const decimals = format.numberDecimals ?? 2;
  switch (format.numberFormat) {
    case "number":
      return num.toLocaleString(undefined, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
      });
    case "currency":
      return num.toLocaleString(undefined, {
        style: "currency",
        currency: "USD",
        minimumFractionDigits: decimals,
      });
    case "percentage":
      return `${(num * 100).toFixed(decimals)}%`;
    case "date":
      return rawInput ?? rawDisplay;
    default:
      return rawDisplay;
  }
}

function borderCss(side?: { style?: string; color?: string }): string | undefined {
  if (!side) return undefined;
  const width =
    side.style === "thick" ? "3px" : side.style === "medium" ? "2px" : "1px";
  const style = side.style === "dashed" ? "dashed" : "solid";
  return `${width} ${style} ${side.color ?? "#334155"}`;
}

export function cellStyleFromFormat(format?: CellFormat): CSSProperties {
  if (!format) return {};
  const borders = format.borders;
  return {
    fontFamily: format.fontFamily,
    fontSize: format.fontSize ? `${format.fontSize}pt` : undefined,
    fontWeight: format.bold ? "bold" : undefined,
    fontStyle: format.italic ? "italic" : undefined,
    textDecoration:
      [format.underline ? "underline" : "", format.strikethrough ? "line-through" : ""]
        .filter(Boolean)
        .join(" ") || undefined,
    color: format.fontColor,
    backgroundColor: format.fillColor,
    textAlign: format.horizontalAlign,
    verticalAlign: format.verticalAlign,
    whiteSpace: format.wrap ? "normal" : "nowrap",
    borderTop: borderCss(borders?.top),
    borderRight: borderCss(borders?.right),
    borderBottom: borderCss(borders?.bottom),
    borderLeft: borderCss(borders?.left),
  };
}