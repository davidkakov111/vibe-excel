import type ExcelJS from "exceljs";
import {
  DEFAULT_OFFICE_THEME_HEX,
  hexFromExcelColor,
  type ExcelColorLike,
} from "@/lib/excel/theme-colors";
import type { CellBorder, CellBorders, CellFormat, BorderStyle } from "@/lib/workbook/types";

export function argbFromHex(hex?: string): string | undefined {
  if (!hex) return undefined;
  const normalized = hex.replace("#", "");
  if (normalized.length === 6) return `FF${normalized.toUpperCase()}`;
  return undefined;
}

export function hexFromArgb(argb?: string): string | undefined {
  if (!argb || argb.length < 6) return undefined;
  const slice = argb.length === 8 ? argb.slice(2) : argb;
  return `#${slice.toUpperCase()}`;
}

function mapExcelBorderStyle(style?: string): BorderStyle | undefined {
  if (!style) return undefined;
  if (style === "thick") return "thick";
  if (style === "medium" || style === "double") return "medium";
  if (style.includes("dash")) return "dashed";
  return "thin";
}

export function readBorderSideFromExcel(
  side?: Partial<ExcelJS.Border>,
  palette: readonly string[] = DEFAULT_OFFICE_THEME_HEX,
): CellBorder | undefined {
  if (!side?.style) return undefined;
  return {
    style: mapExcelBorderStyle(side.style),
    color: hexFromExcelColor(side.color as ExcelColorLike, palette),
  };
}

export function readFormat(
  cell: ExcelJS.Cell,
  palette: readonly string[] = DEFAULT_OFFICE_THEME_HEX,
): CellFormat | undefined {
  const format: CellFormat = {};
  if (cell.font) {
    format.fontFamily = cell.font.name;
    format.fontSize = cell.font.size;
    format.bold = cell.font.bold;
    format.italic = cell.font.italic;
    format.underline = Boolean(cell.font.underline);
    format.strikethrough = cell.font.strike;
    format.fontColor = hexFromExcelColor(cell.font.color as ExcelColorLike, palette);
  }
  const fill = cell.fill;
  if (fill && fill.type === "pattern") {
    const patternFill = fill as ExcelJS.FillPattern;
    const color =
      hexFromExcelColor(patternFill.fgColor as ExcelColorLike, palette) ??
      hexFromExcelColor(patternFill.bgColor as ExcelColorLike, palette);
    if (color) format.fillColor = color;
  }
  if (cell.alignment) {
    format.horizontalAlign = cell.alignment.horizontal as CellFormat["horizontalAlign"];
    format.verticalAlign = cell.alignment.vertical as CellFormat["verticalAlign"];
    format.wrap = cell.alignment.wrapText;
    const rot = cell.alignment.textRotation;
    format.rotation = typeof rot === "number" ? rot : undefined;
  }
  const border = cell.border;
  if (border) {
    const borders: CellBorders = {
      top: readBorderSideFromExcel(border.top, palette),
      right: readBorderSideFromExcel(border.right, palette),
      bottom: readBorderSideFromExcel(border.bottom, palette),
      left: readBorderSideFromExcel(border.left, palette),
    };
    if (Object.values(borders).some(Boolean)) {
      format.borders = borders;
    }
  }
  return Object.keys(format).length ? format : undefined;
}

function excelBorderFromFormat(side?: CellBorder): Partial<ExcelJS.Border> | undefined {
  if (!side?.style) return undefined;
  const style =
    side.style === "thick"
      ? "thick"
      : side.style === "medium"
        ? "medium"
        : side.style === "dashed"
          ? "dashDot"
          : "thin";
  return {
    style,
    color: side.color ? { argb: argbFromHex(side.color) } : undefined,
  };
}

export function applyExcelStyle(cell: ExcelJS.Cell, format?: CellFormat): void {
  if (!format) return;
  cell.font = {
    name: format.fontFamily,
    size: format.fontSize,
    bold: format.bold,
    italic: format.italic,
    underline: format.underline,
    strike: format.strikethrough,
    color: format.fontColor ? { argb: argbFromHex(format.fontColor) } : undefined,
  };
  if (format.fillColor) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: argbFromHex(format.fillColor) },
    };
  }
  cell.alignment = {
    horizontal: format.horizontalAlign,
    vertical: format.verticalAlign,
    wrapText: format.wrap,
    textRotation: format.rotation,
  };
  if (format.numberFormat && format.numberFormat !== "general") {
    const decimals = format.numberDecimals ?? 2;
    const zeros = "0".repeat(decimals);
    const decimalPart = decimals > 0 ? `.${zeros}` : "";
    switch (format.numberFormat) {
      case "number":
        cell.numFmt = `#,##0${decimalPart}`;
        break;
      case "currency":
        cell.numFmt = `$#,##0${decimalPart}`;
        break;
      case "percentage":
        cell.numFmt = `0${decimalPart}%`;
        break;
      case "text":
        cell.numFmt = "@";
        break;
      default:
        break;
    }
  }
  if (format.borders) {
    cell.border = {
      top: excelBorderFromFormat(format.borders.top),
      right: excelBorderFromFormat(format.borders.right),
      bottom: excelBorderFromFormat(format.borders.bottom),
      left: excelBorderFromFormat(format.borders.left),
    };
  }
}