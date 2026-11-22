import type ExcelJS from "exceljs";
import { cellInRange, parseA1Range } from "@/lib/excel/a1-range";
import { hexFromExcelColor } from "@/lib/excel/theme-colors";
import { DEFAULT_OFFICE_THEME_HEX } from "@/lib/excel/theme-colors";
import type { CellFormat, ConditionalFormattingSet, Worksheet } from "@/lib/workbook/types";
import { getCellInput } from "@/lib/workbook/types";
import { readBorderSideFromExcel } from "@/lib/excel/style-map";

interface CfRule {
  type?: string;
  operator?: string;
  priority?: number;
  formulae?: (string | number)[];
  text?: string;
  stopIfTrue?: boolean;
  style?: Record<string, unknown>;
}

function numericValue(input: string | undefined, display: string): number | null {
  const raw = (input ?? display).trim();
  if (!raw) return null;
  const num = Number(raw);
  if (!Number.isNaN(num) && /^-?\d+(\.\d+)?$/.test(raw)) return num;
  const fromDisplay = Number(display);
  return Number.isNaN(fromDisplay) ? null : fromDisplay;
}

function countValueInRange(
  sheet: Worksheet,
  range: { startRow: number; endRow: number; startCol: number; endCol: number },
  needle: string,
): number {
  let count = 0;
  for (let row = range.startRow; row <= range.endRow; row++) {
    for (let col = range.startCol; col <= range.endCol; col++) {
      const value = getCellInput(sheet, row, col).trim();
      if (value === needle) count++;
    }
  }
  return count;
}

function ruleMatches(
  rule: CfRule,
  input: string | undefined,
  display: string,
  sheet?: Worksheet,
  range?: { startRow: number; endRow: number; startCol: number; endCol: number },
): boolean {
  const type = rule.type;
  const text = (input ?? display).toString();
  const num = numericValue(input, display);

  if (type === "containsText") {
    const needle = String(rule.text ?? rule.formulae?.[0] ?? "");
    const op = rule.operator;
    if (op === "beginsWith") return text.startsWith(needle);
    if (op === "endsWith") return text.endsWith(needle);
    return text.includes(needle);
  }

  if (type === "cellIs") {
    const op = rule.operator;
    const formulae = rule.formulae;
    const a = formulae?.[0];
    const b = formulae?.[1];
    const cmp = typeof a === "number" ? a : Number(String(a).replace(/^=/, ""));
    if (Number.isNaN(cmp) || num === null) {
      const str = String(input ?? display);
      const target = String(a ?? "").replace(/^=/, "").replace(/^"|"$/g, "");
      switch (op) {
        case "equal":
          return str === target;
        case "notEqual":
          return str !== target;
        default:
          return false;
      }
    }
    const cmp2 = typeof b === "number" ? b : Number(String(b ?? "").replace(/^=/, ""));
    switch (op) {
      case "equal":
        return num === cmp;
      case "notEqual":
        return num !== cmp;
      case "greaterThan":
        return num > cmp;
      case "lessThan":
        return num < cmp;
      case "greaterThanOrEqual":
        return num >= cmp;
      case "lessThanOrEqual":
        return num <= cmp;
      case "between":
        return num >= cmp && num <= cmp2;
      case "notBetween":
        return num < cmp || num > cmp2;
      default:
        return false;
    }
  }

  if (type === "expression") {
    const formulae = rule.formulae;
    const expr = String(formulae?.[0] ?? "").replace(/^=/, "").trim();
    if (expr === "TRUE()" || expr === "TRUE") return true;
    if (expr === "FALSE()" || expr === "FALSE") return false;
    return false;
  }

  if (type === "duplicateValues" && sheet && range) {
    const needle = (input ?? display).trim();
    if (!needle) return false;
    return countValueInRange(sheet, range, needle) > 1;
  }

  return false;
}

export function formatFromExcelStyle(
  style: Record<string, unknown> | undefined,
  palette: readonly string[] = DEFAULT_OFFICE_THEME_HEX,
): CellFormat | undefined {
  if (!style) return undefined;
  const format: CellFormat = {};
  const font = style.font as ExcelJS.Font | undefined;
  if (font) {
    format.fontFamily = font.name;
    format.fontSize = font.size;
    format.bold = font.bold;
    format.italic = font.italic;
    format.underline = Boolean(font.underline);
    format.strikethrough = font.strike;
    format.fontColor = hexFromExcelColor(font.color as { argb?: string; theme?: number; tint?: number }, palette);
  }
  const fill = style.fill as ExcelJS.Fill | undefined;
  if (fill && fill.type === "pattern") {
    const patternFill = fill as ExcelJS.FillPattern;
    const color =
      hexFromExcelColor(patternFill.fgColor as { argb?: string; theme?: number; tint?: number }, palette) ??
      hexFromExcelColor(patternFill.bgColor as { argb?: string; theme?: number; tint?: number }, palette);
    if (color) format.fillColor = color;
  }
  const alignment = style.alignment as ExcelJS.Alignment | undefined;
  if (alignment) {
    format.horizontalAlign = alignment.horizontal as CellFormat["horizontalAlign"];
    format.verticalAlign = alignment.vertical as CellFormat["verticalAlign"];
    format.wrap = alignment.wrapText;
  }
  const border = style.border as ExcelJS.Borders | undefined;
  if (border) {
    const borders = {
      top: readBorderSideFromExcel(border.top),
      right: readBorderSideFromExcel(border.right),
      bottom: readBorderSideFromExcel(border.bottom),
      left: readBorderSideFromExcel(border.left),
    };
    if (Object.values(borders).some(Boolean)) format.borders = borders;
  }
  return Object.keys(format).length ? format : undefined;
}

function mergeFormats(base: CellFormat | undefined, overlay: CellFormat | undefined): CellFormat | undefined {
  if (!overlay) return base;
  if (!base) return overlay;
  return {
    ...base,
    ...overlay,
    borders: overlay.borders ?? base.borders,
  };
}

export function effectiveCellFormat(
  sheet: Worksheet,
  row: number,
  col: number,
  base: CellFormat | undefined,
  ctx: { input?: string; display: string },
  palette: readonly string[] = DEFAULT_OFFICE_THEME_HEX,
): CellFormat | undefined {
  const sets = sheet.conditionalFormats ?? [];
  if (!sets.length) return base;

  let applied: CellFormat | undefined = base;
  for (const set of sets) {
    const range = parseA1Range(set.ref);
    if (!range || !cellInRange(row, col, range)) continue;
    const rules = [...set.rules].sort(
      (a, b) => (a.priority as number) - (b.priority as number),
    );
    for (const rule of rules) {
      const typed = rule as CfRule;
      if (!ruleMatches(typed, ctx.input, ctx.display, sheet, range)) continue;
      const overlay = formatFromExcelStyle(typed.style, palette);
      applied = mergeFormats(applied, overlay);
      if (typed.stopIfTrue) break;
    }
  }
  return applied;
}

export function normalizeConditionalFormats(
  raw: ConditionalFormattingSet[] | undefined,
): ConditionalFormattingSet[] {
  return raw ?? [];
}