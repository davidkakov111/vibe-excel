import type { GridSelection } from "@/lib/workbook/selection";
import { indexToColumnLabel } from "@/lib/spreadsheet-types";
import type { ConditionalFormattingSet } from "@/lib/workbook/types";

export type CfPreset =
  | "greaterThan"
  | "lessThan"
  | "equal"
  | "textContains"
  | "duplicate"
  | "formula";

export function selectionToA1(sel: GridSelection): string {
  const startRow = Math.min(sel.row, sel.endRow);
  const endRow = Math.max(sel.row, sel.endRow);
  const startCol = Math.min(sel.col, sel.endCol);
  const endCol = Math.max(sel.col, sel.endCol);
  const start = `${indexToColumnLabel(startCol)}${startRow + 1}`;
  const end = `${indexToColumnLabel(endCol)}${endRow + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function argbHex(color: string): { argb: string } {
  const hex = color.startsWith("#") ? color.slice(1) : color;
  return { argb: `FF${hex.padStart(6, "0").slice(-6)}` };
}

export function buildConditionalRule(args: {
  preset: CfPreset;
  value?: string;
  formula?: string;
  fillColor?: string;
  fontColor?: string;
  bold?: boolean;
  priority: number;
}): Record<string, unknown> {
  const style: Record<string, unknown> = {
    font: {
      ...(args.bold ? { bold: true } : {}),
      ...(args.fontColor ? { color: argbHex(args.fontColor) } : {}),
    },
    fill: args.fillColor
      ? {
          type: "pattern",
          pattern: "solid",
          fgColor: argbHex(args.fillColor),
        }
      : undefined,
  };

  switch (args.preset) {
    case "greaterThan":
      return {
        type: "cellIs",
        operator: "greaterThan",
        formulae: [Number(args.value ?? 0)],
        priority: args.priority,
        style,
      };
    case "lessThan":
      return {
        type: "cellIs",
        operator: "lessThan",
        formulae: [Number(args.value ?? 0)],
        priority: args.priority,
        style,
      };
    case "equal":
      return {
        type: "cellIs",
        operator: "equal",
        formulae: [args.value ?? ""],
        priority: args.priority,
        style,
      };
    case "textContains":
      return {
        type: "containsText",
        operator: "containsText",
        text: args.value ?? "",
        priority: args.priority,
        style,
      };
    case "duplicate":
      return {
        type: "duplicateValues",
        priority: args.priority,
        style,
      };
    case "formula":
      return {
        type: "expression",
        formulae: [args.formula?.startsWith("=") ? args.formula : `=${args.formula ?? "TRUE"}`],
        priority: args.priority,
        style,
      };
    default:
      return { type: "expression", formulae: ["=TRUE"], priority: args.priority, style };
  }
}

export function upsertConditionalFormatSet(
  sets: ConditionalFormattingSet[],
  ref: string,
  rules: Array<Record<string, unknown>>,
): ConditionalFormattingSet[] {
  const next = [...(sets ?? [])];
  const index = next.findIndex((s) => s.ref === ref);
  const block: ConditionalFormattingSet = { ref, rules };
  if (index >= 0) next[index] = block;
  else next.push(block);
  return next;
}

export function removeConditionalFormatSet(
  sets: ConditionalFormattingSet[],
  ref: string,
): ConditionalFormattingSet[] {
  return (sets ?? []).filter((s) => s.ref !== ref);
}

export function reorderConditionalFormatRules(
  sets: ConditionalFormattingSet[],
  ref: string,
  ruleIndex: number,
  direction: "up" | "down",
): ConditionalFormattingSet[] {
  return (sets ?? []).map((set) => {
    if (set.ref !== ref) return set;
    const rules = [...set.rules];
    const target = direction === "up" ? ruleIndex - 1 : ruleIndex + 1;
    if (target < 0 || target >= rules.length) return set;
    const tmp = rules[ruleIndex]!;
    rules[ruleIndex] = rules[target]!;
    rules[target] = tmp;
    return {
      ...set,
      rules: rules.map((r, i) => ({ ...r, priority: i + 1 })),
    };
  });
}