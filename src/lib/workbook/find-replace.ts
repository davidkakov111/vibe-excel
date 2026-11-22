import type { Worksheet } from "@/lib/workbook/types";
import { getCellInput, setCellInput } from "@/lib/workbook/types";

export interface FindMatch {
  row: number;
  col: number;
}

export function findInSheet(
  sheet: Worksheet,
  query: string,
  matchCase: boolean,
): FindMatch[] {
  if (!query) return [];
  const matches: FindMatch[] = [];
  const needle = matchCase ? query : query.toLowerCase();
  for (const [key, data] of Object.entries(sheet.cells)) {
    const input = data.input ?? "";
    if (!input) continue;
    const hay = matchCase ? input : input.toLowerCase();
    if (hay.includes(needle)) {
      const [row, col] = key.split(":").map(Number);
      matches.push({ row, col });
    }
  }
  matches.sort((a, b) => a.row - b.row || a.col - b.col);
  return matches;
}

export function replaceInSheet(
  sheet: Worksheet,
  query: string,
  replacement: string,
  matchCase: boolean,
  replaceAll: boolean,
  startFrom?: FindMatch,
): { sheet: Worksheet; count: number; next?: FindMatch } {
  if (!query) return { sheet, count: 0 };
  const matches = findInSheet(sheet, query, matchCase);
  if (!matches.length) return { sheet, count: 0 };

  let nextMatches = matches;
  if (startFrom && !replaceAll) {
    nextMatches = matches.filter(
      (m) =>
        m.row > startFrom.row ||
        (m.row === startFrom.row && m.col >= startFrom.col),
    );
    if (!nextMatches.length) return { sheet, count: 0 };
  }

  const targets = replaceAll ? matches : nextMatches.slice(0, 1);
  let nextSheet = sheet;
  for (const { row, col } of targets) {
    const input = getCellInput(nextSheet, row, col);
    const regex = new RegExp(
      query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
      matchCase ? "g" : "gi",
    );
    const updated = input.replace(regex, replacement);
    nextSheet = setCellInput(nextSheet, row, col, updated);
  }

  const remaining = findInSheet(nextSheet, query, matchCase);
  return {
    sheet: nextSheet,
    count: targets.length,
    next: remaining[0],
  };
}