import type { CellData, CellFormat } from "@/lib/workbook/types";

export const CLIPBOARD_MIME = "application/vnd.ai-test.cells+json";

export interface ClipboardCell {
  input?: string;
  format?: CellFormat;
}

export interface ClipboardPayload {
  version: 1;
  cells: ClipboardCell[][];
}

export function serializeClipboardPayload(payload: ClipboardPayload): string {
  return JSON.stringify(payload);
}

export function parseClipboardPayload(raw: string): ClipboardPayload | null {
  try {
    const parsed = JSON.parse(raw) as ClipboardPayload;
    if (parsed?.version !== 1 || !Array.isArray(parsed.cells)) return null;
    return parsed;
  } catch {
    return null;
  }
}

export function plainTextFromPayload(payload: ClipboardPayload): string {
  return payload.cells
    .map((row) =>
      row.map((cell) => cell.input ?? "").join("\t"),
    )
    .join("\n");
}

export function payloadFromPlainTsv(text: string): ClipboardPayload {
  const lines = text.split(/\r?\n/).filter((line, i, arr) => !(i === arr.length - 1 && line === ""));
  return {
    version: 1,
    cells: lines.map((line) =>
      line.split("\t").map((input) => ({ input })),
    ),
  };
}

export function mergeCellPaste(
  existing: CellData | undefined,
  patch: ClipboardCell,
): CellData | undefined {
  const hasInput = patch.input !== undefined;
  const input = hasInput ? patch.input : existing?.input;
  const format =
    patch.format || existing?.format
      ? { ...existing?.format, ...patch.format }
      : undefined;
  const trimmed = (input ?? "").trim();
  if (
    !trimmed &&
    !format &&
    !existing?.comment &&
    !existing?.validation
  ) {
    return undefined;
  }
  const next: CellData = { ...existing };
  if (hasInput) next.input = input;
  if (format) next.format = format;
  return next;
}