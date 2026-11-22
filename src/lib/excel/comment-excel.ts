import type ExcelJS from "exceljs";
import {
  createCommentThread,
  normalizeCellComment,
  threadAuthorLabel,
  threadPreviewText,
  type CellCommentThread,
} from "@/lib/workbook/comment";

export function commentThreadToExcelNote(
  thread: CellCommentThread,
): ExcelJS.Comment {
  const author = threadAuthorLabel(thread);
  const lines = thread.messages.map(
    (m) => `${m.authorName} (${formatShortTime(m.createdAt)}):\n${m.text}`,
  );
  const body = lines.join("\n\n");
  const statusLine =
    thread.status === "resolved" ? "\n\n— Resolved —" : "";
  return {
    texts: [
      { text: `${author}:`, font: { bold: true, size: 9, name: "Calibri" } },
      {
        text: `\n${body}${statusLine}`,
        font: { size: 9, name: "Calibri" },
      },
    ],
  };
}

function formatShortTime(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  return d.toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

export function excelNoteToCommentThread(
  note: string | ExcelJS.Comment,
): CellCommentThread | null {
  let text = "";
  if (typeof note === "string") {
    text = note;
  } else if (note.texts?.length) {
    text = note.texts.map((t) => t.text ?? "").join("");
  }
  text = text.trim();
  if (!text) return null;

  const cleaned = text.replace(/\n+— Resolved —$/i, "").trim();
  return createCommentThread("import", "Excel import", cleaned);
}

export { normalizeCellComment, threadPreviewText };