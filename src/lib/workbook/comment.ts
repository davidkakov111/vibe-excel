import { nanoid } from "nanoid";

/** One message in a cell comment thread. */
export interface CellCommentMessage {
  id: string;
  userId: string;
  authorName: string;
  text: string;
  createdAt: string;
}

/** Threaded comment (open or resolved). */
export interface CellCommentThread {
  id: string;
  status: "open" | "resolved";
  messages: CellCommentMessage[];
  createdAt: string;
  updatedAt: string;
}

/** @deprecated legacy single-bubble shape still stored in older JSON */
export interface LegacyCellComment {
  id: string;
  userId: string;
  authorName: string;
  text: string;
  createdAt: string;
  updatedAt: string;
}

export type StoredCellComment = CellCommentThread | LegacyCellComment;

export function isCommentThread(
  value: StoredCellComment,
): value is CellCommentThread {
  return Array.isArray((value as CellCommentThread).messages);
}

export function normalizeCellComment(
  raw: StoredCellComment | null | undefined,
): CellCommentThread | null {
  if (!raw) return null;
  if (isCommentThread(raw)) {
    return {
      ...raw,
      messages: raw.messages.map((m) => ({ ...m })),
    };
  }
  const legacy = raw;
  return {
    id: legacy.id,
    status: "open",
    createdAt: legacy.createdAt,
    updatedAt: legacy.updatedAt,
    messages: [
      {
        id: legacy.id,
        userId: legacy.userId,
        authorName: legacy.authorName,
        text: legacy.text,
        createdAt: legacy.createdAt,
      },
    ],
  };
}

export function threadPreviewText(thread: CellCommentThread): string {
  const last = thread.messages[thread.messages.length - 1];
  return last?.text ?? "";
}

export function threadAuthorLabel(thread: CellCommentThread): string {
  const first = thread.messages[0];
  return first?.authorName ?? "Comment";
}

export function createCommentThread(
  userId: string,
  authorName: string,
  text: string,
): CellCommentThread {
  const now = new Date().toISOString();
  const messageId = nanoid();
  return {
    id: nanoid(),
    status: "open",
    createdAt: now,
    updatedAt: now,
    messages: [
      {
        id: messageId,
        userId,
        authorName,
        text: text.trim(),
        createdAt: now,
      },
    ],
  };
}

export function addCommentReply(
  thread: CellCommentThread,
  userId: string,
  authorName: string,
  text: string,
): CellCommentThread {
  const trimmed = text.trim();
  if (!trimmed) return thread;
  const now = new Date().toISOString();
  return {
    ...thread,
    status: "open",
    updatedAt: now,
    messages: [
      ...thread.messages,
      {
        id: nanoid(),
        userId,
        authorName,
        text: trimmed,
        createdAt: now,
      },
    ],
  };
}

export function setCommentResolved(
  thread: CellCommentThread,
  resolved: boolean,
): CellCommentThread {
  return {
    ...thread,
    status: resolved ? "resolved" : "open",
    updatedAt: new Date().toISOString(),
  };
}