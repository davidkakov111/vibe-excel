import type { GridSelection } from "@/lib/workbook/selection";

export type RemotePresence = {
  clientId: string;
  userId: string;
  userName: string;
  color: string;
  sheetId: string;
  selection: GridSelection | null;
};

export const PRESENCE_STALE_MS = 2_500;

export function userInitial(userName: string): string {
  const trimmed = userName.trim();
  if (!trimmed) return "?";
  const parts = trimmed.split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    return `${parts[0]![0] ?? ""}${parts[1]![0] ?? ""}`.toUpperCase();
  }
  return trimmed.slice(0, 1).toUpperCase();
}

export function normalizeSelectionBounds(sel: GridSelection): GridSelection {
  return {
    row: Math.min(sel.row, sel.endRow),
    col: Math.min(sel.col, sel.endCol),
    endRow: Math.max(sel.row, sel.endRow),
    endCol: Math.max(sel.col, sel.endCol),
  };
}

export function cellInRemoteSelection(
  row: number,
  col: number,
  sheetId: string,
  presence: RemotePresence,
): boolean {
  if (presence.sheetId !== sheetId || !presence.selection) return false;
  const s = normalizeSelectionBounds(presence.selection);
  return row >= s.row && row <= s.endRow && col >= s.col && col <= s.endCol;
}

/** Latest presence per user (drops duplicate tabs / stale reconnects in UI). */
export function uniquePresenceByUser(
  others: RemotePresence[],
  lastSeenMs: Map<string, number>,
  now = Date.now(),
): RemotePresence[] {
  const fresh = others.filter((p) => {
    const seen = lastSeenMs.get(p.clientId) ?? 0;
    return now - seen <= PRESENCE_STALE_MS;
  });
  const byUser = new Map<string, RemotePresence>();
  for (const p of fresh) {
    const existing = byUser.get(p.userId);
    if (!existing) {
      byUser.set(p.userId, p);
      continue;
    }
    const existingSeen = lastSeenMs.get(existing.clientId) ?? 0;
    const nextSeen = lastSeenMs.get(p.clientId) ?? 0;
    if (nextSeen >= existingSeen) byUser.set(p.userId, p);
  }
  return [...byUser.values()];
}

export function remoteHighlightsAtCell(
  row: number,
  col: number,
  sheetId: string,
  others: RemotePresence[],
): { userId: string; color: string; initial: string }[] {
  return others
    .filter((p) => cellInRemoteSelection(row, col, sheetId, p))
    .map((p) => ({
      userId: p.userId,
      color: p.color,
      initial: userInitial(p.userName),
    }));
}