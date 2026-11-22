import type { WorkbookDocument } from "@/lib/workbook/types";

export type CollaborationEvent =
  | {
      type: "ops";
      revision: number;
      ops: unknown[];
      clientId: string;
      userId: string;
    }
  | {
      type: "workbook_restore";
      revision: number;
      workbook: WorkbookDocument;
      clientId: string;
      userId: string;
      versionId?: string;
    }
  | {
      type: "presence";
      clientId: string;
      userId: string;
      userName: string;
      color: string;
      sheetId: string;
      selection: {
        row: number;
        col: number;
        endRow: number;
        endCol: number;
      } | null;
    }
  | { type: "presence_leave"; clientId: string }
  | { type: "revision"; revision: number };

type Listener = (event: CollaborationEvent) => void;

const globalStore = globalThis as unknown as {
  __collaborationRooms?: Map<string, Set<Listener>>;
};

function rooms(): Map<string, Set<Listener>> {
  if (!globalStore.__collaborationRooms) {
    globalStore.__collaborationRooms = new Map();
  }
  return globalStore.__collaborationRooms;
}

export function subscribeSpreadsheet(
  spreadsheetId: string,
  listener: Listener,
): () => void {
  const map = rooms();
  let set = map.get(spreadsheetId);
  if (!set) {
    set = new Set();
    map.set(spreadsheetId, set);
  }
  set.add(listener);
  return () => {
    set!.delete(listener);
    if (set!.size === 0) map.delete(spreadsheetId);
  };
}

export function publishSpreadsheetEvent(
  spreadsheetId: string,
  event: CollaborationEvent,
): void {
  const set = rooms().get(spreadsheetId);
  if (!set) return;
  for (const listener of set) {
    try {
      listener(event);
    } catch {
      // ignore broken SSE consumer
    }
  }
}