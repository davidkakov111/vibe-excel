import { describe, expect, it } from "vitest";
import {
  cellInRemoteSelection,
  remoteHighlightsAtCell,
  uniquePresenceByUser,
  userInitial,
  type RemotePresence,
} from "@/lib/collaboration/presence";

describe("presence", () => {
  it("builds initials from name", () => {
    expect(userInitial("Ada Lovelace")).toBe("AL");
    expect(userInitial("Bob")).toBe("B");
  });

  it("prunes stale and dedupes by user", () => {
    const now = 10_000;
    const lastSeen = new Map([
      ["a", now],
      ["b", now - 5_000],
      ["c", now],
    ]);
    const others: RemotePresence[] = [
      {
        clientId: "a",
        userId: "u1",
        userName: "Ada",
        color: "#f00",
        sheetId: "s1",
        selection: { row: 0, col: 0, endRow: 0, endCol: 0 },
      },
      {
        clientId: "c",
        userId: "u1",
        userName: "Ada",
        color: "#00f",
        sheetId: "s1",
        selection: { row: 1, col: 1, endRow: 1, endCol: 1 },
      },
      {
        clientId: "b",
        userId: "u2",
        userName: "Bob",
        color: "#0f0",
        sheetId: "s1",
        selection: null,
      },
    ];
    const unique = uniquePresenceByUser(others, lastSeen, now);
    expect(unique).toHaveLength(1);
    expect(unique[0]!.clientId).toBe("c");
  });

  it("marks cells in remote selection", () => {
    const p: RemotePresence = {
      clientId: "x",
      userId: "u1",
      userName: "Ada Lovelace",
      color: "#2563eb",
      sheetId: "sheet-a",
      selection: { row: 0, col: 0, endRow: 1, endCol: 1 },
    };
    expect(cellInRemoteSelection(0, 1, "sheet-a", p)).toBe(true);
    expect(cellInRemoteSelection(2, 0, "sheet-a", p)).toBe(false);
    const highlights = remoteHighlightsAtCell(0, 0, "sheet-a", [p]);
    expect(highlights[0]?.initial).toBe("AL");
    expect(highlights[0]?.color).toBe("#2563eb");
  });
});