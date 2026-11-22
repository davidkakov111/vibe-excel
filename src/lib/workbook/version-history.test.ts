import { describe, expect, it } from "vitest";
import { shouldCaptureVersion, VERSION_SNAPSHOT_INTERVAL_MS } from "@/lib/workbook/version-history";

describe("version-history", () => {
  it("captures first version after edits", () => {
    expect(
      shouldCaptureVersion({
        lastVersionAt: null,
        lastVersionRevision: null,
        currentRevision: 3,
      }),
    ).toBe(true);
  });

  it("throttles by time and revision count", () => {
    const now = new Date();
    const recent = new Date(now.getTime() - 60_000);
    expect(
      shouldCaptureVersion({
        lastVersionAt: recent,
        lastVersionRevision: 10,
        currentRevision: 12,
        now,
      }),
    ).toBe(false);

    expect(
      shouldCaptureVersion({
        lastVersionAt: recent,
        lastVersionRevision: 10,
        currentRevision: 30,
        now,
      }),
    ).toBe(true);

    expect(
      shouldCaptureVersion({
        lastVersionAt: new Date(now.getTime() - VERSION_SNAPSHOT_INTERVAL_MS - 1),
        lastVersionRevision: 10,
        currentRevision: 12,
        now,
      }),
    ).toBe(true);
  });
});