import type { WorkbookDocument } from "@/lib/workbook/types";

/** Minimum ms between automatic version snapshots for the same spreadsheet. */
export const VERSION_SNAPSHOT_INTERVAL_MS = 2 * 60 * 1000;

/** Create a snapshot after this many collaborative revisions (in addition to time gate). */
export const VERSION_SNAPSHOT_EVERY_N_REVISIONS = 5;

export function shouldCaptureVersion(args: {
  lastVersionAt: Date | null;
  lastVersionRevision: number | null;
  currentRevision: number;
  now?: Date;
}): boolean {
  const now = args.now ?? new Date();
  if (args.lastVersionRevision === null) {
    return args.currentRevision > 0;
  }
  if (
    args.currentRevision - args.lastVersionRevision >=
    VERSION_SNAPSHOT_EVERY_N_REVISIONS
  ) {
    return true;
  }
  if (!args.lastVersionAt) return true;
  return now.getTime() - args.lastVersionAt.getTime() >= VERSION_SNAPSHOT_INTERVAL_MS;
}

export type VersionListItem = {
  id: string;
  revision: number;
  createdAt: string;
  authorName: string | null;
  label: string;
};

export function versionLabel(revision: number, createdAt: Date): string {
  return `Revision ${revision} · ${createdAt.toISOString()}`;
}

export function assertWorkbookSnapshot(workbook: unknown): WorkbookDocument {
  return workbook as WorkbookDocument;
}