import { prisma } from "@/lib/prisma";
import { shouldCaptureVersion } from "@/lib/workbook/version-history";
import type { WorkbookDocument } from "@/lib/workbook/types";

export async function createSpreadsheetVersion(args: {
  spreadsheetId: string;
  revision: number;
  workbook: WorkbookDocument;
  userId: string;
  authorName?: string | null;
  label?: string;
  /** When true, always write a row (user-chosen snapshot). */
  force?: boolean;
}): Promise<void> {
  const latest = await prisma.spreadsheetVersion.findFirst({
    where: { spreadsheetId: args.spreadsheetId },
    orderBy: { createdAt: "desc" },
    select: { createdAt: true, revision: true },
  });

  if (
    !args.force &&
    !shouldCaptureVersion({
      lastVersionAt: latest?.createdAt ?? null,
      lastVersionRevision: latest?.revision ?? null,
      currentRevision: args.revision,
    })
  ) {
    return;
  }

  await prisma.spreadsheetVersion.create({
    data: {
      spreadsheetId: args.spreadsheetId,
      revision: args.revision,
      workbook: args.workbook as object,
      createdById: args.userId,
      authorName: args.authorName ?? null,
      label: args.label ?? "Auto-save",
    },
  });
}

/** Best-effort automatic snapshot after collaborative saves (throttled). */
export async function maybeCreateSpreadsheetVersion(
  args: Omit<Parameters<typeof createSpreadsheetVersion>[0], "force">,
): Promise<void> {
  await createSpreadsheetVersion(args);
}

export async function listSpreadsheetVersions(spreadsheetId: string, limit = 40) {
  return prisma.spreadsheetVersion.findMany({
    where: { spreadsheetId },
    orderBy: { createdAt: "desc" },
    take: limit,
    select: {
      id: true,
      revision: true,
      createdAt: true,
      authorName: true,
      label: true,
      createdById: true,
    },
  });
}