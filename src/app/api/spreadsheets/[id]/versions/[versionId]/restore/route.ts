import { NextResponse } from "next/server";
import { z } from "zod";
import { publishSpreadsheetEvent } from "@/lib/collaboration/hub";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { workbookFromDb } from "@/lib/spreadsheet-db";
import { canEdit, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { prisma } from "@/lib/prisma";
import { assertWorkbookSnapshot } from "@/lib/workbook/version-history";
import type { WorkbookDocument } from "@/lib/workbook/types";

type RouteContext = { params: Promise<{ id: string; versionId: string }> };

const bodySchema = z.object({
  baseRevision: z.number().int().min(0),
  clientId: z.string().min(1),
});

function activeSheetMetrics(workbook: WorkbookDocument) {
  const active =
    workbook.sheets.find((s) => s.id === workbook.activeSheetId) ??
    workbook.sheets[0]!;
  return { rowCount: active.rowCount, colCount: active.colCount };
}

export async function POST(request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id, versionId } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }
  if (!canEdit(access)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }
  const parsed = bodySchema.safeParse(body);
  if (!parsed.success) {
    return NextResponse.json(
      { error: parsed.error.issues[0]?.message ?? "Invalid input" },
      { status: 400 },
    );
  }

  const version = await prisma.spreadsheetVersion.findFirst({
    where: { id: versionId, spreadsheetId: id },
  });
  if (!version) {
    return NextResponse.json({ error: "Version not found" }, { status: 404 });
  }

  const current = access.spreadsheet;
  if (current.revision !== parsed.data.baseRevision) {
    return NextResponse.json(
      {
        error: "Revision conflict",
        revision: current.revision,
        workbook: workbookFromDb(
          current.cells,
          current.rowCount,
          current.colCount,
          current.name,
        ),
      },
      { status: 409 },
    );
  }

  const restored = assertWorkbookSnapshot(version.workbook);
  const { rowCount, colCount } = activeSheetMetrics(restored);
  const nextRevision = current.revision + 1;

  const user = await prisma.user.findUnique({
    where: { id: session.userId },
    select: { name: true, email: true },
  });
  const authorName = user?.name ?? user?.email ?? "User";

  const currentWorkbook = workbookFromDb(
    current.cells,
    current.rowCount,
    current.colCount,
    current.name,
  );

  await prisma.$transaction([
    prisma.spreadsheetVersion.create({
      data: {
        spreadsheetId: id,
        revision: current.revision,
        workbook: currentWorkbook as object,
        createdById: session.userId,
        authorName,
        label: "Before restore",
      },
    }),
    prisma.spreadsheet.updateMany({
      where: {
        id: current.id,
        revision: parsed.data.baseRevision,
        OR: [
          { userId: session.userId },
          {
            shares: {
              some: { userId: session.userId, permission: "EDIT" },
            },
          },
        ],
      },
      data: {
        cells: restored as object,
        revision: nextRevision,
        rowCount,
        colCount,
      },
    }),
    prisma.spreadsheetVersion.create({
      data: {
        spreadsheetId: id,
        revision: nextRevision,
        workbook: restored as object,
        createdById: session.userId,
        authorName,
        label: `Restored · ${version.label}`,
      },
    }),
  ]);

  const updated = await prisma.spreadsheet.findUnique({ where: { id: current.id } });
  if (!updated || updated.revision !== nextRevision) {
    const fresh = await prisma.spreadsheet.findUnique({ where: { id: current.id } });
    if (!fresh) {
      return NextResponse.json({ error: "Not found" }, { status: 404 });
    }
    return NextResponse.json(
      {
        error: "Revision conflict",
        revision: fresh.revision,
        workbook: workbookFromDb(
          fresh.cells,
          fresh.rowCount,
          fresh.colCount,
          fresh.name,
        ),
      },
      { status: 409 },
    );
  }

  publishSpreadsheetEvent(id, {
    type: "workbook_restore",
    revision: nextRevision,
    workbook: restored,
    clientId: parsed.data.clientId,
    userId: session.userId,
    versionId,
  });

  return NextResponse.json({ revision: nextRevision, workbook: restored });
}