import { NextResponse } from "next/server";
import { z } from "zod";
import { publishSpreadsheetEvent } from "@/lib/collaboration/hub";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { workbookFromDb } from "@/lib/spreadsheet-db";
import { applyWorkbookOps, type WorkbookOp } from "@/lib/workbook/ops";
import { canEdit, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { prisma } from "@/lib/prisma";
import { maybeCreateSpreadsheetVersion } from "@/lib/spreadsheet-versions";
import type { WorkbookDocument } from "@/lib/workbook/types";

type RouteContext = { params: Promise<{ id: string }> };

const opSchema = z.discriminatedUnion("type", [
  z.object({
    type: z.literal("setCell"),
    sheetId: z.string(),
    row: z.number().int().min(0),
    col: z.number().int().min(0),
    input: z.string().nullable(),
  }),
  z.object({
    type: z.literal("setFormat"),
    sheetId: z.string(),
    row: z.number().int().min(0),
    col: z.number().int().min(0),
    format: z.record(z.string(), z.unknown()).nullable(),
  }),
  z.object({
    type: z.literal("setComment"),
    sheetId: z.string(),
    row: z.number().int().min(0),
    col: z.number().int().min(0),
    comment: z
      .object({
        id: z.string(),
        status: z.enum(["open", "resolved"]),
        createdAt: z.string(),
        updatedAt: z.string(),
        messages: z
          .array(
            z.object({
              id: z.string(),
              userId: z.string(),
              authorName: z.string(),
              text: z.string(),
              createdAt: z.string(),
            }),
          )
          .min(1),
      })
      .nullable(),
  }),
  z.object({
    type: z.literal("sortRange"),
    sheetId: z.string(),
    startRow: z.number().int().min(0),
    endRow: z.number().int().min(0),
    startCol: z.number().int().min(0),
    endCol: z.number().int().min(0),
    sortCol: z.number().int().min(0),
    ascending: z.boolean(),
  }),
  z.object({
    type: z.literal("restoreRange"),
    sheetId: z.string(),
    cells: z.record(z.string(), z.unknown().nullable()),
  }),
  z.object({
    type: z.literal("patchSheet"),
    sheetId: z.string(),
    charts: z.array(z.record(z.string(), z.unknown())).optional(),
    merges: z
      .array(
        z.object({
          startRow: z.number(),
          startCol: z.number(),
          endRow: z.number(),
          endCol: z.number(),
        }),
      )
      .optional(),
    rowCount: z.number().int().optional(),
    colCount: z.number().int().optional(),
    rowHeights: z.record(z.string(), z.number()).optional(),
    colWidths: z.record(z.string(), z.number()).optional(),
    cells: z.record(z.string(), z.unknown().nullable()).optional(),
    conditionalFormats: z
      .array(
        z.object({
          ref: z.string(),
          rules: z.array(z.record(z.string(), z.unknown())),
        }),
      )
      .optional(),
  }),
]);

const bodySchema = z.object({
  baseRevision: z.number().int().min(0),
  clientId: z.string().min(1),
  ops: z.array(opSchema).min(1).max(200),
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

  const { id } = await context.params;
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

  const current = access.spreadsheet;
  if (current.revision !== parsed.data.baseRevision) {
    const workbook = workbookFromDb(
      current.cells,
      current.rowCount,
      current.colCount,
      current.name,
    );
    return NextResponse.json(
      {
        error: "Revision conflict",
        revision: current.revision,
        workbook,
      },
      { status: 409 },
    );
  }

  let workbook = workbookFromDb(
    current.cells,
    current.rowCount,
    current.colCount,
    current.name,
  );
  workbook = applyWorkbookOps(workbook, parsed.data.ops as WorkbookOp[]);
  const { rowCount, colCount } = activeSheetMetrics(workbook);
  const nextRevision = current.revision + 1;

  const updated = await prisma.spreadsheet.updateMany({
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
      cells: workbook as object,
      revision: nextRevision,
      rowCount,
      colCount,
    },
  });

  if (updated.count === 0) {
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
    type: "ops",
    revision: nextRevision,
    ops: parsed.data.ops,
    clientId: parsed.data.clientId,
    userId: session.userId,
  });

  const author = await prisma.user.findUnique({
    where: { id: session.userId },
    select: { name: true, email: true },
  });
  void maybeCreateSpreadsheetVersion({
    spreadsheetId: id,
    revision: nextRevision,
    workbook,
    userId: session.userId,
    authorName: author?.name ?? author?.email ?? null,
  }).catch(() => {
    // version snapshots are best-effort
  });

  return NextResponse.json({ revision: nextRevision, workbook });
}