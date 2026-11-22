import { NextResponse } from "next/server";
import { z } from "zod";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { workbookFromDb } from "@/lib/spreadsheet-db";
import {
  canEdit,
  canView,
  getSpreadsheetAccess,
  isOwner,
} from "@/lib/spreadsheet-access";
import { prisma } from "@/lib/prisma";
import type { WorkbookDocument } from "@/lib/workbook/types";
import { WORKBOOK_VERSION } from "@/lib/workbook/types";

type RouteContext = { params: Promise<{ id: string }> };

const workbookSchema = z.object({
  version: z.literal(WORKBOOK_VERSION),
  activeSheetId: z.string(),
  sheets: z.array(z.any()),
  settings: z.object({
    showGridlines: z.boolean(),
    zoom: z.number(),
  }),
});

const patchSchema = z.object({
  name: z.string().trim().min(1).max(200).optional(),
  workbook: workbookSchema.optional(),
  cells: z.record(z.string(), z.string()).optional(),
  rowCount: z.number().int().min(1).max(5000).optional(),
  colCount: z.number().int().min(1).max(500).optional(),
  baseRevision: z.number().int().min(0).optional(),
});

function spreadsheetPayload(
  spreadsheet: {
    id: string;
    name: string;
    cells: unknown;
    rowCount: number;
    colCount: number;
    revision: number;
    createdAt: Date;
    updatedAt: Date;
    userId: string;
  },
  access: { role: string },
) {
  const workbook = workbookFromDb(
    spreadsheet.cells,
    spreadsheet.rowCount,
    spreadsheet.colCount,
    spreadsheet.name,
  );
  const active =
    workbook.sheets.find((s) => s.id === workbook.activeSheetId) ??
    workbook.sheets[0];
  return {
    id: spreadsheet.id,
    name: spreadsheet.name,
    workbook,
    rowCount: active?.rowCount ?? spreadsheet.rowCount,
    colCount: active?.colCount ?? spreadsheet.colCount,
    revision: spreadsheet.revision,
    role: access.role,
    createdAt: spreadsheet.createdAt,
    updatedAt: spreadsheet.updatedAt,
  };
}

export async function GET(_request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) {
    return session;
  }

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !canView(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  return NextResponse.json({
    spreadsheet: spreadsheetPayload(access.spreadsheet, access),
  });
}

export async function PATCH(request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) {
    return session;
  }

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !canEdit(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }
  const existing = access.spreadsheet;

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }

  const parsed = patchSchema.safeParse(body);
  if (!parsed.success) {
    return NextResponse.json(
      { error: parsed.error.issues[0]?.message ?? "Invalid input" },
      { status: 400 },
    );
  }

  let workbook: WorkbookDocument = workbookFromDb(
    existing.cells,
    existing.rowCount,
    existing.colCount,
    existing.name,
  );

  if (parsed.data.workbook) {
    workbook = parsed.data.workbook as WorkbookDocument;
  } else if (parsed.data.cells) {
    const active =
      workbook.sheets.find((s) => s.id === workbook.activeSheetId) ??
      workbook.sheets[0]!;
    const nextCells = { ...active.cells };
    for (const [key, value] of Object.entries(parsed.data.cells)) {
      if (value === "") delete nextCells[key];
      else nextCells[key] = { input: value };
    }
    active.cells = nextCells;
  }

  if (parsed.data.rowCount !== undefined || parsed.data.colCount !== undefined) {
    const active =
      workbook.sheets.find((s) => s.id === workbook.activeSheetId) ??
      workbook.sheets[0]!;
    if (parsed.data.rowCount !== undefined) active.rowCount = parsed.data.rowCount;
    if (parsed.data.colCount !== undefined) active.colCount = parsed.data.colCount;
  }

  const active =
    workbook.sheets.find((s) => s.id === workbook.activeSheetId) ??
    workbook.sheets[0]!;

  const updated = await prisma.spreadsheet.updateMany({
    where: {
      id: existing.id,
      ...(parsed.data.baseRevision !== undefined
        ? { revision: parsed.data.baseRevision }
        : {}),
      OR: [
        { userId: session.userId },
        {
          shares: { some: { userId: session.userId, permission: "EDIT" } },
        },
      ],
    },
    data: {
      ...(parsed.data.name && isOwner(access) ? { name: parsed.data.name } : {}),
      rowCount: active.rowCount,
      colCount: active.colCount,
      cells: workbook as object,
      revision: { increment: 1 },
    },
  });
  if (updated.count === 0) {
    if (parsed.data.baseRevision !== undefined) {
      const fresh = await prisma.spreadsheet.findUnique({ where: { id: existing.id } });
      if (fresh) {
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
    }
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  const refreshed = await getSpreadsheetAccess(session.userId, existing.id);
  if (!refreshed) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  return NextResponse.json({
    spreadsheet: spreadsheetPayload(refreshed.spreadsheet, refreshed),
  });
}

export async function DELETE(_request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) {
    return session;
  }

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !isOwner(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  await prisma.spreadsheet.delete({ where: { id: access.spreadsheet.id } });
  return NextResponse.json({ ok: true });
}