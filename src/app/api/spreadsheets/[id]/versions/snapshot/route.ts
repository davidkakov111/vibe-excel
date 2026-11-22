import { NextResponse } from "next/server";
import { z } from "zod";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { workbookFromDb } from "@/lib/spreadsheet-db";
import { canEdit, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { createSpreadsheetVersion } from "@/lib/spreadsheet-versions";
import { prisma } from "@/lib/prisma";

type RouteContext = { params: Promise<{ id: string }> };

const bodySchema = z.object({
  label: z.string().trim().min(1).max(120).optional(),
});

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

  let body: unknown = {};
  try {
    if (request.headers.get("content-length") !== "0") {
      body = await request.json();
    }
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
  const workbook = workbookFromDb(
    current.cells,
    current.rowCount,
    current.colCount,
    current.name,
  );

  const author = await prisma.user.findUnique({
    where: { id: session.userId },
    select: { name: true, email: true },
  });

  await createSpreadsheetVersion({
    spreadsheetId: id,
    revision: current.revision,
    workbook,
    userId: session.userId,
    authorName: author?.name ?? author?.email ?? null,
    label: parsed.data.label ?? "Saved version",
    force: true,
  });

  return NextResponse.json({ ok: true, revision: current.revision });
}