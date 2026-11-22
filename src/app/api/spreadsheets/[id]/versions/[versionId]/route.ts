import { NextResponse } from "next/server";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { canView, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { prisma } from "@/lib/prisma";
import { assertWorkbookSnapshot } from "@/lib/workbook/version-history";

type RouteContext = { params: Promise<{ id: string; versionId: string }> };

export async function GET(_request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id, versionId } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !canView(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  const version = await prisma.spreadsheetVersion.findFirst({
    where: { id: versionId, spreadsheetId: id },
  });
  if (!version) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  const workbook = assertWorkbookSnapshot(version.workbook);
  return NextResponse.json({
    id: version.id,
    revision: version.revision,
    createdAt: version.createdAt.toISOString(),
    authorName: version.authorName,
    label: version.label,
    workbook,
  });
}