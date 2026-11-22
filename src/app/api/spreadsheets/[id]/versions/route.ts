import { NextResponse } from "next/server";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { canView, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { listSpreadsheetVersions } from "@/lib/spreadsheet-versions";

type RouteContext = { params: Promise<{ id: string }> };

export async function GET(_request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !canView(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  const versions = await listSpreadsheetVersions(id);
  return NextResponse.json({
    versions: versions.map((v) => ({
      id: v.id,
      revision: v.revision,
      createdAt: v.createdAt.toISOString(),
      authorName: v.authorName,
      label: v.label,
      createdById: v.createdById,
    })),
  });
}