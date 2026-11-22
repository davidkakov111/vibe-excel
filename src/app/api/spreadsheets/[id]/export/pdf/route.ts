import { NextResponse } from "next/server";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { canView, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { workbookToPdfBuffer } from "@/lib/excel/workbook-pdf";
import { workbookFromDb } from "@/lib/spreadsheet-db";

type RouteContext = { params: Promise<{ id: string }> };

function pdfFileName(name: string): string {
  const base = name.trim() || "spreadsheet";
  return `${base.replace(/[^\w.-]+/g, "_")}.pdf`;
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
  const spreadsheet = access.spreadsheet;

  const workbook = workbookFromDb(
    spreadsheet.cells,
    spreadsheet.rowCount,
    spreadsheet.colCount,
    spreadsheet.name,
  );

  const buffer = await workbookToPdfBuffer(workbook);

  return new NextResponse(new Uint8Array(buffer), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${pdfFileName(spreadsheet.name)}"`,
    },
  });
}