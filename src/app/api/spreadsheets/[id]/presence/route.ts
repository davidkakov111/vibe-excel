import { NextResponse } from "next/server";
import { z } from "zod";
import { publishSpreadsheetEvent } from "@/lib/collaboration/hub";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { canView, getSpreadsheetAccess } from "@/lib/spreadsheet-access";

type RouteContext = { params: Promise<{ id: string }> };

const presenceSchema = z.object({
  clientId: z.string(),
  userId: z.string().optional(),
  userName: z.string().optional(),
  color: z.string().optional(),
  sheetId: z.string().optional(),
  selection: z
    .object({
      row: z.number(),
      col: z.number(),
      endRow: z.number(),
      endCol: z.number(),
    })
    .nullable()
    .optional(),
  leave: z.boolean().optional(),
});

export async function POST(request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !canView(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }

  const parsed = presenceSchema.safeParse(body);
  if (!parsed.success) {
    return NextResponse.json({ error: "Invalid input" }, { status: 400 });
  }

  if (parsed.data.leave) {
    publishSpreadsheetEvent(id, {
      type: "presence_leave",
      clientId: parsed.data.clientId,
    });
    return NextResponse.json({ ok: true });
  }

  if (
    !parsed.data.userId ||
    !parsed.data.userName ||
    !parsed.data.color ||
    !parsed.data.sheetId
  ) {
    return NextResponse.json({ error: "Invalid input" }, { status: 400 });
  }

  if (parsed.data.userId !== session.userId) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  publishSpreadsheetEvent(id, {
    type: "presence",
    clientId: parsed.data.clientId,
    userId: parsed.data.userId,
    userName: parsed.data.userName,
    color: parsed.data.color,
    sheetId: parsed.data.sheetId,
    selection: parsed.data.selection ?? null,
  });

  return NextResponse.json({ ok: true });
}