import { subscribeSpreadsheet } from "@/lib/collaboration/hub";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { canView, getSpreadsheetAccess } from "@/lib/spreadsheet-access";
import { NextResponse } from "next/server";

type RouteContext = { params: Promise<{ id: string }> };

export const dynamic = "force-dynamic";
export const runtime = "nodejs";

export async function GET(_request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !canView(access)) {
    return NextResponse.json({ error: "Not found" }, { status: 404 });
  }

  const encoder = new TextEncoder();
  let unsubscribe = () => {};

  const stream = new ReadableStream({
    start(controller) {
      controller.enqueue(
        encoder.encode(
          `data: ${JSON.stringify({
            type: "revision",
            revision: access.spreadsheet.revision,
          })}\n\n`,
        ),
      );

      unsubscribe = subscribeSpreadsheet(id, (event) => {
        try {
          controller.enqueue(
            encoder.encode(`data: ${JSON.stringify(event)}\n\n`),
          );
        } catch {
          unsubscribe();
        }
      });
    },
    cancel() {
      unsubscribe();
    },
  });

  return new Response(stream, {
    headers: {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
    },
  });
}