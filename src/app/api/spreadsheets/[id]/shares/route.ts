import { NextResponse } from "next/server";
import { z } from "zod";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { getSpreadsheetAccess, isOwner } from "@/lib/spreadsheet-access";
import { prisma } from "@/lib/prisma";
import { nanoid } from "nanoid";

type RouteContext = { params: Promise<{ id: string }> };

const shareSchema = z.object({
  email: z.string().email(),
  permission: z.enum(["VIEW", "EDIT"]),
});

export async function GET(_request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !isOwner(access)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  const shares = await prisma.spreadsheetShare.findMany({
    where: { spreadsheetId: id },
    include: { user: { select: { id: true, email: true, name: true } } },
    orderBy: { createdAt: "asc" },
  });

  return NextResponse.json({
    shares: shares.map((s) => ({
      id: s.id,
      userId: s.userId,
      email: s.user.email,
      name: s.user.name,
      permission: s.permission,
      createdAt: s.createdAt,
    })),
  });
}

export async function POST(request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !isOwner(access)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }

  const parsed = shareSchema.safeParse(body);
  if (!parsed.success) {
    return NextResponse.json(
      { error: parsed.error.issues[0]?.message ?? "Invalid input" },
      { status: 400 },
    );
  }

  const target = await prisma.user.findUnique({
    where: { email: parsed.data.email.toLowerCase() },
  });
  if (!target) {
    return NextResponse.json({ error: "User not found" }, { status: 404 });
  }
  if (target.id === session.userId) {
    return NextResponse.json({ error: "Cannot share with yourself" }, { status: 400 });
  }

  const share = await prisma.spreadsheetShare.upsert({
    where: {
      spreadsheetId_userId: {
        spreadsheetId: id,
        userId: target.id,
      },
    },
    create: {
      id: nanoid(),
      spreadsheetId: id,
      userId: target.id,
      permission: parsed.data.permission,
    },
    update: { permission: parsed.data.permission },
    include: { user: { select: { id: true, email: true, name: true } } },
  });

  return NextResponse.json({
    share: {
      id: share.id,
      userId: share.userId,
      email: share.user.email,
      name: share.user.name,
      permission: share.permission,
      createdAt: share.createdAt,
    },
  });
}

export async function DELETE(request: Request, context: RouteContext) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) return session;

  const { id } = await context.params;
  const access = await getSpreadsheetAccess(session.userId, id);
  if (!access || !isOwner(access)) {
    return NextResponse.json({ error: "Forbidden" }, { status: 403 });
  }

  const { searchParams } = new URL(request.url);
  const userId = searchParams.get("userId");
  if (!userId) {
    return NextResponse.json({ error: "Missing userId" }, { status: 400 });
  }

  await prisma.spreadsheetShare.deleteMany({
    where: { spreadsheetId: id, userId },
  });

  return NextResponse.json({ ok: true });
}