import { NextResponse } from "next/server";
import { z } from "zod";
import { requireSessionUserId } from "@/lib/auth-helpers";
import { parseExcelWorkbook } from "@/lib/excel/workbook-io";
import { createEmptyWorkbook } from "@/lib/workbook/defaults";
import { prisma } from "@/lib/prisma";

const createSchema = z.object({
  name: z.string().trim().min(1).max(200).optional(),
});

export async function GET() {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) {
    return session;
  }

  const owned = await prisma.spreadsheet.findMany({
    where: { userId: session.userId },
    orderBy: { updatedAt: "desc" },
    select: {
      id: true,
      name: true,
      createdAt: true,
      updatedAt: true,
    },
  });

  const shared = await prisma.spreadsheetShare.findMany({
    where: { userId: session.userId },
    include: {
      spreadsheet: {
        select: { id: true, name: true, createdAt: true, updatedAt: true },
      },
    },
    orderBy: { createdAt: "desc" },
  });

  return NextResponse.json({
    owned,
    shared: shared.map((entry) => ({
      ...entry.spreadsheet,
      permission: entry.permission,
    })),
  });
}

export async function POST(request: Request) {
  const session = await requireSessionUserId();
  if (session instanceof NextResponse) {
    return session;
  }

  const contentType = request.headers.get("content-type") ?? "";

  if (contentType.includes("multipart/form-data")) {
    const formData = await request.formData();
    const file = formData.get("file");
    const nameField = formData.get("name");

    if (!(file instanceof File)) {
      return NextResponse.json({ error: "Missing Excel file" }, { status: 400 });
    }

    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".xls")) {
      return NextResponse.json(
        { error: "Only .xlsx and .xls files are supported" },
        { status: 400 },
      );
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const workbook = await parseExcelWorkbook(buffer);
    const active = workbook.sheets[0]!;
    const name =
      (typeof nameField === "string" && nameField.trim()) ||
      file.name.replace(/\.(xlsx|xls)$/i, "") ||
      "Imported spreadsheet";

    const spreadsheet = await prisma.spreadsheet.create({
      data: {
        userId: session.userId,
        name,
        rowCount: active.rowCount,
        colCount: active.colCount,
        cells: workbook as object,
      },
    });

    return NextResponse.json({ spreadsheet }, { status: 201 });
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON body" }, { status: 400 });
  }

  const parsed = createSchema.safeParse(body ?? {});
  if (!parsed.success) {
    return NextResponse.json(
      { error: parsed.error.issues[0]?.message ?? "Invalid input" },
      { status: 400 },
    );
  }

  const workbook = createEmptyWorkbook();
  const active = workbook.sheets[0]!;
  const spreadsheet = await prisma.spreadsheet.create({
    data: {
      userId: session.userId,
      name: parsed.data.name ?? "Untitled spreadsheet",
      rowCount: active.rowCount,
      colCount: active.colCount,
      cells: workbook as object,
    },
  });

  return NextResponse.json({ spreadsheet }, { status: 201 });
}