import type { SharePermission, Spreadsheet } from "@prisma/client";
import { prisma } from "@/lib/prisma";

export type SpreadsheetRole = "owner" | "edit" | "view";

export type SpreadsheetAccess = {
  spreadsheet: Spreadsheet;
  role: SpreadsheetRole;
  permission: SharePermission | "OWNER";
};

export async function getSpreadsheetAccess(
  userId: string,
  spreadsheetId: string,
): Promise<SpreadsheetAccess | null> {
  const owned = await prisma.spreadsheet.findFirst({
    where: { id: spreadsheetId, userId },
  });
  if (owned) {
    return { spreadsheet: owned, role: "owner", permission: "OWNER" };
  }

  const share = await prisma.spreadsheetShare.findFirst({
    where: { spreadsheetId, userId },
    include: { spreadsheet: true },
  });
  if (!share) return null;

  return {
    spreadsheet: share.spreadsheet,
    role: share.permission === "EDIT" ? "edit" : "view",
    permission: share.permission,
  };
}

export function canEdit(access: SpreadsheetAccess): boolean {
  return access.role === "owner" || access.role === "edit";
}

export function canView(_access: SpreadsheetAccess): boolean {
  return true;
}

export function isOwner(access: SpreadsheetAccess): boolean {
  return access.role === "owner";
}