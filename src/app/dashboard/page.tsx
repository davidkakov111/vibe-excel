import { DashboardClient } from "@/components/dashboard-client";
import { auth } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { redirect } from "next/navigation";

export default async function DashboardPage() {
  const session = await auth();
  if (!session?.user?.id) {
    redirect("/login");
  }

  const owned = await prisma.spreadsheet.findMany({
    where: { userId: session.user.id },
    orderBy: { updatedAt: "desc" },
    select: {
      id: true,
      name: true,
      createdAt: true,
      updatedAt: true,
    },
  });

  const shared = await prisma.spreadsheetShare.findMany({
    where: { userId: session.user.id },
    include: {
      spreadsheet: {
        select: { id: true, name: true, createdAt: true, updatedAt: true },
      },
    },
    orderBy: { createdAt: "desc" },
  });

  return (
    <main className="min-h-screen bg-zinc-50">
      <DashboardClient
        userName={session.user.name}
        userEmail={session.user.email}
        initialOwned={owned.map((sheet) => ({
          ...sheet,
          createdAt: sheet.createdAt.toISOString(),
          updatedAt: sheet.updatedAt.toISOString(),
        }))}
        initialShared={shared.map((entry) => ({
          ...entry.spreadsheet,
          permission: entry.permission,
          createdAt: entry.spreadsheet.createdAt.toISOString(),
          updatedAt: entry.spreadsheet.updatedAt.toISOString(),
        }))}
      />
    </main>
  );
}