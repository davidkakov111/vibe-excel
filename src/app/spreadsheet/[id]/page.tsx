import { WorkbookEditor } from "@/components/spreadsheet/workbook-editor";
import { auth } from "@/lib/auth";
import { getSpreadsheetAccess, canView, canEdit, isOwner } from "@/lib/spreadsheet-access";
import { workbookFromDb } from "@/lib/spreadsheet-db";
import { notFound, redirect } from "next/navigation";

type PageProps = { params: Promise<{ id: string }> };

export default async function SpreadsheetPage({ params }: PageProps) {
  const session = await auth();
  if (!session?.user?.id) {
    redirect("/login");
  }

  const { id } = await params;
  const access = await getSpreadsheetAccess(session.user.id, id);
  if (!access || !canView(access)) {
    notFound();
  }

  const spreadsheet = access.spreadsheet;
  const workbook = workbookFromDb(
    spreadsheet.cells,
    spreadsheet.rowCount,
    spreadsheet.colCount,
    spreadsheet.name,
  );

  return (
    <main className="min-h-screen bg-[#e6e6e6]">
      <WorkbookEditor
        spreadsheetId={spreadsheet.id}
        initialName={spreadsheet.name}
        initialWorkbook={workbook}
        initialRevision={spreadsheet.revision}
        canEdit={canEdit(access)}
        isOwner={isOwner(access)}
        userId={session.user.id}
        userName={session.user.name ?? session.user.email ?? "User"}
      />
    </main>
  );
}