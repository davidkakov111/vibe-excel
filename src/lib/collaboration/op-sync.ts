import { applyWorkbookOps, type WorkbookOp } from "@/lib/workbook/ops";
import type { WorkbookDocument } from "@/lib/workbook/types";

/** After a revision conflict, merge server truth with ops the client still needs. */
export function reconcileWorkbookWithOps(
  serverWorkbook: WorkbookDocument,
  ops: WorkbookOp[],
): WorkbookDocument {
  return applyWorkbookOps(serverWorkbook, ops);
}

export type OpsPostResult =
  | { ok: true; revision: number; workbook?: WorkbookDocument }
  | { ok: false; conflict: true; revision: number; workbook: WorkbookDocument }
  | { ok: false; conflict: false };

export async function postOpsWithRetry(args: {
  spreadsheetId: string;
  clientId: string;
  baseRevision: number;
  ops: WorkbookOp[];
  maxAttempts?: number;
  onRevision?: (revision: number) => void;
}): Promise<OpsPostResult> {
  let revision = args.baseRevision;
  for (let attempt = 0; attempt < (args.maxAttempts ?? 5); attempt++) {
    const response = await fetch(`/api/spreadsheets/${args.spreadsheetId}/ops`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        baseRevision: revision,
        clientId: args.clientId,
        ops: args.ops,
      }),
    });

    if (response.ok) {
      const payload = (await response.json()) as {
        revision: number;
        workbook?: WorkbookDocument;
      };
      return { ok: true, revision: payload.revision, workbook: payload.workbook };
    }

    if (response.status === 409) {
      const payload = (await response.json()) as {
        revision: number;
        workbook: WorkbookDocument;
      };
      revision = payload.revision;
      args.onRevision?.(revision);
      if (attempt + 1 >= (args.maxAttempts ?? 5)) {
        return {
          ok: false,
          conflict: true,
          revision: payload.revision,
          workbook: payload.workbook,
        };
      }
      continue;
    }

    return { ok: false, conflict: false };
  }

  return { ok: false, conflict: false };
}