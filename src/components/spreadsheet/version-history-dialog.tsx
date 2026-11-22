"use client";

import { useCallback, useEffect, useState } from "react";
import type { WorkbookDocument } from "@/lib/workbook/types";

type VersionRow = {
  id: string;
  revision: number;
  createdAt: string;
  authorName: string | null;
  label: string;
};

interface VersionHistoryDialogProps {
  spreadsheetId: string;
  open: boolean;
  clientId: string;
  baseRevision: number;
  onClose: () => void;
  onRestored: (workbook: WorkbookDocument, revision: number) => void;
  readOnly?: boolean;
}

export function VersionHistoryDialog({
  spreadsheetId,
  open,
  clientId,
  baseRevision,
  onClose,
  onRestored,
  readOnly,
}: VersionHistoryDialogProps) {
  const [versions, setVersions] = useState<VersionRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [preview, setPreview] = useState<{
    label: string;
    workbook: WorkbookDocument;
  } | null>(null);
  const [restoring, setRestoring] = useState(false);
  const [savingSnapshot, setSavingSnapshot] = useState(false);

  const loadList = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/spreadsheets/${spreadsheetId}/versions`);
      if (!res.ok) throw new Error("Failed to load versions");
      const data = (await res.json()) as { versions: VersionRow[] };
      setVersions(data.versions);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Failed to load versions");
    } finally {
      setLoading(false);
    }
  }, [spreadsheetId]);

  useEffect(() => {
    if (!open) return;
    void loadList();
    setSelectedId(null);
    setPreview(null);
  }, [loadList, open]);

  const loadPreview = async (id: string) => {
    setSelectedId(id);
    setPreview(null);
    setError(null);
    try {
      const res = await fetch(
        `/api/spreadsheets/${spreadsheetId}/versions/${id}`,
      );
      if (!res.ok) throw new Error("Failed to load preview");
      const data = (await res.json()) as {
        label: string;
        workbook: WorkbookDocument;
      };
      setPreview({ label: data.label, workbook: data.workbook });
    } catch (e) {
      setError(e instanceof Error ? e.message : "Preview failed");
    }
  };

  const saveCurrentVersion = async () => {
    if (readOnly) return;
    const label = window.prompt(
      "Version label (optional)",
      "Saved version",
    );
    if (label === null) return;
    setSavingSnapshot(true);
    setError(null);
    try {
      const res = await fetch(
        `/api/spreadsheets/${spreadsheetId}/versions/snapshot`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            label: label.trim() || "Saved version",
          }),
        },
      );
      if (!res.ok) {
        const data = (await res.json().catch(() => null)) as { error?: string };
        throw new Error(data?.error ?? "Could not save version");
      }
      await loadList();
    } catch (e) {
      setError(e instanceof Error ? e.message : "Could not save version");
    } finally {
      setSavingSnapshot(false);
    }
  };

  const restore = async () => {
    if (!selectedId || readOnly) return;
    setRestoring(true);
    setError(null);
    try {
      const res = await fetch(
        `/api/spreadsheets/${spreadsheetId}/versions/${selectedId}/restore`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ baseRevision, clientId }),
        },
      );
      const data = (await res.json()) as {
        error?: string;
        workbook?: WorkbookDocument;
        revision?: number;
      };
      if (!res.ok) {
        throw new Error(data.error ?? "Restore failed");
      }
      if (data.workbook && data.revision !== undefined) {
        onRestored(data.workbook, data.revision);
        onClose();
      }
    } catch (e) {
      setError(e instanceof Error ? e.message : "Restore failed");
    } finally {
      setRestoring(false);
    }
  };

  if (!open) return null;

  const previewSheet =
    preview?.workbook.sheets.find(
      (s) => s.id === preview.workbook.activeSheetId,
    ) ?? preview?.workbook.sheets[0];

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-labelledby="vh-title"
    >
      <div className="flex max-h-[90vh] w-full max-w-3xl flex-col overflow-hidden rounded-lg bg-white shadow-xl">
        <div className="flex items-center justify-between border-b border-zinc-200 px-4 py-3">
          <div>
            <h2 id="vh-title" className="text-lg font-semibold text-zinc-900">
              Version history
            </h2>
            <p className="text-xs text-zinc-500">
              Auto-snapshots about every 5 edits or 2 minutes (not every cell
              keystroke). Use &quot;Save current version&quot; anytime.
            </p>
          </div>
          <div className="flex shrink-0 gap-2">
            <button
              type="button"
              disabled={readOnly || savingSnapshot}
              onClick={() => void saveCurrentVersion()}
              className="rounded bg-[#217346] px-2 py-1 text-sm text-white disabled:opacity-50"
            >
              {savingSnapshot ? "Saving…" : "Save current version"}
            </button>
            <button
              type="button"
              onClick={onClose}
              className="rounded px-2 py-1 text-sm text-zinc-600 hover:bg-zinc-100"
            >
              Close
            </button>
          </div>
        </div>
        {error ? (
          <p className="bg-red-50 px-4 py-2 text-sm text-red-700">{error}</p>
        ) : null}
        <div className="flex min-h-0 flex-1 divide-x divide-zinc-200">
          <div className="w-1/2 overflow-y-auto p-3">
            {loading ? (
              <p className="text-sm text-zinc-500">Loading…</p>
            ) : (
              <ul className="space-y-1">
                {versions.map((v) => (
                  <li key={v.id}>
                    <button
                      type="button"
                      className={`w-full rounded border px-2 py-2 text-left text-xs ${
                        selectedId === v.id
                          ? "border-[#217346] bg-emerald-50"
                          : "border-zinc-200 hover:bg-zinc-50"
                      }`}
                      onClick={() => void loadPreview(v.id)}
                    >
                      <span className="block font-medium text-zinc-900">
                        {v.label}
                      </span>
                      <span className="text-zinc-600">
                        Rev {v.revision} ·{" "}
                        {new Date(v.createdAt).toLocaleString()}
                        {v.authorName ? ` · ${v.authorName}` : ""}
                      </span>
                    </button>
                  </li>
                ))}
                {!versions.length && !loading ? (
                  <li className="text-sm text-zinc-500">
                    No saved versions yet. Snapshots are created periodically
                    while you edit.
                  </li>
                ) : null}
              </ul>
            )}
          </div>
          <div className="flex w-1/2 flex-col p-3">
            <h3 className="mb-2 text-sm font-medium text-zinc-800">Preview</h3>
            {!preview ? (
              <p className="text-sm text-zinc-500">
                Select a version to preview sheet names and size.
              </p>
            ) : (
              <div className="min-h-0 flex-1 space-y-2 overflow-y-auto text-sm">
                <p className="font-medium">{preview.label}</p>
                <p>
                  {preview.workbook.sheets.length} sheet
                  {preview.workbook.sheets.length === 1 ? "" : "s"}:{" "}
                  {preview.workbook.sheets.map((s) => s.name).join(", ")}
                </p>
                {previewSheet ? (
                  <p className="text-zinc-600">
                    Active sheet &quot;{previewSheet.name}&quot;:{" "}
                    {previewSheet.rowCount}×{previewSheet.colCount} cells,{" "}
                    {Object.keys(previewSheet.cells).length} populated,{" "}
                    {previewSheet.conditionalFormats?.length ?? 0} conditional
                    format region(s).
                  </p>
                ) : null}
              </div>
            )}
            <button
              type="button"
              disabled={!selectedId || restoring || readOnly}
              onClick={() => void restore()}
              className="mt-3 rounded bg-[#217346] px-3 py-2 text-sm font-medium text-white disabled:opacity-50"
            >
              {restoring ? "Restoring…" : "Restore this version"}
            </button>
            {readOnly ? (
              <p className="mt-2 text-xs text-zinc-500">View-only — cannot restore.</p>
            ) : (
              <p className="mt-2 text-xs text-zinc-500">
                Restore saves a &quot;Before restore&quot; snapshot, applies the
                version for everyone connected, and bumps the revision.
              </p>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}