"use client";

import { FormEvent, useEffect, useState } from "react";

type ShareRow = {
  id: string;
  userId: string;
  email: string;
  name: string | null;
  permission: "VIEW" | "EDIT";
};

export function ShareDialog({
  spreadsheetId,
  open,
  onClose,
}: {
  spreadsheetId: string;
  open: boolean;
  onClose: () => void;
}) {
  const [shares, setShares] = useState<ShareRow[]>([]);
  const [email, setEmail] = useState("");
  const [permission, setPermission] = useState<"VIEW" | "EDIT">("VIEW");
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    void fetch(`/api/spreadsheets/${spreadsheetId}/shares`)
      .then((r) => r.json())
      .then((d) => setShares(d.shares ?? []))
      .catch(() => setShares([]));
  }, [open, spreadsheetId]);

  if (!open) return null;

  async function onSubmit(event: FormEvent) {
    event.preventDefault();
    setError(null);
    const response = await fetch(`/api/spreadsheets/${spreadsheetId}/shares`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, permission }),
    });
    const payload = await response.json();
    if (!response.ok) {
      setError(payload.error ?? "Share failed");
      return;
    }
    setShares((prev) => {
      const without = prev.filter((s) => s.userId !== payload.share.userId);
      return [...without, payload.share];
    });
    setEmail("");
  }

  async function revoke(userId: string) {
    await fetch(
      `/api/spreadsheets/${spreadsheetId}/shares?userId=${encodeURIComponent(userId)}`,
      { method: "DELETE" },
    );
    setShares((prev) => prev.filter((s) => s.userId !== userId));
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/30 p-4">
      <div className="w-full max-w-md rounded border border-zinc-300 bg-white p-4 shadow-lg">
        <h2 className="text-lg font-semibold text-zinc-900">Share spreadsheet</h2>
        <form onSubmit={onSubmit} className="mt-3 space-y-2">
          <input
            type="email"
            required
            placeholder="colleague@example.com"
            className="w-full rounded border px-2 py-1.5 text-sm"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
          />
          <select
            className="w-full rounded border px-2 py-1.5 text-sm"
            value={permission}
            onChange={(e) => setPermission(e.target.value as "VIEW" | "EDIT")}
          >
            <option value="VIEW">Can view</option>
            <option value="EDIT">Can edit</option>
          </select>
          <button
            type="submit"
            className="rounded bg-[#217346] px-3 py-1.5 text-sm text-white"
          >
            Share
          </button>
        </form>
        {error ? <p className="mt-2 text-sm text-red-600">{error}</p> : null}
        <ul className="mt-4 max-h-48 space-y-2 overflow-y-auto text-sm">
          {shares.map((share) => (
            <li
              key={share.id}
              className="flex items-center justify-between gap-2 border-b border-zinc-100 pb-2"
            >
              <span>
                {share.name ?? share.email} ({share.permission.toLowerCase()})
              </span>
              <button
                type="button"
                className="text-red-600 underline"
                onClick={() => void revoke(share.userId)}
              >
                Revoke
              </button>
            </li>
          ))}
        </ul>
        <button
          type="button"
          className="mt-4 text-sm text-zinc-600 underline"
          onClick={onClose}
        >
          Close
        </button>
      </div>
    </div>
  );
}