"use client";

import Link from "next/link";
import { signOut } from "next-auth/react";
import { FormEvent, useRef, useState, type ReactNode } from "react";

interface SpreadsheetSummary {
  id: string;
  name: string;
  createdAt: string;
  updatedAt: string;
}

interface SharedSummary extends SpreadsheetSummary {
  permission: "VIEW" | "EDIT";
}

interface DashboardClientProps {
  userName: string | null | undefined;
  userEmail: string | null | undefined;
  initialOwned: SpreadsheetSummary[];
  initialShared: SharedSummary[];
}

export function DashboardClient({
  userName,
  userEmail,
  initialOwned,
  initialShared,
}: DashboardClientProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [owned, setOwned] = useState<SpreadsheetSummary[]>(initialOwned);
  const [shared, setShared] = useState<SharedSummary[]>(initialShared);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function refreshList() {
    const response = await fetch("/api/spreadsheets");
    if (!response.ok) {
      return;
    }
    const data = (await response.json()) as {
      owned: SpreadsheetSummary[];
      shared: SharedSummary[];
    };
    setOwned(
      data.owned.map((sheet) => ({
        ...sheet,
        createdAt: String(sheet.createdAt),
        updatedAt: String(sheet.updatedAt),
      })),
    );
    setShared(
      data.shared.map((sheet) => ({
        ...sheet,
        createdAt: String(sheet.createdAt),
        updatedAt: String(sheet.updatedAt),
      })),
    );
  }

  async function handleCreateBlank() {
    setBusy(true);
    setError(null);
    try {
      const response = await fetch("/api/spreadsheets", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: "Untitled spreadsheet" }),
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to create spreadsheet");
      }
      window.location.assign(`/spreadsheet/${payload.spreadsheet.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to create spreadsheet");
    } finally {
      setBusy(false);
    }
  }

  async function handleUpload(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const file = fileInputRef.current?.files?.[0];
    if (!file) {
      setError("Choose an Excel file to upload");
      return;
    }

    setBusy(true);
    setError(null);
    try {
      const formData = new FormData();
      formData.append("file", file);
      const response = await fetch("/api/spreadsheets", {
        method: "POST",
        body: formData,
      });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to upload spreadsheet");
      }
      window.location.assign(`/spreadsheet/${payload.spreadsheet.id}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to upload spreadsheet");
    } finally {
      setBusy(false);
    }
  }

  async function handleDelete(id: string) {
    if (!confirm("Delete this spreadsheet? This cannot be undone.")) {
      return;
    }
    setBusy(true);
    setError(null);
    try {
      const response = await fetch(`/api/spreadsheets/${id}`, {
        method: "DELETE",
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => null);
        throw new Error(payload?.error ?? "Failed to delete spreadsheet");
      }
      await refreshList();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to delete spreadsheet");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="mx-auto w-full max-w-5xl px-4 py-8">
      <header className="mb-8 flex flex-wrap items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-semibold text-zinc-900">Dashboard</h1>
          <p className="text-sm text-zinc-600">
            Signed in as {userName || userEmail}
          </p>
        </div>
        <button
          type="button"
          onClick={() => signOut({ callbackUrl: "/login" })}
          className="rounded-lg border border-zinc-300 px-4 py-2 text-sm hover:bg-zinc-50"
        >
          Log out
        </button>
      </header>

      <section className="mb-8 grid gap-4 rounded-xl border border-zinc-200 bg-white p-5 shadow-sm md:grid-cols-2">
        <DashboardPanel title="New spreadsheet">
          <p className="mb-3 text-sm text-zinc-600">
            Start from a blank 100×26 grid with automatic saving.
          </p>
          <button
            type="button"
            disabled={busy}
            onClick={() => void handleCreateBlank()}
            className="rounded-lg bg-blue-600 px-4 py-2 text-sm font-medium text-white hover:bg-blue-700 disabled:opacity-60"
          >
            Create blank spreadsheet
          </button>
        </DashboardPanel>

        <DashboardPanel title="Upload Excel file">
          <form onSubmit={handleUpload} className="space-y-3">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              className="block w-full text-sm text-zinc-700 file:mr-3 file:rounded-md file:border-0 file:bg-zinc-100 file:px-3 file:py-2 file:text-sm"
            />
            <button
              type="submit"
              disabled={busy}
              className="rounded-lg border border-zinc-300 px-4 py-2 text-sm font-medium hover:bg-zinc-50 disabled:opacity-60"
            >
              Upload and open
            </button>
          </form>
        </DashboardPanel>
      </section>

      {error ? (
        <p className="mb-4 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-sm text-red-700">
          {error}
        </p>
      ) : null}

      <section className="rounded-xl border border-zinc-200 bg-white shadow-sm">
        <h2 className="border-b border-zinc-200 px-5 py-4 text-lg font-medium">
          Your spreadsheets
        </h2>
        {owned.length === 0 ? (
          <p className="px-5 py-8 text-sm text-zinc-500">
            No spreadsheets yet. Create a blank sheet or upload an Excel file.
          </p>
        ) : (
          <ul className="divide-y divide-zinc-200">
            {owned.map((sheet) => (
              <li
                key={sheet.id}
                className="flex flex-wrap items-center justify-between gap-3 px-5 py-4"
              >
                <div>
                  <p className="font-medium text-zinc-900">{sheet.name}</p>
                  <p className="text-xs text-zinc-500">
                    Updated {new Date(sheet.updatedAt).toLocaleString()}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <Link
                    href={`/spreadsheet/${sheet.id}`}
                    className="rounded-md bg-blue-600 px-3 py-1.5 text-sm text-white hover:bg-blue-700"
                  >
                    Open
                  </Link>
                  <button
                    type="button"
                    disabled={busy}
                    onClick={() => void handleDelete(sheet.id)}
                    className="rounded-md border border-red-200 px-3 py-1.5 text-sm text-red-700 hover:bg-red-50 disabled:opacity-60"
                  >
                    Delete
                  </button>
                </div>
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="mt-6 rounded-xl border border-zinc-200 bg-white shadow-sm">
        <h2 className="border-b border-zinc-200 px-5 py-4 text-lg font-medium">
          Shared with you
        </h2>
        {shared.length === 0 ? (
          <p className="px-5 py-8 text-sm text-zinc-500">
            No shared spreadsheets yet.
          </p>
        ) : (
          <ul className="divide-y divide-zinc-200">
            {shared.map((sheet) => (
              <li
                key={sheet.id}
                className="flex flex-wrap items-center justify-between gap-3 px-5 py-4"
              >
                <div>
                  <p className="font-medium text-zinc-900">{sheet.name}</p>
                  <p className="text-xs text-zinc-500">
                    {sheet.permission === "EDIT" ? "Can edit" : "View only"} ·
                    Updated {new Date(sheet.updatedAt).toLocaleString()}
                  </p>
                </div>
                <Link
                  href={`/spreadsheet/${sheet.id}`}
                  className="rounded-md bg-blue-600 px-3 py-1.5 text-sm text-white hover:bg-blue-700"
                >
                  Open
                </Link>
              </li>
            ))}
          </ul>
        )}
      </section>
    </div>
  );
}

function DashboardPanel({
  title,
  children,
}: {
  title: string;
  children: ReactNode;
}) {
  return (
    <div>
      <h2 className="mb-2 text-base font-medium text-zinc-900">{title}</h2>
      {children}
    </div>
  );
}