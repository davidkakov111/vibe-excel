"use client";

import { useEffect, useRef, useState } from "react";

interface FindReplaceDialogProps {
  open: boolean;
  onClose: () => void;
  onReplace: (
    query: string,
    replacement: string,
    matchCase: boolean,
    replaceAll: boolean,
  ) => void;
}

export function FindReplaceDialog({
  open,
  onClose,
  onReplace,
}: FindReplaceDialogProps) {
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const findRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (open) {
      findRef.current?.focus();
    }
  }, [open]);

  if (!open) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center bg-black/30 pt-24"
      onKeyDown={(event) => event.stopPropagation()}
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose();
      }}
    >
      <div
        className="w-full max-w-md rounded-xl border border-zinc-200 bg-white p-4 text-zinc-900 shadow-lg"
        role="dialog"
        aria-label="Find and replace"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <h2 className="mb-3 text-lg font-semibold">Find and replace</h2>
        <label className="mb-2 block text-sm">
          Find
          <input
            ref={findRef}
            className="mt-1 w-full rounded border border-zinc-300 bg-white px-2 py-1 text-zinc-900"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(event) => event.stopPropagation()}
          />
        </label>
        <label className="mb-2 block text-sm">
          Replace with
          <input
            className="mt-1 w-full rounded border border-zinc-300 bg-white px-2 py-1 text-zinc-900"
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            onKeyDown={(event) => event.stopPropagation()}
          />
        </label>
        <label className="mb-3 flex items-center gap-2 text-sm">
          <input
            type="checkbox"
            checked={matchCase}
            onChange={(e) => setMatchCase(e.target.checked)}
          />
          Match case
        </label>
        <div className="flex flex-wrap gap-2">
          <button
            type="button"
            className="rounded bg-blue-600 px-3 py-1.5 text-sm text-white"
            onClick={() => onReplace(query, replacement, matchCase, false)}
          >
            Replace
          </button>
          <button
            type="button"
            className="rounded border border-zinc-300 px-3 py-1.5 text-sm"
            onClick={() => onReplace(query, replacement, matchCase, true)}
          >
            Replace all
          </button>
          <button
            type="button"
            className="rounded border border-zinc-300 px-3 py-1.5 text-sm"
            onClick={onClose}
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
}