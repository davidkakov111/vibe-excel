"use client";

import { useState } from "react";
import type { CellCommentThread } from "@/lib/workbook/comment";
import { threadAuthorLabel, threadPreviewText } from "@/lib/workbook/comment";

export function CellCommentPanel({
  row,
  col,
  thread,
  readOnly,
  currentUserId,
  onStart,
  onReply,
  onResolve,
  onReopen,
  onDelete,
  onClose,
}: {
  row: number;
  col: number;
  thread: CellCommentThread | null;
  readOnly: boolean;
  currentUserId: string;
  onStart: (text: string) => void;
  onReply: (text: string) => void;
  onResolve: () => void;
  onReopen: () => void;
  onDelete: () => void;
  onClose: () => void;
}) {
  const [draft, setDraft] = useState("");
  const resolved = thread?.status === "resolved";

  if (!thread) {
    return (
      <div
        className="absolute z-[40] w-72 rounded-md border border-zinc-300 bg-white p-3 text-zinc-900 shadow-xl"
        role="dialog"
      >
        <p className="mb-2 text-xs font-semibold">New comment</p>
        <textarea
          className="mb-2 w-full resize-none rounded border border-zinc-300 px-2 py-1 text-xs"
          rows={4}
          value={draft}
          disabled={readOnly}
          onChange={(e) => setDraft(e.target.value)}
        />
        <div className="flex gap-2">
          <button
            type="button"
            className="rounded bg-[#217346] px-2 py-1 text-[11px] text-white disabled:opacity-40"
            disabled={readOnly || !draft.trim()}
            onClick={() => {
              onStart(draft);
              setDraft("");
            }}
          >
            Post
          </button>
          <button
            type="button"
            className="rounded border border-zinc-300 px-2 py-1 text-[11px]"
            onClick={onClose}
          >
            Cancel
          </button>
        </div>
      </div>
    );
  }

  return (
    <div
      className={`absolute z-[40] w-72 rounded-md border border-zinc-300 bg-white text-zinc-900 shadow-xl ${
        resolved ? "opacity-85" : ""
      }`}
      role="dialog"
      aria-label={`Comment at row ${row + 1}, column ${col + 1}`}
    >
      <div className="flex items-center justify-between border-b border-zinc-200 bg-[#fff8e6] px-3 py-2">
        <div>
          <p className="text-xs font-semibold text-zinc-800">
            {threadAuthorLabel(thread)}
            {resolved ? (
              <span className="ml-2 rounded bg-zinc-200 px-1.5 py-0.5 text-[10px] font-medium text-zinc-600">
                Resolved
              </span>
            ) : null}
          </p>
          <p className="truncate text-[10px] text-zinc-500">
            {threadPreviewText(thread).slice(0, 60)}
          </p>
        </div>
        <button
          type="button"
          className="text-zinc-500 hover:text-zinc-800"
          onClick={onClose}
          aria-label="Close"
        >
          ×
        </button>
      </div>

      <div className="max-h-52 overflow-y-auto px-3 py-2 text-xs">
        <ul className="space-y-3">
          {thread.messages.map((msg) => (
            <li key={msg.id} className={resolved ? "opacity-70" : ""}>
              <p className="font-semibold text-zinc-800">
                {msg.authorName}
                {msg.userId === currentUserId ? (
                  <span className="ml-1 text-[10px] font-normal text-zinc-500">
                    (you)
                  </span>
                ) : null}
              </p>
              <p className="whitespace-pre-wrap text-zinc-700">{msg.text}</p>
              <p className="text-[10px] text-zinc-400">
                {new Date(msg.createdAt).toLocaleString()}
              </p>
            </li>
          ))}
        </ul>
      </div>

      {!readOnly ? (
        <div className="space-y-2 border-t border-zinc-200 px-3 py-2">
          <textarea
            className="w-full resize-none rounded border border-zinc-300 px-2 py-1 text-xs outline-none focus:border-[#217346]"
            rows={3}
            placeholder={resolved ? "Reopen to reply…" : "Reply…"}
            value={draft}
            disabled={resolved}
            onChange={(e) => setDraft(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                e.preventDefault();
                if (!draft.trim() || resolved) return;
                onReply(draft);
                setDraft("");
              }
            }}
          />
          <div className="flex flex-wrap gap-1">
            <button
              type="button"
              className="rounded bg-[#217346] px-2 py-1 text-[11px] font-medium text-white disabled:opacity-40"
              disabled={resolved || !draft.trim()}
              onClick={() => {
                onReply(draft);
                setDraft("");
              }}
            >
              Reply
            </button>
            {resolved ? (
              <button
                type="button"
                className="rounded border border-zinc-300 px-2 py-1 text-[11px]"
                onClick={onReopen}
              >
                Reopen
              </button>
            ) : (
              <button
                type="button"
                className="rounded border border-zinc-300 px-2 py-1 text-[11px]"
                onClick={onResolve}
              >
                Resolve
              </button>
            )}
            <button
              type="button"
              className="rounded border border-red-200 px-2 py-1 text-[11px] text-red-700"
              onClick={onDelete}
            >
              Delete
            </button>
          </div>
        </div>
      ) : null}
    </div>
  );
}