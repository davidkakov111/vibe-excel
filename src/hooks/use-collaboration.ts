"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { CollaborationEvent } from "@/lib/collaboration/hub";
import {
  uniquePresenceByUser,
  type RemotePresence,
} from "@/lib/collaboration/presence";

export type { RemotePresence };
import {
  postOpsWithRetry,
  reconcileWorkbookWithOps,
} from "@/lib/collaboration/op-sync";
import type { GridSelection } from "@/lib/workbook/selection";
import type { WorkbookOp } from "@/lib/workbook/ops";
import type { WorkbookDocument } from "@/lib/workbook/types";

const PALETTE = [
  "#e11d48",
  "#2563eb",
  "#16a34a",
  "#ca8a04",
  "#9333ea",
  "#0d9488",
  "#ea580c",
];

function colorForUser(userId: string): string {
  let hash = 0;
  for (let i = 0; i < userId.length; i++) hash = (hash + userId.charCodeAt(i)) % 997;
  return PALETTE[hash % PALETTE.length]!;
}

export function useCollaboration(args: {
  spreadsheetId: string;
  clientId: string;
  userId: string;
  userName: string;
  enabled: boolean;
  revision: number;
  onRevision: (revision: number) => void;
  onRemoteOps: (ops: WorkbookOp[], revision: number) => void;
  /** Server workbook merged with the failed ops (keeps local intent e.g. sort). */
  onReconcile: (workbook: WorkbookDocument, revision: number) => void;
  onRemoteWorkbookRestore?: (workbook: WorkbookDocument, revision: number) => void;
  activeSheetId: string;
  selection: GridSelection;
}) {
  const {
    spreadsheetId,
    clientId,
    userId,
    userName,
    enabled,
    revision,
    onRevision,
    onRemoteOps,
    onReconcile,
    onRemoteWorkbookRestore,
    activeSheetId,
    selection,
  } = args;

  const revisionRef = useRef(revision);
  revisionRef.current = revision;
  const opChainRef = useRef<Promise<boolean>>(Promise.resolve(true));
  const onReconcileRef = useRef(onReconcile);
  onReconcileRef.current = onReconcile;

  const [othersRaw, setOthersRaw] = useState<RemotePresence[]>([]);
  const lastSeenRef = useRef<Map<string, number>>(new Map());
  const [tick, setTick] = useState(0);
  const myColor = useMemo(() => colorForUser(userId), [userId]);

  const applyPresenceEvent = useCallback((event: CollaborationEvent) => {
    if (event.type === "presence_leave") {
      lastSeenRef.current.delete(event.clientId);
      setOthersRaw((prev) => prev.filter((p) => p.clientId !== event.clientId));
      return;
    }
    if (event.type !== "presence") return;
    if (event.clientId === clientId) return;
    lastSeenRef.current.set(event.clientId, Date.now());
    setOthersRaw((prev) => {
      const next = prev.filter((p) => p.clientId !== event.clientId);
      next.push({
        clientId: event.clientId,
        userId: event.userId,
        userName: event.userName,
        color: event.color,
        sheetId: event.sheetId,
        selection: event.selection,
      });
      return next;
    });
  }, [clientId]);

  useEffect(() => {
    if (!enabled) return;
    let closed = false;
    let source: EventSource | null = null;
    let retry = 0;
    let retryTimer: ReturnType<typeof setTimeout> | undefined;

    const connect = () => {
      if (closed) return;
      source = new EventSource(`/api/spreadsheets/${spreadsheetId}/stream`);
      source.onopen = () => {
        retry = 0;
      };
      source.onmessage = (message) => {
        const event = JSON.parse(message.data) as CollaborationEvent;
        if (event.type === "revision") {
          onRevision(event.revision);
          return;
        }
        if (event.type === "ops") {
          if (event.clientId === clientId) {
            onRevision(event.revision);
            return;
          }
          onRemoteOps(event.ops as WorkbookOp[], event.revision);
          onRevision(event.revision);
          return;
        }
        if (event.type === "workbook_restore") {
          if (event.clientId === clientId) {
            onRevision(event.revision);
            return;
          }
          onRemoteWorkbookRestore?.(event.workbook, event.revision);
          onRevision(event.revision);
          return;
        }
        applyPresenceEvent(event);
      };
      source.onerror = () => {
        source?.close();
        source = null;
        if (closed) return;
        const delay = Math.min(1000 * 2 ** retry, 10_000);
        retry += 1;
        retryTimer = setTimeout(connect, delay);
      };
    };

    connect();

    return () => {
      closed = true;
      if (retryTimer) clearTimeout(retryTimer);
      source?.close();
    };
  }, [
    applyPresenceEvent,
    clientId,
    enabled,
    onRemoteOps,
    onRemoteWorkbookRestore,
    onRevision,
    spreadsheetId,
  ]);

  useEffect(() => {
    if (!enabled) return;
    const timer = setInterval(() => setTick((t) => t + 1), 1_000);
    return () => clearInterval(timer);
  }, [enabled]);

  useEffect(() => {
    if (!enabled) return;
    const payload = {
      clientId,
      userId,
      userName,
      color: myColor,
      sheetId: activeSheetId,
      selection,
    };
    const post = () => {
      void fetch(`/api/spreadsheets/${spreadsheetId}/presence`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
    };
    post();
    const timer = setInterval(post, 800);
    return () => {
      clearInterval(timer);
      const body = JSON.stringify({ clientId, leave: true });
      if (typeof navigator !== "undefined" && navigator.sendBeacon) {
        navigator.sendBeacon(
          `/api/spreadsheets/${spreadsheetId}/presence`,
          new Blob([body], { type: "application/json" }),
        );
      } else {
        void fetch(`/api/spreadsheets/${spreadsheetId}/presence`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body,
          keepalive: true,
        });
      }
    };
  }, [
    activeSheetId,
    clientId,
    enabled,
    myColor,
    selection,
    spreadsheetId,
    userId,
    userName,
  ]);

  const others = useMemo(
    () => uniquePresenceByUser(othersRaw, lastSeenRef.current, Date.now()),
    // tick drives stale pruning
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [othersRaw, tick],
  );

  const pushOps = useCallback(
    (ops: WorkbookOp[]): Promise<boolean> => {
      const task = async (): Promise<boolean> => {
        let result = await postOpsWithRetry({
          spreadsheetId,
          clientId,
          baseRevision: revisionRef.current,
          ops,
          onRevision,
        });

        if (!result.ok && result.conflict) {
          const merged = reconcileWorkbookWithOps(result.workbook, ops);
          onReconcileRef.current(merged, result.revision);
          revisionRef.current = result.revision;
          result = await postOpsWithRetry({
            spreadsheetId,
            clientId,
            baseRevision: result.revision,
            ops,
            onRevision,
            maxAttempts: 3,
          });
        }

        if (result.ok) {
          onRevision(result.revision);
          return true;
        }

        if (result.conflict) {
          const merged = reconcileWorkbookWithOps(result.workbook, ops);
          onReconcileRef.current(merged, result.revision);
          onRevision(result.revision);
          return false;
        }

        return false;
      };

      const chained = opChainRef.current.then(task, task);
      opChainRef.current = chained.catch(() => false);
      return chained;
    },
    [clientId, onRevision, spreadsheetId],
  );

  return { pushOps, others, myColor };
}