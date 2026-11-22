"use client";

import { useMemo, useState } from "react";
import type { GridSelection } from "@/lib/workbook/selection";
import type { ConditionalFormattingSet } from "@/lib/workbook/types";
import {
  buildConditionalRule,
  removeConditionalFormatSet,
  reorderConditionalFormatRules,
  selectionToA1,
  upsertConditionalFormatSet,
  type CfPreset,
} from "@/lib/workbook/cf-rule-build";

interface ConditionalFormatDialogProps {
  open: boolean;
  onClose: () => void;
  selection: GridSelection;
  sets: ConditionalFormattingSet[];
  onApply: (sets: ConditionalFormattingSet[]) => void;
  readOnly?: boolean;
}

export function ConditionalFormatDialog({
  open,
  onClose,
  selection,
  sets,
  onApply,
  readOnly,
}: ConditionalFormatDialogProps) {
  const ref = useMemo(() => selectionToA1(selection), [selection]);
  const block = sets.find((s) => s.ref === ref);
  const [preset, setPreset] = useState<CfPreset>("greaterThan");
  const [value, setValue] = useState("0");
  const [formula, setFormula] = useState("=MOD(ROW(),2)=0");
  const [fillColor, setFillColor] = useState("#fff2cc");
  const [fontColor, setFontColor] = useState("#000000");
  const [bold, setBold] = useState(false);

  if (!open) return null;

  const addRule = () => {
    if (readOnly) return;
    const priority = (block?.rules.length ?? 0) + 1;
    const rule = buildConditionalRule({
      preset,
      value: preset === "formula" ? undefined : value,
      formula: preset === "formula" ? formula : undefined,
      fillColor,
      fontColor,
      bold,
      priority,
    });
    const next = upsertConditionalFormatSet(sets, ref, [...(block?.rules ?? []), rule]);
    onApply(next);
  };

  const deleteRule = (index: number) => {
    if (readOnly) return;
    const rules = (block?.rules ?? []).filter((_, i) => i !== index);
    const next =
      rules.length === 0
        ? removeConditionalFormatSet(sets, ref)
        : upsertConditionalFormatSet(sets, ref, rules);
    onApply(next);
  };

  const moveRule = (index: number, direction: "up" | "down") => {
    if (readOnly) return;
    onApply(reorderConditionalFormatRules(sets, ref, index, direction));
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4"
      role="dialog"
      aria-labelledby="cf-title"
    >
      <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-lg bg-white p-4 shadow-xl">
        <div className="mb-3 flex items-center justify-between gap-2">
          <h2 id="cf-title" className="text-lg font-semibold text-zinc-900">
            Conditional formatting
          </h2>
          <button
            type="button"
            className="rounded px-2 py-1 text-sm text-zinc-600 hover:bg-zinc-100"
            onClick={onClose}
          >
            Close
          </button>
        </div>
        <p className="mb-3 text-sm text-zinc-600">
          Range <span className="font-mono font-medium">{ref}</span> — rules run
          in list order (first match can stack styles; lower priority number
          first).
        </p>

        <div className="mb-4 space-y-2 rounded border border-zinc-200 p-3">
          <label className="block text-xs font-medium text-zinc-700">Rule type</label>
          <select
            className="w-full rounded border border-zinc-300 px-2 py-1 text-sm"
            value={preset}
            onChange={(e) => setPreset(e.target.value as CfPreset)}
            disabled={readOnly}
          >
            <option value="greaterThan">Cell value — greater than</option>
            <option value="lessThan">Cell value — less than</option>
            <option value="equal">Cell value — equal to</option>
            <option value="textContains">Text contains</option>
            <option value="duplicate">Duplicate values (in rule range)</option>
            <option value="formula">Formula</option>
          </select>
          {preset === "formula" ? (
            <input
              className="w-full rounded border border-zinc-300 px-2 py-1 font-mono text-sm"
              value={formula}
              onChange={(e) => setFormula(e.target.value)}
              disabled={readOnly}
              aria-label="Formula"
            />
          ) : preset === "duplicate" ? (
            <p className="text-xs text-zinc-500">
              Highlights cells whose value appears more than once inside this
              rule range.
            </p>
          ) : (
            <input
              className="w-full rounded border border-zinc-300 px-2 py-1 text-sm"
              value={value}
              onChange={(e) => setValue(e.target.value)}
              disabled={readOnly}
              aria-label="Comparison value"
            />
          )}
          <div className="flex flex-wrap gap-3 pt-1">
            <label className="flex items-center gap-1 text-xs">
              Fill
              <input
                type="color"
                value={fillColor}
                onChange={(e) => setFillColor(e.target.value)}
                disabled={readOnly}
              />
            </label>
            <label className="flex items-center gap-1 text-xs">
              Text
              <input
                type="color"
                value={fontColor}
                onChange={(e) => setFontColor(e.target.value)}
                disabled={readOnly}
              />
            </label>
            <label className="flex items-center gap-1 text-xs">
              <input
                type="checkbox"
                checked={bold}
                onChange={(e) => setBold(e.target.checked)}
                disabled={readOnly}
              />
              Bold
            </label>
          </div>
          <button
            type="button"
            className="rounded bg-[#217346] px-3 py-1.5 text-sm text-white disabled:opacity-50"
            onClick={addRule}
            disabled={readOnly}
          >
            Add rule to range
          </button>
        </div>

        <h3 className="mb-2 text-sm font-medium text-zinc-800">Rules on this range</h3>
        <ul className="space-y-2">
          {(block?.rules ?? []).map((rule, index) => (
            <li
              key={index}
              className="flex items-center justify-between gap-2 rounded border border-zinc-200 px-2 py-1 text-xs"
            >
              <span className="truncate font-mono">
                #{index + 1} {String(rule.type)} {String(rule.operator ?? "")}{" "}
                {String(rule.text ?? rule.formulae ?? "")}
              </span>
              <span className="flex shrink-0 gap-1">
                <IconBtn title="Move up" onClick={() => moveRule(index, "up")} disabled={readOnly}>
                  ↑
                </IconBtn>
                <IconBtn title="Move down" onClick={() => moveRule(index, "down")} disabled={readOnly}>
                  ↓
                </IconBtn>
                <IconBtn title="Delete" onClick={() => deleteRule(index)} disabled={readOnly}>
                  ✕
                </IconBtn>
              </span>
            </li>
          ))}
          {!block?.rules.length ? (
            <li className="text-xs text-zinc-500">No rules for this range yet.</li>
          ) : null}
        </ul>
      </div>
    </div>
  );
}

function IconBtn({
  children,
  onClick,
  title,
  disabled,
}: {
  children: React.ReactNode;
  onClick: () => void;
  title: string;
  disabled?: boolean;
}) {
  return (
    <button
      type="button"
      title={title}
      disabled={disabled}
      onClick={onClick}
      className="rounded border border-zinc-300 px-1.5 py-0.5 hover:bg-zinc-50 disabled:opacity-40"
    >
      {children}
    </button>
  );
}