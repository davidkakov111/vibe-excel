"use client";

interface ChartHelpDialogProps {
  open: boolean;
  onClose: () => void;
}

export function ChartHelpDialog({ open, onClose }: ChartHelpDialogProps) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/40 p-4">
      <div
        className="max-w-md rounded-lg bg-white p-5 text-zinc-900 shadow-xl"
        role="dialog"
        aria-labelledby="chart-help-title"
      >
        <h2 id="chart-help-title" className="text-lg font-semibold text-[#217346]">
          How to add a chart
        </h2>
        <ol className="mt-3 list-decimal space-y-2 pl-5 text-sm">
          <li>
            Type a <strong>header row</strong> (e.g. Month in A1, Sales in B1).
          </li>
          <li>
            Enter <strong>data underneath</strong> (Jan–Mar in column A, numbers in
            column B).
          </li>
          <li>
            <strong>Select the whole block</strong> including headers (e.g. A1:B4).
          </li>
          <li>
            On the green ribbon, click <strong>Bar chart</strong>, Line, Pie, or
            Area.
          </li>
          <li>Scroll below the grid to see the chart preview.</li>
        </ol>
        <button
          type="button"
          className="mt-4 w-full rounded bg-[#217346] px-3 py-2 text-sm font-medium text-white hover:bg-[#185c37]"
          onClick={onClose}
        >
          Got it
        </button>
      </div>
    </div>
  );
}