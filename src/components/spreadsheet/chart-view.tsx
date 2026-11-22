"use client";

import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import type { ChartSeries } from "@/lib/workbook/types";
import type { ChartPoint } from "@/lib/workbook/chart-data";

const COLORS = ["#2563eb", "#16a34a", "#dc2626", "#ca8a04", "#7c3aed"];

function pieChartRows(data: ChartPoint[]) {
  return data
    .filter((d) => d.value !== 0)
    .map((d) => ({
      name: d.name,
      value: d.value,
      sliceValue: Math.abs(d.value),
    }));
}

export function ChartView({
  chart,
  data,
  title,
}: {
  chart: ChartSeries;
  data: ChartPoint[];
  title?: string;
}) {
  const pieData = chart.type === "pie" ? pieChartRows(data) : data;

  return (
    <div className="flex h-full min-h-0 w-full flex-col">
      {title ? (
        <p className="shrink-0 truncate px-1 pt-0.5 text-[10px] font-semibold text-zinc-700">
          {title}
        </p>
      ) : null}
      <div className="min-h-0 flex-1">
        <ResponsiveContainer width="100%" height="100%">
          {chart.type === "bar" ? (
            <BarChart data={data}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" tick={{ fontSize: 9 }} />
              <YAxis tick={{ fontSize: 9 }} width={28} />
              <Tooltip />
              <Bar dataKey="value" fill="#2563eb" />
            </BarChart>
          ) : chart.type === "line" ? (
            <LineChart data={data}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" tick={{ fontSize: 9 }} />
              <YAxis tick={{ fontSize: 9 }} width={28} />
              <Tooltip />
              <Line type="monotone" dataKey="value" stroke="#2563eb" strokeWidth={2} />
            </LineChart>
          ) : chart.type === "area" ? (
            <AreaChart data={data}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="name" tick={{ fontSize: 9 }} />
              <YAxis tick={{ fontSize: 9 }} width={28} />
              <Tooltip />
              <Area type="monotone" dataKey="value" stroke="#2563eb" fill="#93c5fd" />
            </AreaChart>
          ) : pieData.length === 0 ? (
            <div className="flex h-full items-center justify-center text-[10px] text-zinc-500">
              No numeric values
            </div>
          ) : (
            <PieChart>
              <Pie
                data={pieData}
                dataKey="sliceValue"
                nameKey="name"
                label={{ fontSize: 9 }}
              >
                {pieData.map((_, index) => (
                  <Cell key={index} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip formatter={(_v, _n, item) => item.payload?.value} />
              <Legend wrapperStyle={{ fontSize: 10 }} />
            </PieChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  );
}