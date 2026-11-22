import { rangeToA1 } from "@/lib/excel/a1-range";
import type { ChartPoint } from "@/lib/workbook/chart-data";
import type { ChartSeries } from "@/lib/workbook/types";
import zlib from "node:zlib";

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    c = CRC_TABLE[(c ^ buf[i]!) & 0xff]! ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Buffer): Buffer {
  const typeBuf = Buffer.from(type, "ascii");
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const crcBuf = Buffer.concat([typeBuf, data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(crcBuf), 0);
  return Buffer.concat([len, typeBuf, data, crc]);
}

function encodePng(rgba: Uint8Array, width: number, height: number): Buffer {
  const rowSize = width * 4 + 1;
  const raw = Buffer.alloc(rowSize * height);
  for (let y = 0; y < height; y++) {
    raw[y * rowSize] = 0;
    for (let x = 0; x < width; x++) {
      const src = (y * width + x) * 4;
      const dst = y * rowSize + 1 + x * 4;
      raw[dst] = rgba[src]!;
      raw[dst + 1] = rgba[src + 1]!;
      raw[dst + 2] = rgba[src + 2]!;
      raw[dst + 3] = rgba[src + 3]!;
    }
  }
  const compressed = zlib.deflateSync(raw, { level: 6 });
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", compressed),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

function setPixel(
  buf: Uint8Array,
  width: number,
  x: number,
  y: number,
  r: number,
  g: number,
  b: number,
  a = 255,
) {
  if (x < 0 || y < 0 || x >= width) return;
  const height = buf.length / (width * 4);
  if (y >= height) return;
  const i = (y * width + x) * 4;
  if (a >= 255) {
    buf[i] = r;
    buf[i + 1] = g;
    buf[i + 2] = b;
    buf[i + 3] = 255;
    return;
  }
  const alpha = a / 255;
  buf[i] = Math.round(buf[i]! * (1 - alpha) + r * alpha);
  buf[i + 1] = Math.round(buf[i + 1]! * (1 - alpha) + g * alpha);
  buf[i + 2] = Math.round(buf[i + 2]! * (1 - alpha) + b * alpha);
  buf[i + 3] = 255;
}

function fillRect(
  buf: Uint8Array,
  width: number,
  height: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  r: number,
  g: number,
  b: number,
) {
  for (let y = y0; y <= y1 && y < height; y++) {
    for (let x = x0; x <= x1 && x < width; x++) {
      setPixel(buf, width, x, y, r, g, b);
    }
  }
}

function drawGrid(
  buf: Uint8Array,
  width: number,
  height: number,
  left: number,
  right: number,
  top: number,
  bottom: number,
) {
  for (let y = top; y <= bottom; y += 24) {
    for (let x = left; x <= right; x++) {
      setPixel(buf, width, x, y, 226, 232, 240);
    }
  }
  for (let x = left; x <= right; x += 40) {
    for (let y = top; y <= bottom; y++) {
      setPixel(buf, width, x, y, 226, 232, 240);
    }
  }
}

function drawChar5x7(
  buf: Uint8Array,
  w: number,
  x: number,
  y: number,
  ch: string,
  r: number,
  g: number,
  b: number,
) {
  const code = ch.charCodeAt(0);
  if (code < 32 || code > 126) return;
  const glyph = FONT_5X7[code - 32];
  if (!glyph) return;
  for (let row = 0; row < 7; row++) {
    const bits = glyph[row] ?? 0;
    for (let col = 0; col < 5; col++) {
      if (bits & (1 << (4 - col))) {
        setPixel(buf, w, x + col, y + row, r, g, b);
      }
    }
  }
}

function drawText(
  buf: Uint8Array,
  w: number,
  x: number,
  y: number,
  text: string,
  r: number,
  g: number,
  b: number,
  maxChars = 24,
) {
  const label = text.slice(0, maxChars);
  let cx = x;
  for (const ch of label) {
    drawChar5x7(buf, w, cx, y, ch, r, g, b);
    cx += ch === " " ? 4 : 6;
  }
}

/** Minimal 5×7 glyphs (A–Z, 0–9, punctuation). */
const FONT_5X7: number[][] = [
  [0x00, 0x00, 0x00, 0x00, 0x00], // space
  [0x00, 0x00, 0x5f, 0x00, 0x00], // !
  [0x00, 0x07, 0x00, 0x07, 0x00], // "
  [0x14, 0x7f, 0x14, 0x7f, 0x14], // #
  [0x24, 0x2a, 0x7f, 0x2a, 0x12], // $
  [0x23, 0x13, 0x08, 0x64, 0x62], // %
  [0x36, 0x49, 0x55, 0x22, 0x50], // &
  [0x00, 0x05, 0x03, 0x00, 0x00], // '
  [0x00, 0x1c, 0x22, 0x41, 0x00], // (
  [0x00, 0x41, 0x22, 0x1c, 0x00], // )
  [0x14, 0x08, 0x3e, 0x08, 0x14], // *
  [0x08, 0x08, 0x3e, 0x08, 0x08], // +
  [0x00, 0x50, 0x30, 0x00, 0x00], // ,
  [0x08, 0x08, 0x08, 0x08, 0x08], // -
  [0x00, 0x60, 0x60, 0x00, 0x00], // .
  [0x20, 0x10, 0x08, 0x04, 0x02], // /
  [0x3e, 0x51, 0x49, 0x45, 0x3e], // 0
  [0x00, 0x42, 0x7f, 0x40, 0x00], // 1
  [0x42, 0x61, 0x51, 0x49, 0x46], // 2
  [0x21, 0x41, 0x45, 0x4b, 0x31], // 3
  [0x18, 0x14, 0x12, 0x7f, 0x10], // 4
  [0x27, 0x45, 0x45, 0x45, 0x39], // 5
  [0x3c, 0x4a, 0x49, 0x49, 0x30], // 6
  [0x01, 0x71, 0x09, 0x05, 0x03], // 7
  [0x36, 0x49, 0x49, 0x49, 0x36], // 8
  [0x06, 0x49, 0x49, 0x29, 0x1e], // 9
  [0x00, 0x36, 0x36, 0x00, 0x00], // :
  [0x00, 0x56, 0x36, 0x00, 0x00], // ;
  [0x08, 0x14, 0x22, 0x41, 0x00], // <
  [0x14, 0x14, 0x14, 0x14, 0x14], // =
  [0x00, 0x41, 0x22, 0x14, 0x08], // >
  [0x02, 0x01, 0x51, 0x09, 0x06], // ?
  [0x32, 0x49, 0x79, 0x41, 0x3e], // @
  [0x7e, 0x11, 0x11, 0x11, 0x7e], // A
  [0x7f, 0x49, 0x49, 0x49, 0x36], // B
  [0x3e, 0x41, 0x41, 0x41, 0x22], // C
  [0x7f, 0x41, 0x41, 0x22, 0x1c], // D
  [0x7f, 0x49, 0x49, 0x49, 0x41], // E
  [0x7f, 0x09, 0x09, 0x09, 0x01], // F
  [0x3e, 0x41, 0x49, 0x49, 0x7a], // G
  [0x7f, 0x08, 0x08, 0x08, 0x7f], // H
  [0x00, 0x41, 0x7f, 0x41, 0x00], // I
  [0x20, 0x40, 0x41, 0x3f, 0x01], // J
  [0x7f, 0x08, 0x14, 0x22, 0x41], // K
  [0x7f, 0x40, 0x40, 0x40, 0x40], // L
  [0x7f, 0x02, 0x0c, 0x02, 0x7f], // M
  [0x7f, 0x04, 0x08, 0x10, 0x7f], // N
  [0x3e, 0x41, 0x41, 0x41, 0x3e], // O
  [0x7f, 0x09, 0x09, 0x09, 0x06], // P
  [0x3e, 0x41, 0x51, 0x21, 0x5e], // Q
  [0x7f, 0x09, 0x19, 0x29, 0x46], // R
  [0x46, 0x49, 0x49, 0x49, 0x31], // S
  [0x01, 0x01, 0x7f, 0x01, 0x01], // T
  [0x3f, 0x40, 0x40, 0x40, 0x3f], // U
  [0x1f, 0x20, 0x40, 0x20, 0x1f], // V
  [0x3f, 0x40, 0x38, 0x40, 0x3f], // W
  [0x63, 0x14, 0x08, 0x14, 0x63], // X
  [0x07, 0x08, 0x70, 0x08, 0x07], // Y
  [0x61, 0x51, 0x49, 0x45, 0x43], // Z
];

function drawTitleBand(
  buf: Uint8Array,
  width: number,
  height: number,
  title: string,
) {
  fillRect(buf, width, height, 0, 0, width - 1, 44, 248, 250, 252);
  fillRect(buf, width, height, 0, 44, width - 1, 46, 226, 232, 240);
  drawText(buf, width, 16, 16, title.toUpperCase(), 51, 65, 85, 48);
}

function formatTick(value: number): string {
  if (Math.abs(value) >= 1000) return `${Math.round(value / 100) / 10}k`;
  if (Number.isInteger(value)) return String(value);
  return value.toFixed(1);
}

function drawCartesianAxes(
  buf: Uint8Array,
  width: number,
  plotLeft: number,
  plotRight: number,
  plotTop: number,
  plotBottom: number,
  data: ChartPoint[],
  minVal: number,
  maxVal: number,
  chartType: "bar" | "line" | "area",
) {
  const axisColor = { r: 71, g: 85, b: 105 };
  for (let y = plotTop; y <= plotBottom; y++) {
    setPixel(buf, width, plotLeft, y, axisColor.r, axisColor.g, axisColor.b);
  }
  for (let x = plotLeft; x <= plotRight; x++) {
    setPixel(buf, width, x, plotBottom, axisColor.r, axisColor.g, axisColor.b);
  }

  const ticks = 4;
  for (let t = 0; t <= ticks; t++) {
    const fraction = t / ticks;
    const value = maxVal - (maxVal - minVal) * fraction;
    const y = Math.round(plotTop + fraction * (plotBottom - plotTop));
    for (let x = plotLeft - 3; x < plotLeft; x++) {
      setPixel(buf, width, x, y, axisColor.r, axisColor.g, axisColor.b);
    }
    drawText(
      buf,
      width,
      plotLeft - 44,
      y - 3,
      formatTick(value),
      100,
      116,
      139,
      6,
    );
  }

  if (chartType === "bar") {
    const barGap = 10;
    const barWidth = Math.max(
      14,
      Math.floor(
        (plotRight - plotLeft - barGap * (data.length + 1)) / Math.max(data.length, 1),
      ),
    );
    let x = plotLeft + barGap;
    data.forEach((d) => {
      const center = x + Math.floor(barWidth / 2) - 12;
      drawText(buf, width, center, plotBottom + 8, d.name, 71, 85, 105, 10);
      x += barWidth + barGap;
    });
    return;
  }

  const step = (plotRight - plotLeft) / Math.max(data.length - 1, 1);
  data.forEach((d, i) => {
    const px = Math.round(plotLeft + i * step) - 12;
    drawText(buf, width, px, plotBottom + 8, d.name, 71, 85, 105, 10);
  });
}

const BAR_COLORS = [
  [37, 99, 235],
  [22, 163, 74],
  [220, 38, 38],
  [202, 138, 4],
  [124, 58, 237],
];

function drawPie(
  buf: Uint8Array,
  width: number,
  height: number,
  cx: number,
  cy: number,
  radius: number,
  data: ChartPoint[],
) {
  const slices = data.filter((d) => d.value !== 0);
  const total = slices.reduce((sum, d) => sum + Math.abs(d.value), 0) || 1;
  let angle = -Math.PI / 2;
  slices.forEach((d, i) => {
    const slice = (Math.abs(d.value) / total) * Math.PI * 2;
    const color = BAR_COLORS[i % BAR_COLORS.length]!;
    const steps = Math.max(24, Math.ceil(slice * 40));
    for (let s = 0; s <= steps; s++) {
      const a0 = angle + (slice * s) / steps;
      const a1 = angle + (slice * (s + 1)) / steps;
      const x0 = Math.round(cx + Math.cos(a0) * radius);
      const y0 = Math.round(cy + Math.sin(a0) * radius);
      const x1 = Math.round(cx + Math.cos(a1) * radius);
      const y1 = Math.round(cy + Math.sin(a1) * radius);
      fillTriangle(buf, width, height, cx, cy, x0, y0, x1, y1, color[0]!, color[1]!, color[2]!);
    }
    angle += slice;
  });
}

function fillTriangle(
  buf: Uint8Array,
  width: number,
  height: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  r: number,
  g: number,
  b: number,
) {
  const minX = Math.max(0, Math.min(x0, x1, x2));
  const maxX = Math.min(width - 1, Math.max(x0, x1, x2));
  const minY = Math.max(0, Math.min(y0, y1, y2));
  const maxY = Math.min(height - 1, Math.max(y0, y1, y2));
  for (let y = minY; y <= maxY; y++) {
    for (let x = minX; x <= maxX; x++) {
      if (pointInTriangle(x, y, x0, y0, x1, y1, x2, y2)) {
        setPixel(buf, width, x, y, r, g, b);
      }
    }
  }
}

function pointInTriangle(
  px: number,
  py: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): boolean {
  const d =
    (y1 - y2) * (x0 - x2) + (x2 - x1) * (y0 - y2);
  if (d === 0) return false;
  const a =
    ((y1 - y2) * (px - x2) + (x2 - x1) * (py - y2)) / d;
  const b =
    ((y2 - y0) * (px - x2) + (x0 - x2) * (py - y2)) / d;
  const c = 1 - a - b;
  return a >= 0 && b >= 0 && c >= 0;
}

export function renderChartPng(
  chart: ChartSeries,
  data: ChartPoint[],
  width = 720,
  height = 420,
): Buffer | null {
  if (!data.length) return null;
  const rgba = new Uint8Array(width * height * 4);
  fillRect(rgba, width, height, 0, 0, width - 1, height - 1, 255, 255, 255);
  drawTitleBand(rgba, width, height, chart.title);

  const plotLeft = 56;
  const plotRight = width - 24;
  const plotTop = 64;
  const plotBottom = height - 48;
  const values = data.map((d) => d.value);
  const minVal = Math.min(0, ...values);
  const maxVal = Math.max(...values, 1);
  const valueSpan = maxVal - minVal || 1;
  const valueY = (v: number) =>
    plotBottom - Math.round(((plotBottom - plotTop) * (v - minVal)) / valueSpan);

  if (chart.type === "pie") {
    const cx = Math.round((plotLeft + plotRight) / 2);
    const cy = Math.round((plotTop + plotBottom) / 2);
    const radius = Math.min(plotRight - plotLeft, plotBottom - plotTop) / 2 - 12;
    drawPie(rgba, width, height, cx, cy, radius, data);
    return encodePng(rgba, width, height);
  }

  drawGrid(rgba, width, height, plotLeft, plotRight, plotTop, plotBottom);
  const axisChartType =
    chart.type === "bar" ? "bar" : chart.type === "line" ? "line" : "area";
  drawCartesianAxes(
    rgba,
    width,
    plotLeft,
    plotRight,
    plotTop,
    plotBottom,
    data,
    minVal,
    maxVal,
    axisChartType,
  );

  if (chart.type === "line" || chart.type === "area") {
    const step = (plotRight - plotLeft) / Math.max(data.length - 1, 1);
    let prevX = plotLeft;
    let prevY = valueY(values[0] ?? 0);
    for (let i = 0; i < data.length; i++) {
      const d = data[i]!;
      const px = Math.round(plotLeft + i * step);
      const py = valueY(d.value);
      if (i > 0) {
        const minX = Math.min(prevX, px);
        const maxX = Math.max(prevX, px);
        for (let x = minX; x <= maxX; x++) {
          const t = maxX === minX ? 0 : (x - minX) / (maxX - minX);
          const yLine = Math.round(prevY * (1 - t) + py * t);
          setPixel(rgba, width, x, yLine, 37, 99, 235);
          if (chart.type === "area") {
            for (let yy = yLine; yy <= plotBottom; yy++) {
              setPixel(rgba, width, x, yy, 147, 197, 253, 140);
            }
          }
        }
      }
      setPixel(rgba, width, px, py, 30, 64, 175);
      prevX = px;
      prevY = py;
    }
  } else {
    const barGap = 10;
    const barWidth = Math.max(
      14,
      Math.floor(
        (plotRight - plotLeft - barGap * (data.length + 1)) / data.length,
      ),
    );
    let x = plotLeft + barGap;
    data.forEach((d, i) => {
      const topY = valueY(Math.max(d.value, minVal));
      const bottomY = valueY(Math.min(0, minVal));
      const barTop = Math.min(topY, bottomY);
      const barBottom = Math.max(topY, bottomY);
      const color = BAR_COLORS[i % BAR_COLORS.length]!;
      fillRect(
        rgba,
        width,
        height,
        x,
        barTop,
        x + barWidth - 1,
        barBottom,
        color[0]!,
        color[1]!,
        color[2]!,
      );
      x += barWidth + barGap;
    });
  }

  return encodePng(rgba, width, height);
}

export function chartAnchorRef(anchor: ChartSeries["anchor"]): string {
  return rangeToA1(anchor);
}