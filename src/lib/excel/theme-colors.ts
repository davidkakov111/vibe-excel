/** Default Office theme color palette (theme indices 0–11). */
export const DEFAULT_OFFICE_THEME_HEX = [
  "#FFFFFF", // 0 lt1
  "#000000", // 1 dk1
  "#1F497D", // 2 dk2
  "#EEECE1", // 3 lt2
  "#4F81BD", // 4 accent1
  "#C0504D", // 5 accent2
  "#9BBB59", // 6 accent3
  "#8064A2", // 7 accent4
  "#4BACC6", // 8 accent5
  "#F79646", // 9 accent6
  "#0000FF", // 10 hlink
  "#800080", // 11 folHlink
] as const;

const THEME_ORDER = [
  "lt1",
  "dk1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
] as const;

export function parseThemePaletteFromWorkbookXml(themeXml?: string): string[] {
  if (!themeXml) return [...DEFAULT_OFFICE_THEME_HEX];
  const byName: Record<string, string> = {};
  const scheme = themeXml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/);
  if (!scheme) return [...DEFAULT_OFFICE_THEME_HEX];
  const block = scheme[0]!;
  for (const name of THEME_ORDER) {
    const segment = block.match(
      new RegExp(`<a:${name}>[\\s\\S]*?<\\/a:${name}>`, "i"),
    );
    if (!segment) continue;
    const srgb = segment[0].match(/<a:srgbClr val="([0-9A-Fa-f]{6})"/);
    const sys = segment[0].match(/lastClr="([0-9A-Fa-f]{6})"/);
    const hex = srgb?.[1] ?? sys?.[1];
    if (hex) byName[name] = `#${hex.toUpperCase()}`;
  }
  return THEME_ORDER.map((name, i) => byName[name] ?? DEFAULT_OFFICE_THEME_HEX[i]!);
}

function hexToRgb(hex: string): [number, number, number] {
  const n = hex.replace("#", "");
  return [
    parseInt(n.slice(0, 2), 16),
    parseInt(n.slice(2, 4), 16),
    parseInt(n.slice(4, 6), 16),
  ];
}

function rgbToHex(r: number, g: number, b: number): string {
  const clamp = (v: number) => Math.max(0, Math.min(255, Math.round(v)));
  const h = (c: number) => clamp(c).toString(16).padStart(2, "0");
  return `#${h(r)}${h(g)}${h(b)}`.toUpperCase();
}

/** Apply Excel tint (-1 … 1) to an RGB hex color. */
export function applyTint(hex: string, tint?: number): string {
  if (tint === undefined || tint === 0) return hex.toUpperCase();
  let [r, g, b] = hexToRgb(hex);
  if (tint < 0) {
    const factor = 1 + tint;
    r *= factor;
    g *= factor;
    b *= factor;
  } else {
    r = r * (1 - tint) + 255 * tint;
    g = g * (1 - tint) + 255 * tint;
    b = b * (1 - tint) + 255 * tint;
  }
  return rgbToHex(r, g, b);
}

export function resolveThemeColor(
  themeIndex: number,
  tint?: number,
  palette: readonly string[] = DEFAULT_OFFICE_THEME_HEX,
): string | undefined {
  const base = palette[themeIndex];
  if (!base) return undefined;
  return applyTint(base, tint);
}

export interface ExcelColorLike {
  argb?: string;
  theme?: number;
  tint?: number;
}

export function hexFromExcelColor(
  color: ExcelColorLike | undefined,
  palette: readonly string[],
): string | undefined {
  if (!color) return undefined;
  if (color.argb) {
    const slice = color.argb.length === 8 ? color.argb.slice(2) : color.argb;
    return `#${slice.toUpperCase()}`;
  }
  if (color.theme !== undefined) {
    return resolveThemeColor(color.theme, color.tint, palette);
  }
  return undefined;
}