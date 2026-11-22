import { describe, expect, it } from "vitest";
import {
  applyTint,
  DEFAULT_OFFICE_THEME_HEX,
  hexFromExcelColor,
  resolveThemeColor,
} from "@/lib/excel/theme-colors";

describe("theme-colors", () => {
  it("resolves theme index without tint", () => {
    expect(resolveThemeColor(4)).toBe("#4F81BD");
  });

  it("lightens with positive tint", () => {
    const tinted = applyTint("#4F81BD", 0.25);
    expect(tinted).not.toBe("#4F81BD");
    expect(tinted.startsWith("#")).toBe(true);
  });

  it("prefers argb over theme", () => {
    expect(hexFromExcelColor({ argb: "FFFF0000", theme: 4 }, [])).toBe("#FF0000");
  });

  it("falls back to theme when argb missing", () => {
    expect(hexFromExcelColor({ theme: 5 }, DEFAULT_OFFICE_THEME_HEX)).toBe(
      "#C0504D",
    );
  });
});