import { describe, it, expect } from "vitest";
import { resolveColor, ColorContext } from "../src/domain/pptx/color";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const fill = (inner: string) =>
  new DOMParser().parseFromString(`<a:solidFill xmlns:a="${A}">${inner}</a:solidFill>`, "application/xml").documentElement;

const ctx: ColorContext = {
  scheme: { dk1: "#000000", lt1: "#FFFFFF", accent1: "#4F81BD" },
  map: { tx1: "dk1", bg1: "lt1" },
};

describe("resolveColor", () => {
  it("reads srgbClr", () => expect(resolveColor(fill('<a:srgbClr val="ff0000"/>'), ctx)).toBe("#FF0000"));
  it("maps schemeClr through clrMap", () => expect(resolveColor(fill('<a:schemeClr val="tx1"/>'), ctx)).toBe("#000000"));
  it("reads scheme slots directly", () => expect(resolveColor(fill('<a:schemeClr val="accent1"/>'), ctx)).toBe("#4F81BD"));
  it("uses sysClr lastClr", () => expect(resolveColor(fill('<a:sysClr val="window" lastClr="FFFFFF"/>'), ctx)).toBe("#FFFFFF"));
  it("knows a few preset colors", () => expect(resolveColor(fill('<a:prstClr val="black"/>'), ctx)).toBe("#000000"));
  it("returns null for unknown or missing colors", () => {
    expect(resolveColor(fill('<a:schemeClr val="phClr"/>'), ctx)).toBeNull();
    expect(resolveColor(null, ctx)).toBeNull();
    expect(resolveColor(fill(""), ctx)).toBeNull();
  });
  it("applies lumMod/lumOff", () => {
    // bg1 at 85% luminance = #D9D9D9 (standard PowerPoint "White, darker 15%")
    expect(resolveColor(fill('<a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr>'), ctx)).toBe("#D9D9D9");
    // tx1 lumMod 65 lumOff 35 = #595959 ("Black, lighter 35%")
    expect(
      resolveColor(fill('<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>'), ctx)
    ).toBe("#595959");
  });
  it("applies tint and shade", () => {
    expect(resolveColor(fill('<a:srgbClr val="000000"><a:tint val="75000"/></a:srgbClr>'), ctx)).toBe("#404040");
    expect(resolveColor(fill('<a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr>'), ctx)).toBe("#808080");
  });
  it("appends alpha as an 8-digit hex", () => {
    expect(resolveColor(fill('<a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr>'), ctx)).toBe("#0000FF80");
    expect(resolveColor(fill('<a:srgbClr val="0000FF"><a:alpha val="100000"/></a:srgbClr>'), ctx)).toBe("#0000FF");
  });
});
