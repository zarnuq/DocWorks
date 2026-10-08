import { describe, it, expect } from "vitest";
import {
  fillCss,
  readFill,
  readLine,
  readShadow,
  themeStyles,
  writeFill,
  writeLine,
  writeShadow,
} from "../src/domain/pptx/shape-style";
import type { ColorContext } from "../src/domain/pptx/color";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const ctx: ColorContext = { scheme: { accent1: "#4472C4", dk1: "#000000", lt1: "#FFFFFF" }, map: {} };
const xml = (s: string) => new DOMParser().parseFromString(s, "application/xml").documentElement;
const spPr = (inner = "") => xml(`<p:spPr xmlns:p="${P}" xmlns:a="${A}"><a:xfrm/><a:prstGeom prst="rect"/>${inner}</p:spPr>`);
const style = (inner: string) => xml(`<p:style xmlns:p="${P}" xmlns:a="${A}">${inner}</p:style>`);
const order = (el: Element) => Array.from(el.children).map((c) => c.localName);

const theme = themeStyles(
  xml(
    `<a:theme xmlns:a="${A}"><a:themeElements><a:fmtScheme>` +
      `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
      `<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>` +
      `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
      `<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>` +
      `<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"/></a:lnStyleLst>` +
      `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>` +
      `<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>` +
      `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>` +
      `</a:fmtScheme></a:themeElements></a:theme>`
  ).ownerDocument
);

describe("readFill", () => {
  it("reads solid and gradient fills", () => {
    expect(readFill(spPr(`<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`), null, ctx, theme)).toEqual({
      type: "solid",
      color: "#FF000080",
    });
    expect(
      readFill(
        spPr(`<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill>`),
        null,
        ctx,
        theme
      )
    ).toEqual({ type: "gradient", radial: false, angle: 45, stops: [{ pos: 0, color: "#FF0000" }, { pos: 1, color: "#0000FF" }] });
  });

  it("reads noFill as no fill even when the style has one", () => {
    expect(readFill(spPr("<a:noFill/>"), style(`<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>`), ctx, theme)).toBeNull();
  });

  it("resolves theme fill styles with the reference color", () => {
    const st = style(`<a:fillRef idx="2"><a:schemeClr val="accent1"/></a:fillRef>`);
    expect(readFill(spPr(), st, ctx, theme)).toMatchObject({ type: "gradient", angle: 90, stops: [{ pos: 0 }, { pos: 1, color: "#4472C4" }] });
    expect(readFill(spPr(), style(`<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>`), ctx, theme)).toBeNull();
  });

  it("converts fills to CSS, mirroring gradients on flipped shapes", () => {
    expect(fillCss({ type: "solid", color: "#FF0000" })).toBe("#FF0000");
    const g = { type: "gradient" as const, radial: false, angle: 0, stops: [{ pos: 0, color: "#000000" }, { pos: 1, color: "#FFFFFF" }] };
    expect(fillCss(g)).toBe("linear-gradient(90deg, #000000 0%, #FFFFFF 100%)");
    expect(fillCss(g, true)).toBe("linear-gradient(270deg, #000000 0%, #FFFFFF 100%)");
  });
});

describe("readLine", () => {
  it("reads color, width, dash and arrowheads", () => {
    const ln = readLine(
      spPr(`<a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle" w="lg" len="sm"/></a:ln>`),
      null,
      ctx,
      theme
    );
    expect(ln).toEqual({ color: "#00FF00", width: 25400, dash: "dash", cap: "rnd", head: null, tail: { type: "triangle", w: "lg", len: "sm" } });
  });

  it("fills in width and color from the theme line style", () => {
    const ln = readLine(spPr(), style(`<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>`), ctx, theme);
    expect(ln).toMatchObject({ color: "#4472C4", width: 6350, dash: "solid" });
  });

  it("reads noFill as no line", () => {
    expect(readLine(spPr(`<a:ln><a:noFill/></a:ln>`), style(`<a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>`), ctx, theme)).toBeNull();
  });
});

describe("readShadow", () => {
  it("reads an explicit outer shadow", () => {
    const sh = readShadow(
      spPr(`<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`),
      null,
      ctx,
      theme
    );
    expect(sh).toEqual({ color: "#00000066", blur: 50800, dist: 38100, dir: 45 });
  });

  it("uses the theme effect style unless the shape has its own effect list", () => {
    const st = style(`<a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef>`);
    expect(readShadow(spPr(), st, ctx, theme)).toMatchObject({ dir: 90, dist: 19050 });
    expect(readShadow(spPr("<a:effectLst/>"), st, ctx, theme)).toBeNull();
  });
});

describe("writers", () => {
  it("writes fills in schema order, replacing the old fill", () => {
    const el = spPr(`<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="12700"/>`);
    writeFill(el, { type: "gradient", radial: false, angle: 90, stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 1, color: "#00000080" }] });
    expect(order(el)).toEqual(["xfrm", "prstGeom", "gradFill", "ln"]);
    expect(readFill(el, null, ctx, theme)).toEqual({
      type: "gradient",
      radial: false,
      angle: 90,
      stops: [{ pos: 0, color: "#FFFFFF" }, { pos: 1, color: "#00000080" }],
    });
    writeFill(el, null);
    expect(order(el)).toEqual(["xfrm", "prstGeom", "noFill", "ln"]);
    writeFill(el, { type: "solid", color: "#123456" });
    expect(readFill(el, null, ctx, theme)).toEqual({ type: "solid", color: "#123456" });
  });

  it("patches line properties and keeps the rest", () => {
    const el = spPr(`<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:round/></a:ln><a:effectLst/>`);
    writeLine(el, { dash: "dot", tail: { type: "stealth", w: "med", len: "med" } });
    expect(order(el.getElementsByTagNameNS(A, "ln")[0])).toEqual(["solidFill", "prstDash", "round", "tailEnd"]);
    writeLine(el, { color: "#0000FF", width: 38100 });
    expect(readLine(el, null, ctx, theme)).toMatchObject({ color: "#0000FF", width: 38100, dash: "dot", tail: { type: "stealth" } });
    writeLine(el, { tail: null });
    expect(readLine(el, null, ctx, theme)!.tail).toBeNull();
    writeLine(el, null);
    expect(readLine(el, null, ctx, theme)).toBeNull();
    expect(order(el)).toEqual(["xfrm", "prstGeom", "ln", "effectLst"]);
  });

  it("adds a line element in the right place when there is none", () => {
    const el = spPr(`<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:effectLst/>`);
    writeLine(el, { color: "#000000" });
    expect(order(el)).toEqual(["xfrm", "prstGeom", "solidFill", "ln", "effectLst"]);
  });

  it("toggles the outer shadow and keeps other effects", () => {
    const el = spPr(`<a:effectLst><a:glow rad="1"><a:srgbClr val="FF0000"/></a:glow><a:softEdge rad="1"/></a:effectLst>`);
    writeShadow(el, { color: "#00000080", blur: 1, dist: 2, dir: 45 });
    expect(order(el.getElementsByTagNameNS(A, "effectLst")[0])).toEqual(["glow", "outerShdw", "softEdge"]);
    expect(readShadow(el, null, ctx, theme)).toEqual({ color: "#00000080", blur: 1, dist: 2, dir: 45 });
    writeShadow(el, null);
    expect(order(el.getElementsByTagNameNS(A, "effectLst")[0])).toEqual(["glow", "softEdge"]);
  });
});
