import { describe, it, expect } from "vitest";
import { parseSlide, readParagraphs, ShapeModel, SlideModel, TextStyle } from "../src/domain/pptx/slide-model";
import { loadSample } from "./helpers";
import { fillCss } from "../src/domain/pptx/shape-style";

const byName = (s: SlideModel, name: string) => {
  const shape = s.shapes.find((x) => x.name === name);
  if (!shape) throw new Error(`no shape ${name}: ${s.shapes.map((x) => x.name).join(", ")}`);
  return shape;
};
const css = (sh: ShapeModel) => fillCss(sh.fill, sh.flipH, sh.flipV);
const textOf = (sh: ShapeModel) => sh.text!.paras.map((p) => p.runs.map((r) => r.text).join(""));

async function slide(i: number) {
  const doc = await loadSample();
  return parseSlide(doc, doc.slides()[i].path);
}

describe("parseSlide: placeholders inherit from layout and master", () => {
  it("takes geometry from the layout when the slide has no xfrm", async () => {
    const s = await slide(0);
    expect(byName(s, "Title 1").rect).toEqual({ x: 685800, y: 2130425, w: 7772400, h: 1470025 });
  });

  it("takes geometry from the master when the layout has none", async () => {
    const s = await slide(1);
    // layout 2 title has no xfrm of its own -> master title placeholder
    expect(byName(s, "Title 1").rect).toEqual({ x: 457200, y: 274638, w: 8229600, h: 1143000 });
  });

  it("resolves title text styles from master titleStyle", async () => {
    const title = byName(await slide(0), "Title 1");
    const run = title.text!.paras[0].runs[0];
    expect(run).toMatchObject({ text: "Hello Deck", sz: 44, font: "Calibri", color: "#000000", b: false });
    expect(title.text!.paras[0].align).toBe("ctr");
    expect(title.text!.paras[0].bullet).toBeNull();
    expect(title.text!.anchor).toBe("ctr");
  });

  it("layers layout lstStyle over master bodyStyle", async () => {
    const sub = byName(await slide(0), "Subtitle 2");
    const p = sub.text!.paras[0];
    expect(p.runs[0]).toMatchObject({ text: "Subtitle here", sz: 32, color: "#404040" }); // tx1 tint 75%
    expect(p.align).toBe("ctr");
    expect(p.bullet).toBeNull(); // layout buNone beats master buChar
  });

  it("reads body bullets, levels and run formatting", async () => {
    const body = byName(await slide(1), "Content Placeholder 2");
    const [p0, p1] = body.text!.paras;
    expect(p0).toMatchObject({ lvl: 0, bullet: "•", marL: 342900, indent: -342900 });
    expect(p0.runs[0]).toMatchObject({ text: "First point", sz: 32 });
    expect(p1.lvl).toBe(1);
    expect(p1.runs).toHaveLength(2);
    expect(p1.runs[0]).toMatchObject({ text: "Bold red", b: true, sz: 20, color: "#FF0000", i: false });
    expect(p1.runs[1]).toMatchObject({ text: " italic", i: true, b: false, sz: 28 }); // lvl2 default 28pt
  });

  it("reads speaker notes", async () => {
    expect((await slide(0)).notes).toBe("Speaker notes");
    expect((await slide(1)).notes).toBe("");
  });

  it("resolves the master background", async () => {
    expect((await slide(0)).background.color).toBe("#FFFFFF");
  });
});

describe("parseSlide: free shapes", () => {
  it("reads a text box with explicit geometry and alignment", async () => {
    const tb = byName(await slide(2), "TextBox 1");
    expect(tb.kind).toBe("sp");
    expect(tb.rect).toEqual({ x: 914400, y: 914400, w: 3657600, h: 914400 });
    expect(textOf(tb)).toEqual(["Box text"]);
    expect(tb.text!.paras[0].align).toBe("ctr");
    expect(tb.text!.paras[0].runs[0].sz).toBe(18); // otherStyle / defaultTextStyle
    expect(tb.fill).toBeNull();
    expect(tb.text!.anchor).toBe("t");
    expect(tb.text!.insets).toEqual({ l: 91440, t: 45720, r: 91440, b: 45720 });
  });

  it("reads solid fills and shape-style font color", async () => {
    const rect = byName(await slide(2), "Rectangle 2");
    expect(css(rect)).toBe("#00FF00");
    expect(rect.geom).toBe("rect");
    expect(rect.text!.paras[0].runs[0].color).toBe("#FFFFFF"); // fontRef lt1
    expect(rect.text!.anchor).toBe("ctr");
  });

  it("falls back to the theme fill style the fillRef points at", async () => {
    const oval = byName(await slide(2), "Oval 3");
    expect(oval.geom).toBe("ellipse");
    // fillRef idx 3 = the theme's third fill style, a gradient of the accent1 placeholder color
    expect(oval.fill).toMatchObject({ type: "gradient", angle: 270, stops: [{ pos: 0, color: "#4F81BD" }, { pos: 1 }] });
    expect(oval.line?.color).toBe("#4B7BB4"); // theme line style 1: accent1 shaded 95%
  });

  it("reads pictures as media paths", async () => {
    const pic = byName(await slide(2), "Picture 4");
    expect(pic.kind).toBe("pic");
    expect(pic.image).toBe("ppt/media/image1.png");
    expect(pic.rect).toEqual({ x: 914400, y: 2743200, w: 914400, h: 914400 });
  });

  it("reads tables as text grids", async () => {
    const tbl = byName(await slide(2), "Table 5");
    expect(tbl.kind).toBe("table");
    expect(tbl.table).toEqual([
      ["A1", "B1"],
      ["A2", "B2"],
    ]);
    expect(tbl.rect).toEqual({ x: 914400, y: 4572000, w: 3657600, h: 914400 });
  });

  it("flattens group children with the group transform", async () => {
    const g = byName(await slide(2), "TextBox 7");
    expect(g.grouped).toBe(true);
    expect(g.rect.x).toBe(6400800);
    expect(textOf(g)).toEqual(["Grouped"]);
  });

  it("marks only slide-owned shapes as editable", async () => {
    const s = await slide(2);
    const own = s.shapes.filter((x) => x.el.namespaceURI !== "http://schemas.microsoft.com/office/drawing/2008/diagram");
    expect(own.every((x) => !x.inherited)).toBe(true);
    expect(byName(s, "TextBox 1").grouped).toBe(false);
  });

  it("turns gradient fills into CSS gradients with alpha", async () => {
    expect(css(byName(await slide(2), "Gradient"))).toBe("linear-gradient(90deg, #FF0000 0%, #0000FF80 100%)");
  });

  it("reads custom geometry as SVG paths", async () => {
    const f = byName(await slide(2), "Freeform");
    expect(f.geom).toBe("custom");
    expect(f.paths).toEqual([{ w: 914400, h: 914400, d: "M0 0 L914400 914400 L0 914400 Z", fill: "norm", stroke: true }]);
    expect(f.fill).toMatchObject({ type: "gradient", stops: [{ color: "#4F81BD" }, {}] });
  });

  it("reads lines with flips as a stroked path", async () => {
    const l = byName(await slide(2), "Line");
    expect(l.geom).toBe("line");
    expect(l.flipH).toBe(true);
    expect(l.flipV).toBe(false);
    const { w, h } = l.rect;
    expect(l.paths).toEqual([{ w, h, d: `M0 0 L${w} ${h}`, fill: "norm", stroke: true }]);
    expect(l.line?.color).toBe("#4F81BD");
  });

  it("flags placeholders", async () => {
    expect(byName(await slide(0), "Title 1").placeholder).toBe(true);
    expect(byName(await slide(2), "TextBox 1").placeholder).toBe(false);
    expect(byName(await slide(2), "Rectangle 2").paths![0].d).toBe("M0 0 L1828800 0 L1828800 914400 L0 914400 Z");
  });

  it("fills useBgFill shapes with the slide background (PowerPoint Designer layouts)", async () => {
    const s = await slide(2);
    expect(css(byName(s, "BgFill"))).toBe(s.background.color);
  });

  it("mirrors gradient direction for flipped shapes", async () => {
    // ang 50deg, flipH -> 130deg in DrawingML terms -> CSS 220deg
    expect(css(byName(await slide(2), "FlippedGradient"))).toBe("linear-gradient(220deg, #000000 0%, #FFFFFF 100%)");
  });

  it("evaluates formula guides in freeform points", async () => {
    expect(byName(await slide(2), "GuidedFreeform").paths![0].d).toBe("M0 0 L457200 914400 L0 914400 Z");
  });
});

describe("SmartArt", () => {
  const diagram = async () => (await slide(2)).shapes.filter((s) => s.el.namespaceURI === "http://schemas.microsoft.com/office/drawing/2008/diagram");

  it("renders the cached diagram drawing instead of a placeholder frame", async () => {
    const s = await slide(2);
    expect(s.shapes.some((x) => x.name === "Diagram")).toBe(false);
    expect(await diagram()).toHaveLength(2);
  });

  it("places diagram shapes relative to the frame", async () => {
    const [chevron, box] = await diagram();
    expect(chevron.rect).toEqual({ x: 1010000, y: 2020000, w: 1000000, h: 400000 });
    expect(box.rect).toEqual({ x: 1010000, y: 2420000, w: 900000, h: 800000 });
  });

  it("reads diagram text, fills and the separate text box", async () => {
    const [chevron, box] = await diagram();
    expect(chevron.geom).toBe("chevron");
    expect(css(chevron)).toBe("#C0504D");
    expect(chevron.text!.paras[0].runs[0]).toMatchObject({ text: "Step", sz: 28, b: true, color: "#FFFFFF" });
    expect(chevron.textRect).toEqual({ x: 120000, y: 0, w: 760000, h: 400000 });
    expect(css(box)).toMatch(/^#[0-9A-F]{6}E6$/); // tinted accent2 at 90% alpha
  });

  it("keeps diagram shapes read-only (PowerPoint rebuilds them from the diagram data)", async () => {
    expect((await diagram()).every((s) => s.inherited)).toBe(true);
  });
});

describe("readParagraphs: spacing", () => {
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const xml = (s: string) => new DOMParser().parseFromString(s, "application/xml").documentElement;
  const master = xml(
    `<a:lstStyle xmlns:a="${A}"><a:lvl1pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef>` +
      `<a:defRPr sz="2800"/></a:lvl1pPr></a:lstStyle>`
  );
  const style: TextStyle = { chain: [null, null, null, master], fontRefColor: null, colors: {} as TextStyle["colors"], fonts: { major: "Calibri", minor: "Calibri" } };
  const body = xml(
    `<p:txBody xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="${A}">` +
      `<a:p><a:r><a:rPr sz="1700"/><a:t>x</a:t></a:r></a:p>` +
      `<a:p><a:pPr><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr><a:endParaRPr sz="1200"/></a:p>` +
      `</p:txBody>`
  );

  it("inherits line and paragraph spacing from the style chain", () => {
    const [p] = readParagraphs(body, style);
    expect(p).toMatchObject({ lnSpc: { pct: 0.9 }, spcBef: { pts: 10 }, spcAft: null });
  });

  it("lets the paragraph override spacing and records the end-of-paragraph size", () => {
    const p = readParagraphs(body, style)[1];
    expect(p).toMatchObject({ lnSpc: { pts: 24 }, spcBef: { pts: 10 }, spcAft: { pct: 0.5 }, endSz: 12 });
    expect(readParagraphs(body, style)[0].endSz).toBe(28);
  });
});
