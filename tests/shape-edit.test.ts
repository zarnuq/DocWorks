import { describe, it, expect } from "vitest";
import { addShape, connect, connections, lineBox, reorder, rerouteConnectors, setAdjust, setRect, setTransform, shapeId } from "../src/domain/pptx/slide-edit";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { fillCss } from "../src/domain/pptx/shape-style";
import { loadSample, roundTrip } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";

async function slide() {
  const doc = await loadSample();
  const path = doc.slides()[1].path;
  return { doc, path, xml: doc.xml(path)! };
}
const names = (xml: Document) =>
  Array.from(xml.getElementsByTagNameNS(P, "spTree")[0].children)
    .map((c) => c.getElementsByTagNameNS(P, "cNvPr")[0]?.getAttribute("name"))
    .filter(Boolean);

describe("addShape", () => {
  it("adds a preset shape that takes its look from the theme and holds text", async () => {
    const { doc, path, xml } = await slide();
    const el = addShape(xml, "ellipse", { x: 100, y: 200, w: 3000, h: 2000 });
    expect(el.localName).toBe("sp");
    const back = await roundTrip(doc);
    const shape = parseSlide(back, path).shapes.find((s) => s.name === el.getElementsByTagNameNS(P, "cNvPr")[0].getAttribute("name"))!;
    expect(shape.geom).toBe("ellipse");
    expect(shape.rect).toEqual({ x: 100, y: 200, w: 3000, h: 2000 });
    expect(fillCss(shape.fill)).toBe("#4F81BD");
    expect(shape.text).not.toBeNull();
  });

  it("adds lines and connectors as connector shapes without text", async () => {
    const { xml } = await slide();
    const el = addShape(xml, "straightConnector1", { x: 0, y: 0, w: 10, h: 0 }, { tailArrow: true });
    expect(el.localName).toBe("cxnSp");
    expect(el.getElementsByTagNameNS(P, "txBody")).toHaveLength(0);
    expect(el.getElementsByTagNameNS(A, "tailEnd")[0].getAttribute("type")).toBe("triangle");
  });

  it("gives every new shape a unique id", async () => {
    const { xml } = await slide();
    const a = addShape(xml, "rect", { x: 0, y: 0, w: 1, h: 1 });
    const b = addShape(xml, "rect", { x: 0, y: 0, w: 1, h: 1 });
    expect(shapeId(a)).not.toBe(shapeId(b));
  });
});

describe("shape transforms", () => {
  it("writes rotation and flips on the xfrm", async () => {
    const { doc, path, xml } = await slide();
    const el = addShape(xml, "rect", { x: 0, y: 0, w: 10, h: 10 });
    setTransform(el, { rot: 30, flipH: true });
    const xfrm = el.getElementsByTagNameNS(A, "xfrm")[0];
    expect(xfrm.getAttribute("rot")).toBe("1800000");
    expect(xfrm.getAttribute("flipH")).toBe("1");
    setTransform(el, { rot: 0, flipH: false });
    expect(xfrm.hasAttribute("rot")).toBe(false);
    expect(xfrm.hasAttribute("flipH")).toBe(false);
    setTransform(el, { flipV: true });
    const shape = parseSlide(doc, path).shapes.find((s) => s.el === el)!;
    expect(shape.flipV).toBe(true);
  });

  it("writes adjust values the geometry reads back", async () => {
    const { doc, path, xml } = await slide();
    const el = addShape(xml, "chevron", { x: 0, y: 0, w: 1000, h: 400 });
    setAdjust(el, "adj", 30000);
    setAdjust(el, "adj", 25000);
    expect(el.getElementsByTagNameNS(A, "gd")).toHaveLength(1);
    const shape = parseSlide(doc, path).shapes.find((s) => s.el === el)!;
    expect(shape.adj).toEqual({ adj: 25000 });
    expect(shape.paths![0].d).toBe("M0 0 L900 0 L1000 200 L900 400 L0 400 L100 200 Z");
  });
});

describe("reorder", () => {
  it("moves shapes forward, backward, to the front and to the back", async () => {
    const { xml } = await slide();
    const [a, b, c] = ["A", "B", "C"].map((n) => {
      const el = addShape(xml, "rect", { x: 0, y: 0, w: 1, h: 1 });
      el.getElementsByTagNameNS(P, "cNvPr")[0].setAttribute("name", n);
      return el;
    });
    const tail = () => names(xml).slice(-3);
    expect(tail()).toEqual(["A", "B", "C"]);
    reorder(a, "forward");
    expect(tail()).toEqual(["B", "A", "C"]);
    reorder(c, "backward");
    expect(tail()).toEqual(["B", "C", "A"]);
    reorder(a, "back");
    expect(names(xml)[0]).toBe("A");
    reorder(a, "front");
    expect(names(xml).at(-1)).toBe("A");
    void b;
  });
});

describe("connectors", () => {
  it("records and clears the shapes a connector is glued to", async () => {
    const { xml } = await slide();
    const cxn = addShape(xml, "straightConnector1", { x: 0, y: 0, w: 10, h: 10 });
    connect(cxn, "st", { id: "5", idx: 3 });
    connect(cxn, "end", { id: "7", idx: 1 });
    expect(connections(cxn)).toEqual({ st: { id: "5", idx: 3 }, end: { id: "7", idx: 1 } });
    const nv = cxn.getElementsByTagNameNS(P, "cNvCxnSpPr")[0];
    expect(Array.from(nv.children).map((c) => c.localName)).toEqual(["stCxn", "endCxn"]);
    connect(cxn, "st", null);
    expect(connections(cxn)).toEqual({ st: null, end: { id: "7", idx: 1 } });
  });

  it("boxes a line between two points with flips", () => {
    expect(lineBox({ x: 10, y: 20 }, { x: 0, y: 0 })).toEqual({ rect: { x: 0, y: 0, w: 10, h: 20 }, flipH: true, flipV: true });
    expect(lineBox({ x: 0, y: 5 }, { x: 10, y: 5 })).toEqual({ rect: { x: 0, y: 5, w: 10, h: 0 }, flipH: false, flipV: false });
  });
});

describe("rerouteConnectors", () => {
  it("keeps glued connector ends on their shapes' connection sites", async () => {
    const { doc, path, xml } = await slide();
    const a = addShape(xml, "rect", { x: 0, y: 0, w: 100, h: 100 });
    const b = addShape(xml, "rect", { x: 300, y: 0, w: 100, h: 100 });
    const cxn = addShape(xml, "straightConnector1", { x: 0, y: 0, w: 1, h: 1 });
    // rect sites: 0 top, 1 left, 2 bottom, 3 right
    connect(cxn, "st", { id: shapeId(a), idx: 3 });
    connect(cxn, "end", { id: shapeId(b), idx: 1 });
    const model = () => parseSlide(doc, path);
    const box = () => model().shapes.find((s) => s.el === cxn)!;

    expect(rerouteConnectors(xml, model().shapes)).toBe(true);
    expect(box().rect).toEqual({ x: 100, y: 50, w: 200, h: 0 });

    setRect(b, { x: 300, y: 200, w: 100, h: 100 });
    rerouteConnectors(xml, model().shapes);
    expect(box().rect).toEqual({ x: 100, y: 50, w: 200, h: 200 });

    setRect(b, { x: -300, y: -200, w: 100, h: 100 });
    rerouteConnectors(xml, model().shapes);
    expect(box()).toMatchObject({ rect: { x: -300, y: -150, w: 400, h: 200 }, flipH: true, flipV: true });
    expect(rerouteConnectors(xml, model().shapes)).toBe(false);
  });
});
