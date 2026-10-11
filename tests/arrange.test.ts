import { describe, it, expect } from "vitest";
import { addShape, shapeId, setRect } from "../src/domain/pptx/slide-edit";
import { alignRects, copyItems, distributeRects, groupItems, itemRect, pasteItems, setItemRect, snapRect, topLevel, ungroupItem } from "../src/domain/pptx/arrange";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { fixture, loadSample, roundTrip } from "./helpers";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function slide() {
  const doc = await loadSample();
  const path = doc.slides()[1].path;
  const xml = doc.xml(path)!;
  const box = (x: number, y: number, w = 100, h = 100) => addShape(xml, "rect", { x, y, w, h });
  return { doc, path, xml, box, rectOf: (el: Element) => parseSlide(doc, path).shapes.find((s) => s.el === el)!.rect };
}

describe("groups", () => {
  it("groups shapes under one p:grpSp sized to their bounds, and ungroups them back", async () => {
    const { doc, path, xml, box, rectOf } = await slide();
    const a = box(100, 100);
    const b = box(400, 300, 200, 100);
    const g = groupItems(xml, [b, a]);
    expect(g.localName).toBe("grpSp");
    expect(Array.from(g.children).map((c) => c.localName)).toEqual(["nvGrpSpPr", "grpSpPr", "sp", "sp"]);
    expect(itemRect(g)).toEqual({ x: 100, y: 100, w: 500, h: 300 });
    expect(topLevel(a)).toBe(g);
    expect(parseSlide(doc, path).shapes.filter((s) => s.grouped).map((s) => s.el)).toEqual([a, b]);

    // Moving and resizing the group moves and scales its children.
    setItemRect(g, { x: 200, y: 100, w: 1000, h: 300 });
    expect(rectOf(a)).toEqual({ x: 200, y: 100, w: 200, h: 100 });
    expect(rectOf(b)).toEqual({ x: 800, y: 300, w: 400, h: 100 });

    const back = await roundTrip(doc);
    expect(back.xml(path)!.getElementsByTagNameNS(P, "grpSp")).toHaveLength(1);

    const children = ungroupItem(g);
    expect(children).toEqual([a, b]);
    expect(xml.getElementsByTagNameNS(P, "grpSp")).toHaveLength(0);
    expect(rectOf(a)).toEqual({ x: 200, y: 100, w: 200, h: 100 });
    expect(rectOf(b)).toEqual({ x: 800, y: 300, w: 400, h: 100 });
    expect(a.parentElement!.localName).toBe("spTree");
  });
});

describe("align and distribute", () => {
  const rects = [
    { x: 0, y: 0, w: 100, h: 50 },
    { x: 300, y: 100, w: 50, h: 100 },
    { x: 100, y: 400, w: 100, h: 100 },
  ];

  it("aligns to the selection bounds", () => {
    expect(alignRects(rects, "left").map((r) => r.x)).toEqual([0, 0, 0]);
    expect(alignRects(rects, "right").map((r) => r.x + r.w)).toEqual([350, 350, 350]);
    expect(alignRects(rects, "center").map((r) => r.x + r.w / 2)).toEqual([175, 175, 175]);
    expect(alignRects(rects, "top").map((r) => r.y)).toEqual([0, 0, 0]);
    expect(alignRects(rects, "middle").map((r) => r.y + r.h / 2)).toEqual([250, 250, 250]);
    expect(alignRects(rects, "bottom").map((r) => r.y + r.h)).toEqual([500, 500, 500]);
  });

  it("aligns a single shape to the slide", () => {
    expect(alignRects([rects[0]], "center", { x: 0, y: 0, w: 1000, h: 800 })[0].x).toBe(450);
  });

  it("spaces shapes evenly between the outermost ones", () => {
    const h = distributeRects(rects, "horizontal");
    // order by x: 0 (w100), 100 (w100), 300 (w50); total span 350, gaps = (350 - 250) / 2 = 50
    expect(h.map((r) => r.x)).toEqual([0, 300, 150]);
    const v = distributeRects(rects, "vertical");
    expect(v.map((r) => r.y)).toEqual([0, 175, 400]);
  });
});

describe("snapRect", () => {
  it("snaps edges and centers to other shapes and the slide, within tolerance", () => {
    const others = [{ x: 500, y: 500, w: 100, h: 100 }];
    const slideSize = { cx: 1000, cy: 800 };
    const s = snapRect({ x: 497, y: 100, w: 50, h: 50 }, others, slideSize, 5);
    expect(s.dx).toBe(3); // left edge -> other's left edge (500)
    expect(s.guides).toContainEqual({ x: 500 });
    const c = snapRect({ x: 473, y: 373, w: 50, h: 50 }, others, slideSize, 5);
    expect([c.dx, c.dy]).toEqual([2, 2]); // center -> slide center (500, 400)
    expect(snapRect({ x: 200, y: 200, w: 10, h: 10 }, others, slideSize, 5)).toEqual({ dx: 0, dy: 0, guides: [] });
  });
});

describe("copy and paste", () => {
  it("pastes copies with fresh ids, remapped image relationships and kept glue", async () => {
    const doc = await loadSample();
    const from = doc.slides()[2].path;
    const to = doc.slides()[1].path;
    const pic = Array.from(doc.xml(from)!.getElementsByTagNameNS(P, "pic"))[0];
    const a = addShape(doc.xml(from)!, "rect", { x: 0, y: 0, w: 10, h: 10 });
    const b = addShape(doc.xml(from)!, "rect", { x: 50, y: 0, w: 10, h: 10 });
    const c = addShape(doc.xml(from)!, "straightConnector1", { x: 10, y: 5, w: 40, h: 0 });
    c.getElementsByTagNameNS(P, "cNvCxnSpPr")[0].innerHTML =
      `<a:stCxn xmlns:a="${A}" id="${shapeId(a)}" idx="3"/><a:endCxn xmlns:a="${A}" id="${shapeId(b)}" idx="1"/>`;

    const clip = await copyItems(doc, from, [pic, a, b, c]);
    const { pasted, skipped } = pasteItems(doc, to, clip, { dx: 0, dy: 0 });
    expect(skipped).toBe(0);
    expect(pasted).toHaveLength(4);
    const ids = new Set(Array.from(doc.xml(to)!.getElementsByTagNameNS(P, "cNvPr")).map((e) => e.getAttribute("id")));
    expect(ids.size).toBe(doc.xml(to)!.getElementsByTagNameNS(P, "cNvPr").length);

    const embed = pasted[0].getElementsByTagNameNS(A, "blip")[0].getAttributeNS(R, "embed")!;
    expect(doc.relTarget(to, embed)).toBe("ppt/media/image1.png");
    const st = pasted[3].getElementsByTagNameNS(A, "stCxn")[0].getAttribute("id");
    expect(st).toBe(shapeId(pasted[1]));

    const back = await roundTrip(doc);
    expect(parseSlide(back, to).shapes.some((s) => s.image === "ppt/media/image1.png")).toBe(true);
  });

  it("offsets pastes onto the same slide and copies images across presentations", async () => {
    const src = await loadSample();
    const from = src.slides()[2].path;
    const pic = Array.from(src.xml(from)!.getElementsByTagNameNS(P, "pic"))[0];
    const clip = await copyItems(src, from, [pic]);

    const same = pasteItems(src, from, clip, { dx: 91440, dy: 91440 }).pasted[0];
    expect(itemRect(same).x).toBe(itemRect(pic).x + 91440);

    const other = await PptxDocument.load(fixture("sample.pptx"));
    const target = other.slides()[0].path;
    const [copy] = pasteItems(other, target, clip, { dx: 0, dy: 0 }).pasted;
    const rId = copy.getElementsByTagNameNS(A, "blip")[0].getAttributeNS(R, "embed")!;
    const media = other.relTarget(target, rId)!;
    expect(media).toMatch(/^ppt\/media\/image\d+\.png$/);
    const back = await roundTrip(other);
    expect(back.mediaPaths()).toContain(media);
  });

  it("copies chart resources independently", async () => {
    const doc = await loadSample();
    const from = doc.slides()[2].path;
    const frame = Array.from(doc.xml(from)!.getElementsByTagNameNS(P, "graphicFrame")).find((f) => !f.getElementsByTagNameNS(A, "tbl").length)!;
    const clip = await copyItems(doc, from, [frame]);
    const result = pasteItems(doc, doc.slides()[1].path, clip, { dx: 0, dy: 0 });
    expect(result.pasted).toHaveLength(1);
    expect(result.skipped).toBe(0);
  });

  it("moves pasted rects with setRect after the fact", async () => {
    const { box, rectOf } = await slide();
    const a = box(10, 10);
    setRect(a, { x: 20, y: 20, w: 100, h: 100 });
    expect(rectOf(a).x).toBe(20);
  });
});

it("copies SmartArt with its cached drawing between presentations", async () => {
 const doc=await loadSample();const from=doc.slides()[2].path;
 const frame=Array.from(doc.xml(from)!.getElementsByTagNameNS(P,"graphicFrame")).find(f=>f.getElementsByTagNameNS("*","relIds").length)!;
 const other=await loadSample();const target=other.slides()[0].path;
 const clip=await copyItems(doc,from,[frame]);const result=pasteItems(other,target,clip,{dx:0,dy:0});
 expect(result.skipped).toBe(0);expect(result.pasted).toHaveLength(1);
 const model=parseSlide(other,target);expect(model.shapes.filter(s=>s.el.namespaceURI!==P&&s.inherited)).toHaveLength(2);
});
