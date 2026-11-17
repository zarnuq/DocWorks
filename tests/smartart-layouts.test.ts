import { describe, expect, it } from "vitest";
import { addSmartArtItem, removeSmartArtItem, smartArtItems, smartArtLayout } from "../src/domain/pptx/smartart";

const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const pt = (id: string, text: string, type = "") => `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ""}><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt>`;
const pres = (id: string, of: string) => `<dgm:pt modelId="${id}" type="pres"><dgm:prSet presAssocID="${of}" presName="node"/><dgm:spPr/></dgm:pt>`;
/** An item under `parent` with its two transitions, and presentation points for itself and both transitions. */
const item = (id: string) => pt(id, id.toUpperCase()) + `<dgm:pt modelId="p${id}" type="parTrans" cxnId="c${id}"/><dgm:pt modelId="s${id}" type="sibTrans" cxnId="c${id}"/>` + pres(`v${id}`, id) + pres(`l${id}`, `p${id}`) + pres(`g${id}`, `s${id}`);
const link = (parent: string, id: string, ord: number) => `<dgm:cxn modelId="c${id}" srcId="${parent}" destId="${id}" srcOrd="${ord}" destOrd="0" parTransId="p${id}" sibTransId="s${id}"/><dgm:cxn modelId="po${id}" type="presOf" srcId="${id}" destId="v${id}" srcOrd="0" destOrd="0"/>`;
const xml = (s: string) => new DOMParser().parseFromString(s, "application/xml");
const dataOf = (points: string, links: string) => xml(`<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>${pt("doc", "", "doc")}${points}</dgm:ptLst><dgm:cxnLst>${links}</dgm:cxnLst></dgm:dataModel>`);
const drawingOf = (shapes: string) => xml(`<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree>${shapes}</dsp:spTree></dsp:drawing>`);
const sp = (id: string, x: number, y: number, w: number, h: number, inner = "", rot = 0) =>
  `<dsp:sp modelId="${id}"><dsp:spPr><a:xfrm${rot ? ` rot="${rot}"` : ""}><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>${inner.startsWith("<a:custGeom") ? inner : ""}</dsp:spPr>${inner.startsWith("<a:custGeom") ? "" : inner}</dsp:sp>`;
const text = (t: string) => `<dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2000"/><a:t>${t}</a:t></a:r></a:p></dsp:txBody>`;

interface Shown { id: string; text: string; x: number; y: number; w: number; h: number; cx: number; cy: number; rot: number; el: Element }
const shapes = (dr: Document): Shown[] => Array.from(dr.getElementsByTagNameNS(DSP, "sp")).map((el) => {
  const off = el.getElementsByTagNameNS(A, "off")[0], ext = el.getElementsByTagNameNS(A, "ext")[0], xfrm = el.getElementsByTagNameNS(A, "xfrm")[0];
  const [x, y, w, h] = [Number(off.getAttribute("x")), Number(off.getAttribute("y")), Number(ext.getAttribute("cx")), Number(ext.getAttribute("cy"))];
  return { id: el.getAttribute("modelId")!, text: el.getElementsByTagNameNS(A, "t")[0]?.textContent ?? "", x, y, w, h, cx: x + w / 2, cy: y + h / 2, rot: Number(xfrm.getAttribute("rot") ?? 0), el };
});

describe("SmartArt cycles", () => {
  // Four boxes on a circle (centre 5000,5000, radius 3000) from the top, clockwise, an arrow after each (the last too).
  const C = 5000, R = 3000, ids = ["a", "b", "c", "d"];
  const at = (deg: number, r = R) => ({ x: C + r * Math.cos((deg * Math.PI) / 180), y: C + r * Math.sin((deg * Math.PI) / 180) });
  const cycle = () => ({
    data: dataOf(ids.map(item).join(""), ids.map((id, i) => link("doc", id, i)).join("")),
    drawing: drawingOf(ids.map((id, i) => {
      const p = at(-90 + i * 90), g = at(-45 + i * 90);
      return sp(`v${id}`, p.x - 500, p.y - 500, 1000, 1000, text(id.toUpperCase())) + sp(`g${id}`, g.x - 200, g.y - 200, 400, 400, "", ((i * 90 + 360) % 360) * 60000 + 60000);
    }).join("")),
  });
  const angle = (s: Shown) => (Math.atan2(s.cy - C, s.cx - C) * 180) / Math.PI;
  const turn = (a: number, b: number) => (((b - a) % 360) + 360) % 360;

  it("is told apart from a row", () => {
    const { data, drawing } = cycle();
    expect(smartArtLayout(data, drawing)).toBe("cycle");
  });

  it("adds an item on the circle: every item moves round, shrinks to fit, and each keeps an arrow after it", () => {
    const { data, drawing } = cycle();
    addSmartArtItem(data, drawing, "b", "New");
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["A", "B", "New", "C", "D"]);
    const all = shapes(drawing);
    const boxes = ["A", "B", "New", "C", "D"].map((t) => all.find((s) => s.text === t)!);
    // On the same circle, 72° apart, starting where the first item was.
    for (const b of boxes) expect(Math.hypot(b.cx - C, b.cy - C)).toBeCloseTo(R, -1);
    expect(angle(boxes[0])).toBeCloseTo(-90, 0);
    boxes.slice(1).forEach((b, i) => expect(turn(angle(boxes[i]), angle(b))).toBeCloseTo(72, 0));
    // Neighbours are closer: the boxes shrink by the chords' ratio, and their text with them.
    const s = Math.sin(Math.PI / 5) / Math.sin(Math.PI / 4);
    for (const b of boxes) expect(b.w).toBeCloseTo(1000 * s, -1);
    expect(Number(boxes[0].el.getElementsByTagNameNS(A, "rPr")[0].getAttribute("sz"))).toBeLessThan(2000);
    // Five arrows, each half-way round to the next item and turned by as much as it moved.
    const arrows = all.filter((x) => !x.text);
    expect(arrows).toHaveLength(5);
    const sorted = [...arrows].sort((p, q) => turn(-90, angle(p)) - turn(-90, angle(q)));
    sorted.forEach((arrow, i) => expect(turn(-90, angle(arrow))).toBeCloseTo(36 + 72 * i, 0));
    expect(turn(0, sorted[0].rot / 60000)).toBeCloseTo(turn(0, 1 + (36 - 45)), 0);
    // Every shape still stands for a point in the data.
    const points = new Set(Array.from(data.getElementsByTagNameNS(DGM, "pt")).map((p) => p.getAttribute("modelId")));
    for (const x of all) expect(points.has(x.id)).toBe(true);
  });

  it("keeps the last item its size when a two-item cycle loses one", () => {
    const { data, drawing } = cycle();
    removeSmartArtItem(data, drawing, "d");
    removeSmartArtItem(data, drawing, "c");
    removeSmartArtItem(data, drawing, "b");
    const [left] = shapes(drawing).filter((s) => s.text);
    expect([left.text, left.w, left.h]).toEqual(["A", 1000, 1000]);
    expect(left.el.getElementsByTagNameNS(A, "rPr")[0].getAttribute("sz")).toBe("2000");
  });

  it("removes an item: the rest spread round, keeping their size and arrows", () => {
    const { data, drawing } = cycle();
    removeSmartArtItem(data, drawing, "d");
    const all = shapes(drawing);
    const boxes = ["A", "B", "C"].map((t) => all.find((s) => s.text === t)!);
    boxes.slice(1).forEach((b, i) => expect(turn(angle(boxes[i]), angle(b))).toBeCloseTo(120, 0));
    expect(boxes.map((b) => b.w)).toEqual([1000, 1000, 1000]);
    // The last item of a cycle keeps the arrow back to the first.
    expect(all.filter((x) => !x.text)).toHaveLength(3);
  });
});

describe("SmartArt layouts measured from PowerPoint's own drawings", () => {
  // Radial (radial4): the first item in the middle, its children on an arc over it, an arrow from the middle to each;
  // a second top-level item PowerPoint doesn't draw.
  const radial = () => ({
    data: dataOf(["m", "a", "b", "hidden"].map(item).join(""), link("doc", "m", 0) + link("doc", "hidden", 1) + link("m", "a", 0) + link("m", "b", 1)),
    drawing: drawingOf(
      sp("vm", 2085974, 1700046, 1924050, 1924050, text("M")) +
      sp("va", 4664, 439903, 1827847, 1462278, text("A")) + sp("la", 778662, 1340653, 1547443, 548354, "", 12900000) +
      sp("vb", 4263487, 439903, 1827847, 1462278, text("B")) + sp("lb", 3769893, 1340653, 1547443, 548354, "", 19500000)
    ),
  });
  const C = { x: 2085974 + 962025, y: 1700046 + 962025 };
  const deg = (s: Shown) => (((Math.atan2(s.cy - C.y, s.cx - C.x) * 180) / Math.PI) + 360) % 360;

  it("adds an item round the middle one: the arc keeps its ends, arrows turn with their items", () => {
    const { data, drawing } = radial();
    expect(smartArtLayout(data, drawing)).toBe("radial");
    addSmartArtItem(data, drawing, "a", "New");
    expect(smartArtItems(data).map((i) => [i.text, i.depth])).toEqual([["M", 0], ["A", 1], ["New", 1], ["B", 1], ["HIDDEN", 0]]);
    const all = shapes(drawing);
    const [a, fresh, b, m] = ["A", "New", "B", "M"].map((t) => all.find((s) => s.text === t)!);
    expect([deg(a), deg(fresh), deg(b)].map(Math.round)).toEqual([215, 270, 325]);
    const r = Math.hypot(a.cx - C.x, a.cy - C.y);
    expect(Math.hypot(fresh.cx - C.x, fresh.cy - C.y)).toBeCloseTo(r, -2);
    // There is room: nothing shrinks, the middle stays.
    expect([fresh.w, fresh.h, m.x, m.y]).toEqual([1827847, 1462278, 2085974, 1700046]);
    const arrows = all.filter((s) => !s.text);
    expect(arrows).toHaveLength(3);
    const up = arrows.find((s) => !["la", "lb"].includes(s.id))!;
    expect(Math.round(deg(up))).toBe(270);
    expect(Math.abs(up.rot - (12900000 + 55 * 60000))).toBeLessThan(60000);
    // Removing it again puts the two back where they were.
    removeSmartArtItem(data, drawing, smartArtItems(data)[2].id);
    const back = shapes(drawing);
    expect([back.find((s) => s.text === "A")!.x, back.find((s) => s.text === "B")!.x].map(Math.round)).toEqual([4664, 4263487]);
    expect(() => removeSmartArtItem(data, drawing, "m")).toThrow(/around the middle/i);
  });

  // Vertical process (process4): arrow callouts overlapping a little, the last item a shorter box with its children in it.
  const stack = () => ({
    data: dataOf(["a", "b", "c", "c1", "c2"].map(item).join(""), ["a", "b", "c"].map((id, i) => link("doc", id, i)).join("") + link("c", "c1", 0) + link("c", "c2", 1)),
    drawing: drawingOf(
      sp("va", 0, 799, 8229600, 1719839, text("A"), 10800000) + sp("vb", 0, 1703865, 8229600, 1719839, text("B"), 10800000) +
      sp("vc", 0, 3406931, 8229600, 1118231, text("C")) + sp("vc1", 0, 3988412, 4114800, 514386, text("C1")) + sp("vc2", 4114800, 3988412, 4114800, 514386, text("C2"))
    ),
  });

  it("adds to a column whose items aren't the same size, moving an item's children with it", () => {
    const { data, drawing } = stack();
    expect(smartArtLayout(data, drawing)).toBe("row");
    addSmartArtItem(data, drawing, "a", "New");
    const all = shapes(drawing);
    const [a, fresh, b, c, c1] = ["A", "New", "B", "C", "C1"].map((t) => all.find((s) => s.text === t)!);
    // Same top and bottom as before, in order, every item scaled alike.
    expect(a.y).toBe(799);
    expect(c.y + c.h).toBeCloseTo(3406931 + 1118231, -1);
    expect([a.y < fresh.y, fresh.y < b.y, b.y < c.y]).toEqual([true, true, true]);
    const s = a.h / 1719839;
    expect(s).toBeLessThan(1);
    expect(fresh.h).toBeCloseTo(a.h, -1);
    expect(c.h / 1118231).toBeCloseTo(s, 3);
    // C's children are still inside it, where they were relative to it.
    expect((c1.y - c.y) / c.h).toBeCloseTo((3988412 - 3406931) / 1118231, 2);
    expect(a.rot).toBe(10800000);
  });
});

describe("SmartArt hierarchies", () => {
  // R over A and B, A over A1. Boxes 1,000,000 × 600,000, one level per 1,000,000, lines from parent to child.
  const W = 1_000_000, H = 600_000;
  const line = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="0" b="0"/><a:pathLst><a:path><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="0" y="400000"/></a:lnTo></a:path></a:pathLst></a:custGeom>`;
  const place: Record<string, [x: number, level: number]> = { r: [1_500_000, 0], a: [500_000, 1], b: [2_500_000, 1], a1: [500_000, 2] };
  const tree = (move: Partial<typeof place> = {}) => {
    const at = { ...place, ...move } as typeof place;
    return {
      data: dataOf(["r", "a", "b", "a1"].map(item).join(""), link("doc", "r", 0) + link("r", "a", 0) + link("r", "b", 1) + link("a", "a1", 0)),
      drawing: drawingOf(
        Object.entries(at).map(([id, [x, level]]) => sp(`v${id}`, x, level * 1_000_000, W, H, text(id.toUpperCase()))).join("") +
        // The lines' own geometry doesn't matter: they are redrawn.
        ["a", "b", "a1"].map((id) => sp(`l${id}`, at[id][0], at[id][1] * 1_000_000 - 400_000, 91440, 400_000, line)).join("")
      ),
    };
  };
  const box = (dr: Document, t: string) => shapes(dr).find((s) => s.text === t)!;
  const lineOf = (dr: Document, id: string) => shapes(dr).find((s) => s.id === id)!;
  const path = (s: Shown) => Array.from(s.el.getElementsByTagNameNS(A, "pt")).map((p) => [Number(p.getAttribute("x")), Number(p.getAttribute("y"))]);

  it("is told apart from a row, and lists every level's items", () => {
    const { data, drawing } = tree();
    expect(smartArtLayout(data, drawing)).toBe("tree");
    expect(smartArtItems(data).map((i) => [i.text, i.depth])).toEqual([["R", 0], ["A", 1], ["A1", 2], ["B", 1]]);
  });

  it("adds a sibling at any level: the tree is laid out again in its own space, lines redrawn", () => {
    const { data, drawing } = tree();
    addSmartArtItem(data, drawing, "b", "C");
    expect(smartArtItems(data).map((i) => [i.text, i.depth])).toEqual([["R", 0], ["A", 1], ["A1", 2], ["B", 1], ["C", 1]]);
    // Three boxes and two gaps need 5,000,000; the tree had 3,000,000: everything is 0.6 its size.
    const [r, a, b, c, a1] = ["R", "A", "B", "C", "A1"].map((t) => box(drawing, t));
    expect([a.x, b.x, c.x].map(Math.round)).toEqual([500_000, 1_700_000, 2_900_000]);
    expect([a.w, a.h].map(Math.round)).toEqual([600_000, 360_000]);
    expect([r.y, a.y, b.y, c.y, a1.y].map(Math.round)).toEqual([0, 600_000, 600_000, 600_000, 1_200_000]);
    // A parent sits over the middle of its children.
    expect(r.cx).toBeCloseTo(b.cx, -1);
    expect(a1.cx).toBeCloseTo(a.cx, -1);
    expect(Number(r.el.getElementsByTagNameNS(A, "rPr")[0].getAttribute("sz"))).toBeLessThan(2000);
    // The new box has a line from its parent's bottom middle to its own top middle, with an elbow half-way.
    const lines = shapes(drawing).filter((s) => !s.text);
    expect(lines).toHaveLength(4);
    const fresh = lines.find((l) => !["la", "lb", "la1"].includes(l.id))!;
    expect([fresh.x, fresh.y, fresh.w, fresh.h].map(Math.round)).toEqual([Math.round(r.cx), 360_000, Math.round(c.cx - r.cx), 240_000]);
    expect(path(fresh)).toEqual([[0, 0], [0, 120_000], [fresh.w, 120_000], [fresh.w, 240_000]]);
    // A child straight under its parent gets a straight line.
    const down = lineOf(drawing, "la1");
    expect(path(down).map(([x]) => x)).toEqual([45720, 45720]);
    expect(Math.round(down.cx)).toBe(Math.round(a.cx));
    const points = new Set(Array.from(data.getElementsByTagNameNS(DGM, "pt")).map((p) => p.getAttribute("modelId")));
    for (const s of shapes(drawing)) expect(points.has(s.id)).toBe(true);
  });

  it("adds an item below another, on the next level", () => {
    const { data, drawing } = tree();
    const id = addSmartArtItem(data, drawing, "b", "B1", true);
    expect(smartArtItems(data).map((i) => [i.text, i.depth])).toEqual([["R", 0], ["A", 1], ["A1", 2], ["B", 1], ["B1", 2]]);
    const cxn = Array.from(data.getElementsByTagNameNS(DGM, "cxn")).find((c) => c.getAttribute("destId") === id)!;
    expect([cxn.getAttribute("srcId"), cxn.getAttribute("srcOrd")]).toEqual(["b", "0"]);
    // Nothing else moves: the tree still fits.
    expect([box(drawing, "A").x, box(drawing, "B").x, box(drawing, "R").x]).toEqual([500_000, 2_500_000, 1_500_000]);
    const b1 = box(drawing, "B1");
    expect([b1.x, b1.y, b1.w, b1.h]).toEqual([2_500_000, 2_000_000, W, H]);
    expect(shapes(drawing).filter((s) => !s.text)).toHaveLength(4);
    // An item that already has children gets the new one after them.
    addSmartArtItem(data, drawing, "a", "A2", true);
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["R", "A", "A1", "A2", "B", "B1"]);
  });

  it("removes an item with everything under it, and closes the gap", () => {
    const { data, drawing } = tree();
    removeSmartArtItem(data, drawing, "a");
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["R", "B"]);
    expect(shapes(drawing).map((s) => s.id).sort()).toEqual(["lb", "vb", "vr"]);
    // B is centred under R, in the middle of the space the tree had.
    expect([box(drawing, "R").x, box(drawing, "B").x]).toEqual([1_500_000, 1_500_000]);
    expect(path(lineOf(drawing, "lb")).map(([x]) => x)).toEqual([45720, 45720]);
    expect(() => removeSmartArtItem(data, drawing, "r")).toThrow(/at least one/i);
  });

  it("doesn't lay out a tree with an assistant (an organisation chart's box between the levels)", () => {
    const { data, drawing } = tree();
    // An assistant of R: a point of type "asst" with its own box beside the stem and its own line.
    const add = (doc: Document, parent: Element, markup: string, ns: string) => {
      for (const el of Array.from(xml(`<x xmlns:dgm="${DGM}" xmlns:a="${A}" xmlns:dsp="${DSP}">${markup}</x>`).documentElement.children)) parent.appendChild(doc.importNode(el, true));
      void ns;
    };
    add(data, data.getElementsByTagNameNS(DGM, "ptLst")[0], item("x").replace('modelId="x"', 'modelId="x" type="asst"'), DGM);
    add(data, data.getElementsByTagNameNS(DGM, "cxnLst")[0], link("r", "x", 2), DGM);
    add(drawing, drawing.getElementsByTagNameNS(DSP, "spTree")[0], sp("vx", 2_700_000, 600_000, W, 300_000, text("X")), DSP);
    expect(smartArtLayout(data, drawing)).not.toBe("tree");
    expect(() => addSmartArtItem(data, drawing, "a", "New")).toThrow(/top-level/i);
    expect(shapes(drawing).find((s) => s.text === "X")).toMatchObject({ x: 2_700_000, y: 600_000 });
  });

  it("leaves trees it can't lay out (levels not in rows) to their top-level items", () => {
    // A1 hangs beside A, half a level down: an organisation chart's hanging layout.
    const { data, drawing } = tree({ a1: [900_000, 1.5] });
    expect(smartArtLayout(data, drawing)).toBe("row");
    expect(() => addSmartArtItem(data, drawing, "a", "X")).toThrow(/top-level/i);
    expect(() => addSmartArtItem(data, drawing, "r", "X", true)).toThrow(/top-level/i);
    expect(smartArtItems(data)).toHaveLength(4);
  });
});
