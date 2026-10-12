import { describe, expect, it } from "vitest";
import { addSmartArtItem, removeSmartArtItem, setSmartArtText, smartArtItems } from "../src/domain/pptx/smartart";

const DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const t = (text: string) => `<dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>${text}</a:t></a:r></a:p></dgm:t>`;
const pt = (id: string, text: string, type = "") => `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ""}><dgm:prSet/><dgm:spPr/>${t(text)}</dgm:pt>`;
const pres = (id: string, of: string) => `<dgm:pt modelId="${id}" type="pres"><dgm:prSet presAssocID="${of}" presName="text"/><dgm:spPr/></dgm:pt>`;
const cxn = (src: string, dest: string, ord: number) => `<dgm:cxn modelId="c${src}${dest}" srcId="${src}" destId="${dest}" srcOrd="${ord}" destOrd="0"/>`;
// A list "Plan" (with child "Budget") and "Build", in that order though stored out of order.
const data = () => new DOMParser().parseFromString(
  `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>${pt("doc", "", "doc")}${pt("b", "Build")}${pt("a", "Plan")}${pt("a1", "Budget")}${pres("pa", "a")}${pres("pb", "b")}</dgm:ptLst>` +
  `<dgm:cxnLst>${cxn("doc", "b", 1)}${cxn("doc", "a", 0)}${cxn("a", "a1", 0)}</dgm:cxnLst></dgm:dataModel>`, "application/xml");
const drawing = () => new DOMParser().parseFromString(
  `<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:sp modelId="pa"><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="2400"/><a:t>Plan</a:t></a:r></a:p></dsp:txBody></dsp:sp><dsp:sp modelId="pb"><dsp:txBody><a:bodyPr/><a:p><a:r><a:t>Build</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`, "application/xml");

describe("SmartArt text", () => {
  it("lists the items in order, with their level", () => {
    expect(smartArtItems(data())).toEqual([
      { id: "a", text: "Plan", depth: 0 },
      { id: "a1", text: "Budget", depth: 1 },
      { id: "b", text: "Build", depth: 0 },
    ]);
  });

  it("changes an item's text in the data and in the drawing shown, keeping its formatting", () => {
    const d = data(), dr = drawing();
    setSmartArtText(d, dr, "a", "Plan it\nthen check");
    expect(smartArtItems(d)[0].text).toBe("Plan it\nthen check");
    const shown = dr.getElementsByTagNameNS(DSP, "sp")[0];
    expect(Array.from(shown.getElementsByTagNameNS(A, "p")).map((p) => p.textContent)).toEqual(["Plan it", "then check"]);
    // The run keeps its size; the other item is untouched.
    expect(shown.getElementsByTagNameNS(A, "rPr")[0].getAttribute("sz")).toBe("2400");
    expect(dr.getElementsByTagNameNS(DSP, "sp")[1].textContent).toBe("Build");
    const model = Array.from(d.getElementsByTagNameNS(DGM, "pt")).find((p) => p.getAttribute("modelId") === "a")!;
    expect(model.getElementsByTagNameNS(A, "rPr")[0].getAttribute("b")).toBe("1");
  });
});

describe("SmartArt items", () => {
  // A process of three items in a row (x 0..400), each a box with an arrow after it (but the last).
  const items = ["a", "b", "c"];
  const flow = () => {
    const pts = items.map((id, i) => pt(id, id.toUpperCase()) + `<dgm:pt modelId="p${id}" type="parTrans" cxnId="c${id}"/><dgm:pt modelId="s${id}" type="sibTrans" cxnId="c${id}"/>` + pres(`v${id}`, id) + pres(`v_s${id}`, `s${id}`)).join("");
    const cxns = items.map((id, i) => `<dgm:cxn modelId="c${id}" srcId="doc" destId="${id}" srcOrd="${i}" destOrd="0" parTransId="p${id}" sibTransId="s${id}"/><dgm:cxn modelId="po${id}" type="presOf" srcId="${id}" destId="v${id}" srcOrd="0" destOrd="0"/>`).join("");
    const box = (id: string, x: number, w: number, text = "") =>
      `<dsp:sp modelId="${id}"><dsp:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="${w}" cy="100"/></a:xfrm></dsp:spPr>${text ? `<dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="2000"/><a:t>${text}</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="${x}" y="0"/><a:ext cx="${w}" cy="100"/></dsp:txXfrm>` : ""}</dsp:sp>`;
    const shapes = items.map((id, i) => box(`v${id}`, i * 150, 100, id.toUpperCase()) + (i < 2 ? box(`v_s${id}`, i * 150 + 110, 30) : "")).join("");
    return {
      data: new DOMParser().parseFromString(`<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>${pt("doc", "", "doc")}${pts}</dgm:ptLst><dgm:cxnLst>${cxns}</dgm:cxnLst></dgm:dataModel>`, "application/xml"),
      drawing: new DOMParser().parseFromString(`<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree>${shapes}</dsp:spTree></dsp:drawing>`, "application/xml"),
    };
  };
  const boxes = (dr: Document) => Array.from(dr.getElementsByTagNameNS(DSP, "sp")).map((sp) => {
    const off = sp.getElementsByTagNameNS(A, "off")[0], ext = sp.getElementsByTagNameNS(A, "ext")[0];
    return { text: sp.textContent ?? "", x: Number(off.getAttribute("x")), w: Number(ext.getAttribute("cx")), h: Number(ext.getAttribute("cy")) };
  });

  it("adds an item after another, laying the row out again in the same space", () => {
    const { data, drawing } = flow();
    addSmartArtItem(data, drawing, "c", "D");
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["A", "B", "C", "D"]);
    const b = boxes(drawing);
    const shown = b.filter((x) => x.text);
    expect(shown.map((x) => x.text)).toEqual(["A", "B", "C", "D"]);
    // Same span (0..400), equal widths, in order, an arrow in each of the three gaps.
    expect(shown[0].x).toBe(0);
    expect(shown[3].x + shown[3].w).toBeCloseTo(400, 0);
    expect(new Set(shown.map((x) => x.w)).size).toBe(1);
    expect(shown.every((x, i) => i === 0 || x.x > shown[i - 1].x + shown[i - 1].w)).toBe(true);
    const arrows = b.filter((x) => !x.text);
    expect(arrows).toHaveLength(3);
    arrows.forEach((a, i) => expect(a.x).toBeGreaterThan(shown[i].x + shown[i].w - 1));
    // The shrunk boxes keep their proportions, and their text shrinks with them.
    expect(shown[0].h).toBeLessThan(100);
    expect(Number(drawing.getElementsByTagNameNS(A, "rPr")[0].getAttribute("sz"))).toBeLessThan(2000);
  });

  it("adds in the middle and keeps the data model's order and connections consistent", () => {
    const { data, drawing } = flow();
    const id = addSmartArtItem(data, drawing, "a", "A2");
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["A", "A2", "B", "C"]);
    const cxn = Array.from(data.getElementsByTagNameNS(DGM, "cxn")).find((c) => c.getAttribute("destId") === id)!;
    const ids = new Set(Array.from(data.getElementsByTagNameNS(DGM, "pt")).map((p) => p.getAttribute("modelId")));
    expect(ids.has(cxn.getAttribute("parTransId"))).toBe(true);
    expect(ids.has(cxn.getAttribute("sibTransId"))).toBe(true);
    expect(boxes(drawing).filter((x) => x.text).map((x) => x.text)).toEqual(["A", "A2", "B", "C"]);
  });

  it("removes an item with its shapes, its connections and the arrow the new last item no longer needs", () => {
    const { data, drawing } = flow();
    removeSmartArtItem(data, drawing, "c");
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["A", "B"]);
    const b = boxes(drawing);
    expect(b.filter((x) => x.text).map((x) => x.text)).toEqual(["A", "B"]);
    expect(b.filter((x) => !x.text)).toHaveLength(1);
    expect(Array.from(data.getElementsByTagNameNS(DGM, "pt")).some((p) => ["c", "pc", "sc", "vc", "v_sc"].includes(p.getAttribute("modelId")!))).toBe(false);
    expect(Array.from(data.getElementsByTagNameNS(DGM, "cxn")).some((c) => c.getAttribute("srcId") === "c" || c.getAttribute("destId") === "c")).toBe(false);
    // Fewer items keep their size, centred in the old span.
    const shown = b.filter((x) => x.text);
    expect(shown.map((x) => x.w)).toEqual([100, 100]);
    expect(shown[0].x).toBe(400 - (shown[1].x + shown[1].w));
    expect(() => { removeSmartArtItem(data, drawing, "a"); removeSmartArtItem(data, drawing, "b"); }).toThrow(/at least one/i);
  });

  it("removes an item it added, and the added item's shapes are linked to the data like the others", () => {
    const { data, drawing } = flow();
    const id = addSmartArtItem(data, drawing, "c", "D");
    removeSmartArtItem(data, drawing, id);
    expect(smartArtItems(data).map((i) => i.text)).toEqual(["A", "B", "C"]);
    const b = boxes(drawing);
    expect(b.filter((x) => x.text).map((x) => x.text)).toEqual(["A", "B", "C"]);
    expect(b.filter((x) => !x.text)).toHaveLength(2);
    // Every drawing shape stands for a presentation point in the data.
    const pres = new Set(Array.from(data.getElementsByTagNameNS(DGM, "pt")).map((p) => p.getAttribute("modelId")));
    addSmartArtItem(data, drawing, "a", "A2");
    const all = new Set(Array.from(data.getElementsByTagNameNS(DGM, "pt")).map((p) => p.getAttribute("modelId")));
    for (const sp of Array.from(drawing.getElementsByTagNameNS(DSP, "sp"))) expect(all.has(sp.getAttribute("modelId"))).toBe(true);
    expect(all.size).toBeGreaterThan(pres.size);
  });

  it("refuses diagrams that aren't laid out in a row or column", () => {
    const { data, drawing } = flow();
    // "B" (its box and text box, after A's and A's arrow) moved down: not a row.
    for (const i of [3, 4]) drawing.getElementsByTagNameNS(A, "off")[i].setAttribute("y", "500");
    expect(() => addSmartArtItem(data, drawing, "a", "X")).toThrow(/row or column/i);
    expect(smartArtItems(data)).toHaveLength(3);
  });
});
