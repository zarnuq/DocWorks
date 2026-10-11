import { describe, it, expect, beforeEach } from "vitest";
import { scalePath, renderSlide } from "../src/presentation/pptx/slide-renderer";
import { parseSlide, SlideModel } from "../src/domain/pptx/slide-model";
import { loadSample } from "./helpers";

const WIDTH = 960; // 9144000 EMU wide -> 1px = 9525 EMU
let host: HTMLElement;
beforeEach(() => {
  document.body.innerHTML = "";
  host = document.body.appendChild(document.createElement("div"));
});

async function render(i: number, editable = true, mutate?: (m: SlideModel) => void) {
  const doc = await loadSample();
  const model = parseSlide(doc, doc.slides()[i].path);
  mutate?.(model);
  const shapes = renderSlide(host, model, doc.slideSize, { width: WIDTH, editable, imageUrl: (p) => `blob:${p}` });
  const byName = (n: string) => shapes.find((s) => s.shape.name === n)!;
  return { model, shapes, byName };
}

describe("renderSlide", () => {
  it("tags slide shapes with their ids for animations, not inherited ones", async () => {
    const { shapes } = await render(0, false);
    const own = shapes.find((s) => !s.shape.inherited && s.shape.kind !== "cell")!;
    expect(own.el.dataset.spids).toBe(own.shape.el.getElementsByTagNameNS("*", "cNvPr")[0].getAttribute("id"));
    for (const s of shapes.filter((x) => x.shape.inherited)) expect(s.el.dataset.spids).toBeUndefined();
  });

  it("sizes arrowheads like PowerPoint: thin lines count as about 2 pt wide, so heads stay readable", async () => {
    const marker = async (width: number) => {
      host.innerHTML = "";
      await render(2, false, (m) => {
        const s = m.shapes.find((x) => x.paths?.length)!;
        s.line = { color: "#FF0000", width, dash: "solid", cap: "flat", head: null, tail: { type: "triangle", w: "med", len: "med" } };
      });
      const mk = host.querySelector("marker")!;
      return [mk.getAttribute("markerUnits"), Number(mk.getAttribute("markerWidth")).toFixed(2), Number(mk.getAttribute("markerHeight")).toFixed(2)];
    };
    // 0.75 pt line: medium head is 3 x 0.7 mm (25200 EMU), 1 px = 9525 EMU here.
    expect(await marker(9525)).toEqual(["userSpaceOnUse", (75600 / 9525).toFixed(2), (75600 / 9525).toFixed(2)]);
    // A 4 pt line: 3 x its own width.
    expect(await marker(50800)).toEqual(["userSpaceOnUse", (152400 / 9525).toFixed(2), (152400 / 9525).toFixed(2)]);
  });

  it("sizes the slide to the requested width keeping aspect ratio", async () => {
    await render(0);
    const slide = host.querySelector<HTMLElement>(".pptx-slide")!;
    expect(slide.style.width).toBe("960px");
    expect(slide.style.height).toBe("720px");
    expect(slide.style.backgroundColor).toBe("rgb(255, 255, 255)");
  });

  it("positions shapes in px", async () => {
    const { byName } = await render(2);
    const el = byName("TextBox 1").el;
    expect([el.style.left, el.style.top, el.style.width, el.style.height]).toEqual(["96px", "96px", "384px", "96px"]);
  });

  it("scales font sizes from points (12700 EMU per pt)", async () => {
    const { byName } = await render(2);
    const p = byName("TextBox 1").textEl!.querySelector("p")!;
    expect(parseFloat(p.style.fontSize)).toBeCloseTo((18 * 12700) / 9525);
    expect(byName("TextBox 1").pxPerPt).toBeCloseTo(12700 / 9525);
  });

  it("applies normAutofit font scale", async () => {
    const { byName } = await render(2, true, (m) => (m.shapes.find((s) => s.name === "TextBox 1")!.text!.fontScale = 0.5));
    expect(byName("TextBox 1").pxPerPt).toBeCloseTo((12700 / 9525) * 0.5);
  });

  it("draws fills, borders, ellipses and rotation", async () => {
    const { byName } = await render(2, true, (m) => (m.shapes.find((s) => s.name === "Rectangle 2")!.rot = 30));
    const rect = byName("Rectangle 2").el;
    expect(rect.querySelector("svg path")!.getAttribute("fill")).toBe("#00FF00");
    expect(rect.style.transform).toBe("rotate(30deg)");
    const oval = byName("Oval 3").el.querySelector("svg path")!;
    expect(oval.getAttribute("d")).toMatch(/^M0 48 A48 48 0 0 1 /);
    expect(oval.getAttribute("stroke")).toBe("#4B7BB4");
  });

  it("anchors text vertically", async () => {
    const { byName } = await render(2);
    expect(byName("Rectangle 2").textEl!.style.justifyContent).toBe("center");
    expect(byName("TextBox 1").textEl!.style.justifyContent).toBe("flex-start");
  });

  it("renders pictures through the image resolver", async () => {
    const { byName } = await render(2);
    expect(byName("Picture 4").el.querySelector("img")!.getAttribute("src")).toBe("blob:ppt/media/image1.png");
  });

  it("renders tables with one editable text box per cell, in model order", async () => {
    const { byName, shapes, model } = await render(2);
    const table = byName("Table 5");
    const tds = Array.from(table.el.querySelectorAll("td"));
    expect(tds.map((td) => td.textContent)).toEqual(["A1", "B1", "A2", "B2"]);
    expect(Array.from(table.el.querySelectorAll("col")).map((c) => c.style.width)).toEqual(["192px", "192px"]);
    const i = shapes.indexOf(table);
    expect(shapes.slice(i + 1, i + 5).map((r) => [r.shape.kind, r.el, r.textEl!.contentEditable])).toEqual(tds.map((td) => ["cell", td, "true"]));
    expect(shapes.map((r) => r.shape)).toEqual(model.shapes);
  });

  it("draws merged cells once, with the style's fill and borders", async () => {
    const { byName } = await render(2, true, (m) => {
      const t = m.shapes.find((s) => s.name === "Table 5")!.table!;
      Object.assign(t.cells[0][0], { gridSpan: 2, fill: { type: "solid", color: "#FF0000" }, borders: { l: null, r: null, t: null, b: { color: "#00FF00", width: 38100 } } });
      t.cells[0][1].hidden = true;
    });
    const tds = Array.from(byName("Table 5").el.querySelectorAll("td"));
    expect(tds).toHaveLength(3);
    expect(tds[0].colSpan).toBe(2);
    expect(tds[0].style.background).toBe("rgb(255, 0, 0)");
    expect(tds[0].style.borderBottom).toBe("4px solid rgb(0, 255, 0)");
  });

  it("makes slide-owned text editable only when editing", async () => {
    const { byName } = await render(2);
    expect(byName("TextBox 1").textEl!.contentEditable).toBe("true");
    host.innerHTML = "";
    const thumb = await render(2, false);
    expect(thumb.byName("TextBox 1").textEl!.contentEditable).not.toBe("true");
  });

  it("never makes inherited shapes editable", async () => {
    const { shapes } = await render(0, true, (m) => {
      m.shapes[0].inherited = true;
    });
    expect(shapes[0].textEl?.contentEditable).not.toBe("true");
    expect(shapes[0].el.classList.contains("is-inherited")).toBe(true);
  });

  it("marks empty placeholders for a prompt", async () => {
    const { byName } = await render(0, true, (m) => (m.shapes.find((s) => s.name === "Title 1")!.text!.paras[0].runs = []));
    expect(byName("Title 1").textEl!.classList.contains("is-empty")).toBe(true);
  });

  it("uses a background image when present", async () => {
    await render(0, true, (m) => (m.background = { color: null, image: "ppt/media/bg.png" }));
    expect(host.querySelector<HTMLElement>(".pptx-slide")!.style.backgroundImage).toContain("blob:ppt/media/bg.png");
  });

  it("draws custom geometry as SVG instead of a box", async () => {
    const { byName } = await render(2);
    const el = byName("Freeform").el;
    const path = el.querySelector("svg path")!;
    // drawn in pixels: 914400 EMU at 9525 EMU/px
    expect(path.getAttribute("d")).toBe("M0 0 L96 96 L0 96 Z");
    expect(path.getAttribute("fill")).toMatch(/^url\(#dw-grad-\d+\)$/);
    expect(el.style.background).toBe("");
    expect(el.style.borderStyle).toBe("");
  });

  it("strokes lines and mirrors flipped ones", async () => {
    const { byName } = await render(2);
    const svg = byName("Line").el.querySelector("svg")!;
    expect(svg.querySelector("path")!.getAttribute("stroke")).toBe("#4F81BD");
    expect(svg.querySelector("path")!.getAttribute("fill")).toBe("none");
    expect(svg.style.transform).toBe("scale(-1, 1)");
  });

  it("paints gradient fills", async () => {
    const { byName } = await render(2);
    const svg = byName("Gradient").el.querySelector("svg")!;
    const stops = Array.from(svg.querySelectorAll("linearGradient stop")).map((s) => [s.getAttribute("stop-color"), s.getAttribute("stop-opacity")]);
    expect(stops).toEqual([["#FF0000", "1"], ["#0000FF", String(0x80 / 255)]]);
    expect(svg.querySelector("path")!.getAttribute("fill")).toBe(`url(#${svg.querySelector("linearGradient")!.id})`);
  });

  it("only prompts on empty placeholders", async () => {
    const { byName } = await render(2);
    expect(byName("Oval 3").textEl!.hasAttribute("data-prompt")).toBe(false);
  });

  it("paints background-filled shapes with the slide background", async () => {
    const { byName } = await render(2);
    expect(byName("BgFill").el.querySelector("svg path")!.getAttribute("fill")).toBe("#FFFFFF");
  });

  it("draws SmartArt shapes with their outline and inset text box", async () => {
    const { shapes } = await render(2);
    const chevron = shapes.find((r) => r.shape.geom === "chevron")!;
    expect(chevron.el.querySelector("svg path")!.getAttribute("d")).toBe(
      scalePath("M0 0 L880000 0 L1000000 200000 L880000 400000 L0 400000 L120000 200000 Z", 1 / 9525, 1 / 9525)
    );
    expect(chevron.textEl!.style.left).toBe(`${120000 / 9525}px`);
    expect(chevron.textEl!.contentEditable).not.toBe("true");
  });
});
