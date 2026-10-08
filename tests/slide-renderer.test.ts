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

  it("renders tables", async () => {
    const { byName } = await render(2);
    const cells = Array.from(byName("Table 5").el.querySelectorAll("td")).map((td) => td.textContent);
    expect(cells).toEqual(["A1", "B1", "A2", "B2"]);
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
    expect(byName("Oval 3").textEl!.classList.contains("is-empty")).toBe(false);
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
