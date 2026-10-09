import { describe, it, expect } from "vitest";
import { findAll, replaceAll, replaceMatch } from "../src/domain/pptx/find";
import { setCrop } from "../src/domain/pptx/slide-edit";
import { itemRect } from "../src/domain/pptx/arrange";
import { addTable } from "../src/domain/pptx/table";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { renderSlide } from "../src/presentation/pptx/slide-renderer";
import { rangeInParagraph } from "../src/presentation/pptx/text-html";
import { buildShapePanel, PanelControl } from "../src/presentation/pptx/shape-panel";
import { loadSample, roundTrip, texts } from "./helpers";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** Put `runs` (one a:r each) into the first paragraph of a text box on slide 2. */
async function withRuns(...runs: string[]) {
  const doc = await loadSample();
  const path = doc.slides()[1].path;
  const sp = Array.from(doc.xml(path)!.getElementsByTagNameNS(P, "sp"))[0];
  const p = sp.getElementsByTagNameNS(A, "p")[0];
  p.innerHTML = runs.map((t) => `<a:r xmlns:a="${A}"><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r>`).join("");
  return { doc, path, sp, p };
}

describe("find and replace", () => {
  it("finds text across runs, slides and table cells", async () => {
    const { doc, sp } = await withRuns("Hel", "lo wor", "ld");
    const tc = doc.xml(doc.slides()[2].path)!.getElementsByTagNameNS(A, "tc")[0];
    tc.getElementsByTagNameNS(A, "t")[0].textContent = "hello world table";
    const hits = findAll(doc, "hello world", { matchCase: false });
    expect(hits.map((h) => [h.slide, h.owner, h.start, h.end])).toEqual([
      [1, sp, 0, 11],
      [2, tc, 0, 11],
    ]);
    expect(findAll(doc, "hello world", { matchCase: true })).toHaveLength(1);
  });

  it("replaces a match that spans runs, keeping the formatting of the run it starts in", async () => {
    const { doc, p } = await withRuns("Hel", "lo wor", "ld");
    p.getElementsByTagNameNS(A, "rPr")[1].setAttribute("b", "1");
    const [hit] = findAll(doc, "lo world", { matchCase: true });
    replaceMatch(hit, "p!");
    expect(texts(p)).toEqual(["Hel", "p!"]);
    expect(p.getElementsByTagNameNS(A, "r")).toHaveLength(2);
    expect(p.getElementsByTagNameNS(A, "rPr")[1].getAttribute("b")).toBe("1");
  });

  it("replaces every match and counts them", async () => {
    const { doc, p } = await withRuns("a cat, a CAT, a c", "at");
    expect(replaceAll(doc, "cat", "dog", { matchCase: false })).toBe(3);
    expect(texts(p).join("")).toBe("a dog, a dog, a dog");
    const back = await roundTrip(doc);
    expect(findAll(back, "dog", { matchCase: true })).toHaveLength(3);
  });

  it("ignores an empty query", async () => {
    const { doc } = await withRuns("x");
    expect(findAll(doc, "", { matchCase: false })).toEqual([]);
  });
});

describe("image crop", () => {
  const pic = async () => {
    const doc = await loadSample();
    const path = doc.slides()[2].path;
    const el = Array.from(doc.xml(path)!.getElementsByTagNameNS(P, "pic"))[0];
    const model = () => parseSlide(doc, path).shapes.find((s) => s.el === el)!;
    return { doc, path, el, model };
  };

  it("crops keeping the picture's scale, and uncrops back", async () => {
    const { el, model } = await pic();
    const before = itemRect(el);
    setCrop(el, { l: 0.25, t: 0, r: 0, b: 0.5 });
    expect(model().crop).toEqual({ l: 0.25, t: 0, r: 0, b: 0.5 });
    expect(itemRect(el)).toEqual({ x: before.x + before.w * 0.25, y: before.y, w: before.w * 0.75, h: before.h * 0.5 });
    setCrop(el, { l: 0, t: 0, r: 0, b: 0 });
    expect(itemRect(el)).toEqual(before);
    expect(el.getElementsByTagNameNS(A, "srcRect")).toHaveLength(0);
    expect(model().crop).toBeNull();
  });

  it("survives a file whose crop leaves nothing visible", async () => {
    const { el } = await pic();
    setCrop(el, { l: 0.5, t: 0, r: 0, b: 0 });
    el.getElementsByTagNameNS(A, "srcRect")[0].setAttribute("r", "50000");
    setCrop(el, { l: 0, t: 0, r: 0, b: 0 });
    expect(Object.values(itemRect(el)).every(Number.isFinite)).toBe(true);
  });

  it("draws only the cropped part of the image", async () => {
    const { el, model, doc } = await pic();
    setCrop(el, { l: 0.5, t: 0, r: 0, b: 0 });
    const host = document.createElement("div");
    const shapes = renderSlide(host, { ...parseSlide(doc, doc.slides()[2].path) }, doc.slideSize, { width: 960, editable: false, imageUrl: () => "x" });
    const img = shapes.find((s) => s.shape.el === el)!.el.querySelector("img")!;
    expect([img.style.width, img.style.height, img.style.left, img.style.top]).toEqual(["200%", "100%", "-100%", "0%"]);
    expect(model().crop!.l).toBe(0.5);
  });

  it("has crop fields in the picture's format panel", async () => {
    const { el, model } = await pic();
    const host = document.createElement("div");
    const ctl: PanelControl = {
      lockAspect: false,
      edit(fn) {
        fn(el);
        buildShapePanel(host, model(), ctl);
      },
    };
    buildShapePanel(host, model(), ctl);
    const f = host.querySelector<HTMLInputElement>('[data-field="Crop left %"]')!;
    f.value = "10";
    f.dispatchEvent(new Event("change"));
    expect(model().crop).toMatchObject({ l: 0.1 });
    Array.from(host.querySelectorAll("button")).find((b) => b.textContent === "Reset crop")!.click();
    expect(model().crop).toBeNull();
  });
});

describe("table text is searchable", () => {
  it("finds text typed into a new table", async () => {
    const doc = await loadSample();
    const xml = doc.xml(doc.slides()[0].path)!;
    const frame = addTable(xml, 1, 1, { x: 0, y: 0, w: 100, h: 100 }, null);
    frame.getElementsByTagNameNS(A, "p")[0].innerHTML = `<a:r xmlns:a="${A}"><a:t>needle</a:t></a:r>`;
    expect(findAll(doc, "needle", { matchCase: true })[0].owner.localName).toBe("tc");
  });
});

describe("rangeInParagraph", () => {
  it("maps paragraph offsets to a DOM range, counting line breaks as one character", () => {
    const p = document.createElement("p");
    p.innerHTML = "<span>ab</span><br><span>cd<b>ef</b></span>";
    const r = rangeInParagraph(p, 1, 6);
    expect(r.toString()).toBe("bcde");
    expect(rangeInParagraph(p, 3, 5).toString()).toBe("cd");
  });
});
