import { describe, expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addTextBox, writeParagraphs } from "../src/domain/pptx/slide-edit";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { linkLabel, linkTarget, setShapeLink } from "../src/domain/pptx/links";
import { roundTrip } from "./helpers";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

async function deck() {
  const doc = await PptxDocument.load(await blankPptx());
  doc.addSlide(0);
  doc.addSlide(1);
  const [a, b, c] = doc.slides().map((s) => s.path);
  const box = addTextBox(doc.xml(a)!, { x: 0, y: 0, w: 3000000, h: 1000000 }, "Go");
  return { doc, a, b, c, box };
}
const shapeOf = (doc: PptxDocument, path: string, el: Element) => parseSlide(doc, path).shapes.find((s) => s.el === el)!;

describe("links to slides", () => {
  it("writes a text link to another slide as a slide relationship with Office's jump action", async () => {
    const { doc, a, c, box } = await deck();
    const t = shapeOf(doc, a, box).text!;
    writeParagraphs(t, [{ ...t.paras[0], runs: [{ ...t.paras[0].runs[0], link: `#slide:${c}` }] }]);
    const back = await roundTrip(doc);
    const h = back.xml(a)!.getElementsByTagNameNS(A, "hlinkClick")[0];
    expect(h.getAttribute("action")).toBe("ppaction://hlinksldjump");
    expect(back.rels(a).find((r) => r.id === h.getAttribute("r:id"))).toMatchObject({ type: "slide", target: c, external: false });
    expect(parseSlide(back, a).shapes.find((s) => s.text?.paras[0]?.runs[0]?.link)!.text!.paras[0].runs[0].link).toBe(`#slide:${c}`);
  });

  it("writes next / previous / first / last slide jumps without a relationship", async () => {
    const { doc, a, box } = await deck();
    const t = shapeOf(doc, a, box).text!;
    writeParagraphs(t, [{ ...t.paras[0], runs: [{ ...t.paras[0].runs[0], link: "#jump:nextslide" }] }]);
    const h = doc.xml(a)!.getElementsByTagNameNS(A, "hlinkClick")[0];
    expect([h.getAttribute("r:id"), h.getAttribute("action")]).toEqual(["", "ppaction://hlinkshowjump?jump=nextslide"]);
    expect(shapeOf(doc, a, box).text!.paras[0].runs[0].link).toBe("#jump:nextslide");
  });

  it("links a whole shape (cNvPr hlinkClick first), reads it back and removes it", async () => {
    const { doc, a, b, box } = await deck();
    setShapeLink(doc, a, box, `#slide:${b}`);
    const cNvPr = box.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "cNvPr")[0];
    expect(cNvPr.firstElementChild?.localName).toBe("hlinkClick");
    expect(shapeOf(doc, a, box).link).toBe(`#slide:${b}`);
    setShapeLink(doc, a, box, "https://example.com/");
    expect(shapeOf(doc, a, box).link).toBe("https://example.com/");
    setShapeLink(doc, a, box, null);
    expect(shapeOf(doc, a, box).link).toBeNull();
    expect(cNvPr.getElementsByTagNameNS(A, "hlinkClick")).toHaveLength(0);
  });

  it("resolves a link to the slide index to show", async () => {
    const { doc, b } = await deck();
    const paths = doc.slides().map((s) => s.path);
    expect(linkTarget(`#slide:${b}`, 0, paths)).toBe(1);
    expect(linkTarget("#jump:nextslide", 1, paths)).toBe(2);
    expect(linkTarget("#jump:previousslide", 0, paths)).toBe(0);
    expect(linkTarget("#jump:firstslide", 2, paths)).toBe(0);
    expect(linkTarget("#jump:lastslide", 0, paths)).toBe(2);
    expect(linkTarget("#slide:ppt/slides/missing.xml", 0, paths)).toBeNull();
    expect(linkTarget("https://example.com/", 0, paths)).toBeNull();
  });

  it("labels links for the dialog", async () => {
    const { doc, c } = await deck();
    const paths = doc.slides().map((s) => s.path);
    expect(linkLabel(`#slide:${c}`, paths)).toBe("Slide 3");
    expect(linkLabel("#jump:lastslide", paths)).toBe("Last slide");
    expect(linkLabel("https://x.y/", paths)).toBe("https://x.y/");
  });

  it("drops links to a deleted slide", async () => {
    const { doc, a, c, box } = await deck();
    setShapeLink(doc, a, box, `#slide:${c}`);
    doc.deleteSlide(2);
    expect(box.getElementsByTagNameNS(A, "hlinkClick")).toHaveLength(0);
    expect(doc.rels(a).some((r) => r.type === "slide")).toBe(false);
  });

  it("keeps slide links when copying within a deck, drops them when pasting into another", async () => {
    const { doc, a, c, box } = await deck();
    setShapeLink(doc, a, box, `#slide:${c}`);
    const clip = await doc.copySlides([0]);
    const at = doc.pasteSlides(clip, 2);
    const copy = doc.slides()[at].path;
    const copied = parseSlide(doc, copy).shapes.find((s) => s.link)!;
    expect(copied.link).toBe(`#slide:${c}`);
    expect(doc.slides()).toHaveLength(4);
    const other = await PptxDocument.load(await blankPptx());
    const index = other.pasteSlides(clip, 0);
    const pasted = other.slides()[index].path;
    expect(other.xml(pasted)!.getElementsByTagNameNS(A, "hlinkClick")).toHaveLength(0);
    expect(other.slides()).toHaveLength(2);
  });
});

describe("drawing linked shapes", () => {
  it("marks a linked shape with its link, so clicks on it follow the link", async () => {
    const { renderSlide } = await import("../src/presentation/pptx/slide-renderer");
    const { doc, a, b, box } = await deck();
    setShapeLink(doc, a, box, `#slide:${b}`);
    const host = document.body.createDiv();
    const shapes = renderSlide(host, parseSlide(doc, a), doc.slideSize, { width: 960, editable: false, imageUrl: () => null });
    expect(shapes.find((r) => r.shape.el === box)!.el.dataset.link).toBe(`#slide:${b}`);
    host.remove();
  });
});
