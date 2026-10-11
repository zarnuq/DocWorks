import { describe, it, expect } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { dropIndex, PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { addPicture, addShape, setRect, writePlainParagraphs } from "../src/domain/pptx/slide-edit";
import { itemRect } from "../src/domain/pptx/arrange";
import { loadSample, roundTrip, texts } from "./helpers";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

const blank = async () => PptxDocument.load(await blankPptx());
const phTypes = (doc: PptxDocument, i = 0) =>
  parseSlide(doc, doc.slides()[i].path).shapes.map((s) => s.el.getElementsByTagNameNS(P, "ph")[0]?.getAttribute("type") ?? "body");
const placeholderTxBody = (doc: PptxDocument, i: number, type: string) =>
  Array.from(doc.xml(doc.slides()[i].path)!.getElementsByTagNameNS(P, "sp"))
    .find((sp) => sp.getElementsByTagNameNS(P, "ph")[0]?.getAttribute("type") === type)!
    .getElementsByTagNameNS(P, "txBody")[0];

describe("layouts", () => {
  it("lists the master's layouts by name", async () => {
    const doc = await blank();
    expect(doc.layouts().map((l) => [l.name, l.type])).toEqual([
      ["Title Slide", "title"],
      ["Title and Content", "obj"],
      ["Blank", "blank"],
    ]);
    expect(doc.layoutOf(doc.slides()[0].path)).toBe(doc.layouts()[0].path);
  });

  it("changes a slide's layout, keeping placeholder text and adding the new layout's placeholders", async () => {
    const doc = await blank();
    writePlainParagraphs(placeholderTxBody(doc, 0, "ctrTitle"), "My talk");
    doc.setLayout(0, doc.layouts()[1].path);
    const back = await roundTrip(doc);
    expect(back.layoutOf(back.slides()[0].path)).toBe(back.layouts()[1].path);
    expect(phTypes(back)).toEqual(["title", "body"]);
    expect(texts(back.xml(back.slides()[0].path)!)).toEqual(["My talk"]);
  });

  it("drops empty placeholders the new layout lacks and pins filled ones where they were", async () => {
    const doc = await blank();
    writePlainParagraphs(placeholderTxBody(doc, 0, "ctrTitle"), "Keep me");
    const before = parseSlide(doc, doc.slides()[0].path).shapes[0].rect;
    doc.setLayout(0, doc.layouts()[2].path);
    const shapes = parseSlide(doc, doc.slides()[0].path).shapes;
    expect(shapes).toHaveLength(1);
    expect(shapes[0].rect).toEqual(before);
    expect(texts(shapes[0].el)).toEqual(["Keep me"]);
  });
});

describe("layout positions", () => {
  it("moves matched placeholders to the new layout's position", async () => {
    const doc = await blank();
    const title = parseSlide(doc, doc.slides()[0].path).shapes[0].el;
    setRect(title, { x: 1, y: 2, w: 3, h: 4 });
    doc.setLayout(0, doc.layouts()[1].path);
    expect(parseSlide(doc, doc.slides()[0].path).shapes[0].rect).toEqual({ x: 838200, y: 365125, w: 10515600, h: 1325563 });
  });
});

describe("background", () => {
  it("sets a solid or gradient slide background and resets it to the master's", async () => {
    const doc = await blank();
    const path = doc.slides()[0].path;
    doc.setBackground(0, { type: "solid", color: "#FF0000" });
    const cSld = doc.xml(path)!.getElementsByTagNameNS(P, "cSld")[0];
    expect(cSld.firstElementChild!.localName).toBe("bg");
    expect(parseSlide(doc, path).background.color).toBe("#FF0000");
    expect(parseSlide(doc, path).ownBackground).toEqual({ type: "solid", color: "#FF0000" });

    doc.setBackground(0, { type: "gradient", radial: false, angle: 90, stops: [{ pos: 0, color: "#000000" }, { pos: 1, color: "#FFFFFF" }] });
    expect(cSld.getElementsByTagNameNS(A, "gradFill")).toHaveLength(1);
    expect(cSld.getElementsByTagNameNS(A, "solidFill")).toHaveLength(0);
    expect(parseSlide(doc, path).background.color).toMatch(/linear-gradient/);

    doc.setBackground(0, null);
    expect(cSld.getElementsByTagNameNS(P, "bg")).toHaveLength(0);
    expect(parseSlide(doc, path).background.color).toBe("#FFFFFF");
    expect(parseSlide(doc, path).ownBackground).toBeNull();
    await roundTrip(doc);
  });
});

describe("slide size", () => {
  const layoutTitle = (doc: PptxDocument) => {
    const sp = Array.from(doc.xml(doc.layouts()[1].path)!.getElementsByTagNameNS(P, "sp"))[0];
    return itemRect(sp);
  };

  it("stretches layout and master placeholders to the new size", async () => {
    const doc = await blank();
    const before = layoutTitle(doc);
    doc.setSlideSize(9144000, 6858000);
    expect(doc.slideSize).toEqual({ cx: 9144000, cy: 6858000 });
    const r = layoutTitle(doc);
    expect(r.x).toBeCloseTo(before.x * 0.75, -1);
    expect(r.w).toBeCloseTo(before.w * 0.75, -1);
    expect([r.y, r.h]).toEqual([before.y, before.h]);
    // Going back restores the layout.
    doc.setSlideSize(12192000, 6858000);
    const again = layoutTitle(doc);
    expect(again.x).toBeCloseTo(before.x, -1);
    expect(again.w).toBeCloseTo(before.w, -1);
    const back = await roundTrip(doc);
    expect(back.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
  });

  it("scales slide content and pictures uniformly and centers them", async () => {
    const doc = await blank();
    const path = doc.slides()[0].path;
    const box = addShape(doc.xml(path)!, "ellipse", { x: 1000000, y: 2000000, w: 4000000, h: 2000000 });
    const logo = addPicture(doc.xml(doc.layouts()[0].path)!, { x: 0, y: 0, w: 1000000, h: 1000000 }, "rId99", "logo.png");
    doc.setSlideSize(9144000, 6858000);
    // As large as it can be while staying on the slide: here its left edge meets the slide's.
    const r = itemRect(box);
    expect(r.x).toBe(0);
    expect(r.w / r.h).toBeCloseTo(2, 3);
    expect(r.w).toBeGreaterThan(3000000);
    expect(r.w).toBeLessThan(4000000);
    expect(r.y + r.h / 2 - 6858000 / 2).toBeCloseTo((3000000 - 6858000 / 2) * (r.w / 4000000), -2);
    // A picture keeps its aspect ratio; its center moves with the stretch.
    expect(itemRect(logo)).toEqual({ x: 0, y: 125000, w: 750000, h: 750000 });
  });

  it("shrinks text with the boxes, and switching back restores the slide instead of shrinking it again", async () => {
    const doc = await blank();
    const path = doc.slides()[0].path;
    const box = addShape(doc.xml(path)!, "rect", { x: 0, y: 1000000, w: 12192000, h: 2000000 });
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const rPr = doc.xml(path)!.createElementNS(A, "a:endParaRPr");
    rPr.setAttribute("sz", "4000");
    box.getElementsByTagNameNS(A, "p")[0].appendChild(rPr);
    const before = itemRect(box);
    doc.setSlideSize(9144000, 6858000);
    expect(itemRect(box).w).toBe(9144000);
    expect(rPr.getAttribute("sz")).toBe("3000");
    doc.setSlideSize(12192000, 6858000);
    const after = itemRect(box);
    expect(after.x).toBeCloseTo(before.x, -2);
    expect(after.y).toBeCloseTo(before.y, -2);
    expect(after.w).toBeCloseTo(before.w, -2);
    expect(after.h).toBeCloseTo(before.h, -2);
    expect(rPr.getAttribute("sz")).toBe("4000");
  });
});

describe("sections", () => {
  async function fourSlides() {
    const doc = await blank();
    for (let i = 0; i < 3; i++) doc.addSlide(i);
    return doc;
  }
  const sectionSlides = (doc: PptxDocument) =>
    Array.from(doc.xml("ppt/presentation.xml")!.getElementsByTagNameNS(P14, "section")).map((s) => [
      s.getAttribute("name"),
      Array.from(s.getElementsByTagNameNS(P14, "sldId")).map((id) => id.getAttribute("id")),
    ]);

  it("adds a section, putting earlier slides in a default section", async () => {
    const doc = await fourSlides();
    expect(doc.sections()).toEqual([]);
    doc.addSection(2, "Part two");
    expect(doc.sections().map((s) => [s.name, s.start])).toEqual([
      ["Default Section", 0],
      ["Part two", 2],
    ]);
    const ids = doc.slides().map((s) => s.id);
    expect(sectionSlides(await roundTrip(doc))).toEqual([
      ["Default Section", ids.slice(0, 2)],
      ["Part two", ids.slice(2)],
    ]);
  });

  it("renames and removes sections, merging their slides into the one before", async () => {
    const doc = await fourSlides();
    doc.addSection(0, "Intro");
    doc.addSection(2, "Body");
    doc.renameSection(1, "Main");
    expect(doc.sections().map((s) => s.name)).toEqual(["Intro", "Main"]);
    doc.removeSection(1);
    expect(doc.sections().map((s) => [s.name, s.start])).toEqual([["Intro", 0]]);
    expect(sectionSlides(doc)[0][1]).toHaveLength(4);
    doc.removeSection(0);
    expect(doc.sections()).toEqual([]);
    expect(doc.xml("ppt/presentation.xml")!.getElementsByTagNameNS(P14, "sectionLst")).toHaveLength(0);
  });

  it("keeps sections in step with adding, duplicating, deleting and moving slides", async () => {
    const doc = await fourSlides();
    doc.addSection(0, "A");
    doc.addSection(2, "B");
    doc.addSlide(2); // into B, after its first slide
    doc.duplicateSlide(0); // into A
    expect(doc.sections().map((s) => s.start)).toEqual([0, 3]);
    doc.deleteSlide(0);
    expect(doc.sections().map((s) => s.start)).toEqual([0, 2]);
    // Move the last slide to the front: it joins A.
    doc.moveSlide(doc.slides().length - 1, 0);
    expect(doc.sections().map((s) => s.start)).toEqual([0, 3]);
    const all = sectionSlides(doc).flatMap(([, ids]) => ids);
    expect(all).toEqual(doc.slides().map((s) => s.id));
  });

  it("can hold empty sections", async () => {
    const doc = await fourSlides();
    doc.addSection(0, "A");
    doc.addSection(4, "Empty end");
    expect(doc.sections().map((s) => [s.name, s.start])).toEqual([
      ["A", 0],
      ["Empty end", 4],
    ]);
    expect(sectionSlides(doc)[1][1]).toEqual([]);
  });
});

describe("sample deck", () => {
  it("changes layouts in a real deck without losing slides", async () => {
    const doc = await loadSample();
    const count = doc.slides().length;
    for (const l of doc.layouts()) doc.setLayout(1, l.path);
    const back = await roundTrip(doc);
    expect(back.slides()).toHaveLength(count);
  });
});

describe("dropIndex", () => {
  it("turns a drop above or below a thumbnail into a moveSlide target", () => {
    expect(dropIndex(0, 2, false)).toBe(1); // above slide 3: lands second
    expect(dropIndex(0, 2, true)).toBe(2); // below slide 3
    expect(dropIndex(3, 0, false)).toBe(0); // above the first
    expect(dropIndex(3, 1, true)).toBe(2);
    expect(dropIndex(2, 2, true)).toBe(2); // onto itself
    expect(dropIndex(2, 3, false)).toBe(2); // just below itself
  });
});
