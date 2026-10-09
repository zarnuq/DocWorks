import { describe, it, expect } from "vitest";
import { blankDocx, blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { roundTrip, W } from "./helpers";

const phType = (el: Element) => el.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "ph")[0]?.getAttribute("type") ?? "body";

describe("blank presentation", () => {
  it("opens as a 16:9 deck with one title slide", async () => {
    const doc = await PptxDocument.load(await blankPptx());
    expect(doc.slideSize).toEqual({ cx: 12192000, cy: 6858000 });
    expect(doc.slides()).toHaveLength(1);
    const shapes = parseSlide(doc, doc.slides()[0].path).shapes;
    expect(shapes.map((s) => phType(s.el))).toEqual(["ctrTitle", "subTitle"]);
    expect(shapes.every((s) => s.rect.w > 0 && s.rect.h > 0)).toBe(true);
  });

  it("adds title-and-content slides after the title slide", async () => {
    const doc = await PptxDocument.load(await blankPptx());
    doc.addSlide(0);
    const back = await roundTrip(doc);
    const shapes = parseSlide(back, back.slides()[1].path).shapes;
    expect(shapes.map((s) => phType(s.el))).toEqual(["title", "body"]);
  });
});

describe("blank document", () => {
  it("opens with one empty paragraph and Normal style", async () => {
    const doc = await DocxDocument.load(await blankDocx());
    expect(doc.body.getElementsByTagNameNS(W, "p")).toHaveLength(1);
    doc.ensureParagraphStyles(["Heading1"]);
    const ids = Array.from(doc.stylesXml()!.getElementsByTagNameNS(W, "style")).map((s) => s.getAttributeNS(W, "styleId"));
    expect(ids).toEqual(expect.arrayContaining(["Normal", "Heading1"]));
  });
});
