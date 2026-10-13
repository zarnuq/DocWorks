import { describe, expect, it } from "vitest";
import { blankPptx } from "../src/infrastructure/blank";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";
import { addShape, addTextBox, writeParagraphs } from "../src/domain/pptx/slide-edit";
import { parseSlide } from "../src/domain/pptx/slide-model";
import { writeFill, writeLine, writeShadow } from "../src/domain/pptx/shape-style";
import { applyFormat, copyFormat } from "../src/domain/pptx/paint-format";

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const spPr = (el: Element) => el.getElementsByTagNameNS(P, "spPr")[0];

async function shapes() {
  const doc = await PptxDocument.load(await blankPptx());
  const path = doc.slides()[0].path;
  const xml = doc.xml(path)!;
  const src = addShape(xml, "ellipse", { x: 0, y: 0, w: 1000000, h: 1000000 });
  const dst = addTextBox(xml, { x: 2000000, y: 0, w: 2000000, h: 500000 }, "Target text");
  const model = (el: Element) => parseSlide(doc, path).shapes.find((s) => s.el === el)!;
  return { src, dst, model };
}

describe("paint format", () => {
  it("copies fill, outline, shadow and text style onto another shape", async () => {
    const { src, dst, model } = await shapes();
    writeFill(spPr(src), { type: "solid", color: "#FF0000" });
    writeLine(spPr(src), { color: "#0000FF", width: 38100, dash: "dash" });
    writeShadow(spPr(src), { color: "#00000080", blur: 50800, dist: 38100, dir: 45 });
    const t = model(src).text!;
    writeParagraphs(t, [{ ...t.paras[0], align: "ctr", runs: [{ ...t.paras[0].defaults, text: "Source", b: true, sz: 24, color: "#FF00FF", font: "Arial" }] }]);
    const format = copyFormat(model(src));
    applyFormat(model(dst), format);
    const after = model(dst);
    expect(after.fill).toEqual({ type: "solid", color: "#FF0000" });
    expect(after.line).toMatchObject({ color: "#0000FF", width: 38100, dash: "dash" });
    expect(after.shadow).toMatchObject({ dist: 38100 });
    expect(after.text!.paras[0]).toMatchObject({ align: "ctr" });
    expect(after.text!.paras[0].runs[0]).toMatchObject({ text: "Target text", b: true, sz: 24, color: "#FF00FF", font: "Arial" });
  });

  it("removes fill and outline the source doesn't have", async () => {
    const { src, dst, model } = await shapes();
    writeFill(spPr(dst), { type: "solid", color: "#00FF00" });
    writeFill(spPr(src), null);
    writeLine(spPr(src), null);
    applyFormat(model(dst), copyFormat(model(src)));
    expect(model(dst).fill).toBeNull();
    expect(model(dst).line).toBeNull();
  });
});
