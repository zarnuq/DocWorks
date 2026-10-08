import { describe, it, expect } from "vitest";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { blockPos, readSample } from "./helpers";

async function setup() {
  const s = await readSample();
  return { ...s, styles: new DocxStyles(s.docx.stylesXml(), s.docx.numberingXml()) };
}

describe("DocxStyles.paragraphCss", () => {
  it("resolves a style through basedOn and docDefaults", async () => {
    const { styles } = await setup();
    const css = styles.paragraphCss("Heading1");
    expect(css).toContain("font-size: 14pt");
    expect(css).toContain("font-weight: bold");
    expect(css).toContain("color: #365F91");
    expect(css).toContain("margin-top: 24pt");
  });

  it("uses the default paragraph style when none is set", async () => {
    const { styles } = await setup();
    const css = styles.paragraphCss(null);
    expect(css).toContain("font-size: 11pt");
    expect(css).toContain("margin-bottom: 10pt");
    expect(css).toContain("line-height: 1.15");
  });

  it("returns nothing for unknown styles without defaults", () => {
    expect(new DocxStyles(null, null).paragraphCss("Nope")).toBe("");
  });
});

describe("DocxStyles.listMarks", () => {
  it("labels style-based bullets and numbers", async () => {
    const { doc, styles } = await setup();
    const marks = styles.listMarks(doc);
    expect(marks.get(blockPos(doc, 3))?.text).toBe("•");
    expect(marks.get(blockPos(doc, 4))?.text).toBe("•");
    expect(marks.get(blockPos(doc, 5))?.text).toBe("1.");
    expect(marks.has(blockPos(doc, 0))).toBe(false);
  });

  it("counts per list and level, restarting deeper levels", async () => {
    const { docx } = await setup();
    const numId = docx.ensureList("ordered");
    const s = new DocxStyles(docx.stylesXml(), docx.numberingXml());
    const p = (ilvl: number) => schema.nodes.paragraph.create({ numId, ilvl }, schema.text("x"));
    const doc = schema.nodes.doc.create(null, [p(0), p(1), p(1), p(0), p(1)]);
    const labels = Array.from(s.listMarks(doc).values()).map((m) => m.text);
    expect(labels).toEqual(["1.", "a.", "b.", "2.", "a."]);
  });

  it("indents by the level's w:ind", async () => {
    const { docx } = await setup();
    const numId = docx.ensureList("bullet");
    const s = new DocxStyles(docx.stylesXml(), docx.numberingXml());
    const doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create({ numId, ilvl: 1 }, schema.text("x"))]);
    expect(s.listMarks(doc).get(0)).toEqual({ text: "◦", indentPt: 72, hangingPt: 18 });
  });
});
