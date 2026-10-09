import { describe, it, expect } from "vitest";
import { DocxStyles, singleLine } from "../src/domain/docx/docx-styles";
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

  it("knows Word's single line height for common fonts (it depends on the font, unlike CSS line-height: 1)", () => {
    expect(singleLine("Garamond")).toBe(1.33);
    expect(singleLine("Calibri")).toBe(1.22);
    expect(singleLine("Times New Roman")).toBe(1.15);
    expect(singleLine("Some Unknown Font")).toBe(1.17);
    expect(singleLine(null)).toBe(1.17);
  });

  it("uses the default paragraph style when none is set", async () => {
    const { styles } = await setup();
    const css = styles.paragraphCss(null);
    expect(css).toContain("font-size: 11pt");
    expect(css).toContain("margin-bottom: 10pt");
    // 1.15 lines (276/240) of the font's single line.
    expect(css).toContain("line-height: 1.35");
  });

  it("layers a table style's paragraph settings between the defaults and the paragraph style", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const xml = new DOMParser().parseFromString(
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="252" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
        `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>` +
        `<w:style w:type="paragraph" w:styleId="Spaced"><w:pPr><w:spacing w:after="240"/></w:pPr></w:style>` +
        `<w:style w:type="table" w:styleId="TableGrid"><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:style></w:styles>`,
      "application/xml"
    );
    const styles = new DocxStyles(xml, null);
    expect(styles.paragraphCss(null)).toContain("margin-bottom: 8pt");
    // In a Table Grid table: no space after, single spacing (Word's usual tight table text).
    expect(styles.paragraphCss(null, "TableGrid")).toContain("margin-bottom: 0pt");
    expect(styles.paragraphCss(null, "TableGrid")).toContain("line-height: 1.17");
    // A paragraph style that sets its own spacing still wins.
    expect(styles.paragraphCss("Spaced", "TableGrid")).toContain("margin-bottom: 12pt");
  });

  it("returns only single line spacing for unknown styles without defaults", () => {
    expect(new DocxStyles(null, null).paragraphCss("Nope")).toBe("line-height: 1.17"); // Word still uses single spacing
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
