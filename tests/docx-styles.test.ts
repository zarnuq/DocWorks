import { describe, it, expect } from "vitest";
import { DocxStyles, lineHeight, singleLine } from "../src/domain/docx/docx-styles";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DOMSerializer } from "prosemirror-model";
import { readDocx } from "../src/domain/docx/docx-reader";
import { styleRenderedParagraphs } from "../src/presentation/docx/docx-editor";
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
    // From the fonts' own metrics (scripts/make_line_heights.py), not rounded to two places.
    expect(singleLine("Calibri")).toBe(1.2207);
    expect(singleLine("Times New Roman")).toBe(1.1499);
    expect(singleLine("Verdana")).toBe(1.2153);
    expect(singleLine("Comic Sans MS")).toBe(1.3936);
    expect(singleLine("Some Unknown Font")).toBe(1.17);
    expect(singleLine(null)).toBe(1.17);
    // Multiple line spacing keeps that precision: a page of lines adds up.
    expect(lineHeight(1.15, "Calibri")).toBe(1.4038);
  });

  it("uses the default paragraph style when none is set", async () => {
    const { styles } = await setup();
    const css = styles.paragraphCss(null);
    expect(css).toContain("font-size: 11pt");
    expect(css).toContain("margin-bottom: 10pt");
    // 1.15 lines (276/240) of the font's single line.
    expect(css).toContain("line-height: 1.3455");
  });

  it("indents paragraphs from their style (table of contents levels)", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const xml = new DOMParser().parseFromString(
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="TOC2"><w:pPr><w:ind w:left="220" w:right="440"/></w:pPr></w:style>` +
        `<w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="TOC2"/></w:style></w:styles>`,
      "application/xml"
    );
    const styles = new DocxStyles(xml, null);
    expect(styles.paragraphCss("Child")).toContain("margin-left: 11pt; margin-right: 22pt");
    expect(styles.paragraphCss(null)).not.toContain("margin-left");
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

describe("line height follows the text, not the style", () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  // Style: 12pt (unknown font, single line 1.17), 279/240 lines - like a Google Docs export.
  const styles = new DocxStyles(
    new DOMParser().parseFromString(
      `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:line="279" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
      "application/xml"
    ),
    null
  );
  const render = (body: string) => {
    const xml = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${body}</w:body>`, "application/xml").documentElement;
    const { doc } = readDocx(xml, { imageSrc: () => null, hyperlinkHref: () => null });
    const div = document.createElement("div");
    div.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content, { document }));
    styleRenderedParagraphs(div, doc, styles);
    return Array.from(div.querySelectorAll("p")).map((p) => p.getAttribute("style") ?? "");
  };
  const arial10 = `<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr>`;

  it("sizes the line by the paragraph mark's font (10pt Arial text in a 12pt style)", () => {
    const [css] = render(`<w:p><w:pPr>${arial10}</w:pPr><w:r>${arial10}<w:t>Text</w:t></w:r></w:p>`);
    // 279/240 lines of Arial's single line (1.1499) at 10pt, not of 12pt.
    expect(css).toMatch(/font-size: 10pt; font-family: "Arial", "Liberation Sans", Arimo, sans-serif; line-height: 1\.3368\b/);
  });

  it("uses the text's size when every run sets it and the mark doesn't", () => {
    const [css] = render(`<w:p><w:r>${arial10}<w:t>a</w:t></w:r><w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>b</w:t></w:r></w:p>`);
    expect(css).toContain("font-size: 11pt");
  });

  it("doesn't size text with no size of its own by the paragraph mark (it keeps the style's)", () => {
    // Word: the mark's 14pt only affects the mark; the text is 12pt from the style.
    const [css] = render(`<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>`);
    expect(css).not.toContain("font-size: 14pt");
  });

  it("sizes an empty paragraph by its mark", () => {
    const [css] = render(`<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>`);
    expect(css).toContain("font-size: 14pt");
  });

  it("keeps the style's size when some text has none of its own", () => {
    const [css] = render(`<w:p><w:r>${arial10}<w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>`);
    expect(css).not.toContain("font-size: 10pt");
  });
});

describe("table borders from the table style", () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const styles = new DocxStyles(
    new DOMParser().parseFromString(
      `<w:styles xmlns:w="${W}"><w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders>` +
        ["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4"/>`).join("") +
        `</w:tblBorders></w:tblPr></w:style><w:style w:type="table" w:styleId="Plain"><w:basedOn w:val="TableNormal"/></w:style>` +
        `<w:style w:type="table" w:styleId="Roomy"><w:tblPr><w:tblCellMar><w:top w:w="72" w:type="dxa"/><w:left w:w="144" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
        `<w:style w:type="table" w:styleId="RoomyGrid"><w:basedOn w:val="Roomy"/><w:tblPr><w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`,
      "application/xml"
    ),
    null
  );

  it("has no borders without a style (Word's Table Normal), else the style's", () => {
    expect(styles.tableBorders(null)).toBe("none");
    expect(styles.tableBorders("Plain")).toBe("none");
    expect(styles.tableBorders("TableGrid")).toBe("all");
  });

  it("pads cells by the table style's cell margins, nearest style first", () => {
    expect(styles.cellPaddingCss(null)).toBe("");
    expect(styles.cellPaddingCss("TableGrid")).toBe("");
    expect(styles.cellPaddingCss("Roomy")).toBe("--docx-cell-top: 3.6pt; --docx-cell-left: 7.2pt");
    expect(styles.cellPaddingCss("RoomyGrid")).toBe("--docx-cell-top: 3.6pt; --docx-cell-left: 0pt");
  });

  it("gives rendered tables their style's cell padding", () => {
    const xml = new DOMParser().parseFromString(`<w:body xmlns:w="${W}"><w:tbl><w:tblPr><w:tblStyle w:val="Roomy"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body>`, "application/xml").documentElement;
    const { doc } = readDocx(xml, { imageSrc: () => null, hyperlinkHref: () => null });
    const div = document.createElement("div");
    div.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content, { document }));
    styleRenderedParagraphs(div, doc, styles);
    expect(div.querySelector("table")!.style.getPropertyValue("--docx-cell-left")).toBe("7.2pt");
  });

  it("draws a table without borders of its own with its style's (a header's empty layout table shows none)", () => {
    const xml = new DOMParser().parseFromString(
      `<w:body xmlns:w="${W}"><w:tbl><w:tblPr/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body>`,
      "application/xml"
    ).documentElement;
    const { doc } = readDocx(xml, { imageSrc: () => null, hyperlinkHref: () => null });
    const div = document.createElement("div");
    div.appendChild(DOMSerializer.fromSchema(schema).serializeFragment(doc.content, { document }));
    styleRenderedParagraphs(div, doc, styles);
    expect(Array.from(div.querySelectorAll("table")).map((t) => t.className)).toEqual(["docx-borders-none", "docx-borders-all"]);
  });
});
