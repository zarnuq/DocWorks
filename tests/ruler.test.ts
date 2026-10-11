import { describe, expect, it } from "vitest";
import { EditorState } from "prosemirror-state";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { cmd, tabAdvance } from "../src/presentation/docx/docx-editor";
import { NS } from "../src/core/utils/xml";
import { outlineEntries, rulerSection, setRulerMargins } from "../src/presentation/docx/docx-ruler";
import { docxSchema as s } from "../src/domain/docx/schema";

const context = { imageSrc: () => null, hyperlinkHref: () => null };
const parse = (xml: string) => new DOMParser().parseFromString(`<w:body xmlns:w="${NS.w}">${xml}</w:body>`, "application/xml").documentElement;
describe("ruler document properties", () => {
  it("positions text at left, center, right and decimal tab stops", () => {
    const stop = (align: string) => [{ pt: 72, align, leader: null }];
    expect(tabAdvance(stop("left"), 30, 40, 20)).toBe(66);
    expect(tabAdvance(stop("center"), 30, 40, 20)).toBe(46);
    expect(tabAdvance(stop("right"), 30, 40, 20)).toBe(26);
    expect(tabAdvance(stop("decimal"), 30, 40, 20)).toBe(46);
    expect(tabAdvance(stop("left"), 100, 40, 20)).toBe(44);
  });
  it("reads exact fractional and negative indents, right indent and explicit tab stops", () => {
    const body = parse(`<w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="1440"/><w:tab w:val="decimal" w:pos="3000" w:leader="dot"/></w:tabs><w:ind w:left="180" w:right="360" w:hanging="120"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>`);
    const { doc, origins } = readDocx(body, context);
    expect(doc.firstChild!.attrs).toMatchObject({ indent: 0.25, rightIndent: 18, firstLine: -6, tabs: [{ pt: 72, align: "center", leader: null }, { pt: 150, align: "decimal", leader: "dot" }] });
    let state = EditorState.create({ doc });
    cmd.setParagraph({ indent: -0.125, rightIndent: 27, tabs: [{ pt: 90, align: "right", leader: null }] })(state, tr => { state = state.apply(tr); });
    writeBody(body, state.doc, origins, { nextDocPrId: () => 1 });
    const again = readDocx(body, context).doc.firstChild!;
    expect(again.attrs).toMatchObject({ indent: -0.125, rightIndent: 27, tabs: [{ pt: 90, align: "right", leader: null }] });
  });
  it("changes margins in the cursor's section, preserving columns and other sections", () => {
    const body = parse(`<w:p><w:r><w:t>First</w:t></w:r><w:pPr><w:sectPr><w:pgMar w:left="1440" w:right="1440"/><w:cols w:num="2"/></w:sectPr></w:pPr></w:p><w:p><w:r><w:t>Last</w:t></w:r></w:p><w:sectPr><w:pgMar w:left="720" w:right="720"/></w:sectPr>`);
    const { doc } = readDocx(body, context);
    let state = EditorState.create({ doc });
    expect(rulerSection(state).setup.left).toBe(96);
    setRulerMargins(state, tr => { state = state.apply(tr); }, { left: 120, right: 72 });
    expect(rulerSection(state).setup).toMatchObject({ left: 120, right: 72 });
    expect(state.doc.attrs.sectPr).toContain('w:left="720"');
    expect(rulerSection(state).xml).toContain('w:num="2"');
  });
  it("provides heading positions for outline navigation, including headings inside columns", () => {
    const doc = s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text("Text")),
      s.nodes.columns.create(null, [s.nodes.paragraph.create({ style: "Heading2" }, s.text("Topic"))])]);
    expect(outlineEntries(doc)).toEqual([{ pos: 8, level: 2, text: "Topic" }]);
  });
});
