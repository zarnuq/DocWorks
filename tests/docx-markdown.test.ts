import { describe, expect, it } from "vitest";
import { Node as PMNode } from "prosemirror-model";
import { docxToMarkdown } from "../src/domain/docx/to-markdown";
import { docxSchema as s } from "../src/domain/docx/schema";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { blankDocx } from "../src/infrastructure/blank";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { noteToDocx } from "../src/domain/docx/from-html";
import { latexToOmml } from "../src/domain/docx/math";
import { fixture } from "./helpers";

const p = (text: string, attrs = {}) => s.nodes.paragraph.create(attrs, text ? s.text(text) : null);
const convert = async (...nodes: PMNode[]) => {
  const docx = await DocxDocument.load(await blankDocx());
  return (await docxToMarkdown(docx, "Note attachments", s.nodes.doc.create(null, nodes))).markdown;
};

describe("DOCX to Markdown", () => {
  it("restores checked and unchecked task lists, including nested items", async () => {
    const root = createDiv();
    root.innerHTML = '<ul><li><input type="checkbox" checked>Done<ul><li><input type="checkbox">Child</li></ul></li><li>☑ ordinary symbol</li></ul>';
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => null));
    expect((await docxToMarkdown(docx, "images")).markdown).toBe('- [x] Done\n    - [ ] Child\n- ☑ ordinary symbol\n');
  });

  it("preserves aligned table columns", async () => {
    const root = createDiv();
    root.innerHTML = '<table><tr><th>Name</th><th align="right">Value</th><th align="center">Unit</th></tr><tr><td>A</td><td align="right">1</td><td align="center">m</td></tr></table>';
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => null));
    expect((await docxToMarkdown(docx, "images")).markdown).toContain('| --- | ---: | :---: |');
  });

  it("uses one blank line around display equations", async () => {
    const math = s.nodes.opaque.create({ xml: latexToOmml("x^2", true), text: "" });
    expect(await convert(p("Before"), s.nodes.paragraph.create(null, math), p("After"))).toBe('Before\n\n$$\nx^2\n$$\n\nAfter\n');
  });

  it("preserves code language, callout type/title/folding, properties and horizontal rules", async () => {
    const root = createDiv();
    root.dataset.exportFrontmatter = '---\ntags: [demo, northwind]\n---\n';
    root.innerHTML = '<div class="callout" data-callout="tip" data-callout-fold="+"><div class="callout-title-inner">Useful <em>tip</em></div><div class="callout-content"><p>Body</p><ul><li>Item</li></ul></div></div><pre><code class="language-python">print(1)\n</code></pre><hr><p>After</p>';
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => null));
    expect((await docxToMarkdown(docx, "images")).markdown).toBe('---\ntags: [demo, northwind]\n---\n\n> [!tip]+ Useful *tip*\n>\n> Body\n>\n> - Item\n\n```python\nprint(1)\n```\n\n---\n\nAfter\n');
    const read = readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: () => null });
    writeBody(docx.body, read.doc, read.origins, docx);
    const reloaded = await DocxDocument.load(await docx.save());
    expect((await docxToMarkdown(reloaded, "images")).markdown).toBe((await docxToMarkdown(docx, "images")).markdown);
  });
  it("keeps nesting expressed by Word's list styles and indentation", async () => {
    const docx = await DocxDocument.load(fixture("sample.docx"));
    const result = await docxToMarkdown(docx, "images");
    expect(result.markdown).toContain("- First bullet\n    - Nested bullet\n1. Numbered");
    expect(result.markdown).toContain("\n\n- Direct nested\n");
  });

  it("preserves headings, paragraphs and quotes", async () => {
    expect(await convert(p("Title", { style: "Title" }), p("Section", { style: "Heading2" }), p("Body"), p("Quoted", { style: "Quote" })))
      .toBe("# Title\n\n## Section\n\nBody\n\n> Quoted\n");
  });

  it("escapes literal Markdown syntax instead of changing its meaning", async () => {
    expect(await convert(p("# literal *stars* [link] $cash$"), p("1. ordinary"), p("---")))
      .toBe("\\# literal \\*stars\\* \\[link\\] \\$cash\\$\n\n1\\. ordinary\n\n\\---\n");
  });

  it("keeps overlapping formatting, links, inline code and hard breaks", async () => {
    const para = s.nodes.paragraph.create(null, [
      s.text("one ", [s.marks.bold.create()]), s.text("two", [s.marks.bold.create(), s.marks.italic.create()]),
      s.text(" link", [s.marks.link.create({ href: "https://example.com/a b" })]),
      s.nodes.hard_break.create(), s.text("a`b", [s.marks.font.create({ font: "Courier New" })]),
      s.text("gone", [s.marks.strike.create()]),
    ]);
    expect(await convert(para)).toBe("**one *two***[ link](<https://example.com/a%20b>)  \n``a`b``~~gone~~\n");
  });

  it("keeps nested bullet and numbered lists, including restarts", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const ordered = docx.ensureList("ordered"), bullet = docx.ensureList("bullet"), restart = docx.restartList(ordered);
    const doc = s.nodes.doc.create(null, [p("One", { numId: ordered }), p("Child", { numId: bullet, ilvl: 1 }), p("Two", { numId: ordered }), p("Gap"), p("Again", { numId: restart })]);
    expect((await docxToMarkdown(docx, "images", doc)).markdown).toBe("1. One\n    - Child\n2. Two\n\nGap\n\n1. Again\n");
  });

  it("preserves tables, escapes pipes and puts multi-paragraph cells on separate lines", async () => {
    const cell = (...nodes: ReturnType<typeof p>[]) => s.nodes.table_cell.create(null, nodes);
    const row = (...cells: ReturnType<typeof cell>[]) => s.nodes.table_row.create(null, cells);
    const table = s.nodes.table.create(null, [row(cell(p("Name")), cell(p("Value"))), row(cell(p("A|B")), cell(p("First"), p("Second")))]);
    expect(await convert(table)).toBe("| Name | Value |\n| --- | --- |\n| A\\|B | First<br>Second |\n");
  });

  it("expands merged cells to keep table columns lined up", async () => {
    const table = s.nodes.table.create(null, [s.nodes.table_row.create(null, [s.nodes.table_cell.create({ colspan: 2 }, p("Merged")), s.nodes.table_cell.create(null, p("Last"))])]);
    expect(await convert(table)).toBe("| Merged |  | Last |\n| --- | --- | --- |\n");
  });

  it("turns the exported Word code box back into a fenced code block", async () => {
    const root = createDiv();
    root.innerHTML = '<pre><code>const a = `x`;\n```\n</code></pre>';
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => null));
    expect((await docxToMarkdown(docx, "images")).markdown).toBe("````\nconst a = `x`;\n```\n````\n");
  });

  it("extracts referenced images once, with safe relative links and alt text", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const bytes = new Uint8Array([1, 2, 3]);
    const id = docx.addImage(bytes, "png");
    const image = s.nodes.image.create({ rId: id, alt: "A [picture]" });
    const result = await docxToMarkdown(docx, "My note attachments", s.nodes.doc.create(null, s.nodes.paragraph.create(null, [image, image])));
    expect(result.images).toHaveLength(1);
    expect(result.images[0].bytes).toEqual(bytes);
    expect(result.markdown).toBe("![A \\[picture\\]](<My%20note%20attachments/image-1.png>)![A \\[picture\\]](<My%20note%20attachments/image-1.png>)\n");
  });

  it("exports equations and footnotes", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const id = docx.addFootnote("A footnote\nSecond line");
    const math = s.nodes.opaque.create({ xml: latexToOmml("x^2"), text: "" });
    const ref = s.nodes.opaque.create({ xml: `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnoteReference w:id="${id}"/></w:r>`, text: "*" });
    const result = await docxToMarkdown(docx, "images", s.nodes.doc.create(null, s.nodes.paragraph.create(null, [math, ref])));
    expect(result.markdown).toContain("$x^2$[^fn-1]");
    expect(result.markdown).toContain("[^fn-1]: A footnote\n    Second line");
  });

  it("exports the current editor content without changing the source document", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const read = readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: () => null });
    writeBody(docx.body, s.nodes.doc.create(read.doc.attrs, p("Saved")), read.origins, docx);
    const before = new XMLSerializer().serializeToString(docx.body);
    expect((await docxToMarkdown(docx, "images", s.nodes.doc.create(null, p("Current edit")))).markdown).toBe("Current edit\n");
    expect(new XMLSerializer().serializeToString(docx.body)).toBe(before);
  });
});
