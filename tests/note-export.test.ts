import { describe, expect, it } from "vitest";
import { htmlToDoc } from "../src/domain/docx/from-html";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { blankDocx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { noteToDocx } from "../src/domain/docx/from-html";
import { pageSetup } from "../src/domain/docx/pages";

const html = (s: string) => {
  const el = document.createElement("div");
  el.innerHTML = s;
  return el;
};
const lists: string[] = [];
const ctx = { list: (kind: "bullet" | "ordered" | "task") => (lists.push(kind), kind === "bullet" ? "1" : kind === "task" ? "3" : "2"), image: () => null };
const blocks = (s: string) => {
  const out: { type: string; text: string; attrs: Record<string, unknown>; marks: string[] }[] = [];
  htmlToDoc(html(s), ctx).forEach((n) => out.push({ type: n.type.name, text: n.textContent, attrs: n.attrs, marks: [] }));
  return out;
};

describe("note export to .docx", () => {
  it("exports MathJax inline and display equations as editable Word equations", async () => {
    const root = html('<p>Inline <span class="math math-inline"><mjx-container><mjx-math data-latex="\\frac{a}{b}">preview</mjx-math></mjx-container></span> end</p><div class="math math-block"><mjx-container display="true"><mjx-math data-latex="x^2">preview</mjx-math></mjx-container></div>');
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => null));
    const ns = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    expect(docx.body.getElementsByTagNameNS(ns, "f")).toHaveLength(1);
    expect(docx.body.getElementsByTagNameNS(ns, "oMathPara")).toHaveLength(1);
    expect(docx.body.textContent).not.toContain("preview");
    expect(docx.body.textContent).toContain("Inline ab end");
  });

  it("keeps resolved internal links clickable without exporting unsafe schemes", async () => {
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), html('<p><a class="internal-link" href="obsidian://open?vault=notes&amp;file=Other.md%23Heading">Alias</a><a href="javascript:alert(1)">Unsafe</a></p>'), () => null));
    expect(docx.rels(docx.mainPath).filter(r => r.type === "hyperlink").map(r => r.target)).toEqual(["obsidian://open?vault=notes&file=Other.md%23Heading"]);
    expect(docx.body.textContent).toBe("AliasUnsafe");
  });

  it("exports the selected paper, orientation and margins, with an optional title", async () => {
    const bytes = await noteToDocx(await blankDocx(), html("<p>Body</p>"), () => null, { pageSize: "a4", landscape: true, margin: "narrow", title: "My note" });
    const docx = await DocxDocument.load(bytes);
    expect(pageSetup(docx.body)).toMatchObject({ width: 1123, height: 794, top: 48, right: 48, bottom: 48, left: 48 });
    expect(docx.body.textContent).toBe("My noteBody");
    expect(docx.body.firstElementChild?.textContent).toBe("My note");
  });

  it("defaults to Letter with one-inch margins and no added title", async () => {
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), html("<p>Body</p>"), () => null));
    expect(pageSetup(docx.body)).toMatchObject({ width: 816, height: 1056, top: 96, left: 96 });
    expect(docx.body.textContent).toBe("Body");
  });

  it("downscales styled and ordinary text and fits pictures to the selected page", async () => {
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), html('<h1>Heading</h1><p>Body</p><pre>Code</pre><p><img src="image.png"></p>'), () => ({ bytes: new Uint8Array([1]), ext: "png", width: 2000, height: 1000 }), { pageSize: "letter", margin: "none", scale: 50 }));
    const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const sizes = Array.from(docx.stylesXml()!.getElementsByTagNameNS(ns, "sz")).map(n => Number(n.getAttributeNS(ns, "val")));
    expect(sizes).toContain(12); // default 12 pt becomes 6 pt
    expect(new XMLSerializer().serializeToString(docx.body)).toContain('w:sz w:val="10"'); // code 10 pt becomes 5 pt
    const extent = docx.body.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "extent")[0];
    expect(Number(extent.getAttribute("cx"))).toBe(408 * 9525);
    expect(Number(extent.getAttribute("cy"))).toBe(204 * 9525);
    expect(pageSetup(docx.body).left).toBe(0);
  });

  it("turns headings, paragraphs and quotes into Word paragraph styles", () => {
    const b = blocks(`<h1>Title</h1><p>Body</p><h3>Sub</h3><blockquote><p>Quoted</p></blockquote>`);
    expect(b.map((x) => [x.text, x.attrs.style])).toEqual([["Title", "Heading1"], ["Body", null], ["Sub", "Heading3"], ["Quoted", "Quote"]]);
  });

  it("keeps bold, italic, strike, highlight, inline code and links", () => {
    const doc = htmlToDoc(html(`<p><strong>b</strong><em>i</em><del>s</del><mark>h</mark><code>c</code><a href="https://example.com">l</a><a class="internal-link" href="Other">w</a></p>`), ctx);
    const marks: string[][] = [];
    doc.firstChild!.forEach((t) => marks.push(t.marks.map((m) => (m.type.name === "font" ? `font:${m.attrs.font}` : m.type.name === "link" ? `link:${m.attrs.href}` : m.type.name))));
    expect(marks).toEqual([["bold"], ["italic"], ["strike"], ["highlight"], ["font:Courier New"], ["link:https://example.com"], []]);
  });

  it("numbers lists by kind and nesting depth, and ticks task items", () => {
    const b = blocks(`<ul><li>a<ul><li>a1</li></ul></li><li class="task-list-item"><input type="checkbox" checked>done</li></ul><ol><li><p>one</p></li></ol>`);
    expect(b.map((x) => [x.text, x.attrs.numId, x.attrs.ilvl])).toEqual([["a", "1", 0], ["a1", "1", 1], ["☑ done", "3", 0], ["one", "2", 0]]);
  });

  it("draws task boxes without a bullet, in a symbol font so checked and unchecked boxes match", () => {
    const nodes: ReturnType<typeof htmlToDoc>["firstChild"][] = [];
    htmlToDoc(html(`<ul class="contains-task-list"><li class="task-list-item"><input type="checkbox" checked>done</li><li class="task-list-item"><input type="checkbox">todo</li></ul>`), ctx).forEach((n) => nodes.push(n));
    const [done, todo] = nodes.map((n) => n!);
    for (const p of [done, todo]) {
      expect(p.attrs.numId).toBe("3");
      const box = p.content.content.find((n) => n.isText)!;
      expect(box.text).toMatch(/^[☑☐]$/);
      expect(box.marks.map((m) => [m.type.name, m.attrs.font])).toEqual([["font", "Segoe UI Symbol"]]);
    }
    expect([done.textContent, todo.textContent]).toEqual(["☑ done", "☐ todo"]);
  });

  it("puts code blocks in a shaded one-cell box, one paragraph per line, and tables in bordered tables", () => {
    const doc = htmlToDoc(html(`<pre><code>let a = 1;\nlet b = 2;\n</code></pre><table><thead><tr><th>H</th><th>I</th></tr></thead><tbody><tr><td>1</td><td><strong>2</strong></td></tr></tbody></table>`), ctx);
    const [code, table] = [doc.child(0), doc.child(1)];
    expect(code.type).toBe(schema.nodes.table);
    expect(code.firstChild!.firstChild!.attrs.background).toBe("#F1F3F4");
    const lines: string[] = [];
    code.firstChild!.firstChild!.forEach((p) => lines.push(p.textContent));
    expect(lines).toEqual(["let a = 1;", "let b = 2;"]);
    expect(table.attrs.borders).toBe("all");
    expect(table.childCount).toBe(2);
    expect(table.firstChild!.attrs.header).toBe(true);
    expect(table.child(1).child(1).textContent).toBe("2");
  });

  it("writes a Word document the editor reads back", async () => {
    const bytes = await noteToDocx(await blankDocx(), html(`<h2>Plan</h2><ul><li>first</li></ul><p>text <strong>bold</strong></p>`), () => null);
    const docx = await DocxDocument.load(bytes);
    const xml = new XMLSerializer().serializeToString(docx.body);
    expect(xml).toContain("Heading2");
    expect(xml).toContain("first");
    expect(xml).toContain("<w:b/>");
    expect(new XMLSerializer().serializeToString(docx.stylesXml()!)).toContain('w:styleId="Heading2"');
  });

  it("starts each numbered list at 1", async () => {
    const bytes = await noteToDocx(await blankDocx(), html(`<ol><li>a</li><li>b</li></ol><p>between</p><ol><li>c</li></ol>`), () => null);
    const docx = await DocxDocument.load(bytes);
    const ids = Array.from(docx.body.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "numId")).map((n) => n.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "val"));
    expect(ids[0]).toBe(ids[1]);
    expect(ids[2]).not.toBe(ids[0]);
    const numbering = new XMLSerializer().serializeToString(docx.numberingXml()!);
    expect(numbering).toContain(`w:numId="${ids[2]}"`);
    expect(numbering).toMatch(/<w:startOverride w:val="1"\/>/);
  });

  it("keeps table column alignment", () => {
    const doc = htmlToDoc(html(`<table><tbody><tr><td>a</td><td style="text-align: right;">1</td><td align="center">c</td></tr></tbody></table>`), ctx);
    const row = doc.firstChild!.firstChild!;
    expect([0, 1, 2].map((i) => row.child(i).firstChild!.attrs.align)).toEqual([null, "right", "center"]);
  });

  it("writes embedded files as their name, not their preview, and keeps embedded notes' text", () => {
    // Built with DOM calls, as Obsidian does (an HTML string would move the div out of the paragraph).
    const root = html(`<p>Slide: <span class="internal-embed file-embed" src="deck.pptx#2"></span></p>` +
      `<div class="internal-embed markdown-embed" src="Other"><div class="markdown-embed-title">Other</div><div class="markdown-embed-content"><p>Embedded note text</p></div><div class="markdown-embed-link"></div></div>`);
    // The preview goes in the embed element itself (a document's preview has pictures too).
    const preview = root.querySelector<HTMLElement>(".file-embed")!;
    preview.addClass("dw-embed");
    preview.createDiv({ cls: "dw-embed-title", text: "deck" });
    const body = preview.createDiv({ cls: "dw-embed-body", text: "Today" });
    body.createEl("button", { text: "‹" });
    body.createEl("img", { attr: { src: "x.png" } });
    const doc = htmlToDoc(root, ctx);
    const texts: string[] = [];
    doc.forEach((n) => texts.push(n.textContent));
    expect(texts).toEqual(["Slide: deck.pptx#2", "Embedded note text"]);
  });
});

describe("pictures between blocks", () => {
  it("keeps a picture that sits directly between blocks (looked up by the element itself)", () => {
    const root = html(`<p>Before</p><img alt="slide"><p>After</p>`);
    const img = root.querySelector("img")!;
    const doc = htmlToDoc(root, { ...ctx, image: (el: HTMLImageElement) => (el === img ? { rId: "rId9", src: "", width: 10, height: 5 } : null) });
    const pics: string[] = [];
    doc.forEach((block) => block.descendants((n) => void (n.type.name === "image" && pics.push(n.attrs.rId as string))));
    expect(pics).toEqual(["rId9"]);
  });
});
