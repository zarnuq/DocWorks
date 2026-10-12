import { describe, expect, it } from "vitest";
import { blankDocx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { noteToDocx } from "../src/domain/docx/from-html";
import { docxToMarkdown } from "../src/domain/docx/to-markdown";
import { latexToOmml, ommlToLatex } from "../src/domain/docx/math";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { docxSchema as s } from "../src/domain/docx/schema";
import { NS } from "../src/core/utils/xml";

async function roundtrip(html: string, properties?: string) {
  const root = createDiv();
  root.innerHTML = html;
  if (properties) root.dataset.exportFrontmatter = properties;
  const bytes = await noteToDocx(await blankDocx(), root, () => null);
  const docx = await DocxDocument.load(bytes);
  const first = await docxToMarkdown(docx, "images");
  const read = readDocx(docx.body, { imageSrc: () => null, hyperlinkHref: id => docx.hyperlinkHref(id) });
  writeBody(docx.body, read.doc, read.origins, docx);
  const reloaded = await DocxDocument.load(await docx.save());
  expect((await docxToMarkdown(reloaded, "images")).markdown).toBe(first.markdown);
  return { markdown: first.markdown, docx, read };
}

describe("saved note round trips", () => {
  it.each([1, 2, 3, 4, 5, 6])("preserves heading level %s", async level => {
    expect((await roundtrip(`<h${level}>Heading</h${level}><p>Body</p>`)).markdown).toBe(`${"#".repeat(level)} Heading\n\nBody\n`);
  });

  it("keeps emphasis, inline code, links and highlight together", async () => {
    expect((await roundtrip('<p><strong>Bold</strong> <em>italic</em> <del>old</del> <mark>marked</mark> <code>a`b</code> <a href="https://example.com/a">web</a> <a href="obsidian://open?vault=notes&amp;file=Other.md">internal</a></p>')).markdown).toBe('**Bold** *italic* ~~old~~ ==marked== ``a`b`` [web](<https://example.com/a>) [internal](<obsidian://open?vault=notes&file=Other.md>)\n');
  });

  it("preserves a picture's bytes, alt text and relative extracted path", async () => {
    const root = createDiv();
    root.innerHTML = '<p><img alt="Field photo" src="photo.png"></p>';
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => ({ bytes, ext: 'png', width: 400, height: 200 })));
    const result = await docxToMarkdown(docx, 'My images');
    expect(result.markdown).toBe('![Field photo](<My%20images/image-1.png>)\n');
    expect(result.images).toEqual([{ name: 'image-1.png', bytes }]);
    expect(result.warnings).toEqual([]);
  });

  it("preserves list nesting and a restarted ordered list across a saved document", async () => {
    expect((await roundtrip('<ol><li>First<ul><li>Child</li></ul></li><li>Second</li></ol><p>Gap</p><ol><li>Restart</li></ol>')).markdown).toBe('1. First\n    - Child\n2. Second\n\nGap\n\n1. Restart\n');
  });

  it.each(["python", "javascript", "typescript", "bash", "json", "sql", "c++", "c#", "text", "mermaid"])("preserves %s code language and literal content", async language => {
    const { markdown } = await roundtrip(`<pre><code class="language-${language}"># literal $x$ &amp; &lt;tag&gt;\n    indented\n\nlast\n</code></pre>`);
    expect(markdown).toBe(`\`\`\`${language}\n# literal $x$ & <tag>\n    indented\n\nlast\n\`\`\`\n`);
  });

  it("gets the code language from the pre element and enlarges a conflicting fence", async () => {
    expect((await roundtrip('<pre class="language-python"><code>```\nprint(1)\n</code></pre>')).markdown).toBe('````python\n```\nprint(1)\n````\n');
  });

  it("does not invent a language for an unlabeled code box", async () => {
    expect((await roundtrip('<pre><code>plain\n</code></pre>')).markdown).toBe('```\nplain\n```\n');
  });

  it("keeps an empty labeled code block", async () => {
    expect((await roundtrip('<pre><code class="language-python"></code></pre>')).markdown).toBe('```python\n\n```\n');
  });

  it("keeps independently labeled consecutive code blocks", async () => {
    expect((await roundtrip('<pre><code class="language-python">a</code></pre><pre><code class="language-sql">b</code></pre>')).markdown).toBe('```python\na\n```\n\n```sql\nb\n```\n');
  });

  it.each(["note", "tip", "warning", "danger", "custom-type"])("preserves %s callout type and body formatting", async type => {
    const { markdown } = await roundtrip(`<div class="callout" data-callout="${type}"><div class="callout-title-inner">Title</div><div class="callout-content"><p><strong>Bold</strong> and <em>italic</em></p><p>Second</p></div></div>`);
    expect(markdown).toBe(`> [!${type}] Title\n>\n> **Bold** and *italic*\n>\n> Second\n`);
  });

  it.each(["+", "-"])("preserves callout folding %s", async fold => {
    expect((await roundtrip(`<div class="callout" data-callout="tip" data-callout-fold="${fold}"><div class="callout-title-inner">Title</div><div class="callout-content"><p>Body</p></div></div>`)).markdown).toBe(`> [!tip]${fold} Title\n>\n> Body\n`);
  });

  it("keeps a callout with no title or body", async () => {
    expect((await roundtrip('<div class="callout" data-callout="tip"><div class="callout-content"></div></div>')).markdown).toBe('> [!tip]\n');
  });

  it("preserves nested callouts and code languages inside them", async () => {
    const { markdown } = await roundtrip('<div class="callout" data-callout="note"><div class="callout-title-inner">Outer</div><div class="callout-content"><div class="callout" data-callout="tip"><div class="callout-title-inner">Inner</div><div class="callout-content"><pre><code class="language-python">print(1)</code></pre></div></div></div></div>');
    expect(markdown).toBe('> [!note] Outer\n>\n> > [!tip] Inner\n> >\n> > ```python\n> > print(1)\n> > ```\n');
  });

  it.each(["checked", ""])("keeps task state %s and inline formatting", async checked => {
    const { markdown } = await roundtrip(`<ul><li><input type="checkbox" ${checked}><strong>Work</strong> <em>now</em></li></ul>`);
    expect(markdown).toBe(`- [${checked ? "x" : " "}] **Work** *now*\n`);
  });

  it("keeps three nested task levels and returns to the parent level", async () => {
    const { markdown } = await roundtrip('<ul><li><input type="checkbox" checked>A<ul><li><input type="checkbox">B<ul><li><input type="checkbox" checked>C</li></ul></li></ul></li><li><input type="checkbox">D</li></ul>');
    expect(markdown).toBe('- [x] A\n    - [ ] B\n        - [x] C\n- [ ] D\n');
  });

  it("keeps ordinary checkbox glyphs as ordinary text", async () => {
    expect((await roundtrip('<p>☑ Symbol</p><ul><li>☐ Symbol</li></ul>')).markdown).toBe('☑ Symbol\n\n- ☐ Symbol\n');
  });

  it("indents a task's continuation beneath its text", async () => {
    expect((await roundtrip('<ul><li><input type="checkbox">First<br>\nSecond</li></ul>')).markdown).toBe('- [ ] First  \n      Second\n');
  });

  it("keeps empty task items", async () => {
    expect((await roundtrip('<ul><li>\n<input type="checkbox"> \n</li><li>\n<input type="checkbox" checked> \n</li></ul>')).markdown).toBe('- [ ] \n- [x] \n');
  });

  it("does not add a task marker to later plain bullets", async () => {
    expect((await roundtrip('<ul><li><input type="checkbox" checked>Task</li><li>Plain</li></ul>')).markdown).toBe('- [x] Task\n- Plain\n');
  });

  it.each([["left", "---"], ["center", ":---:"], ["right", "---:"]])("preserves %s table alignment", async (align, separator) => {
    const { markdown } = await roundtrip(`<table><tr><th style="text-align:${align}">Column</th></tr><tr><td style="text-align:${align}">Value</td></tr></table>`);
    expect(markdown).toBe(`| Column |\n| ${separator} |\n| Value |\n`);
  });

  it.each(['---\ntags: [demo]\n---\n', '---\ntitle: "A & B <C>"\nmultiline: |\n  one\n  two\n---\n', '---\r\ntags: [demo]\r\n---\r\n', '---\ntags: [demo]\n---'])("preserves properties %j without showing them in Word", async properties => {
    const { markdown, docx } = await roundtrip('<h1>Title</h1>', properties);
    expect(markdown).toBe(`${properties.trimEnd()}\n\n# Title\n`);
    expect(docx.body.textContent).toBe('Title');
  });

  it("does not export malformed document variables as properties", async () => {
    expect((await roundtrip('<p>Body</p>', 'not properties')).markdown).toBe('Body\n');
  });

  it("updates a document variable without duplicating it or disturbing other variables", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    docx.setDocumentVariable('Other', 'unchanged');
    docx.setDocumentVariable('DocWorksFrontmatter', 'first');
    docx.setDocumentVariable('DocWorksFrontmatter', 'second');
    const reloaded = await DocxDocument.load(await docx.save());
    expect(reloaded.documentVariable('Other')).toBe('unchanged');
    expect(reloaded.documentVariable('DocWorksFrontmatter')).toBe('second');
    const settings = reloaded.xml(reloaded.relOfType(reloaded.mainPath, 'settings')!)!;
    expect(settings.getElementsByTagNameNS(NS.w, 'docVar')).toHaveLength(2);
  });

  it("keeps multiple horizontal rules and distinguishes empty paragraphs", async () => {
    expect((await roundtrip('<p>Before</p><hr><p></p><p>Between</p><hr><p>After</p>')).markdown).toBe('Before\n\n---\n\nBetween\n\n---\n\nAfter\n');
  });

  it("uses current edited code content while retaining its language", async () => {
    const { docx, read } = await roundtrip('<pre><code class="language-python">old</code></pre>');
    const control = read.doc.firstChild!, table = control.firstChild!, row = table.firstChild!, cell = row.firstChild!;
    const paragraph = s.nodes.paragraph.create(cell.firstChild!.attrs, s.text('new', cell.firstChild!.firstChild!.marks));
    const editedCell = cell.type.create(cell.attrs, paragraph);
    const editedRow = row.type.create(row.attrs, editedCell);
    const editedTable = table.type.create(table.attrs, editedRow);
    const edited = control.type.create(control.attrs, editedTable);
    writeBody(docx.body, s.nodes.doc.create(read.doc.attrs, edited), read.origins, docx);
    expect((await docxToMarkdown(await DocxDocument.load(await docx.save()), 'images')).markdown).toBe('```python\nnew\n```\n');
  });
});

describe("equation round trips", () => {
  it.each(['matrix', 'pmatrix', 'bmatrix', 'cases', 'aligned', 'gathered'])("keeps rows inside a display %s environment", environment => {
    const latex = `\\begin{${environment}}a&b\\\\c&d\\end{${environment}}`;
    const display = latexToOmml(latex, true);
    const xml = new DOMParser().parseFromString(display, 'application/xml');
    expect(xml.getElementsByTagNameNS(NS.m, 'oMath')).toHaveLength(1);
    expect(ommlToLatex(display)).toBe(ommlToLatex(latexToOmml(latex)));
  });

  it("splits only top-level display line breaks", () => {
    const latex = String.raw`x=1\\\begin{matrix}a&b\\c&d\end{matrix}\\y=2`;
    const xml = new DOMParser().parseFromString(latexToOmml(latex, true), 'application/xml');
    expect(xml.getElementsByTagNameNS(NS.m, 'oMath')).toHaveLength(3);
    expect(xml.getElementsByTagNameNS(NS.m, 'mr')).toHaveLength(2);
  });

  it.each([String.raw`\sum_{i=1}^{n} x_i`, String.raw`\prod_{i=1}^{n} x_i^2`, String.raw`\int_0^1 x^2`, String.raw`\sin x^2`, String.raw`\cos x_i`, String.raw`\sum_{i=1}^{n} {x_i+y_i}`])("keeps the operand scripts in %s", latex => {
    const xml = new DOMParser().parseFromString(latexToOmml(latex), 'application/xml');
    const operator = xml.getElementsByTagNameNS(NS.m, /\\(?:sin|cos)/.test(latex) ? 'func' : 'nary')[0];
    expect(operator.parentElement?.localName).toBe('oMath');
    const operand = Array.from(operator.children).find(e => e.localName === 'e')!;
    expect(operand.querySelector('sSub, sSup, sSubSup')).not.toBeNull();
    const exported = ommlToLatex(latexToOmml(latex));
    expect(ommlToLatex(latexToOmml(exported))).toBe(exported);
  });

  it.each([String.raw`\frac{a}{b}`, String.raw`\sqrt{n}`, String.raw`\bar{x}`, String.raw`C_{\mathrm{after}}`, String.raw`x_i^2`, String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`])("keeps inline and display %s equations without excess spacing", async latex => {
    const root = createDiv();
    const before = root.createEl('p', { text: 'Before' });
    const inline = before.createSpan({ cls: 'math math-inline' });
    const inlineMath = document.createElement('mjx-math');
    inlineMath.setAttribute('data-latex', latex);
    inline.append(inlineMath);
    const display = root.createDiv({ cls: 'math math-block' });
    const displayMath = document.createElement('mjx-math');
    displayMath.setAttribute('data-latex', latex);
    display.append(displayMath);
    root.createEl('p', { text: 'After' });
    const docx = await DocxDocument.load(await noteToDocx(await blankDocx(), root, () => null));
    const { markdown } = await docxToMarkdown(docx, 'images');
    expect(markdown).toContain('Before$');
    expect(markdown).toContain('\n\n$$\n');
    expect(markdown).toContain('\n$$\n\nAfter\n');
    expect(markdown).not.toContain('\n\n\n');
    const equations = docx.body.getElementsByTagNameNS(NS.m, 'oMath');
    expect(equations).toHaveLength(2);
    expect(ommlToLatex(new XMLSerializer().serializeToString(equations[1]))).toBe(ommlToLatex(new XMLSerializer().serializeToString(equations[0])));
  });
});
