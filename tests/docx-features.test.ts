import { describe, it, expect, afterEach, vi } from "vitest";
import { kerningEm, textWidthEm } from "../src/presentation/pptx/text-fit";
import { EditorState, TextSelection, Command } from "prosemirror-state";
import { DecorationSet, EditorView } from "prosemirror-view";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { cmd, commentRanges, footnoteIds, docxPlugins, EditorContext, linkHref, pageNumberNode } from "../src/presentation/docx/docx-editor";
import { pagesKey, pagesPlugin } from "../src/presentation/docx/docx-pages";
import { DOMSerializer, Fragment, Slice } from "prosemirror-model";
import { blankDocx } from "../src/infrastructure/blank";
import { writeBody } from "../src/domain/docx/docx-writer";
import { paragraphProps, readDocx } from "../src/domain/docx/docx-reader";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readSample, W } from "./helpers";
import { headerFooterType, layoutPages, pageSetup, TopBlock, Unit } from "../src/domain/docx/pages";
import { acceptRevisions, rejectRevisions, revisions } from "../src/domain/docx/revisions";
import { countWords, findKey, findRanges, replaceAllCmd, replaceRangeCmd, setFind } from "../src/presentation/docx/docx-find";

const s = schema.nodes;
const m = schema.marks;
const ctx: EditorContext = { ensureList: () => "1", listKind: () => null, styles: new DocxStyles(null, null) };

function state(blocks: ReturnType<typeof s.paragraph.create>[], from = 1, to = from) {
  const doc = s.doc.create(null, blocks);
  return EditorState.create({ doc, selection: TextSelection.create(doc, from, to), plugins: docxPlugins(ctx) });
}
function run(st: EditorState, command: Command): EditorState {
  let next = st;
  expect(command(st, (tr) => (next = st.apply(tr)))).toBe(true);
  return next;
}
const text = (st: EditorState) => {
  const out: string[] = [];
  st.doc.descendants((n) => {
    if (n.isTextblock) out.push(n.textContent);
    return !n.isTextblock;
  });
  return out;
};

describe("find and replace", () => {
  const doc = () =>
    state([
      s.paragraph.create(null, [schema.text("Hello "), schema.text("wor", [m.bold.create()]), schema.text("ld, hello")]),
      s.table.create(null, [s.table_row.create(null, [s.table_cell.create(null, s.paragraph.create(null, schema.text("HELLO cell")))])]),
    ]);

  it("finds matches across marks, paragraphs and table cells", () => {
    const st = doc();
    const hits = findRanges(st.doc, "hello", false);
    expect(hits.map((h) => st.doc.textBetween(h.from, h.to))).toEqual(["Hello", "hello", "HELLO"]);
    expect(findRanges(st.doc, "hello", true)).toHaveLength(1);
    expect(st.doc.textBetween(...(Object.values(findRanges(st.doc, "o wor", false)[0]) as [number, number]))).toBe("o wor");
    expect(findRanges(st.doc, "", false)).toEqual([]);
  });

  it("does not match across an image or other inline atom", () => {
    const st = state([s.paragraph.create(null, [schema.text("ab"), s.opaque.create({ xml: "<x/>", text: "" }), schema.text("cd")])]);
    expect(findRanges(st.doc, "bc", false)).toEqual([]);
  });

  it("replaces one match, keeping the formatting where it starts", () => {
    let st = doc();
    const [, , third] = findRanges(st.doc, "hello", false);
    st = run(st, replaceRangeCmd(third, "Hi"));
    expect(text(st)).toEqual(["Hello world, hello", "Hi cell"]);
    const hit = findRanges(st.doc, "wor", true)[0];
    st = run(st, replaceRangeCmd(hit, "WOR"));
    expect(st.doc.rangeHasMark(hit.from, hit.from + 3, m.bold)).toBe(true);
  });

  it("replaces all matches in one undoable step", () => {
    const st = run(doc(), replaceAllCmd("hello", "bye", false));
    expect(text(st)).toEqual(["bye world, bye", "bye cell"]);
  });

  it("highlights matches and the current one", () => {
    let st = doc();
    st = st.apply(setFind(st.tr, "hello", false, 1));
    const deco = findKey.getState(st)!.decorations as DecorationSet;
    const found = deco.find();
    expect(found).toHaveLength(3);
    expect(found.filter((d) => (d.spec as { current?: boolean }).current)).toHaveLength(1);
    // Edits keep the highlights in step.
    st = st.apply(st.tr.insertText("x", 1));
    expect(findKey.getState(st)!.matches[0].from).toBe(2);
  });
});

describe("links", () => {
  it("adds, changes and removes a link on the selection", () => {
    let st = state([s.paragraph.create(null, schema.text("see the docs"))], 9, 13);
    st = run(st, cmd.setLink("https://a.example/"));
    const linkAt = (pos: number) => st.doc.resolve(pos).marks().find((x) => x.type === m.link)?.attrs;
    expect(linkAt(10)).toMatchObject({ href: "https://a.example/", rId: null });
    // With just a caret inside, the whole link changes.
    st = st.apply(st.tr.setSelection(TextSelection.create(st.doc, 11)));
    st = run(st, cmd.setLink("https://b.example/"));
    expect(linkAt(10)?.href).toBe("https://b.example/");
    expect(linkAt(12)?.href).toBe("https://b.example/");
    st = run(st, cmd.setLink(null));
    expect(linkAt(10)).toBeUndefined();
  });

  it("reports the link at the cursor", () => {
    const st = state([s.paragraph.create(null, [schema.text("x"), schema.text("go", [m.link.create({ href: "https://c/" })])])], 3);
    expect(linkHref(st)).toBe("https://c/");
    expect(linkHref(state([s.paragraph.create(null, schema.text("x"))], 1))).toBeNull();
  });

  it("writes new links with an external relationship that reads back", async () => {
    const r = await readSample();
    const doc = r.doc.replace(0, r.doc.child(0).nodeSize, new Slice(Fragment.from(s.paragraph.create(null, schema.text("link me", [m.link.create({ href: "https://new.example/" })]))), 0, 0));
    writeBody(r.docx.body, doc, r.origins, r.docx);
    const h = r.docx.body.getElementsByTagNameNS(W, "hyperlink")[0];
    const rId = h.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")!;
    expect(r.docx.hyperlinkHref(rId)).toBe("https://new.example/");
    const back = await DocxDocument.load(await r.docx.save());
    const again = readDocx(back.body, { imageSrc: () => null, hyperlinkHref: (id) => back.hyperlinkHref(id) });
    let href: string | null = null;
    again.doc.child(0).forEach((n) => (href ??= n.marks.find((x) => x.type === m.link)?.attrs.href ?? null));
    expect(href).toBe("https://new.example/");
  });
});

describe("word count", () => {
  it("counts words and characters in the document or the selection", () => {
    const st = state([s.paragraph.create(null, schema.text("One two  three")), s.paragraph.create(null, schema.text("four-five six."))]);
    expect(countWords(st.doc)).toEqual({ words: 5, chars: 28 });
    expect(countWords(st.doc, 1, 8)).toEqual({ words: 2, chars: 7 });
    expect(countWords(s.doc.create(null, s.paragraph.create()))).toEqual({ words: 0, chars: 0 });
  });
});

describe("paragraph spacing and indents", () => {
  const pPr = (inner: string) => new DOMParser().parseFromString(`<w:pPr xmlns:w="${W}">${inner}</w:pPr>`, "application/xml").documentElement;

  it("reads space before / after and first-line or hanging indents in points", () => {
    expect(paragraphProps(pPr(`<w:spacing w:before="240" w:after="120"/><w:ind w:left="720" w:firstLine="360"/>`))).toMatchObject({
      spaceBefore: 12,
      spaceAfter: 6,
      firstLine: 18,
      indent: 1,
    });
    expect(paragraphProps(pPr(`<w:ind w:hanging="360"/>`))).toMatchObject({ firstLine: -18, spaceBefore: null, spaceAfter: null });
  });

  it("writes them into w:spacing and w:ind, keeping other attributes", async () => {
    const r = await readSample();
    let doc = r.doc;
    const first = doc.child(0);
    doc = doc.replace(0, first.nodeSize, new Slice(Fragment.from(first.type.create({ ...first.attrs, spaceBefore: 6, spaceAfter: 0, firstLine: -18 }, first.content)), 0, 0));
    writeBody(r.docx.body, doc, r.origins, r.docx);
    const p = r.docx.body.getElementsByTagNameNS(W, "p")[0];
    const spacing = p.getElementsByTagNameNS(W, "spacing")[0];
    expect([spacing.getAttributeNS(W, "before"), spacing.getAttributeNS(W, "after")]).toEqual(["120", "0"]);
    const ind = p.getElementsByTagNameNS(W, "ind")[0];
    expect([ind.getAttributeNS(W, "hanging"), ind.hasAttributeNS(W, "firstLine")]).toEqual(["360", false]);
    const again = readDocx(r.docx.body, { imageSrc: () => null, hyperlinkHref: () => null }).doc.child(0).attrs;
    expect(again).toMatchObject({ spaceBefore: 6, spaceAfter: 0, firstLine: -18 });
  });

  it("sets spacing on the selected paragraphs and shows it", () => {
    let st = state([s.paragraph.create(null, schema.text("a")), s.paragraph.create(null, schema.text("b"))], 1, 5);
    st = run(st, cmd.setParagraph({ spaceBefore: 12, firstLine: 18 }));
    expect([st.doc.child(0).attrs.spaceBefore, st.doc.child(1).attrs.firstLine]).toEqual([12, 18]);
    const deco = (st.plugins.find((p) => p.props.decorations)!.getState(st) as DecorationSet).find(0, 3)[0] as unknown as { type: { attrs: { style: string } } };
    expect(deco.type.attrs.style).toContain("margin-top: 12pt");
    expect(deco.type.attrs.style).toContain("text-indent: 18pt");
  });
});

describe("pages", () => {
  const body = (sect: string) =>
    new DOMParser().parseFromString(`<w:body xmlns:w="${W}"><w:p/>${sect}</w:body>`, "application/xml").documentElement;

  it("reads page size and margins from the final section, in px", () => {
    const p = pageSetup(body(`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="720" w:bottom="1080" w:left="720" w:header="708" w:footer="708"/></w:sectPr>`));
    expect(p).toEqual({ width: 794, height: 1123, top: 96, right: 48, bottom: 72, left: 48, header: 47, footer: 47 });
    // Letter with 1" margins when the document has no section properties.
    expect(pageSetup(body(""))).toMatchObject({ width: 816, height: 1056, top: 96, left: 96 });
  });

  const block = (height: number, breakAfter = false): TopBlock => ({ unit: { kind: "block", height, pos: 0 }, breakAfter });
  /** A paragraph of `n` lines of `lh` px; line k starts at text position 100 * (k + 1). */
  const text = (n: number, lh: number, pos = 0): Unit => ({
    kind: "text",
    height: n * lh,
    pos,
    left: 96,
    lines: Array.from({ length: n }, (_, k) => ({ top: k * lh, bottom: (k + 1) * lh, pos: pos + 100 * (k + 1) })),
  });
  const GAP = 30;

  it("starts a new page before a block that doesn't fit, and after manual page breaks", () => {
    const blocks = [100, 300, 250, 50, 400, 10].map((h) => block(h));
    expect(layoutPages(blocks, 500, GAP).starts).toEqual([2, 4]);
    expect(layoutPages(blocks, 500, GAP).fills).toEqual([100, 200, 90]);
    blocks[0].breakAfter = true;
    expect(layoutPages(blocks, 500, GAP).starts).toEqual([1, 2, 4]);
    // A block taller than a page still gets its own page, never an empty one.
    expect(layoutPages([block(900), block(10)], 500, GAP).starts).toEqual([1]);
  });

  it("splits a long paragraph between lines, with a spacer that carries the page gap", () => {
    const l = layoutPages([block(300), { unit: text(10, 40, 1000), breakAfter: false }], 500, GAP);
    expect(l.starts).toEqual([]);
    // Lines 1-5 fit (300 + 200 = 500); line 6 starts page 2.
    expect(l.spacers).toEqual([{ pos: 1600, height: GAP, gapAt: 0, left: 96, page: 1 }]);
    expect(l.pages).toBe(2);
    expect(l.fills).toEqual([300]);
  });

  it("keeps at least two lines on each side of a break, else moves the paragraph", () => {
    expect(layoutPages([block(420), { unit: text(3, 40), breakAfter: false }], 500, GAP).starts).toEqual([1]);
    // 4 lines with only one fitting: the whole paragraph moves too.
    expect(layoutPages([block(450), { unit: text(4, 40), breakAfter: false }], 500, GAP).starts).toEqual([1]);
  });

  it("splits a table row inside each cell, so the row continues at the top of the next page", () => {
    const tall = text(10, 50, 2000);
    const short = text(2, 50, 3000);
    const table: Unit = { kind: "table", height: 500, pos: 0, rows: [{ height: 500, cells: [{ padTop: 0, padBottom: 0, left: 100, children: [tall] }, { padTop: 0, padBottom: 0, left: 400, children: [short] }] }] };
    const l = layoutPages([block(100), { unit: table, breakAfter: false }, block(50)], 500, GAP);
    // Cell 1: lines 1-8 fit (100 + 400); lines 9-10 continue on page 2. Cell 2 fits.
    expect(l.spacers).toEqual([{ pos: 2900, height: GAP, gapAt: 0, left: 96, page: 1 }]);
    expect(l.pages).toBe(2);
    expect(l.fills).toEqual([500 - 100 - 50]);
  });

  it("moves a row to the next page when none of its cells can put anything before the break", () => {
    const row = (h: number, a: number, b: number) => ({ height: h, cells: [{ padTop: 0, padBottom: 0, left: 100, children: [text(1, h, a)] }, { padTop: 0, padBottom: 0, left: 400, children: [text(1, h, b)] }] });
    const table: Unit = { kind: "table", height: 230, pos: 0, rows: [row(30, 1000, 1500), row(200, 2000, 3000)] };
    const l = layoutPages([block(450), { unit: table, breakAfter: false }], 500, GAP);
    expect(l.spacers).toEqual([
      { pos: 2000, height: 20 + GAP, gapAt: 20, left: 100, page: 1 },
      { pos: 3000, height: 20 + GAP, gapAt: null, left: 400, page: 1 },
    ]);
    expect(l.pages).toBe(2);
    // A table whose first row can't start on the page moves whole, like any block.
    expect(layoutPages([block(450), { unit: { ...table, rows: [row(200, 2000, 3000)], height: 200 }, breakAfter: false }], 500, GAP).starts).toEqual([1]);
  });

  it("can keep every page at Letter size whatever the document says", () => {
    const a4 = body(`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="720" w:bottom="1080" w:left="720" w:header="708" w:footer="708"/></w:sectPr>`);
    expect(pageSetup(a4, "letter")).toEqual({ width: 816, height: 1056, top: 96, right: 48, bottom: 72, left: 48, header: 47, footer: 47 });
    expect(pageSetup(a4, "document").width).toBe(794);
  });
});

describe("page view", () => {
  it("draws a page gap after a manual page break, with the next page's header", async () => {
    const doc = s.doc.create(null, [
      s.paragraph.create(null, [schema.text("one"), s.page_break.create()]),
      s.paragraph.create(null, schema.text("two")),
    ]);
    const host = document.body.appendChild(document.createElement("div"));
    const setup = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
    const view = new EditorView(host, {
      state: EditorState.create({
        doc,
        plugins: [
          ...docxPlugins(ctx),
          pagesPlugin({
            setup: () => setup,
            header: (page) => createSpan({ text: `Header ${page}` }),
            footer: (page) => createSpan({ text: `Footer ${page}` }),
          }),
        ],
      }),
    });
    await new Promise((r) => setTimeout(r, 100));
    const gaps = host.querySelectorAll(".docx-page-gap");
    expect(gaps).toHaveLength(1);
    expect(gaps[0].textContent).toContain("Footer 1");
    expect(gaps[0].textContent).toContain("Header 2");
    expect(host.querySelector(".docx-page-first")!.textContent).toContain("Header 1");
    expect(host.querySelector(".docx-page-last")!.textContent).toContain("Footer 2");
    expect(pagesKey.getState(view.state)!.pages).toBe(2);
    // Each page is padded to its full height (nothing has height under jsdom, so the whole text area).
    const fills = Array.from(host.querySelectorAll<HTMLElement>(".docx-page-fill")).map((f) => f.style.getPropertyValue("--docx-fill"));
    expect(fills).toEqual(["864px", "864px"]);
    view.destroy();
  });
});

describe("text as wide as in the document's font (when that font isn't installed)", () => {
  /** Fake canvas: every character of the installed font is 0.6em wide. */
  class FakeCanvas {
    getContext() {
      return {
        font: "",
        measureText(this: { font: string }, text: string) {
          return { width: [...text].length * 0.6 * parseFloat(/([\d.]+)px/.exec(this.font)![1]) };
        },
      };
    }
  }
  afterEach(() => vi.unstubAllGlobals());
  const spacing = (doc: ReturnType<typeof s.doc.create>, styles = ctx.styles) => {
    const st = EditorState.create({ doc, plugins: docxPlugins({ ...ctx, styles }) });
    return st.plugins
      .flatMap((pl) => (pl.props.decorations?.call(pl, st) as DecorationSet | undefined)?.find() ?? [])
      .map((d) => ({ from: d.from, to: d.to, style: (d as unknown as { type: { attrs?: { style?: string } } }).type.attrs?.style ?? "" }))
      .filter((d) => d.style.startsWith("letter-spacing"));
  };

  it("spreads the width difference over each run, from its font, size and weight", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const doc = s.doc.create(null, [s.paragraph.create(null, [schema.text("Hello ", [m.font.create({ font: "Calibri" }), m.size.create({ pt: 12 })]), schema.text("world", [m.font.create({ font: "Calibri" }), m.size.create({ pt: 12 }), m.bold.create()])])]);
    const [plain, bold] = spacing(doc);
    expect(kerningEm("Hello ", "Calibri", false) + kerningEm("world", "Calibri", true)).not.toBe(0);
    const px = 12 * (96 / 72);
    // Kerned, as Google Docs (and Word's default for new documents) draws them.
    const want = (text: string, b: boolean) => ((textWidthEm(text, "Calibri", b, false)! + kerningEm(text, "Calibri", b) - text.length * 0.6) * px) / text.length;
    expect([plain.from, plain.to, bold.from, bold.to]).toEqual([1, 7, 7, 12]);
    expect(parseFloat(plain.style.slice(15))).toBeCloseTo(want("Hello ", false), 3);
    expect(parseFloat(bold.style.slice(15))).toBeCloseTo(want("world", true), 3);
  });

  it("uses the paragraph style's font and size for runs without their own", () => {
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    const styles = new DocxStyles(
      new DOMParser().parseFromString(
        `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
        "application/xml"
      ),
      null
    );
    const [d] = spacing(s.doc.create(null, [s.paragraph.create(null, schema.text("Hello"))]), styles);
    const px = 10 * (96 / 72);
    expect(parseFloat(d.style.slice(15))).toBeCloseTo(((textWidthEm("Hello", "Arial", false, false)! + kerningEm("Hello", "Arial", false) - 5 * 0.6) * px) / 5, 3);
  });

  it("leaves fonts it has no widths for, and text without a canvas, alone", () => {
    expect(spacing(s.doc.create(null, [s.paragraph.create(null, schema.text("Hello", [m.font.create({ font: "Calibri" })]))]))).toEqual([]);
    vi.stubGlobal("OffscreenCanvas", FakeCanvas);
    expect(spacing(s.doc.create(null, [s.paragraph.create(null, schema.text("Hello", [m.font.create({ font: "Wingdings" })]))]))).toEqual([]);
  });
});

describe("keeping paragraphs and rows together", () => {
  const read = (inner: string) => {
    const body = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${inner}</w:body>`, "application/xml").documentElement;
    return { body, ...readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null }) };
  };
  const writeCtx = { nextDocPrId: () => 1, ensureParagraphStyles: () => {} };

  it("reads keep with next / keep lines together, unset meaning the style decides", () => {
    const { doc } = read(`<w:p><w:pPr><w:keepNext/><w:keepLines w:val="0"/></w:pPr></w:p><w:p/>`);
    expect([doc.child(0).attrs.keepNext, doc.child(0).attrs.keepLines]).toEqual([true, false]);
    expect([doc.child(1).attrs.keepNext, doc.child(1).attrs.keepLines]).toEqual([null, null]);
  });

  it("writes them in schema order, turning a style's setting off with val=0", () => {
    const { body, doc, origins } = read(`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>`);
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
    st = run(st, cmd.setParagraph({ keepNext: false, keepLines: true }));
    writeBody(body, st.doc, origins, writeCtx);
    const pPr = body.getElementsByTagNameNS(W, "pPr")[0];
    expect(Array.from(pPr.children).map((c) => c.localName)).toEqual(["pStyle", "keepNext", "keepLines", "jc"]);
    expect(pPr.children[1].getAttributeNS(W, "val")).toBe("0");
    expect(pPr.children[2].hasAttributeNS(W, "val")).toBe(false);
  });

  it("knows which styles keep with the next paragraph", () => {
    const styles = new DocxStyles(
      new DOMParser().parseFromString(
        `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:keepNext/><w:keepLines/></w:pPr></w:style>` +
          `<w:style w:type="paragraph" w:styleId="Heading2"><w:basedOn w:val="Heading1"/><w:pPr><w:keepLines w:val="0"/></w:pPr></w:style></w:styles>`,
        "application/xml"
      ),
      null
    );
    expect(styles.keep("Heading2")).toEqual({ next: true, lines: false });
    expect(styles.keep(null)).toEqual({ next: false, lines: false });
  });

  it("lets rows break across pages or not, for the whole table (Google Docs' table option)", () => {
    const { body, doc, origins } = read(
      `<w:tbl><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`
    );
    expect(doc.child(0).child(0).attrs.cantSplit).toBe(true);
    let st = EditorState.create({ doc, selection: TextSelection.create(doc, 4) });
    expect(cmd.rowsCanBreak(st)).toBe(false);
    st = run(st, cmd.setRowsCanBreak(false));
    expect([0, 1].map((i) => st.doc.child(0).child(i).attrs.cantSplit)).toEqual([true, true]);
    st = run(st, cmd.setRowsCanBreak(true));
    expect(cmd.rowsCanBreak(st)).toBe(true);
    st = run(st, cmd.setRowsCanBreak(false));
    writeBody(body, st.doc, origins, writeCtx);
    expect(body.getElementsByTagNameNS(W, "cantSplit")).toHaveLength(2);
  });
});

describe("section breaks", () => {
  it("starts a new page where a section ends, unless the next section starts continuous", async () => {
    const sect = (type: string) => `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ""}</w:sectPr>`;
    const xml =
      `<w:body xmlns:w="${W}"><w:p><w:pPr>${sect("")}</w:pPr><w:r><w:t>one</w:t></w:r></w:p>` +
      `<w:p><w:pPr>${sect("continuous")}</w:pPr><w:r><w:t>two</w:t></w:r></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p></w:body>`;
    const body = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    const { doc } = readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null });
    // A section's w:type says how that section starts: section 2 is continuous, section 3 (the body's) a new page.
    expect(doc.child(0).attrs.sectPr).toContain("sectPr");
    const host = document.body.appendChild(document.createElement("div"));
    const setup = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
    const view = new EditorView(host, {
      state: EditorState.create({ doc, plugins: [...docxPlugins(ctx), pagesPlugin({ setup: () => setup, header: () => null, footer: () => null })] }),
    });
    await new Promise((r) => setTimeout(r, 100));
    expect(pagesKey.getState(view.state)!.pages).toBe(2);
    expect(host.querySelector(".docx-page-gap")!.nextElementSibling!.textContent).toBe("three");
    view.destroy();
  });
});

describe("headers and footers", () => {
  it("creates a default header and footer referenced from the final section", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    expect(docx.headerFooter("header")).toBeNull();
    const h = docx.ensureHeaderFooter("header");
    const f = docx.ensureHeaderFooter("footer");
    expect(docx.ensureHeaderFooter("header")).toBe(h);
    const sectPr = Array.from(docx.body.children).pop()!;
    expect(Array.from(sectPr.children).slice(0, 2).map((c) => c.localName)).toEqual(["headerReference", "footerReference"]);
    const back = await DocxDocument.load(await docx.save());
    expect(back.headerFooter("header")!.localName).toBe("hdr");
    expect(back.headerFooter("footer")!.localName).toBe("ftr");
  });

  it("keeps separate first-page and even-page headers, switched on like Google Docs' header options", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    expect(docx.differentFirstPage()).toBe(false);
    expect(docx.differentOddEven()).toBe(false);
    const first = docx.ensureHeaderFooter("header", "first");
    const even = docx.ensureHeaderFooter("footer", "even");
    expect(docx.ensureHeaderFooter("header", "first")).toBe(first);
    expect(docx.headerFooter("header")).toBeNull();
    docx.setDifferentFirstPage(true);
    docx.setDifferentOddEven(true);
    const sectPr = Array.from(docx.body.children).pop()!;
    const refs = Array.from(sectPr.children).filter((c) => c.localName.endsWith("Reference"));
    expect(refs.map((r) => `${r.localName}:${r.getAttributeNS(W, "type")}`)).toEqual(["headerReference:first", "footerReference:even"]);
    // titlePg sits after the page setup, as the schema wants.
    const order = Array.from(sectPr.children).map((c) => c.localName);
    expect(order.indexOf("titlePg")).toBeGreaterThan(order.indexOf("pgMar"));
    const back = await DocxDocument.load(await docx.save());
    expect([back.differentFirstPage(), back.differentOddEven()]).toEqual([true, true]);
    expect(back.headerFooter("header", "first")!.localName).toBe("hdr");
    expect(back.headerFooter("footer", "even")!.localName).toBe("ftr");
    back.setDifferentFirstPage(false);
    back.setDifferentOddEven(false);
    expect([back.differentFirstPage(), back.differentOddEven()]).toEqual([false, false]);
    // The parts stay, as in Word and Google Docs, ready if the option comes back on.
    expect(back.headerFooter("header", "first")).not.toBeNull();
  });

  it("picks the header a page shows: first page, even pages, then the default", () => {
    expect([1, 2, 3].map((p) => headerFooterType(p, true, true))).toEqual(["first", "even", "default"]);
    expect([1, 2].map((p) => headerFooterType(p, false, true))).toEqual(["default", "even"]);
    expect([1, 2].map((p) => headerFooterType(p, true, false))).toEqual(["first", "default"]);
  });

  it("edits header text through the normal reader and writer", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    docx.ensureHeaderFooter("footer");
    const root = docx.headerFooter("footer")!;
    const read = readDocx(root, { imageSrc: () => null, hyperlinkHref: () => null });
    const doc = s.doc.create(null, [s.paragraph.create({ align: "center" }, [schema.text("Page "), pageNumberNode()])]);
    writeBody(root, doc, read.origins, docx);
    const back = await DocxDocument.load(await docx.save());
    const again = readDocx(back.headerFooter("footer")!, { imageSrc: () => null, hyperlinkHref: () => null }).doc;
    expect(again.child(0).attrs.align).toBe("center");
    const field = again.child(0).child(1);
    expect(field.type).toBe(s.opaque);
    expect(back.headerFooter("footer")!.getElementsByTagNameNS(W, "fldSimple")[0].getAttributeNS(W, "instr")).toContain("PAGE");
  });

  it("marks page number fields so the page view can fill them in", () => {
    const dom = DOMSerializer.fromSchema(schema).serializeNode(pageNumberNode()) as HTMLElement;
    expect(dom.dataset.field).toBe("PAGE");
  });
});

describe("word tables", () => {
  it("lays out paragraphs in a table with the table style's spacing (Table Grid: no space after)", () => {
    const body = new DOMParser().parseFromString(
      `<w:body xmlns:w="${W}"><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:body>`,
      "application/xml"
    ).documentElement;
    const { doc } = readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null });
    expect(doc.child(0).attrs.style).toBe("TableGrid");
    const styles = new DocxStyles(
      new DOMParser().parseFromString(
        `<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>` +
          `<w:style w:type="table" w:styleId="TableGrid"><w:pPr><w:spacing w:after="0"/></w:pPr></w:style></w:styles>`,
        "application/xml"
      ),
      null
    );
    const host = document.body.appendChild(document.createElement("div"));
    const view = new EditorView(host, { state: EditorState.create({ doc, plugins: docxPlugins({ ...ctx, styles }) }) });
    expect(host.querySelector("td p")!.getAttribute("style")).toContain("margin-bottom: 0pt");
    expect(host.querySelector(":scope .ProseMirror > p")!.getAttribute("style")).toContain("margin-bottom: 8pt");
    view.destroy();
  });


  const cell = (text: string, attrs: Record<string, unknown> = {}) => s.table_cell.create(attrs, s.paragraph.create(null, text ? schema.text(text) : null));
  const grid = (rows: string[][], attrs: Record<string, unknown> = {}) =>
    s.table.create(attrs, rows.map((r) => s.table_row.create(null, r.map((t) => cell(t)))));
  /** Caret inside cell (row, col) of the first table. */
  const at = (st: EditorState, row: number, col: number) => {
    let pos = -1;
    st.doc.child(0).forEach((tr, trOff, r) => tr.forEach((tc, tcOff, c) => (r === row && c === col ? (pos = 1 + trOff + 1 + tcOff + 2) : null)));
    return st.apply(st.tr.setSelection(TextSelection.create(st.doc, pos)));
  };
  const texts = (st: EditorState) => {
    const out: string[][] = [];
    st.doc.child(0).forEach((tr) => {
      const row: string[] = [];
      tr.forEach((tc) => row.push(tc.textContent));
      out.push(row);
    });
    return out;
  };
  const tableState = (rows: string[][]) => state([grid(rows), s.paragraph.create()], 4);

  it("merges a cell with the one to its right and splits it again", () => {
    let st = at(tableState([["a", "b", "c"], ["d", "e", "f"]]), 0, 0);
    st = run(st, cmd.mergeRight);
    expect(texts(st)).toEqual([["ab", "c"], ["d", "e", "f"]]);
    expect(st.doc.child(0).child(0).child(0).attrs.colspan).toBe(2);
    st = run(at(st, 0, 0), cmd.splitCell);
    expect(texts(st)[0]).toHaveLength(3);
    expect(st.doc.child(0).child(0).child(0).attrs.colspan).toBe(1);
  });

  it("merges a cell with the one below and shows it as one", () => {
    let st = at(tableState([["a", "b"], ["c", "d"], ["e", "f"]]), 0, 1);
    st = run(st, cmd.mergeDown);
    st = run(at(st, 0, 1), cmd.mergeDown);
    const col = (r: number) => st.doc.child(0).child(r).child(1).attrs.vMerge;
    expect([col(0), col(1), col(2)]).toEqual(["restart", "continue", "continue"]);
    expect(st.doc.child(0).child(0).child(1).textContent).toBe("bdf");
    const decos = (st.plugins.find((p) => p.props.decorations)!.getState(st) as DecorationSet).find();
    const attrsOfDeco = decos.map((d) => (d as unknown as { type: { attrs: Record<string, string> } }).type.attrs);
    expect(attrsOfDeco.some((a) => a.rowspan === "3")).toBe(true);
    expect(attrsOfDeco.filter((a) => a.class === "docx-merged")).toHaveLength(2);
    st = run(at(st, 0, 1), cmd.splitCell);
    expect([col(0), col(1), col(2)]).toEqual([null, null, null]);
  });

  it("sets borders, column widths and the repeating header row", () => {
    let st = at(tableState([["a", "b"], ["c", "d"]]), 1, 1);
    st = run(st, cmd.setTableBorders("outer"));
    st = run(st, cmd.setColumnWidth(2880));
    st = run(st, cmd.toggleHeaderRow);
    const t = st.doc.child(0);
    expect(t.attrs.borders).toBe("outer");
    expect(t.attrs.widths[1]).toBe(2880);
    expect(t.child(0).attrs.header).toBe(true);
  });

  it("reads and writes merges, widths, borders and header rows", async () => {
    const r = await readSample();
    let st = EditorState.create({ doc: r.doc.replace(0, r.doc.content.size, new Slice(Fragment.from([grid([["a", "b"], ["c", "d"]]), s.paragraph.create()]), 0, 0)), plugins: docxPlugins(ctx) });
    st = at(st, 0, 0);
    st = run(st, cmd.mergeDown);
    st = run(at(st, 0, 1), cmd.setTableBorders("none"));
    st = run(at(st, 0, 1), cmd.setColumnWidth(1440));
    st = run(at(st, 0, 1), cmd.toggleHeaderRow);
    writeBody(r.docx.body, st.doc, r.origins, r.docx);
    const tbl = r.docx.body.getElementsByTagNameNS(W, "tbl")[0];
    const vMerges = Array.from(tbl.getElementsByTagNameNS(W, "vMerge")).map((v) => v.getAttributeNS(W, "val"));
    expect(vMerges).toEqual(["restart", null]);
    expect(tbl.getElementsByTagNameNS(W, "tblHeader")).toHaveLength(1);
    expect(tbl.getElementsByTagNameNS(W, "gridCol")[1].getAttributeNS(W, "w")).toBe("1440");
    expect(tbl.getElementsByTagNameNS(W, "tblBorders")[0].getElementsByTagNameNS(W, "top")[0].getAttributeNS(W, "val")).toBe("nil");
    const again = readDocx(r.docx.body, { imageSrc: () => null, hyperlinkHref: () => null }).doc.child(0);
    expect(again.attrs).toMatchObject({ borders: "none" });
    expect(again.attrs.widths[1]).toBe(1440);
    expect(again.child(0).attrs.header).toBe(true);
    expect([again.child(0).child(0).attrs.vMerge, again.child(1).child(0).attrs.vMerge]).toEqual(["restart", "continue"]);
  });
});

describe("comments", () => {
  it("round-trips replies and resolution, retaining thread identities when text changes", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const parent = docx.addComment("Question", "Ada");
    const reply = docx.replyToComment(parent, "Answer", "Bob");
    docx.setCommentResolved(parent, true);
    docx.setCommentText(parent, "Question\nupdated");
    docx.setCommentText(reply, "Answer\nupdated");
    let back = await DocxDocument.load(await docx.save());
    expect(back.comments().find(c => c.id === parent)?.resolved).toBe(true);
    expect(back.comments().find(c => c.id === reply)?.parentId).toBe(parent);
    back.setCommentResolved(parent, false);
    back = await DocxDocument.load(await back.save());
    expect(back.comments().find(c => c.id === parent)?.resolved).toBe(false);
    back.deleteComment(parent);
    expect(back.comments()).toEqual([]);
    expect(back.xml("word/commentsExtended.xml")!.documentElement.children).toHaveLength(0);
  });

  it("adds, edits and deletes comments in comments.xml", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    expect(docx.comments()).toEqual([]);
    const id = docx.addComment("Check this", "Ada");
    const id2 = docx.addComment("Second", "Bob");
    expect(id2).not.toBe(id);
    docx.setCommentText(id, "Check this again\nplease");
    let back = await DocxDocument.load(await docx.save());
    expect(back.comments().map((c) => [c.id, c.author, c.initials, c.text])).toEqual([
      [id, "Ada", "A", "Check this again\nplease"],
      [id2, "Bob", "B", "Second"],
    ]);
    back.deleteComment(id);
    back = await DocxDocument.load(await back.save());
    expect(back.comments().map((c) => c.id)).toEqual([id2]);
  });

  it("anchors a comment to the selection and finds its range", () => {
    let st = state([s.paragraph.create(null, schema.text("one two three"))], 5, 8);
    st = run(st, cmd.addCommentAnchors("7"));
    const ranges = commentRanges(st.doc);
    expect(st.doc.textBetween(ranges.get("7")!.from, ranges.get("7")!.to)).toBe("two");
    expect(st.doc.textContent).toBe("one two three");
    const decos = (st.plugins.find((p) => p.props.decorations)!.getState(st) as DecorationSet).find();
    expect(decos.some((d) => (d as unknown as { type: { attrs: Record<string, string> } }).type.attrs?.class === "docx-comment-range")).toBe(true);
    st = run(st, cmd.removeCommentAnchors("7"));
    expect(commentRanges(st.doc).size).toBe(0);
    st.doc.descendants((n) => expect(n.type).not.toBe(s.opaque));
  });

  it("writes anchors Word understands", async () => {
    const r = await readSample();
    let st = EditorState.create({ doc: r.doc, plugins: docxPlugins(ctx), selection: TextSelection.create(r.doc, 2, 5) });
    st = run(st, cmd.addCommentAnchors("0"));
    writeBody(r.docx.body, st.doc, r.origins, r.docx);
    const p = r.docx.body.getElementsByTagNameNS(W, "p")[0];
    expect(Array.from(p.children).map((c) => c.localName)).toEqual(expect.arrayContaining(["commentRangeStart", "commentRangeEnd"]));
    expect(p.getElementsByTagNameNS(W, "commentReference")).toHaveLength(1);
    const again = readDocx(r.docx.body, { imageSrc: () => null, hyperlinkHref: () => null }).doc;
    expect(commentRanges(again).get("0")).toBeDefined();
  });
});

describe("footnotes", () => {
  it("adds, edits and deletes footnotes, creating the part with its separators", async () => {
    const docx = await DocxDocument.load(await blankDocx());
    const id = docx.addFootnote("A source.");
    const back = await DocxDocument.load(await docx.save());
    expect(back.footnotes()).toEqual([{ id, text: "A source." }]);
    const all = Array.from(back.xml("word/footnotes.xml")!.getElementsByTagNameNS(W, "footnote")).map((f) => f.getAttributeNS(W, "type"));
    expect(all).toEqual(["separator", "continuationSeparator", null]);
    back.setFootnoteText(id, "Another source.");
    expect(back.footnotes()[0].text).toBe("Another source.");
    // The footnote keeps its number mark.
    expect(back.xml("word/footnotes.xml")!.getElementsByTagNameNS(W, "footnoteRef")).toHaveLength(1);
    back.deleteFootnote(id);
    expect(back.footnotes()).toEqual([]);
  });

  it("inserts numbered references and removes them", () => {
    let st = state([s.paragraph.create(null, schema.text("claim one, claim two"))], 10);
    st = run(st, cmd.insertFootnote("4"));
    st = st.apply(st.tr.setSelection(TextSelection.create(st.doc, st.doc.child(0).content.size + 1)));
    st = run(st, cmd.insertFootnote("2"));
    expect(footnoteIds(st.doc)).toEqual(["4", "2"]);
    const decos = (st.plugins.find((p) => p.props.decorations)!.getState(st) as DecorationSet).find();
    const notes = decos.map((d) => (d as unknown as { type: { attrs: Record<string, string> } }).type.attrs?.["data-note"]).filter(Boolean);
    expect(notes).toEqual(["1", "2"]);
    st = run(st, cmd.removeFootnoteRef("4"));
    expect(footnoteIds(st.doc)).toEqual(["2"]);
  });
});

describe("table of contents", () => {
  const heading = (style: string, t: string) => s.paragraph.create({ style }, schema.text(t));

  it("inserts a TOC field built from the headings and updates it", async () => {
    let st = state([s.paragraph.create(null, schema.text("intro")), heading("Heading1", "Start"), heading("Heading2", "Detail"), heading("Heading4", "Too deep")], 1);
    st = run(st, cmd.insertToc);
    const toc = st.doc.child(1);
    expect(toc.attrs.kind).toBe("toc");
    expect(text(st).slice(1, 4)).toEqual(["Contents", "Start", "Detail"]);
    expect(toc.child(1).firstChild!.attrs.xml).toContain('w:fldCharType="begin"');
    // Rename a heading, then update.
    const pos = st.doc.child(0).nodeSize + toc.nodeSize + 1;
    st = st.apply(st.tr.insertText("Real ", pos));
    st = run(st, cmd.updateToc);
    expect(text(st).slice(1, 4)).toEqual(["Contents", "Real Start", "Detail"]);
    // Written out as Word's TOC field inside a content control.
    const body = new DOMParser().parseFromString(`<w:body xmlns:w="${W}"/>`, "application/xml").documentElement;
    writeBody(body, st.doc, new Map(), { nextDocPrId: () => 1 });
    const sdt = body.getElementsByTagNameNS(W, "sdt")[0];
    expect(sdt.getElementsByTagNameNS(W, "docPartGallery")[0].getAttributeNS(W, "val")).toBe("Table of Contents");
    expect(sdt.getElementsByTagNameNS(W, "instrText")[0].textContent).toContain('TOC \\o "1-3"');
    expect(Array.from(sdt.getElementsByTagNameNS(W, "pStyle")).map((x) => x.getAttributeNS(W, "val"))).toEqual(["TOCHeading", "TOC1", "TOC2"]);
    expect(Array.from(sdt.getElementsByTagNameNS(W, "fldChar")).map((x) => x.getAttributeNS(W, "fldCharType"))).toEqual(["begin", "separate", "end"]);
  });

  it("recognizes Word's own table of contents so it can be updated", async () => {
    const xml =
      `<w:body xmlns:w="${W}"><w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>` +
      `<w:sdtContent><w:p><w:r><w:t>Old entry</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>New</w:t></w:r></w:p></w:body>`;
    const body = new DOMParser().parseFromString(xml, "application/xml").documentElement;
    const { doc } = readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null });
    expect(doc.child(0).attrs.kind).toBe("toc");
    let st = EditorState.create({ doc, plugins: docxPlugins(ctx) });
    st = run(st, cmd.updateToc);
    expect(text(st).slice(0, 2)).toEqual(["Contents", "New"]);
  });

  it("fills in each heading's page from the page view, after a dot leader to a right tab", async () => {
    const doc = s.doc.create(null, [
      heading("Heading1", "Start"),
      s.paragraph.create(null, [schema.text("body"), s.page_break.create()]),
      heading("Heading2", "Detail"),
    ]);
    const host = document.body.appendChild(document.createElement("div"));
    const setup = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
    const view = new EditorView(host, {
      state: EditorState.create({ doc, selection: TextSelection.create(doc, 1), plugins: [...docxPlugins(ctx), pagesPlugin({ setup: () => setup, header: () => null, footer: () => null })] }),
    });
    await new Promise((r) => setTimeout(r, 100));
    cmd.insertToc(view.state, view.dispatch);
    const toc = view.state.doc.child(1);
    expect([1, 2].map((i) => toc.child(i).textContent)).toEqual(["Start\t1", "Detail\t2"]);
    expect(toc.child(1).attrs.tabLeader).toBe("dot");
    // Written as Word writes them: a right tab with a dot leader at the right margin.
    const body = new DOMParser().parseFromString(`<w:body xmlns:w="${W}"/>`, "application/xml").documentElement;
    writeBody(body, view.state.doc, new Map(), { nextDocPrId: () => 1, textWidth: () => 9360 });
    const tab = body.getElementsByTagNameNS(W, "tabs")[0].firstElementChild!;
    expect(["val", "leader", "pos"].map((k) => tab.getAttributeNS(W, k))).toEqual(["right", "dot", "9360"]);
    view.destroy();
  });

  it("leaves page numbers out without the page view (Word fills them in when it opens the file)", () => {
    let st = state([heading("Heading1", "Start")], 1);
    st = run(st, cmd.insertToc);
    expect(st.doc.child(1).child(1).textContent).toBe("Start");
  });

  it("defines the TOC styles and asks Word to refresh fields on open", async () => {
    const r = await readSample();
    r.docx.prepareToc();
    const ids = Array.from(r.docx.stylesXml()!.getElementsByTagNameNS(W, "style")).map((x) => x.getAttributeNS(W, "styleId"));
    expect(ids).toEqual(expect.arrayContaining(["TOC1", "TOC2", "TOC3", "TOCHeading"]));
    expect(r.docx.xml("word/settings.xml")!.getElementsByTagNameNS(W, "updateFields")[0].getAttributeNS(W, "val")).toBe("true");
  });
});

describe("track changes", () => {
  const body = () =>
    new DOMParser().parseFromString(
      `<w:body xmlns:w="${W}">` +
        `<w:p><w:r><w:t xml:space="preserve">Keep </w:t></w:r><w:ins w:id="1" w:author="Ada" w:date="2024-01-01T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>` +
        `<w:del w:id="2" w:author="Bob" w:date="2024-01-02T00:00:00Z"><w:r><w:delText xml:space="preserve"> removed</w:delText></w:r></w:del>` +
        `<w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="Ada"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t> styled</w:t></w:r></w:p>` +
        `<w:p><w:pPr><w:rPr><w:del w:id="4" w:author="Bob"/></w:rPr></w:pPr><w:r><w:t>joined</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>end</w:t></w:r></w:p></w:body>`,
      "application/xml"
    ).documentElement;
  const paras = (b: Element) => Array.from(b.getElementsByTagNameNS(W, "p")).map((p) => Array.from(p.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join(""));

  it("lists insertions, deletions and formatting changes", () => {
    expect(revisions(body()).map((r) => [r.id, r.kind, r.author, r.text])).toEqual([
      ["1", "insert", "Ada", "added"],
      ["2", "delete", "Bob", " removed"],
      ["3", "format", "Ada", " styled"],
      ["4", "delete", "Bob", "¶"],
    ]);
  });

  it("accepts everything", () => {
    const b = body();
    expect(acceptRevisions(b)).toBe(4);
    expect(paras(b)).toEqual(["Keep added styled", "joinedend"]);
    expect(b.getElementsByTagNameNS(W, "rPrChange")).toHaveLength(0);
    expect(b.getElementsByTagNameNS(W, "b")).toHaveLength(1);
    expect(revisions(b)).toEqual([]);
  });

  it("rejects everything", () => {
    const b = body();
    expect(rejectRevisions(b)).toBe(4);
    expect(paras(b)).toEqual(["Keep  removed styled", "joined", "end"]);
    expect(b.getElementsByTagNameNS(W, "delText")).toHaveLength(0);
    // The old formatting comes back.
    expect(b.getElementsByTagNameNS(W, "i")).toHaveLength(1);
    expect(b.getElementsByTagNameNS(W, "b")).toHaveLength(0);
  });

  it("restores old paragraph properties ahead of a section break", () => {
    const b = new DOMParser().parseFromString(
      `<w:body xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/><w:sectPr/>` +
        `<w:pPrChange w:id="9" w:author="Ada"><w:pPr><w:keepNext/><w:spacing w:after="0"/></w:pPr></w:pPrChange></w:pPr></w:p></w:body>`,
      "application/xml"
    ).documentElement;
    rejectRevisions(b);
    expect(Array.from(b.getElementsByTagNameNS(W, "pPr")[0].children).map((c) => c.localName)).toEqual(["keepNext", "spacing", "sectPr"]);
  });

  it("accepts or rejects one change by id", () => {
    const b = body();
    acceptRevisions(b, new Set(["2"]));
    expect(paras(b)[0]).toBe("Keep added styled");
    rejectRevisions(b, new Set(["1"]));
    expect(paras(b)[0]).toBe("Keep  styled");
    expect(revisions(b).map((r) => r.id)).toEqual(["3", "4"]);
  });

  it("shows insertions and deletions in the editor and keeps them when the paragraph is edited", async () => {
    const b = body();
    const { doc, origins } = readDocx(b, { imageSrc: () => null, hyperlinkHref: () => null });
    const first = doc.child(0);
    const insMark = first.child(1).marks.find((x) => x.type === m.insertion);
    expect(insMark?.attrs.author).toBe("Ada");
    const del = first.child(2);
    expect((DOMSerializer.fromSchema(schema).serializeNode(del) as HTMLElement).classList.contains("docx-del")).toBe(true);
    // Edit the paragraph: the w:ins wrapper must survive.
    const edited = doc.replace(1, 1, new Slice(Fragment.from(schema.text("New ")), 0, 0));
    const holder = b.ownerDocument.createElementNS(W, "w:body");
    writeBody(holder, edited, origins, { nextDocPrId: () => 1 });
    const ins = holder.getElementsByTagNameNS(W, "ins")[0];
    expect(ins.getAttributeNS(W, "author")).toBe("Ada");
    expect(ins.textContent).toBe("added");
  });
});

describe("display fidelity", () => {
  const ser = (n: Parameters<DOMSerializer["serializeNode"]>[0]) => DOMSerializer.fromSchema(schema).serializeNode(n) as HTMLElement;
  const read = (body: string) =>
    readDocx(new DOMParser().parseFromString(`<w:body xmlns:w="${W}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">${body}</w:body>`, "application/xml").documentElement, {
      imageSrc: () => null,
      hyperlinkHref: () => null,
    }).doc;

  it("uses the theme's fonts, with a matching generic fallback", () => {
    const xml = (s: string) => new DOMParser().parseFromString(s, "application/xml");
    const styles = new DocxStyles(
      xml(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/></w:styles>`),
      null,
      xml(`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Garamond"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`)
    );
    expect(styles.paragraphCss(null)).toContain(`font-family: "Garamond", serif`);
  });

  it("draws cell borders a cell switches off or on (an underline under a form field)", () => {
    const doc = read(
      `<w:tbl><w:tblPr><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:bottom w:val="single" w:sz="4"/></w:tcBorders></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`
    );
    const table = ser(doc.child(0));
    expect(table.classList.contains("docx-borders-none")).toBe(true);
    const td = table.querySelector("td")!;
    expect(td.classList.contains("docx-e-t-none")).toBe(true);
    expect(td.classList.contains("docx-e-b-line")).toBe(true);
  });

  it("shows equations instead of dropping them", () => {
    const doc = read(`<w:p><w:r><w:t xml:space="preserve">tree of </w:t></w:r><m:oMath><m:sSub><m:e><m:r><m:t>K</m:t></m:r></m:e><m:sub><m:r><m:t>5</m:t></m:r></m:sub></m:sSub></m:oMath></w:p>`);
    const math = doc.child(0).child(1);
    expect(math.attrs.text).toBe("K5");
    const dom = ser(math);
    expect(dom.textContent).toBe("K5");
    expect(dom.querySelector("msub")?.lastElementChild?.textContent).toBe("5");
  });
});

describe("content controls and fields", () => {
  const read = (inner: string) => {
    const body = new DOMParser().parseFromString(`<w:body xmlns:w="${W}">${inner}</w:body>`, "application/xml").documentElement;
    return { body, ...readDocx(body, { imageSrc: () => null, hyperlinkHref: () => null }) };
  };
  const sdt = (content: string, pr = `<w:alias w:val="Box"/>`) => `<w:sdt><w:sdtPr>${pr}</w:sdtPr><w:sdtEndPr/><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
  const p = (t: string) => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  const writeCtx = { nextDocPrId: () => 1, ensureParagraphStyles: () => {} };

  it("reads a block content control as editable paragraphs", () => {
    const { doc } = read(sdt(p("One") + p("Two")));
    const box = doc.child(0);
    expect(box.type.name).toBe("sdt");
    expect(box.childCount).toBe(2);
    expect(box.child(1).type.name).toBe("paragraph");
    expect(box.child(1).textContent).toBe("Two");
  });

  it("writes an edited content control back inside its properties", () => {
    const { body, doc, origins } = read(sdt(p("One")));
    const st = EditorState.create({ doc }).apply(EditorState.create({ doc }).tr.insertText("!", 5));
    writeBody(body, st.doc, origins, writeCtx);
    const out = body.firstElementChild!;
    expect(out.localName).toBe("sdt");
    expect(out.getElementsByTagNameNS(W, "alias")[0].getAttributeNS(W, "val")).toBe("Box");
    expect(out.getElementsByTagNameNS(W, "sdtEndPr")).toHaveLength(1);
    const content = out.getElementsByTagNameNS(W, "sdtContent")[0];
    expect(content.textContent).toBe("One!");
  });

  it("marks the result of a PAGE / NUMPAGES field so the page view can fill it in", () => {
    const field = (instr: string, result: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const { doc } = read(sdt(`<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>${field("PAGE", "2")}<w:r><w:t xml:space="preserve"> of </w:t></w:r>${field("NUMPAGES", "9")}</w:p>`));
    expect(doc.textContent).toBe("Page 2 of 9");
    const dom = DOMSerializer.fromSchema(schema).serializeFragment(doc.content, { document });
    const div = document.createElement("div");
    div.appendChild(dom);
    expect(Array.from(div.querySelectorAll<HTMLElement>("[data-field]")).map((f) => `${f.dataset.field}=${f.textContent}`)).toEqual(["PAGE=2", "NUMPAGES=9"]);
    expect(div.textContent).not.toContain("PAGE");
  });

  it("does not mark the result of other fields", () => {
    const { doc } = read(`<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGEREF _Toc1 \\h</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`);
    expect(doc.child(0).textContent).toBe("4");
    expect(doc.child(0).lastChild!.marks.some((x) => x.type.name === "field")).toBe(false);
  });

  it("reads Word's table of contents as a content control of styled entries", () => {
    const entry = (style: string, t: string) => `<w:p><w:pPr><w:pStyle w:val="${style}"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9345"/></w:tabs></w:pPr><w:r><w:t>${t}</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>5</w:t></w:r></w:p>`;
    const { doc } = read(sdt(entry("TOC1", "Intro") + entry("TOC2", "Detail"), `<w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj>`));
    const toc = doc.child(0);
    expect(toc.type.name).toBe("sdt");
    expect(toc.attrs.kind).toBe("toc");
    expect(toc.child(1).attrs.style).toBe("TOC2");
    expect(toc.child(1).attrs.tabLeader).toBe("dot");
  });

  it("stretches the last tab of a paragraph with a right tab stop, drawing its leader", () => {
    const { doc } = read(
      `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9345"/></w:tabs></w:pPr><w:r><w:t>1.1</w:t><w:tab/><w:t>Scope</w:t><w:tab/><w:t>6</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`
    );
    const st = EditorState.create({ doc, plugins: docxPlugins(ctx) });
    const decos = st.plugins.flatMap((pl) => (pl.props.decorations?.call(pl, st) as DecorationSet | undefined)?.find() ?? []);
    const fills = decos.filter((d) => (d as unknown as { type: { attrs?: { class?: string } } }).type.attrs?.class?.includes("docx-tab-fill"));
    expect(fills).toHaveLength(1);
    // The second tab of the first paragraph: "1.1\tScope\t6" starts at 1.
    expect(fills[0].from).toBe(1 + "1.1\tScope".length);
    expect((fills[0] as unknown as { type: { attrs: { class: string } } }).type.attrs.class).toContain("docx-leader-dot");
    // Earlier tabs stay tabs (their own element, so the flex line keeps them).
    const tabs = decos.filter((d) => (d as unknown as { type: { attrs?: { class?: string } } }).type.attrs?.class === "docx-tab");
    expect(tabs.map((d) => d.from)).toEqual([1 + "1.1".length]);
  });

  it("builds and updates its own table of contents as a content control", () => {
    const heading = (style: string, t: string) => s.paragraph.create({ style }, schema.text(t));
    let st = state([s.paragraph.create(null, schema.text("intro")), heading("Heading1", "Start"), heading("Heading2", "Detail")], 1);
    st = run(st, cmd.insertToc);
    const toc = st.doc.child(1);
    expect(toc.type.name).toBe("sdt");
    expect(toc.attrs.kind).toBe("toc");
    expect(text(st).slice(1, 4)).toEqual(["Contents", "Start", "Detail"]);
    expect(toc.child(2).attrs.style).toBe("TOC2");
  });

  it("lays out the paragraphs inside a content control one by one", async () => {
    const doc = s.doc.create(null, [
      s.sdt.create(null, [s.paragraph.create(null, [schema.text("one"), s.page_break.create()]), s.paragraph.create(null, schema.text("two"))]),
    ]);
    const host = document.body.appendChild(document.createElement("div"));
    const setup = { width: 816, height: 1056, top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 };
    const view = new EditorView(host, {
      state: EditorState.create({ doc, plugins: [...docxPlugins(ctx), pagesPlugin({ setup: () => setup, header: () => null, footer: () => null })] }),
    });
    await new Promise((r) => setTimeout(r, 100));
    expect(pagesKey.getState(view.state)!.pages).toBe(2);
    // The gap sits between the two paragraphs, inside the control.
    const gap = host.querySelector(".docx-page-gap")!;
    expect(gap.nextElementSibling!.textContent).toBe("two");
    view.destroy();
  });
});
