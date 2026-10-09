import { describe, it, expect } from "vitest";
import { EditorState, TextSelection, Command } from "prosemirror-state";
import { DecorationSet, EditorView } from "prosemirror-view";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { cmd, commentRanges, footnoteIds, docxPlugins, EditorContext, linkHref, pageNumberNode, pagesKey, pagesPlugin } from "../src/presentation/docx/docx-editor";
import { DOMSerializer, Fragment, Slice } from "prosemirror-model";
import { blankDocx } from "../src/infrastructure/blank";
import { writeBody } from "../src/domain/docx/docx-writer";
import { paragraphProps, readDocx } from "../src/domain/docx/docx-reader";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readSample, W } from "./helpers";
import { pageFills, pageSetup, paginate } from "../src/domain/docx/pages";
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

  it("starts a new page before the block that doesn't fit, and after manual page breaks", () => {
    const blocks = [100, 300, 250, 50, 400, 10].map((height) => ({ height, breakAfter: false }));
    expect(paginate(blocks, 500)).toEqual([2, 4]);
    blocks[0].breakAfter = true;
    expect(paginate(blocks, 500)).toEqual([1, 2, 4]);
    // A block taller than a page still gets its own page, never an empty one.
    expect(paginate([{ height: 900, breakAfter: false }, { height: 10, breakAfter: false }], 500)).toEqual([1]);
  });

  it("knows how much of each page is left blank, so pages keep their full height", () => {
    const blocks = [100, 300, 250, 50, 400, 10].map((height) => ({ height, breakAfter: false }));
    expect(pageFills(blocks, [2, 4], 500)).toEqual([100, 200, 90]);
    expect(pageFills([{ height: 900 }], [], 500)).toEqual([0]);
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
    expect(toc.attrs.text).toBe("Contents\nStart\n    Detail");
    expect(toc.attrs.xml).toContain('TOC \\o "1-3"');
    expect(toc.attrs.xml).toContain('w:val="TOC2"');
    // Rename a heading, then update.
    const pos = st.doc.child(0).nodeSize + toc.nodeSize + 1;
    st = st.apply(st.tr.insertText("Real ", pos));
    st = run(st, cmd.updateToc);
    expect(st.doc.child(1).attrs.text).toBe("Contents\nReal Start\n    Detail");
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
    expect(st.doc.child(0).attrs.text).toBe("Contents\nNew");
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
