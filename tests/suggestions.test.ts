import { describe, it, expect } from "vitest";
import { EditorState, TextSelection } from "prosemirror-state";
import { undo, redo, history } from "prosemirror-history";
import { joinBackward, splitBlock } from "prosemirror-commands";
import { cmd } from "../src/presentation/docx/docx-editor";
import { blankDocx } from "../src/infrastructure/blank";
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readDocx } from "../src/domain/docx/docx-reader";
import { writeBody } from "../src/domain/docx/docx-writer";
import { docxSchema as s } from "../src/domain/docx/schema";
import { acceptRevisions, groupRevisions, rejectRevisions, revisions } from "../src/domain/docx/revisions";
import { suggestTransaction } from "../src/domain/docx/suggestions";
import { NS } from "../src/core/utils/xml";

const context = { imageSrc: () => null, hyperlinkHref: () => null };
async function setup(text = "Original text") {
  const pkg = await DocxDocument.load(await blankDocx());
  const original = s.nodes.doc.create(null, s.nodes.paragraph.create(null, s.text(text)));
  writeBody(pkg.body, original, new Map(), pkg);
  const { doc, origins } = readDocx(pkg.body, context);
  let state = EditorState.create({ doc, plugins: [history()] });
  return { pkg, origins, get state() { return state; },
    edit: (tr: typeof state.tr, opts?: { forward?: boolean }) => { state = state.apply(suggestTransaction(tr, origins, pkg, context, "Ada", opts)); },
    command: (fn: typeof undo) => fn(state, tr => { state = state.apply(suggestTransaction(tr, origins, pkg, context, "Ada")); }),
    xml: () => { writeBody(pkg.body, state.doc, origins, pkg); return pkg.body; }
  };
}
const plain = (body: Element) => readDocx(body, context).doc.textContent;

describe("suggestion recording", () => {
  it("doesn't record a format change on an untouched link when typing elsewhere in its paragraph", async () => {
    const pkg = await DocxDocument.load(await blankDocx());
    const p = new DOMParser().parseFromString(`<w:p xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:r><w:t xml:space="preserve">This report covers the </w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>study website</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>`, "application/xml").documentElement;
    pkg.body.insertBefore(pkg.body.ownerDocument.importNode(p, true), pkg.body.firstChild);
    const ctx = { imageSrc: () => null, hyperlinkHref: () => "https://example.org" };
    const { doc, origins } = readDocx(pkg.body, ctx);
    const state = EditorState.create({ doc });
    const next = state.apply(suggestTransaction(state.tr.insertText("big ", 6), origins, pkg, ctx, "Ada"));
    writeBody(pkg.body, next.doc, origins, pkg);
    expect(revisions(pkg.body).map(r => [r.kind, r.text])).toEqual([["insert", "big "]]);
  });

  it("records a deletion as a deletion only when the original runs differ just in namespace declarations", async () => {
    // Chromium writes a run's own xmlns declarations; one rewritten from the editor has none. Same text, same run.
    const pkg = await DocxDocument.load(await blankDocx());
    const run = (t: string, decl: boolean) => `<w:r${decl ? ' xmlns:zz="urn:zz"' : ""}><w:rPr><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${t}</w:t></w:r>`;
    const p = new DOMParser().parseFromString(`<w:p xmlns:w="${NS.w}">${run("conduct an ", false)}${run("integrated audit. You will do this independen", true)}${run("tly.", false)}</w:p>`, "application/xml").documentElement;
    pkg.body.insertBefore(pkg.body.ownerDocument.importNode(p, true), pkg.body.firstChild);
    const { doc, origins } = readDocx(pkg.body, context);
    const state = EditorState.create({ doc });
    const text = doc.textContent;
    const next = state.apply(suggestTransaction(state.tr.delete(1 + text.indexOf("integrated"), 1 + text.indexOf(" this")), origins, pkg, context, "Ada"));
    writeBody(pkg.body, next.doc, origins, pkg);
    expect(revisions(pkg.body).map(r => [r.kind, r.text])).toEqual([["delete", "integrated audit. You will do"]]);
  });

  it("doesn't list an empty run (a blank line's) as a deleted change", async () => {
    const pkg = await DocxDocument.load(await blankDocx());
    const para = (inner: string) => `<w:p xmlns:w="${NS.w}">${inner}</w:p>`;
    for (const xml of [para("<w:r><w:t>Hello world</w:t></w:r>"), para("<w:r/>"), para("<w:r><w:t>foo bar</w:t></w:r>")].reverse()) {
      pkg.body.insertBefore(pkg.body.ownerDocument.importNode(new DOMParser().parseFromString(xml, "application/xml").documentElement, true), pkg.body.firstChild);
    }
    const { doc, origins } = readDocx(pkg.body, context);
    const state = EditorState.create({ doc });
    const next = state.apply(suggestTransaction(state.tr.delete(7, doc.content.size - 5), origins, pkg, context, "Ada"));
    writeBody(pkg.body, next.doc, origins, pkg);
    expect(revisions(pkg.body).filter(r => r.text === "")).toEqual([]);
  });

  it("gives every suggestion its own id when edits touch only part of a long document", async () => {
    const pkg = await DocxDocument.load(await blankDocx());
    const blocks = Array.from({ length: 40 }, (_, i) => s.nodes.paragraph.create(null, s.text(`Paragraph ${i}`)));
    writeBody(pkg.body, s.nodes.doc.create(null, blocks), new Map(), pkg);
    const { doc, origins } = readDocx(pkg.body, context);
    let state = EditorState.create({ doc });
    const edit = (tr: typeof state.tr) => { state = state.apply(suggestTransaction(tr, origins, pkg, context, "Ada")); };
    const start = (i: number) => { let p = 0; for (let k = 0; k < i; k++) p += state.doc.child(k).nodeSize; return p + 1; };
    edit(state.tr.insertText("late ", start(35)));
    edit(state.tr.insertText("early ", start(2)));
    edit(state.tr.delete(start(20), start(20) + 9));
    writeBody(pkg.body, state.doc, origins, pkg);
    const list = revisions(pkg.body);
    expect(list.map(r => [r.kind, r.text])).toEqual([["insert", "early "], ["delete", "Paragraph"], ["insert", "late "]]);
    expect(new Set(list.map(r => r.id)).size).toBe(3);
    const accepted = pkg.body.cloneNode(true) as Element;
    acceptRevisions(accepted);
    expect(plain(accepted)).toContain("early Paragraph 2");
  });

  it("applies a suggestion in place: an edit in a table cell leaves the rest of the table (and its page gaps) where they were", async () => {
    const pkg = await DocxDocument.load(await blankDocx());
    const cell = (t: string) => s.nodes.table_cell.create(null, s.nodes.paragraph.create(null, s.text(t)));
    const table = s.nodes.table.create(null, [s.nodes.table_row.create(null, [cell("a"), cell("b")]), s.nodes.table_row.create(null, [cell("c"), cell("d")])]);
    writeBody(pkg.body, s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text("x")), table, s.nodes.paragraph.create(null, s.text("y"))]), new Map(), pkg);
    const { doc, origins } = readDocx(pkg.body, context);
    const state = EditorState.create({ doc });
    const tablePos = doc.child(0).nodeSize, secondRow = tablePos + 1 + doc.child(1).child(0).nodeSize;
    const tr = suggestTransaction(state.tr.insertText("new ", tablePos + 4), origins, pkg, context, "Ada");
    expect(tr.mapping.mapResult(secondRow).deleted).toBe(false);
    expect(tr.mapping.mapResult(tablePos).deleted).toBe(false);
    expect(tr.doc.child(1).child(0).child(0).textContent).toBe("new a");
  });

  it("records a replacement with author, deletion and insertion and restores either version", async () => {
    const t = await setup();
    t.edit(t.state.tr.insertText("suggested", 10, 14));
    const body = t.xml();
    expect(revisions(body).map(r => [r.kind, r.author, r.text])).toEqual([["delete", "Ada", "text"], ["insert", "Ada", "suggested"]]);
    const accepted = body.cloneNode(true) as Element;
    acceptRevisions(accepted);
    expect(plain(accepted)).toBe("Original suggested");
    rejectRevisions(body);
    expect(plain(body)).toBe("Original text");
  });
  it("keeps undo and redo atomic without recording them as new suggestions", async () => {
    const t = await setup();
    t.edit(t.state.tr.insertText("new ", 1));
    expect(revisions(t.xml())).toHaveLength(1);
    t.command(undo);
    expect(revisions(t.xml())).toHaveLength(0);
    expect(t.state.doc.textContent).toBe("Original text");
    t.command(redo);
    expect(revisions(t.xml())).toHaveLength(1);
  });
  it("records bold and paragraph formatting with rejectable original properties", async () => {
    const t = await setup();
    t.edit(t.state.tr.addMark(1, 9, s.marks.bold.create()));
    t.edit(t.state.tr.setNodeMarkup(0, undefined, { ...t.state.doc.firstChild!.attrs, align: "center" }));
    const body = t.xml();
    expect(revisions(body).filter(r => r.kind === "format")).toHaveLength(2);
    rejectRevisions(body);
    const doc = readDocx(body, context).doc;
    expect(doc.firstChild!.attrs.align).toBe(null);
    expect(doc.firstChild!.firstChild!.marks.some(m => m.type === s.marks.bold)).toBe(false);
  });
  it("records paragraph insertion and preserves an existing revision when typing more", async () => {
    const t = await setup("One");
    t.edit(t.state.tr.insertText("!", 4));
    t.edit(t.state.tr.insertText("?", 5));
    t.edit(t.state.tr.insert(5, s.nodes.paragraph.create(null, s.text("Second"))));
    const body = t.xml();
    const accepted = body.cloneNode(true) as Element;
    acceptRevisions(accepted);
    expect(plain(accepted)).toContain("Second");
    rejectRevisions(body);
    expect(plain(body)).toBe("One");
  });
  it("leaves selection-only transactions untouched and does not mutate backing XML while editing", async () => {
    const t = await setup();
    const original = new XMLSerializer().serializeToString(t.pkg.body);
    const tr = t.state.tr.setSelection(TextSelection.create(t.state.doc, 2));
    expect(suggestTransaction(tr, t.origins, t.pkg, context, "Ada")).toBe(tr);
    t.edit(t.state.tr.insertText("hello", 1));
    expect(new XMLSerializer().serializeToString(t.pkg.body)).toBe(original);
    expect(t.xml().getElementsByTagNameNS(NS.w, "ins").length).toBeGreaterThan(0);
  });
  it("accepts and rejects removal of a trailing paragraph without leaving an extra blank paragraph", async () => {
    const t = await setup("First");
    t.edit(t.state.tr.insert(t.state.doc.content.size, s.nodes.paragraph.create(null, s.text("Last"))));
    const proposed = t.xml().cloneNode(true) as Element;
    rejectRevisions(proposed);
    expect(readDocx(proposed, context).doc.childCount).toBe(1);
    // Start with accepted paragraphs, then suggest removing the last.
    acceptRevisions(t.pkg.body);
    const read = readDocx(t.pkg.body, context);
    const state = EditorState.create({ doc: read.doc });
    const tr = state.tr.delete(state.doc.firstChild!.nodeSize, state.doc.content.size);
    const next = state.apply(suggestTransaction(tr, read.origins, t.pkg, context, "Ada"));
    writeBody(t.pkg.body, next.doc, read.origins, t.pkg);
    const rejected = t.pkg.body.cloneNode(true) as Element;
    rejectRevisions(rejected);
    expect(plain(rejected)).toBe("FirstLast");
    acceptRevisions(t.pkg.body);
    expect(readDocx(t.pkg.body, context).doc.childCount).toBe(1);
    expect(plain(t.pkg.body)).toBe("First");
  });
  it("tracks inserted table rows and removes the table when all inserted rows are rejected", async () => {
    const t = await setup();
    const cell = s.nodes.table_cell.create(null, s.nodes.paragraph.create(null, s.text("Cell")));
    t.edit(t.state.tr.insert(t.state.doc.content.size, s.nodes.table.create(null, s.nodes.table_row.create(null, cell))));
    const body = t.xml();
    expect(revisions(body).some(r => r.text === "row")).toBe(true);
    rejectRevisions(body);
    expect(body.getElementsByTagNameNS(NS.w, "tbl")).toHaveLength(0);
  });
  it("rejects section-margin changes without duplicating section properties", async () => {
    const t = await setup("First");
    const section = `<w:sectPr xmlns:w="${NS.w}"><w:pgMar w:left="1440" w:right="1440"/></w:sectPr>`;
    const node = t.state.doc.firstChild!;
    t.edit(t.state.tr.setNodeMarkup(0, undefined, { ...node.attrs, sectPr: section }));
    const body = t.xml();
    acceptRevisions(body);
    const read = readDocx(body, context);
    const state = EditorState.create({ doc: read.doc });
    const nextSection = section.replace('w:left="1440"', 'w:left="720"');
    const tr = state.tr.setNodeMarkup(0, undefined, { ...state.doc.firstChild!.attrs, sectPr: nextSection });
    const changed = state.apply(suggestTransaction(tr, read.origins, t.pkg, context, "Ada"));
    writeBody(body, changed.doc, read.origins, t.pkg);
    rejectRevisions(body);
    expect(body.firstElementChild!.getElementsByTagNameNS(NS.w, "sectPr")).toHaveLength(1);
    expect(body.firstElementChild!.getElementsByTagNameNS(NS.w, "pgMar")[0].getAttributeNS(NS.w, "left")).toBe("1440");
  });
  it("preserves insertion suggestions when clearing formatting", async () => {
    const t = await setup("First");
    t.edit(t.state.tr.insertText("New ", 1));
    t.edit(t.state.tr.removeMark(1, 5));
    const body = t.xml();
    expect(revisions(body).some(r => r.kind === "insert" && r.text === "New ")).toBe(true);
    rejectRevisions(body);
    expect(plain(body)).toBe("First");
  });
  it("keeps the cursor in an inserted empty paragraph and after consecutive deletions", async () => {
    const t = await setup("First");
    const insert = t.state.tr.insert(t.state.doc.content.size, s.nodes.paragraph.create());
    insert.setSelection(TextSelection.create(insert.doc, insert.doc.content.size - 1));
    t.edit(insert);
    expect(t.state.selection.$from.parent.textContent).toBe("");
    // Select the newly inserted empty paragraph, then type normally there.
    const end = t.state.doc.content.size - 1;
    t.edit(t.state.tr.setSelection(TextSelection.create(t.state.doc, end)).insertText("Last"));
    expect(t.state.selection.$from.parent.textContent).toBe("Last");
    t.edit(t.state.tr.delete(t.state.selection.from - 1, t.state.selection.from));
    t.edit(t.state.tr.delete(t.state.selection.from - 1, t.state.selection.from));
    t.edit(t.state.tr.insertText("!"));
    const accepted = t.xml();
    acceptRevisions(accepted);
    expect(readDocx(accepted, context).doc.lastChild!.textContent).toBe("La!");
  });
});

describe("typing as it happens: one key at a time", () => {
  it("joins letters typed one by one into one insertion", async () => {
    const t = await setup("Hello world");
    t.edit(t.state.tr.setSelection(TextSelection.create(t.state.doc, 7)));
    for (const ch of "new ") t.edit(t.state.tr.insertText(ch));
    expect(t.state.doc.textContent).toBe("Hello new world");
    const ins = revisions(t.xml()).filter(r => r.kind === "insert");
    expect(ins.map(r => r.text)).toEqual(["new "]);
  });
  it("joins characters deleted one by one with Backspace, and with Delete, into one deletion each", async () => {
    const t = await setup("Hello world");
    // Backspace three times from the end of "Hello".
    t.edit(t.state.tr.setSelection(TextSelection.create(t.state.doc, 6)));
    for (let i = 0; i < 3; i++) {
      const pos = t.state.selection.from;
      t.edit(t.state.tr.delete(pos - 1, pos));
    }
    let dels = revisions(t.xml()).filter(r => r.kind === "delete");
    expect(dels.map(r => r.text)).toEqual(["llo"]);
    // Delete twice forward from before "world".
    t.edit(t.state.tr.setSelection(TextSelection.create(t.state.doc, t.state.doc.firstChild!.content.size - 4)));
    for (let i = 0; i < 2; i++) {
      const pos = t.state.selection.from;
      t.edit(t.state.tr.delete(pos, pos + 1), { forward: true });
    }
    dels = revisions(t.xml()).filter(r => r.kind === "delete");
    expect(dels.map(r => r.text)).toEqual(["llo", "wo"]);
  });
});

describe("structural suggestions", () => {
  const p = (t: string) => s.nodes.paragraph.create(null, t ? s.text(t) : undefined);
  const cell = (t: string) => s.nodes.table_cell.create(null, p(t));
  const table = () => s.nodes.table.create(null, [s.nodes.table_row.create(null, [cell("a"), cell("b")]), s.nodes.table_row.create(null, [cell("c"), cell("d")])]);
  /** Text of every paragraph, "|" between them. */
  const lines = (body: Element) => {
    const out: string[] = [];
    readDocx(body, context).doc.descendants(n => (n.isTextblock ? (out.push(n.textContent), false) : true));
    return out.join("|");
  };
  async function suggest(blocks: ReturnType<typeof p>[], cursor: number, run: (st: EditorState, d: (tr: EditorState["tr"]) => void) => unknown) {
    const pkg = await DocxDocument.load(await blankDocx());
    writeBody(pkg.body, s.nodes.doc.create(null, blocks), new Map(), pkg);
    const original = lines(pkg.body);
    const { doc, origins } = readDocx(pkg.body, context);
    let state = EditorState.create({ doc });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, cursor)));
    let plainEdit = state;
    run(plainEdit, tr => { plainEdit = plainEdit.apply(tr); });
    run(state, tr => { state = state.apply(suggestTransaction(tr, origins, pkg, context, "Ada")); });
    writeBody(pkg.body, state.doc, origins, pkg);
    const body = pkg.body;
    const accepted = body.cloneNode(true) as Element, rejected = body.cloneNode(true) as Element;
    acceptRevisions(accepted);
    rejectRevisions(rejected);
    const want = pkg.body.ownerDocument.createElementNS(NS.w, "w:body");
    writeBody(want, plainEdit.doc, new Map(), pkg);
    return { body, state, original, edited: lines(want), accepted, rejected, kinds: revisions(body).map(r => [r.kind, r.text]) };
  }

  it("records Enter inside a paragraph as an inserted paragraph mark only", async () => {
    const r = await suggest([p("Hello world")], 7, (st, d) => splitBlock(st, d));
    expect(r.kinds).toEqual([["insert", "¶"]]);
    expect(lines(r.accepted)).toBe("Hello |world");
    expect(lines(r.rejected)).toBe("Hello world");
    expect(r.state.selection.$from.parent.textContent).toBe("world");
    expect(r.state.selection.$from.parentOffset).toBe(0);
  });
  it("records joining two paragraphs as a deleted paragraph mark only", async () => {
    const r = await suggest([p("Hello"), p("world")], 8, (st, d) => joinBackward(st, d));
    expect(r.kinds).toEqual([["delete", "¶"]]);
    expect(lines(r.accepted)).toBe("Helloworld");
    expect(lines(r.rejected)).toBe("Hello|world");
    expect(r.state.selection.$from.parent.textContent).toBe("world");
    expect(r.state.selection.$from.parentOffset).toBe(0);
  });
  it("records deleting a selection across paragraphs as the deleted text plus the paragraph marks between", async () => {
    // "Hello world" | "middle" | "foo bar": delete from "world" to "bar".
    const blocks = [p("Hello world"), p("middle"), p("foo bar")];
    const r = await suggest(blocks, 1, (st, d) => d(st.tr.delete(7, 26)));
    // The review list names a paragraph's mark before its text.
    expect(r.kinds).toEqual([["delete", "¶"], ["delete", "world"], ["delete", "¶"], ["delete", "middle"], ["delete", "foo "]]);
    expect(lines(r.accepted)).toBe(r.edited);
    expect(lines(r.rejected)).toBe(r.original);
  });
  it("records deleting a selection across a table as deleted text, a deleted table and deleted marks, never an insertion", async () => {
    const blocks = [p("Hello world"), p("one"), table(), p(""), p("foo bar"), p("end")];
    // From "world" to "bar".
    const r = await suggest(blocks, 1, (st, d) => {
      let bar = -1;
      st.doc.descendants((n, pos) => { if (n.isText && n.text === "foo bar") bar = pos + 4; return bar < 0; });
      d(st.tr.delete(7, bar));
    });
    expect(r.kinds.filter(k => k[0] === "insert")).toEqual([]);
    expect(lines(r.accepted)).toBe(r.edited);
    expect(lines(r.rejected)).toBe(r.original);
    // The table's text is struck through too.
    expect(r.kinds.map(k => k[1]).join("|")).toContain("a");
  });
  it("groups one edit's revisions into one suggestion for review, as Google Docs lists it", async () => {
    const r = await suggest([p("Hello world"), p("middle"), p("foo bar")], 1, (st, d) => d(st.tr.delete(7, 26)));
    const groups = groupRevisions(revisions(r.body));
    expect(groups).toHaveLength(1);
    expect(groups[0]).toMatchObject({ kind: "delete", text: "world middle foo " });
    expect(groups[0].ids.size).toBe(revisions(r.body).length);
    const typed = await suggest([p("Hello world"), p("foo bar")], 1, (st, d) => d(st.tr.insertText("big", 7, 18)));
    expect(groupRevisions(revisions(typed.body))).toMatchObject([{ kind: "replace", text: "world foo ", inserted: "big" }]);
  });
  it("keeps separate edits (other times or authors) as separate suggestions", () => {
    const rev = (id: string, date: string, author = "Ada") => ({ id, kind: "delete" as const, author, date, text: id });
    expect(groupRevisions([rev("1", "t1"), rev("2", "t1"), rev("3", "t2"), rev("4", "t2", "Bo"), rev("5", "")]).map(g => [...g.ids])).toEqual([["1", "2"], ["3"], ["4"], ["5"]]);
  });
  it("records typing over a selection across paragraphs as one deletion and one insertion", async () => {
    const r = await suggest([p("Hello world"), p("foo bar")], 1, (st, d) => d(st.tr.insertText("big", 7, 18)));
    expect(r.kinds).toEqual([["delete", "¶"], ["delete", "world"], ["insert", "big"], ["delete", "foo "]]);
    expect(lines(r.accepted)).toBe(r.edited);
    expect(lines(r.rejected)).toBe(r.original);
  });
  it("records an added table column as inserted cells and a grid change", async () => {
    const r = await suggest([p("x"), table()], 6, (st, d) => cmd.addColumn(true)(st, d));
    expect(r.kinds.filter(k => k[1] === "cell")).toHaveLength(2);
    expect(lines(r.accepted)).toBe(r.edited);
    expect(lines(r.rejected)).toBe(r.original);
    expect(r.rejected.getElementsByTagNameNS(NS.w, "gridCol")).toHaveLength(2);
    expect(r.accepted.getElementsByTagNameNS(NS.w, "gridCol")).toHaveLength(3);
    expect(r.accepted.getElementsByTagNameNS(NS.w, "tblGrid")).toHaveLength(1);
  });
  it("records a deleted table column as deleted cells that stay visible until accepted", async () => {
    const r = await suggest([p("x"), table()], 6, (st, d) => cmd.deleteColumn(st, d));
    expect(r.kinds.filter(k => k[1] === "cell")).toHaveLength(2);
    expect(r.body.getElementsByTagNameNS(NS.w, "tc")).toHaveLength(4);
    expect(lines(r.accepted)).toBe("x|b|d");
    expect(lines(r.rejected)).toBe(r.original);
  });
  it("keeps a recorded column deletion through later edits in the table", async () => {
    const r = await suggest([p("x"), table()], 6, (st, d) => cmd.deleteColumn(st, d));
    const read = readDocx(r.body, context);
    const pkg = await DocxDocument.load(await blankDocx());
    let state = EditorState.create({ doc: read.doc });
    let at = -1;
    state.doc.descendants((n, pos) => { if (n.isText && n.text === "d") at = pos; });
    state = state.apply(suggestTransaction(state.tr.insertText("!", at + 1), read.origins, pkg, context, "Ada"));
    const body = r.body.cloneNode(true) as Element;
    writeBody(body, state.doc, read.origins, pkg);
    const rejected = body.cloneNode(true) as Element;
    rejectRevisions(rejected);
    expect(lines(rejected)).toBe(r.original);
    acceptRevisions(body);
    expect(lines(body)).toBe("x|b|d!");
  });
  it("records merged cells so that accepting merges and rejecting restores both cells", async () => {
    const r = await suggest([p("x"), table()], 6, (st, d) => cmd.mergeRight(st, d));
    expect(lines(r.accepted)).toBe(r.edited);
    expect(lines(r.rejected)).toBe(r.original);
    expect(r.rejected.getElementsByTagNameNS(NS.w, "gridSpan")).toHaveLength(0);
  });
  it("records a deleted table as deleted rows, and an inserted table as inserted rows", async () => {
    const del = await suggest([p("x"), table()], 6, (st, d) => cmd.deleteTable(st, d));
    expect(lines(del.accepted)).toBe("x");
    expect(lines(del.rejected)).toBe(del.original);
    const ins = await suggest([p("x")], 1, (st, d) => cmd.insertTable(2, 2)(st, d));
    expect(lines(ins.accepted)).toBe(ins.edited);
    expect(lines(ins.rejected)).toBe("x");
  });
});
