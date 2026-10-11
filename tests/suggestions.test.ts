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
import { acceptRevisions, rejectRevisions, revisions } from "../src/domain/docx/revisions";
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
