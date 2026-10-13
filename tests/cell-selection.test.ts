import { describe, it, expect } from "vitest";
import { EditorState, TextSelection, Transaction } from "prosemirror-state";
import { DecorationSet, EditorView } from "prosemirror-view";
import { toggleMark, deleteSelection } from "prosemirror-commands";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { cmd, docxPlugins, EditorContext } from "../src/presentation/docx/docx-editor";
import { CellSelection } from "../src/presentation/docx/cell-selection";

const s = schema.nodes;
const ctx: EditorContext = { ensureList: () => "B", listKind: () => null, styles: new DocxStyles(null, null) };
const para = (text: string) => s.paragraph.create(null, text ? schema.text(text) : null);
/** "a" paragraph, then a rows × cols table whose cells read "r,c". */
function doc(rows = 2, cols = 3, prefix = "") {
  const table = s.table.create(null, Array.from({ length: rows }, (_, r) => s.table_row.create(null, Array.from({ length: cols }, (_, c) => s.table_cell.create(null, para(`${prefix}${r},${c}`))))));
  return s.doc.create(null, [para("a"), table, para("")]);
}
const make = (d = doc()) => EditorState.create({ doc: d, plugins: docxPlugins(ctx) });
/** Position of cell (r, c) in the first table. */
function cellPos(st: EditorState, r: number, c: number, table = 0): number {
  let found = -1, seen = 0;
  st.doc.descendants((n, p) => {
    if (found >= 0 || n.type !== s.table) return found < 0;
    if (seen++ !== table) return false;
    let q = p + 1;
    for (let i = 0; i < r; i++) q += n.child(i).nodeSize;
    q += 1;
    for (let i = 0; i < c; i++) q += n.child(r).child(i).nodeSize;
    found = q;
    return false;
  });
  return found;
}
const select = (st: EditorState, a: [number, number], b: [number, number]) => st.apply(st.tr.setSelection(CellSelection.create(st.doc, cellPos(st, ...a), cellPos(st, ...b))));
const texts = (st: EditorState, table = 0) => {
  const t: string[][] = [];
  let seen = 0;
  st.doc.descendants((n) => {
    if (n.type !== s.table) return true;
    if (seen++ === table) n.forEach((row) => { const r: string[] = []; row.forEach((c) => r.push(c.textContent)); t.push(r); });
    return false;
  });
  return t;
};
const apply = (st: EditorState, f: (st: EditorState, d: (tr: Transaction) => void) => unknown) => {
  let next = st;
  f(st, (tr) => (next = st.apply(tr)));
  return next;
};

describe("cell selection", () => {
  it("covers the rectangle between its corner cells, whichever way it was dragged", () => {
    const st = select(make(), [0, 2], [1, 1]);
    expect((st.selection as CellSelection).cells.map((p) => st.doc.nodeAt(p)!.textContent)).toEqual(["0,1", "0,2", "1,1", "1,2"]);
  });

  it("is what a drag from one cell into another makes", () => {
    const st = make();
    const plugin = st.plugins.find((p) => p.props.createSelectionBetween)!;
    const view = { state: st } as EditorView;
    const between = (a: number, b: number) => plugin.props.createSelectionBetween!.call(plugin, view, st.doc.resolve(a), st.doc.resolve(b));
    expect(between(cellPos(st, 0, 0) + 2, cellPos(st, 1, 1) + 3)).toBeInstanceOf(CellSelection);
    expect(between(cellPos(st, 0, 0) + 2, cellPos(st, 0, 0) + 3)).toBeNull();
    expect(between(1, cellPos(st, 1, 1) + 3)).toBeNull();
  });

  it("tints only its cells and hides the text selection", () => {
    const st = select(make(), [0, 0], [1, 1]);
    expect(st.selection.visible).toBe(false);
    let tinted = 0;
    for (const p of st.plugins) {
      const set = p.props.decorations?.call(p, st) as DecorationSet | undefined;
      set?.find().forEach((d) => void ((d as unknown as { type: { attrs: { class?: string } } }).type.attrs.class === "is-cell-selected" && tinted++));
    }
    expect(tinted).toBe(4);
  });

  it("Delete clears the selected cells and keeps the table's shape", () => {
    const st = apply(select(make(), [0, 1], [1, 2]), deleteSelection);
    expect(texts(st)).toEqual([["0,0", "", ""], ["1,0", "", ""]]);
  });

  it("typing replaces the selected cells: the text goes in the first", () => {
    const st0 = select(make(), [0, 1], [1, 2]);
    const st = st0.apply(st0.tr.insertText("x"));
    expect(texts(st)).toEqual([["0,0", "x", ""], ["1,0", "", ""]]);
    expect(st.selection).toBeInstanceOf(TextSelection);
  });

  it("bold and colour apply to the selected cells only, not the cells between them in reading order", () => {
    let st = select(make(), [0, 1], [1, 1]);
    st = apply(st, toggleMark(schema.marks.bold));
    st = apply(st, cmd.setMark(schema.marks.color, { color: "#FF0000" }));
    const marked = (r: number, c: number) => st.doc.nodeAt(cellPos(st, r, c))!.firstChild!.firstChild!.marks.map((m) => m.type.name).sort();
    expect([marked(0, 1), marked(1, 1)]).toEqual([["bold", "color"], ["bold", "color"]]);
    expect([marked(0, 2), marked(1, 0)]).toEqual([[], []]);
  });

  it("paragraph formatting applies to the selected cells only", () => {
    const st = apply(select(make(), [0, 1], [1, 1]), cmd.setAlign("center"));
    const align = (r: number, c: number) => st.doc.nodeAt(cellPos(st, r, c))!.firstChild!.attrs.align as string | null;
    expect([align(0, 1), align(1, 1), align(0, 2), align(1, 0)]).toEqual(["center", "center", null, null]);
  });

  it("keeps the same cells across an edit above the table", () => {
    let st = select(make(), [0, 0], [1, 1]);
    st = st.apply(st.tr.insertText("more", 1));
    expect(st.selection).toBeInstanceOf(CellSelection);
    expect((st.selection as CellSelection).cells.map((p) => st.doc.nodeAt(p)!.textContent)).toEqual(["0,0", "0,1", "1,0", "1,1"]);
  });

  it("copies as a table of just the selected cells", () => {
    const st = select(make(), [0, 1], [1, 2]);
    const slice = st.selection.content();
    expect(slice.content.firstChild!.type).toBe(s.table);
    const rows: string[][] = [];
    slice.content.firstChild!.forEach((row) => { const r: string[] = []; row.forEach((c) => r.push(c.textContent)); rows.push(r); });
    expect(rows).toEqual([["0,1", "0,2"], ["1,1", "1,2"]]);
  });
});

describe("pasting cells", () => {
  const paste = (st: EditorState, slice: ReturnType<EditorState["selection"]["content"]>) => {
    let next = st;
    const view = { state: st, dispatch: (tr: Transaction) => (next = st.apply(tr)) } as unknown as EditorView;
    const handled = st.plugins.some((p) => p.props.handlePaste?.call(p, view, new Event("paste") as ClipboardEvent, slice));
    return { handled, st: next };
  };
  const copied = () => select(make(doc(2, 2, "x")), [0, 0], [1, 1]).selection.content();

  it("overwrites cells from the cursor's cell, keeping the table's shape", () => {
    const st0 = make();
    const { handled, st } = paste(st0.apply(st0.tr.setSelection(TextSelection.create(st0.doc, cellPos(st0, 0, 1) + 2))), copied());
    expect(handled).toBe(true);
    expect(texts(st)).toEqual([["0,0", "x0,0", "x0,1"], ["1,0", "x1,0", "x1,1"]]);
  });

  it("starts at a cell selection's first cell, and drops what falls outside the table", () => {
    const { st } = paste(select(make(), [1, 2], [1, 2]), copied());
    expect(texts(st)).toEqual([["0,0", "0,1", "0,2"], ["1,0", "1,1", "x0,0"]]);
  });

  it("leaves other pastes alone", () => {
    const st0 = make();
    const inCell = st0.apply(st0.tr.setSelection(TextSelection.create(st0.doc, cellPos(st0, 0, 1) + 2)));
    expect(paste(inCell, inCell.doc.slice(1, 2)).handled).toBe(false);
    expect(paste(st0, copied()).handled).toBe(false);
  });
});

describe("dragging across cells", () => {
  it("selects cells as the mouse moves from one cell into others, and back to text in the first", () => {
    const host = document.body.createDiv();
    const view = new EditorView(host, { state: make() });
    // jsdom has no layout: ProseMirror's own click handling finds no position and leaves the selection alone.
    document.elementFromPoint = () => null;
    try {
      const tds = Array.from(view.dom.querySelectorAll("td"));
      tds[0].firstChild!.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, button: 0 }));
      tds[4].dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
      expect(view.state.selection).toBeInstanceOf(CellSelection);
      expect((view.state.selection as CellSelection).cells.map((p) => view.state.doc.nodeAt(p)!.textContent)).toEqual(["0,0", "0,1", "1,0", "1,1"]);
      expect(view.dom.querySelectorAll("td.is-cell-selected").length).toBe(4);
      tds[0].dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
      expect(view.state.selection).toBeInstanceOf(TextSelection);
      window.dispatchEvent(new MouseEvent("mouseup"));
      tds[5].dispatchEvent(new MouseEvent("mousemove", { bubbles: true }));
      expect(view.state.selection).toBeInstanceOf(TextSelection);
    } finally {
      view.destroy();
    }
  });
});
