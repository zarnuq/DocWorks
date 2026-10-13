import { describe, it, expect } from "vitest";
import { EditorState, TextSelection, Command } from "prosemirror-state";
import { DecorationSet } from "prosemirror-view";
import { docxSchema as schema } from "../src/domain/docx/schema";
import { DocxStyles } from "../src/domain/docx/docx-styles";
import { cmd, docxPlugins, EditorContext } from "../src/presentation/docx/docx-editor";

const s = schema.nodes;
const para = (text: string, attrs: Record<string, unknown> = {}) => s.paragraph.create(attrs, text ? schema.text(text) : null);
const ctx: EditorContext = {
  ensureList: (kind) => (kind === "bullet" ? "B" : "O"),
  listKind: (numId) => (numId === "B" ? "bullet" : numId === "O" ? "ordered" : null),
  styles: new DocxStyles(null, null),
};

function state(blocks: ReturnType<typeof para>[], from = 1, to = from) {
  const doc = s.doc.create(null, blocks);
  return EditorState.create({ doc, selection: TextSelection.create(doc, from, to), plugins: docxPlugins(ctx) });
}
function run(st: EditorState, command: Command): EditorState {
  let next = st;
  expect(command(st, (tr) => (next = st.apply(tr)))).toBe(true);
  return next;
}
const attrs = (st: EditorState, i = 0) => st.doc.child(i).attrs;

describe("lists", () => {
  it("toggles a bullet list on and off", () => {
    let st = run(state([para("a")]), cmd.toggleList(ctx, "bullet"));
    expect(attrs(st)).toMatchObject({ numId: "B", ilvl: 0 });
    st = run(st, cmd.toggleList(ctx, "bullet"));
    expect(attrs(st)).toMatchObject({ numId: null, ilvl: 0 });
  });

  it("switches list kind and applies to every selected paragraph", () => {
    const st = run(state([para("a", { numId: "B" }), para("b")], 1, 5), cmd.toggleList(ctx, "ordered"));
    expect([attrs(st, 0).numId, attrs(st, 1).numId]).toEqual(["O", "O"]);
  });

  it("nests and un-nests list levels, leaving the list below level 0", () => {
    let st = run(state([para("a", { numId: "B" })]), cmd.shiftLevel(1));
    expect(attrs(st).ilvl).toBe(1);
    st = run(run(st, cmd.shiftLevel(-1)), cmd.shiftLevel(-1));
    expect(attrs(st)).toMatchObject({ numId: null, ilvl: 0 });
  });

  it("indents plain paragraphs instead", () => {
    let st = run(state([para("a")]), cmd.shiftLevel(1));
    expect(attrs(st).indent).toBe(1);
    st = run(run(st, cmd.shiftLevel(-1)), cmd.shiftLevel(-1));
    expect(attrs(st).indent).toBe(0);
  });

  it("Enter at the end of a list item starts the next item (as in Google Docs); after a heading, a normal paragraph", () => {
    let st = run(state([para("a", { numId: "B", ilvl: 1, key: "k1" })], 2), cmd.splitParagraph);
    expect(st.doc.childCount).toBe(2);
    expect(attrs(st, 1)).toMatchObject({ numId: "B", ilvl: 1, key: null });
    st = run(state([para("Title", { style: "Heading2", key: "k2" })], 6), cmd.splitParagraph);
    expect(attrs(st, 1)).toMatchObject({ style: null, key: null });
    // In the middle, both halves keep the paragraph's settings; only the original keeps its key.
    st = run(state([para("ab", { numId: "O", key: "k3" })], 2), cmd.splitParagraph);
    expect([attrs(st, 0).key, attrs(st, 1).key, attrs(st, 1).numId]).toEqual(["k3", null, "O"]);
  });

  it("Enter on an empty list item ends the list", () => {
    const st = run(state([para("", { numId: "B" })]), cmd.exitEmptyListItem);
    expect(attrs(st).numId).toBeNull();
    expect(cmd.exitEmptyListItem(state([para("x", { numId: "B" })], 2), () => {})).toBe(false);
  });
});

describe("paragraph formatting", () => {
  it("sets styles and keeps other attributes", () => {
    const st = run(state([para("a", { key: "k1", align: "center" })]), cmd.setStyle("Heading2"));
    expect(attrs(st)).toMatchObject({ style: "Heading2", key: "k1", align: "center" });
  });

  it("sets alignment and line spacing", () => {
    let st = run(state([para("a")]), cmd.setAlign("right"));
    st = run(st, cmd.setLineSpacing("1.5"));
    expect(attrs(st)).toMatchObject({ align: "right", lineSpacing: "1.5" });
  });
});

describe("marks", () => {
  it("applies a valued mark to the selection", () => {
    const st = run(state([para("hello")], 1, 3), cmd.setMark(schema.marks.color, { color: "#FF0000" }));
    const first = st.doc.child(0).child(0);
    expect(first.text).toBe("he");
    expect(first.marks[0].attrs.color).toBe("#FF0000");
  });

  it("stores the mark for typing when the selection is empty", () => {
    const st = run(state([para("hello")], 2), cmd.setMark(schema.marks.size, { pt: 20 }));
    expect(st.storedMarks?.[0].attrs.pt).toBe(20);
  });

  it("clears formatting but keeps links", () => {
    const text = schema.text("x", [schema.marks.bold.create(), schema.marks.link.create({ href: "h" }), schema.marks.rpr.create({ xml: "<w:rPr/>" })]);
    const st = run(state([s.paragraph.create(null, text)], 1, 2), cmd.clearFormatting);
    expect(st.doc.child(0).child(0).marks.map((m) => m.type.name)).toEqual(["link"]);
  });
});

describe("inserts", () => {
  it("inserts an inline page break at the cursor", () => {
    const st = run(state([para("ab")], 2), cmd.insertPageBreak);
    expect(st.doc.child(0).child(1).type.name).toBe("page_break");
  });

  it("inserts an equation at the cursor and edits it in place", () => {
    const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    let st = run(state([para("ab")], 2), cmd.setEquation(null, "x^2"));
    const eq = st.doc.child(0).child(1);
    expect(eq.type.name).toBe("opaque");
    expect(eq.attrs.xml).toMatch(/^<m:oMath /);
    expect(eq.attrs.text).toBe("x2");
    st = run(st, cmd.setEquation(2, "\\frac{1}{2}"));
    const edited = st.doc.child(0).child(1);
    expect(edited.attrs.xml).toContain("<m:f>");
    expect(edited.attrs.text).toBe("12");
    expect(cmd.equationAt(st, 2)).toBe("\\frac{1}{2}");
    // A display equation stays one.
    st = st.apply(st.tr.setNodeMarkup(2, undefined, { ...edited.attrs, xml: `<m:oMathPara xmlns:m="${M}"><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>` }));
    st = run(st, cmd.setEquation(2, "z"));
    expect(st.doc.child(0).child(1).attrs.xml).toMatch(/^<m:oMathPara /);
  });

  it("removes an equation edited to nothing", () => {
    let st = run(state([para("ab")], 2), cmd.setEquation(null, "x"));
    st = run(st, cmd.setEquation(2, "  "));
    expect(st.doc.child(0).textContent).toBe("ab");
    expect(st.doc.child(0).childCount).toBe(1);
  });

  it("inserts a horizontal rule after the paragraph", () => {
    const st = run(state([para("ab")], 2), cmd.insertRule);
    expect(st.doc.child(1).type.name).toBe("raw_block");
    expect(st.doc.child(1).attrs.kind).toBe("hr");
  });
});

describe("inserting a table", () => {
  it("puts the caret in its first cell, as Google Docs does", () => {
    const st = run(state([para("a")], 2), cmd.insertTable(2, 2));
    const $from = st.selection.$from;
    expect($from.node($from.depth - 1).type).toBe(s.table_cell);
    expect($from.index($from.depth - 3)).toBe(0); // first row
    expect($from.index($from.depth - 2)).toBe(0); // first cell
  });
});

describe("Tab in a table", () => {
  const cellIndex = (st: EditorState) => {
    const $f = st.selection.$from;
    return [$f.index($f.depth - 3), $f.index($f.depth - 2)];
  };
  it("moves to the next cell, back with Shift+Tab, and adds a row from the last cell (as in Google Docs)", () => {
    let st = run(state([para("a")], 2), cmd.insertTable(2, 2));
    st = run(st, cmd.goToCell(1));
    expect(cellIndex(st)).toEqual([0, 1]);
    st = run(st, cmd.goToCell(1));
    expect(cellIndex(st)).toEqual([1, 0]);
    st = run(st, cmd.goToCell(-1));
    expect(cellIndex(st)).toEqual([0, 1]);
    st = run(run(st, cmd.goToCell(1)), cmd.goToCell(1));
    expect(cellIndex(st)).toEqual([1, 1]);
    st = run(st, cmd.goToCell(1));
    expect(cellIndex(st)).toEqual([2, 0]);
    let rows = 0;
    st.doc.descendants((n) => { if (n.type === s.table) rows = n.childCount; });
    expect(rows).toBe(3);
    expect(cmd.goToCell(1)(state([para("x")], 1), () => {})).toBe(false);
  });
});

describe("tables", () => {
  const withTable = () => run(state([para("a")]), cmd.insertTable(2, 3));
  const inCell = (st: EditorState, row: number, col: number) => {
    let pos = -1;
    st.doc.descendants((n, p) => {
      if (pos < 0 && n.type === s.table) {
        let q = p + 1;
        for (let r = 0; r < row; r++) q += n.child(r).nodeSize;
        q += 1;
        for (let c = 0; c < col; c++) q += n.child(row).child(c).nodeSize;
        pos = q + 2;
      }
    });
    return st.apply(st.tr.setSelection(TextSelection.create(st.doc, pos)));
  };
  const table = (st: EditorState) => st.doc.child(1);
  const shape = (st: EditorState) => Array.from({ length: table(st).childCount }, (_, r) => table(st).child(r).childCount);

  it("inserts a table after the current paragraph", () => {
    expect(shape(withTable())).toEqual([3, 3]);
  });

  it("adds and deletes rows", () => {
    let st = run(inCell(withTable(), 0, 0), cmd.addRow(true));
    expect(shape(st)).toEqual([3, 3, 3]);
    st = run(inCell(st, 2, 0), cmd.deleteRow);
    expect(shape(st)).toEqual([3, 3]);
  });

  it("adds and deletes columns", () => {
    let st = run(inCell(withTable(), 0, 1), cmd.addColumn(true));
    expect(shape(st)).toEqual([4, 4]);
    st = run(inCell(st, 1, 0), cmd.deleteColumn);
    expect(shape(st)).toEqual([3, 3]);
  });

  it("shades the current cell", () => {
    const st = run(inCell(withTable(), 1, 2), cmd.setCellBackground("#ABCDEF"));
    expect(table(st).child(1).child(2).attrs.background).toBe("#ABCDEF");
  });

  it("deletes the whole table", () => {
    const st = run(inCell(withTable(), 0, 0), cmd.deleteTable);
    expect(st.doc.childCount).toBe(1);
  });

  it("refuses table commands outside a table", () => {
    expect(cmd.addRow(true)(state([para("a")]))).toBe(false);
  });
});

describe("decorations", () => {
  it("decorates list paragraphs with their marker", () => {
    const st = state([para("a", { numId: "B" }), para("b")]);
    const plugin = st.plugins.find((p) => p.props.decorations)!;
    const set = plugin.props.decorations!.call(plugin, st) as DecorationSet;
    const found = set.find();
    const list = found.find((d) => d.from === 0)!;
    expect((list as unknown as { type: { attrs: Record<string, string> } }).type.attrs["data-list"]).toBe("•");
  });
});

describe("paint format", () => {
  const m = schema.marks;
  const rich = () => s.paragraph.create({ align: "center", spaceAfter: 12 }, schema.text("Source", [m.bold.create(), m.color.create({ color: "#FF0000" })]));

  it("copies the marks at the cursor and paints them over the selected text, replacing its own", () => {
    const src = state([rich(), s.paragraph.create(null, schema.text("target text", [m.italic.create()]))], 3);
    const format = cmd.copyFormat(src);
    const target = src.apply(src.tr.setSelection(TextSelection.create(src.doc, 9, 15)));
    const st = run(target, cmd.paintFormat(format));
    const marks = (pos: number) => st.doc.resolve(pos).marks().map((x) => x.type.name).sort();
    expect(marks(11)).toEqual(["bold", "color"]);
    expect(marks(17)).toEqual(["italic"]);
    // Only part of the paragraph: its paragraph formatting stays.
    expect(attrs(st, 1).align).toBeNull();
  });

  it("copies paragraph formatting too when the paint covers whole paragraphs", () => {
    const src = state([rich(), para("whole")], 3);
    const format = cmd.copyFormat(src);
    const target = src.apply(src.tr.setSelection(TextSelection.create(src.doc, 9, 14)));
    const st = run(target, cmd.paintFormat(format));
    expect(attrs(st, 1)).toMatchObject({ align: "center", spaceAfter: 12 });
  });
});
