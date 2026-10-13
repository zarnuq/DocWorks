/*
 * Google Docs' cell selection: dragging from one table cell into another selects the block of cells between them.
 * Each selected cell is one range, so marks and paragraph commands that walk `selection.ranges` touch only those cells;
 * Delete clears them, typing replaces them, copy gives a table of just them, and pasting cells overwrites cells.
 * The text selection is hidden and the cells are tinted by a decoration instead.
 */
import { Fragment, Node as PMNode, ResolvedPos, Slice } from "prosemirror-model";
import { EditorState, Plugin, Selection, SelectionRange, TextSelection, Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { attrsOf, CellAttrs, docxSchema as schema } from "../../domain/docx/schema";

const s = schema.nodes;
const isCell = (doc: PMNode, pos: number) => doc.nodeAt(pos)?.type === s.table_cell;

/** Table position, row and cell index of the cell at `pos`. */
function locate(doc: PMNode, pos: number) {
  const $cell = doc.resolve(pos);
  const d = $cell.depth;
  return { table: $cell.node(d - 1), tablePos: $cell.before(d - 1), row: $cell.index(d - 1), col: $cell.index(d) };
}

/** Position of the innermost cell around `$pos`, or null. */
export function cellBefore($pos: ResolvedPos): number | null {
  for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type === s.table_cell) return $pos.before(d);
  return null;
}

const sameTable = (doc: PMNode, a: number, b: number) => locate(doc, a).tablePos === locate(doc, b).tablePos;

/** Cells in the rows and grid columns spanned by cells `a` and `b`, in document order. */
function cellRect(doc: PMNode, a: number, b: number): number[] {
  const A = locate(doc, a), B = locate(doc, b);
  const gridOf = (row: PMNode, col: number) => {
    let g = 0;
    for (let i = 0; i < col; i++) g += attrsOf<CellAttrs>(row.child(i)).colspan;
    return [g, g + attrsOf<CellAttrs>(row.child(col)).colspan - 1];
  };
  const [fa, la] = gridOf(A.table.child(A.row), A.col), [fb, lb] = gridOf(B.table.child(B.row), B.col);
  const [g0, g1, r0, r1] = [Math.min(fa, fb), Math.max(la, lb), Math.min(A.row, B.row), Math.max(A.row, B.row)];
  const out: number[] = [];
  let rowPos = A.tablePos + 1;
  A.table.forEach((row, _, r) => {
    let cellPos = rowPos + 1, g = 0;
    row.forEach((cell) => {
      const span = attrsOf<CellAttrs>(cell).colspan;
      if (r >= r0 && r <= r1 && g <= g1 && g + span - 1 >= g0) out.push(cellPos);
      g += span;
      cellPos += cell.nodeSize;
    });
    rowPos += row.nodeSize;
  });
  return out;
}

export class CellSelection extends Selection {
  readonly anchorCell: number;
  readonly headCell: number;
  /** Positions of the selected cells, in document order. */
  readonly cells: number[];

  constructor(anchorCell: number, headCell: number, doc: PMNode) {
    const cells = cellRect(doc, anchorCell, headCell);
    const ranges = cells.map((p) => new SelectionRange(doc.resolve(p + 1), doc.resolve(p + doc.nodeAt(p)!.nodeSize - 1)));
    super(ranges[0].$from, ranges[ranges.length - 1].$to, ranges);
    this.anchorCell = anchorCell;
    this.headCell = headCell;
    this.cells = cells;
  }

  static create(doc: PMNode, anchorCell: number, headCell = anchorCell): CellSelection {
    return new CellSelection(anchorCell, headCell, doc);
  }

  map(doc: PMNode, mapping: Parameters<Selection["map"]>[1]): Selection {
    const a = mapping.mapResult(this.anchorCell), h = mapping.mapResult(this.headCell);
    if (!a.deleted && !h.deleted && isCell(doc, a.pos) && isCell(doc, h.pos) && sameTable(doc, a.pos, h.pos)) return new CellSelection(a.pos, h.pos, doc);
    return Selection.near(doc.resolve(Math.min(mapping.map(this.head), doc.content.size)));
  }

  /** A table of just the selected cells, so copying a block pastes as one. */
  content(): Slice {
    const doc = this.$anchor.doc;
    const { table } = locate(doc, this.cells[0]);
    const rows = new Map<number, PMNode[]>();
    for (const p of this.cells) {
      const { row } = locate(doc, p);
      rows.set(row, [...(rows.get(row) ?? []), doc.nodeAt(p)!]);
    }
    return new Slice(Fragment.from(table.type.create(table.attrs, [...rows.values()].map((cells) => s.table_row.create(null, cells)))), 0, 0);
  }

  /** Clear every selected cell (keeping its first paragraph's formatting), then put `content` in the first. */
  replace(tr: Transaction, content = Slice.empty): void {
    const from = tr.steps.length;
    for (const p of this.cells) {
      const pos = tr.mapping.slice(from).map(p);
      const cell = tr.doc.nodeAt(pos)!;
      const first = cell.firstChild;
      const empty = first?.type === s.paragraph ? s.paragraph.create({ ...first.attrs, key: null }) : s.paragraph.create();
      tr.replaceWith(pos + 1, pos + cell.nodeSize - 1, empty);
    }
    tr.setSelection(TextSelection.create(tr.doc, tr.mapping.slice(from).map(this.cells[0]) + 2));
    if (content.size) tr.replaceSelection(content);
  }

  replaceWith(tr: Transaction, node: PMNode): void {
    this.replace(tr, new Slice(Fragment.from(node), 0, 0));
  }

  eq(other: Selection): boolean {
    return other instanceof CellSelection && other.anchorCell === this.anchorCell && other.headCell === this.headCell;
  }

  toJSON(): { type: string; anchor: number; head: number } {
    return { type: "docxcell", anchor: this.anchorCell, head: this.headCell };
  }

  static fromJSON(doc: PMNode, json: { anchor: number; head: number }): CellSelection {
    return new CellSelection(json.anchor, json.head, doc);
  }
}
CellSelection.prototype.visible = false;
Selection.jsonID("docxcell", CellSelection);

/** Rows of cells in a pasted slice (a table, or rows / cells cut out of one), or null when it holds none. */
function pastedRows(slice: Slice): PMNode[][] | null {
  const rows: PMNode[][] = [];
  const loose: PMNode[] = [];
  slice.content.forEach((n) => {
    if (n.type === s.table_cell) loose.push(n);
    else if (n.type === s.table_row || n.type === s.table) {
      n.descendants((m) => {
        if (m.type !== s.table_row) return m.type === s.table;
        const cells: PMNode[] = [];
        m.forEach((c) => cells.push(c));
        rows.push(cells);
        return false;
      });
      if (n.type === s.table_row) rows.push(Array.from({ length: n.childCount }, (_, i) => n.child(i)));
    }
  });
  if (loose.length) rows.push(loose);
  return rows.length ? rows : null;
}

/** Paste copied cells over the cells from the selection's first cell on, as Word's "overwrite cells"; what falls outside the table is dropped. */
function pasteCells(state: EditorState, slice: Slice): Transaction | null {
  const rows = pastedRows(slice);
  const sel = state.selection;
  const target = sel instanceof CellSelection ? sel.cells[0] : cellBefore(sel.$from);
  if (!rows || target === null) return null;
  const { table, tablePos, row, col } = locate(state.doc, target);
  const jobs: { pos: number; size: number; content: Fragment }[] = [];
  let rowPos = tablePos + 1;
  table.forEach((r, _, i) => {
    const src = rows[i - row];
    let cellPos = rowPos + 1;
    r.forEach((cell, __, j) => {
      if (src && j >= col && src[j - col]) jobs.push({ pos: cellPos, size: cell.nodeSize, content: src[j - col].content });
      cellPos += cell.nodeSize;
    });
    rowPos += r.nodeSize;
  });
  const tr = state.tr;
  for (const job of jobs.reverse()) tr.replaceWith(job.pos + 1, job.pos + job.size - 1, job.content);
  return tr.scrollIntoView();
}

/** The cell selection's mouse handling, tint, paste and drag rules. */
export function cellSelectionPlugin(): Plugin {
  let dragFrom: number | null = null;
  const cellAt = (view: EditorView, target: EventTarget | null): number | null => {
    const td = target instanceof Element ? target.closest("td") : null;
    if (!td || !view.dom.contains(td)) return null;
    const pos = view.posAtDOM(td, 0) - 1;
    return isCell(view.state.doc, pos) ? pos : null;
  };
  return new Plugin({
    props: {
      decorations(state) {
        const sel = state.selection;
        if (!(sel instanceof CellSelection)) return null;
        return DecorationSet.create(state.doc, sel.cells.map((p) => Decoration.node(p, p + state.doc.nodeAt(p)!.nodeSize, { class: "is-cell-selected" })));
      },
      // Reading the browser's selection: while a drag is choosing cells keep them; otherwise two cells make a cell selection (Shift+arrows).
      createSelectionBetween(view, $anchor, $head) {
        if (dragFrom !== null && view.state.selection instanceof CellSelection) return view.state.selection;
        const a = cellBefore($anchor), h = cellBefore($head);
        if (a === null || h === null || a === h || !sameTable($anchor.doc, a, h)) return null;
        return new CellSelection(a, h, $anchor.doc);
      },
      handlePaste(view, _event, slice) {
        const tr = pasteCells(view.state, slice);
        if (!tr) return false;
        view.dispatch(tr);
        return true;
      },
      handleDOMEvents: {
        mousedown(view, e) {
          if (e.button !== 0) return false;
          const start = cellAt(view, e.target);
          if (start === null) return false;
          dragFrom = start;
          const move = (ev: MouseEvent) => {
            const over = cellAt(view, ev.target);
            if (dragFrom === null || over === null || !sameTable(view.state.doc, dragFrom, over)) return;
            const sel = view.state.selection;
            if (over === dragFrom && !(sel instanceof CellSelection)) return;
            const next = over === dragFrom ? TextSelection.near(view.state.doc.resolve(dragFrom + 2)) : new CellSelection(dragFrom, over, view.state.doc);
            if (!next.eq(sel)) view.dispatch(view.state.tr.setSelection(next));
          };
          const stop = () => {
            dragFrom = null;
            window.removeEventListener("mousemove", move);
            window.removeEventListener("mouseup", stop);
          };
          window.addEventListener("mousemove", move);
          window.addEventListener("mouseup", stop);
          return false;
        },
        // Dropping cells would insert them into whatever row they land in; Google Docs doesn't drag a cell selection either.
        dragstart(view, e) {
          if (!(view.state.selection instanceof CellSelection)) return false;
          e.preventDefault();
          return true;
        },
      },
    },
  });
}
