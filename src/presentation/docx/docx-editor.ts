/*
 * ProseMirror commands, keymap, input rules and decorations for the docx
 * editor. No Obsidian imports, so everything here runs under jsdom tests.
 */
import { Attrs, MarkType, Node as PMNode, ResolvedPos } from "prosemirror-model";
import { Command, EditorState, Plugin, Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { baseKeymap, chainCommands, toggleMark } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { InputRule, inputRules } from "prosemirror-inputrules";
import { attrsOf, docxSchema as schema, Align, ParagraphAttrs } from "../../domain/docx/schema";
import { DocxStyles } from "../../domain/docx/docx-styles";
import type { ListKind } from "../../infrastructure/docx/docx-document";

export interface EditorContext {
  /** numId of the list definition used for new bullet / numbered lists. */
  ensureList(kind: ListKind): string;
  listKind(numId: string): ListKind | null;
  styles: DocxStyles;
}

const MAX_LEVEL = 8;
const CODE_FONT = "Courier New";

/** Positions and nodes of the paragraphs touched by the selection. */
function selectedParagraphs(state: EditorState): { pos: number; node: PMNode }[] {
  const { from, to } = state.selection;
  const out: { pos: number; node: PMNode }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type === schema.nodes.paragraph) out.push({ pos, node });
    return node.type !== schema.nodes.paragraph;
  });
  return out;
}

/** Command updating the attrs of every selected paragraph. */
function updateParagraphs(fn: (attrs: ParagraphAttrs) => Partial<ParagraphAttrs>): Command {
  return (state, dispatch) => {
    const paras = selectedParagraphs(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { pos, node } of paras) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...fn(attrsOf<ParagraphAttrs>(node)) });
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

interface CellInfo {
  table: PMNode;
  tablePos: number;
  row: number;
  col: number;
}

function cellAround($pos: ResolvedPos): CellInfo | null {
  for (let d = $pos.depth; d > 2; d--) {
    if ($pos.node(d).type === schema.nodes.table_cell) {
      return { table: $pos.node(d - 2), tablePos: $pos.before(d - 2), row: $pos.index(d - 2), col: $pos.index(d - 1) };
    }
  }
  return null;
}

const newCell = () => schema.nodes.table_cell.create(null, schema.nodes.paragraph.create());

/** Command rebuilding the table around the cursor; `fn` returns new rows (or [] to delete the table). */
function editTable(fn: (rows: PMNode[][], info: CellInfo) => PMNode[][]): Command {
  return (state, dispatch) => {
    const info = cellAround(state.selection.$from);
    if (!info) return false;
    if (dispatch) {
      const rows: PMNode[][] = [];
      info.table.forEach((row) => {
        const cells: PMNode[] = [];
        row.forEach((c) => cells.push(c));
        rows.push(cells);
      });
      const next = fn(rows, info).filter((r) => r.length);
      const end = info.tablePos + info.table.nodeSize;
      const tr = next.length
        ? state.tr.replaceWith(
            info.tablePos,
            end,
            schema.nodes.table.create(
              info.table.attrs,
              next.map((cells, i) => schema.nodes.table_row.create(info.table.maybeChild(i)?.attrs ?? null, cells))
            )
          )
        : state.tr.delete(info.tablePos, end);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function insertBlockAfter(node: PMNode): Command {
  return (state, dispatch) => {
    const $from = state.selection.$from;
    const pos = $from.depth > 0 ? $from.after(1) : $from.pos;
    if (dispatch) dispatch(state.tr.insert(pos, node).scrollIntoView());
    return true;
  };
}

export const cmd = {
  setStyle: (style: string | null): Command => updateParagraphs(() => ({ style })),
  setAlign: (align: Align | null): Command => updateParagraphs(() => ({ align })),
  setLineSpacing: (lineSpacing: string | null): Command => updateParagraphs(() => ({ lineSpacing })),

  toggleList(ctx: EditorContext, kind: ListKind): Command {
    return (state, dispatch) => {
      const paras = selectedParagraphs(state);
      const allOn = paras.length > 0 && paras.every(({ node }) => {
        const { numId } = attrsOf<ParagraphAttrs>(node);
        return numId !== null && ctx.listKind(numId) === kind;
      });
      if (!dispatch) return paras.length > 0;
      const numId = allOn ? null : ctx.ensureList(kind);
      return updateParagraphs((a) => ({ numId, ilvl: numId ? a.ilvl : 0 }))(state, dispatch);
    };
  },

  /** Tab / Shift-Tab: list level inside a list, paragraph indent outside. */
  shiftLevel(delta: number): Command {
    return updateParagraphs((a) => {
      if (!a.numId) return { indent: Math.max(0, a.indent + delta) };
      const ilvl = a.ilvl + delta;
      return ilvl < 0 ? { numId: null, ilvl: 0 } : { ilvl: Math.min(MAX_LEVEL, ilvl) };
    });
  },

  /** Enter in an empty list item ends the list instead of adding another item. */
  exitEmptyListItem: ((state, dispatch) => {
    const { $from, empty } = state.selection;
    const p = $from.parent;
    if (!empty || p.type !== schema.nodes.paragraph || !p.attrs.numId || p.content.size) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(), undefined, { ...p.attrs, numId: null, ilvl: 0 }));
    return true;
  }) as Command,

  /** Apply a mark with attrs (or remove it when attrs is null) to the selection or the next typed text. */
  setMark(type: MarkType, attrs: Attrs | null): Command {
    return (state, dispatch) => {
      const { from, to, empty } = state.selection;
      if (dispatch) {
        const tr = state.tr;
        if (empty) {
          tr.removeStoredMark(type);
          if (attrs) tr.addStoredMark(type.create(attrs));
        } else {
          tr.removeMark(from, to, type);
          if (attrs) tr.addMark(from, to, type.create(attrs));
        }
        dispatch(tr);
      }
      return true;
    };
  },

  /** Remove all direct formatting (and the original run properties), keeping links. */
  clearFormatting: ((state, dispatch) => {
    const { from, to } = state.selection;
    if (dispatch) {
      const tr = state.tr;
      for (const type of Object.values(schema.marks)) if (type !== schema.marks.link) tr.removeMark(from, to, type);
      dispatch(tr.setStoredMarks([]));
    }
    return true;
  }) as Command,

  insertPageBreak: ((state, dispatch) => {
    if (dispatch) dispatch(state.tr.replaceSelectionWith(schema.nodes.page_break.create()).scrollIntoView());
    return true;
  }) as Command,

  insertRule: insertBlockAfter(schema.nodes.raw_block.create({ kind: "hr" })),

  insertTable(rows: number, cols: number): Command {
    const row = () => schema.nodes.table_row.create(null, Array.from({ length: cols }, newCell));
    return insertBlockAfter(schema.nodes.table.create(null, Array.from({ length: rows }, row)));
  },

  addRow: (after: boolean): Command =>
    editTable((rows, { row }) => {
      rows.splice(row + (after ? 1 : 0), 0, rows[row].map(newCell));
      return rows;
    }),

  addColumn: (after: boolean): Command =>
    editTable((rows, { col }) => rows.map((cells) => (cells.splice(Math.min(col + (after ? 1 : 0), cells.length), 0, newCell()), cells))),

  deleteRow: editTable((rows, { row }) => rows.filter((_, i) => i !== row)),
  deleteColumn: editTable((rows, { col }) => rows.map((cells) => cells.filter((_, i) => i !== col))),
  deleteTable: editTable(() => []),

  setCellBackground(background: string | null): Command {
    return (state, dispatch) => {
      const $from = state.selection.$from;
      for (let d = $from.depth; d > 0; d--) {
        if ($from.node(d).type !== schema.nodes.table_cell) continue;
        if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(d), undefined, { ...$from.node(d).attrs, background }));
        return true;
      }
      return false;
    };
  },

  toggleCode: ((state, dispatch) => {
    const { from, to, empty, $from } = state.selection;
    const font = schema.marks.font;
    const active = empty
      ? (state.storedMarks ?? $from.marks()).some((m) => m.type === font && m.attrs.font === CODE_FONT)
      : state.doc.rangeHasMark(from, to, font.create({ font: CODE_FONT }));
    return cmd.setMark(font, active ? null : { font: CODE_FONT })(state, dispatch);
  }) as Command,
};

// -----------------------------------------------------------------------------
// Plugins
// -----------------------------------------------------------------------------

/** Markdown-style shortcuts at the start of a paragraph: "- ", "1. ", "# ", "> ". */
function paragraphRule(re: RegExp, attrs: (m: RegExpMatchArray, ctx: EditorContext) => Partial<Attrs>, ctx: EditorContext): InputRule {
  return new InputRule(re, (state, match, start, end) => {
    const $start = state.doc.resolve(start);
    const p = $start.parent;
    if (p.type !== schema.nodes.paragraph) return null;
    return state.tr.delete(start, end).setNodeMarkup($start.before(), undefined, { ...p.attrs, ...attrs(match, ctx) });
  });
}

function decorations(styles: DocxStyles, doc: PMNode): DecorationSet {
  const marks = styles.listMarks(doc);
  const decos: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    const a = attrsOf<ParagraphAttrs>(node);
    const css = [styles.paragraphCss(a.style)];
    // Direct paragraph formatting must win over the style, so it goes last.
    if (a.align) css.push(`text-align: ${a.align}`);
    if (node.attrs.lineSpacing) css.push(`line-height: ${node.attrs.lineSpacing}`);
    const attrs: Record<string, string> = {};
    const mark = marks.get(pos);
    if (mark) {
      css.push(`padding-left: ${mark.indentPt}pt`, `--docx-hanging: ${mark.hangingPt}pt`);
      attrs.class = "docx-list";
      attrs["data-list"] = mark.text;
    } else if (node.attrs.indent) {
      css.push(`margin-left: ${node.attrs.indent * 36}pt`);
    }
    attrs.style = css.filter(Boolean).join("; ");
    decos.push(Decoration.node(pos, pos + node.nodeSize, attrs));
    return false;
  });
  return DecorationSet.create(doc, decos);
}

export function docxPlugins(ctx: EditorContext, extraKeys: Record<string, Command> = {}): Plugin[] {
  const m = schema.marks;
  return [
    history(),
    inputRules({
      rules: [
        paragraphRule(/^\s*[-*]\s$/, () => ({ numId: ctx.ensureList("bullet"), ilvl: 0 }), ctx),
        paragraphRule(/^1[.)]\s$/, () => ({ numId: ctx.ensureList("ordered"), ilvl: 0 }), ctx),
        paragraphRule(/^(#{1,6})\s$/, (match) => ({ style: `Heading${match[1].length}` }), ctx),
        paragraphRule(/^>\s$/, () => ({ style: "Quote" }), ctx),
      ],
    }),
    keymap({
      ...extraKeys,
      Enter: chainCommands(cmd.exitEmptyListItem, baseKeymap.Enter),
      Tab: cmd.shiftLevel(1),
      "Shift-Tab": cmd.shiftLevel(-1),
      "Mod-z": undo,
      "Mod-y": redo,
      "Mod-Shift-z": redo,
      "Mod-b": toggleMark(m.bold),
      "Mod-i": toggleMark(m.italic),
      "Mod-u": toggleMark(m.underline),
      "Mod-Shift-x": toggleMark(m.strike),
      "Mod-`": cmd.toggleCode,
      "Mod-Enter": cmd.insertPageBreak,
    }),
    keymap(baseKeymap),
    new Plugin({
      state: {
        init: (_, state) => decorations(ctx.styles, state.doc),
        apply: (tr: Transaction, old: DecorationSet) => (tr.docChanged ? decorations(ctx.styles, tr.doc) : old),
      },
      props: {
        decorations(state) {
          return this.getState(state);
        },
      },
    }),
  ];
}

