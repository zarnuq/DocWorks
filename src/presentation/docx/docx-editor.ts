/*
 * ProseMirror commands, keymap, input rules and decorations for the docx
 * editor. No Obsidian imports, so everything here runs under jsdom tests.
 */
import { Attrs, MarkType, Node as PMNode, ResolvedPos } from "prosemirror-model";
import { Command, EditorState, Plugin, PluginKey, Transaction } from "prosemirror-state";
import { Decoration, DecorationSet, EditorView } from "prosemirror-view";
import { layoutPages, PageBlock, PageSetup } from "../../domain/docx/pages";
import { baseKeymap, chainCommands, toggleMark } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { InputRule, inputRules } from "prosemirror-inputrules";
import { attrsOf, docxSchema as schema, HEADING_LEVEL, Align, CellAttrs, ParagraphAttrs, TableAttrs, TableBorders } from "../../domain/docx/schema";
import { DocxStyles, lineHeight } from "../../domain/docx/docx-styles";
import type { ListKind } from "../../infrastructure/docx/docx-document";
import { findPlugin } from "./docx-find";

export interface EditorContext {
  /** numId of the list definition used for new bullet / numbered lists. */
  ensureList(kind: ListKind): string;
  listKind(numId: string): ListKind | null;
  styles: DocxStyles;
}

const MAX_LEVEL = 8;
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
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
      const next = fn(rowsOf(info.table), info).filter((r) => r.length);
      const tr = next.length
        ? replaceTable(state, info, next)
        : state.tr.delete(info.tablePos, info.tablePos + info.table.nodeSize);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** Rows of the table as arrays of cells. */
const rowsOf = (table: PMNode) => {
  const rows: PMNode[][] = [];
  table.forEach((row) => {
    const cells: PMNode[] = [];
    row.forEach((c) => cells.push(c));
    rows.push(cells);
  });
  return rows;
};

/** Grid column where cell `col` of `cells` starts. */
const gridStart = (cells: PMNode[], col: number) => cells.slice(0, col).reduce((n, c) => n + attrsOf<CellAttrs>(c).colspan, 0);
/** Index of the cell starting at grid column `g`, or -1. */
const cellAtGrid = (cells: PMNode[], g: number) => cells.findIndex((_, i) => gridStart(cells, i) === g);

const isEmptyCell = (cell: PMNode) => cell.childCount === 1 && cell.firstChild!.content.size === 0;
/** Cell content followed by another cell's content (skipping an empty side). */
function joinContent(a: PMNode, b: PMNode): PMNode[] {
  const nodes: PMNode[] = [];
  if (!isEmptyCell(a) || isEmptyCell(b)) a.forEach((n) => nodes.push(n));
  if (!isEmptyCell(b)) b.forEach((n) => nodes.push(n));
  return nodes;
}

const emptyContent = () => schema.nodes.paragraph.create();

/** Replace the table around the cursor with new rows and attrs. */
function replaceTable(state: EditorState, info: CellInfo, rows: PMNode[][], attrs = info.table.attrs): Transaction {
  const table = schema.nodes.table.create(
    attrs,
    rows.map((cells, i) => schema.nodes.table_row.create(info.table.maybeChild(i)?.attrs ?? null, cells))
  );
  return state.tr.replaceWith(info.tablePos, info.tablePos + info.table.nodeSize, table);
}

function tableCommand(fn: (rows: PMNode[][], info: CellInfo) => boolean): Command {
  return (state, dispatch) => {
    const info = cellAround(state.selection.$from);
    if (!info) return false;
    const rows = rowsOf(info.table);
    if (!fn(rows, info)) return false;
    if (dispatch) dispatch(replaceTable(state, info, rows).scrollIntoView());
    return true;
  };
}

function setTableAttrs(fn: (a: TableAttrs, info: CellInfo) => Partial<TableAttrs>): Command {
  return (state, dispatch) => {
    const info = cellAround(state.selection.$from);
    if (!info) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup(info.tablePos, undefined, { ...info.table.attrs, ...fn(attrsOf<TableAttrs>(info.table), info) }));
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

/** Start and end of the link mark run around `pos` inside its paragraph. */
function linkExtent(state: EditorState, pos: number): { from: number; to: number } | null {
  const $pos = state.doc.resolve(pos);
  const link = ($pos.marks().find((x) => x.type === schema.marks.link) ?? $pos.nodeAfter?.marks.find((x) => x.type === schema.marks.link)) || null;
  if (!link) return null;
  const start = $pos.start();
  let from = -1;
  let to = -1;
  $pos.parent.forEach((child, offset) => {
    const a = start + offset;
    const has = link.isInSet(child.marks);
    if (has && from < 0 && a <= pos) from = a;
    if (!has && a < pos) from = -1;
    if (has && from >= 0 && (to < 0 || to === a)) to = a + child.nodeSize;
  });
  return from >= 0 && to >= pos ? { from, to } : null;
}

/** A PAGE field (Word fills in the page number), shown as "1" until the page view numbers it. */
export function pageNumberNode(): PMNode {
  return schema.nodes.opaque.create({ xml: `<w:fldSimple xmlns:w="${NS_W}" w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>`, text: "1" });
}

const opaqueXml = (xml: string) => schema.nodes.opaque.create({ xml, text: "" });

/** Comment id -> the text range its w:commentRangeStart / End markers enclose. */
export function commentRanges(doc: PMNode): Map<string, { from: number; to: number }> {
  const starts = new Map<string, number>();
  const out = new Map<string, { from: number; to: number }>();
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.opaque) return true;
    const m = /commentRange(Start|End)\b[^>]*w:id="(-?\d+)"/.exec(node.attrs.xml as string);
    if (m?.[1] === "Start") starts.set(m[2], pos + 1);
    else if (m && starts.has(m[2])) out.set(m[2], { from: starts.get(m[2])!, to: pos });
    return false;
  });
  return out;
}

/** Remove matching opaque markers from the end, keeping earlier positions valid. */
function removeOpaqueNodes(matches: (xml: string) => boolean): Command {
  return (state, dispatch) => {
    const at: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type === schema.nodes.opaque && matches(node.attrs.xml as string)) at.push(pos);
      return true;
    });
    if (!at.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const pos of at.reverse()) tr.delete(pos, pos + 1);
      dispatch(tr);
    }
    return true;
  };
}

const FOOTNOTE_RE = /footnoteReference\b[^>]*w:id="(-?\d+)"/;

/** Footnote ids in the order their references appear. */
export function footnoteIds(doc: PMNode): string[] {
  const ids: string[] = [];
  doc.descendants((node) => {
    const m = node.type === schema.nodes.opaque ? FOOTNOTE_RE.exec(node.attrs.xml as string) : null;
    if (m) ids.push(m[1]);
    return true;
  });
  return ids;
}

/** Headings 1–3 for the table of contents. */
export function tocEntries(doc: PMNode): { level: number; text: string }[] {
  const out: { level: number; text: string }[] = [];
  doc.descendants((node) => {
    const style = node.type === schema.nodes.paragraph ? attrsOf<ParagraphAttrs>(node).style : null;
    const level = style && style !== "Title" ? HEADING_LEVEL[style] : undefined;
    if (level && level <= 3 && node.textContent.trim()) out.push({ level, text: node.textContent.trim() });
    return node.type !== schema.nodes.paragraph;
  });
  return out;
}

/** A Word TOC field (\o "1-3") with the current headings as its cached result; Word refreshes it on open. */
function tocNode(doc: PMNode): PMNode {
  const entries = tocEntries(doc);
  const esc = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const para = (style: string | null, runs: string) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}${runs}</w:p>`;
  const text = (t: string) => `<w:r><w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
  const begin = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`;
  const rows = entries.length ? entries : [{ level: 1, text: "No headings yet" }];
  const body = rows.map((e, i) => para(`TOC${e.level}`, (i === 0 ? begin : "") + text(e.text))).join("");
  const xml =
    `<w:sdt xmlns:w="${NS_W}"><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>` +
    para("TOCHeading", text("Contents")) + body + para(null, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`) + `</w:sdtContent></w:sdt>`;
  const lines = ["Contents", ...rows.map((e) => `${"    ".repeat(e.level - 1)}${e.text}`)];
  return schema.nodes.raw_block.create({ xml, text: lines.join("\n"), kind: "toc" });
}

/** Address of the link at the cursor or selection start, or null. */
export function linkHref(state: EditorState): string | null {
  const { $from } = state.selection;
  const link = [...$from.marks(), ...($from.nodeAfter?.marks ?? [])].find((x) => x.type === schema.marks.link);
  return (link?.attrs.href as string | null | undefined) ?? null;
}

export const cmd = {
  insertToc: ((state, dispatch) => insertBlockAfter(tocNode(state.doc))(state, dispatch)) as Command,

  /** Rebuild every table of contents from the current headings. */
  updateToc: ((state, dispatch) => {
    const at: { pos: number; size: number }[] = [];
    state.doc.forEach((node, pos) => {
      if (node.type === schema.nodes.raw_block && node.attrs.kind === "toc") at.push({ pos, size: node.nodeSize });
    });
    if (!at.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { pos, size } of at.reverse()) tr.replaceWith(pos, pos + size, tocNode(state.doc));
      dispatch(tr);
    }
    return true;
  }) as Command,

  /** Mark the selection as the range of comment `id` (its text lives in comments.xml). */
  addCommentAnchors(id: string): Command {
    return (state, dispatch) => {
      const { from, to, empty } = state.selection;
      if (empty) return false;
      if (dispatch) {
        const tr = state.tr;
        tr.insert(to, [
          opaqueXml(`<w:commentRangeEnd xmlns:w="${NS_W}" w:id="${id}"/>`),
          opaqueXml(`<w:r xmlns:w="${NS_W}"><w:commentReference w:id="${id}"/></w:r>`),
        ]);
        tr.insert(from, opaqueXml(`<w:commentRangeStart xmlns:w="${NS_W}" w:id="${id}"/>`));
        dispatch(tr);
      }
      return true;
    };
  },

  /** Put the reference to footnote `id` after the selection. */
  insertFootnote(id: string): Command {
    return (state, dispatch) => {
      if (dispatch) {
        const xml = `<w:r xmlns:w="${NS_W}"><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
        dispatch(state.tr.insert(state.selection.to, schema.nodes.opaque.create({ xml, text: "*" })).scrollIntoView());
      }
      return true;
    };
  },

  removeFootnoteRef(id: string): Command {
    return removeOpaqueNodes((xml) => FOOTNOTE_RE.exec(xml)?.[1] === id);
  },

  removeCommentAnchors(id: string): Command {
    const re = new RegExp(`(commentRangeStart|commentRangeEnd|commentReference)\\b[^>]*w:id="${id}"`);
    return removeOpaqueNodes((xml) => re.test(xml));
  },

  /**
   * Link the selection to `href` (null removes the link). With only a caret, the
   * link around it changes; outside a link the address is inserted as linked text.
   */
  setLink(href: string | null): Command {
    return (state, dispatch) => {
      const { from, to, empty } = state.selection;
      const range = empty ? linkExtent(state, from) : { from, to };
      if (!range && !href) return false;
      if (dispatch) {
        const tr = state.tr;
        const mark = href ? schema.marks.link.create({ href, rId: null, anchor: null }) : null;
        if (range) {
          tr.removeMark(range.from, range.to, schema.marks.link);
          if (mark) tr.addMark(range.from, range.to, mark);
        } else if (mark && href) {
          tr.insert(from, schema.text(href, [...state.selection.$from.marks(), mark]));
        }
        dispatch(tr.scrollIntoView());
      }
      return true;
    };
  },

  setStyle: (style: string | null): Command => updateParagraphs(() => ({ style })),
  /** Spacing, indents or other paragraph attributes for the selected paragraphs. */
  setParagraph: (patch: Partial<ParagraphAttrs>): Command => updateParagraphs(() => patch),
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
      for (const type of Object.values(schema.marks)) if (type !== schema.marks.link && type !== schema.marks.insertion) tr.removeMark(from, to, type);
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

  /** Merge the cell at the cursor with the cell to its right. */
  mergeRight: tableCommand((rows, { row, col }) => {
    const [cell, next] = [rows[row][col], rows[row][col + 1]];
    if (!next) return false;
    const a = attrsOf<CellAttrs>(cell);
    rows[row].splice(col, 2, cell.type.create({ ...a, colspan: a.colspan + attrsOf<CellAttrs>(next).colspan }, joinContent(cell, next)));
    return true;
  }),

  /** Merge the cell at the cursor (or the merge it starts) with the cell below. */
  mergeDown: tableCommand((rows, { row, col }) => {
    const g = gridStart(rows[row], col);
    const span = attrsOf<CellAttrs>(rows[row][col]).colspan;
    let top = row;
    while (top > 0 && attrsOf<CellAttrs>(rows[top][col]).vMerge === "continue") top--;
    let last = row;
    const at = (r: number) => (rows[r] ? cellAtGrid(rows[r], g) : -1);
    while (at(last + 1) >= 0 && attrsOf<CellAttrs>(rows[last + 1][at(last + 1)]).vMerge === "continue") last++;
    const below = at(last + 1);
    if (below < 0 || attrsOf<CellAttrs>(rows[last + 1][below]).colspan !== span) return false;
    const topCol = cellAtGrid(rows[top], g);
    const origin = rows[top][topCol];
    const next = rows[last + 1][below];
    rows[top][topCol] = origin.type.create({ ...origin.attrs, vMerge: "restart" }, joinContent(origin, next));
    rows[last + 1][below] = next.type.create({ ...next.attrs, vMerge: "continue" }, emptyContent());
    return true;
  }),

  /** Undo a horizontal or vertical merge at the cursor. */
  splitCell: tableCommand((rows, { row, col }) => {
    const cell = rows[row][col];
    const a = attrsOf<CellAttrs>(cell);
    if (a.colspan > 1) {
      rows[row].splice(col, 1, cell.type.create({ ...a, colspan: 1 }, cell.content), ...Array.from({ length: a.colspan - 1 }, () => newCell()));
      return true;
    }
    if (!a.vMerge) return false;
    const g = gridStart(rows[row], col);
    let r = row;
    while (r > 0 && attrsOf<CellAttrs>(rows[r][cellAtGrid(rows[r], g)]).vMerge === "continue") r--;
    for (let k = r; k < rows.length; k++) {
      const i = cellAtGrid(rows[k], g);
      const c = i >= 0 ? rows[k][i] : null;
      if (!c || (k > r && attrsOf<CellAttrs>(c).vMerge !== "continue")) break;
      rows[k][i] = c.type.create({ ...c.attrs, vMerge: null }, c.content);
    }
    return true;
  }),

  setTableBorders: (borders: TableBorders): Command => setTableAttrs(() => ({ borders })),

  /** Width of the cursor's column in twips (new tables share 6.5" evenly first). */
  setColumnWidth: (twips: number): Command =>
    setTableAttrs((a, { table, row, col }) => {
      const cols = rowsOf(table).reduce((n, cells) => Math.max(n, gridStart(cells, cells.length)), 0);
      const widths = a.widths?.length === cols ? [...a.widths] : Array.from({ length: cols }, () => Math.round(9360 / cols));
      const cells = rowsOf(table)[row];
      widths[gridStart(cells, col)] = Math.max(144, Math.round(twips));
      return { widths };
    }),

  /** Make rows from the top down to the cursor's row repeat on every page (or stop that). */
  toggleHeaderRow: ((state, dispatch) => {
    const info = cellAround(state.selection.$from);
    if (!info) return false;
    if (dispatch) {
      const on = !info.table.child(info.row).attrs.header;
      const tr = state.tr;
      let pos = info.tablePos + 1;
      info.table.forEach((row, _o, i) => {
        if (i <= info.row || !on) tr.setNodeMarkup(pos, undefined, { ...row.attrs, header: on && i <= info.row });
        pos += row.nodeSize;
      });
      dispatch(tr);
    }
    return true;
  }) as Command,

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

/** Display CSS for a paragraph: its style (inside a table, layered over the table style), then its direct formatting. */
function paragraphCss(styles: DocxStyles, doc: PMNode, node: PMNode, pos: number): string {
  const a = attrsOf<ParagraphAttrs>(node);
  const $p = doc.resolve(pos);
  let tableStyle: string | null = null;
  for (let d = $p.depth; d > 0 && !tableStyle; d--) if ($p.node(d).type === schema.nodes.table) tableStyle = attrsOf<TableAttrs>($p.node(d)).style;
  const css = [styles.paragraphCss(a.style, tableStyle)];
  // Direct paragraph formatting must win over the style, so it goes last.
  if (a.align) css.push(`text-align: ${a.align}`);
  if (a.lineSpacing) css.push(`line-height: ${lineHeight(Number(a.lineSpacing), styles.fontFor(a.style, tableStyle))}`);
  if (a.spaceBefore !== null) css.push(`margin-top: ${a.spaceBefore}pt`);
  if (a.spaceAfter !== null) css.push(`margin-bottom: ${a.spaceAfter}pt`);
  if (a.firstLine) css.push(`text-indent: ${a.firstLine}pt`);
  return css.filter(Boolean).join("; ");
}

/** Give rendered paragraphs (header / footer bands, outside any editor) their document styling. */
export function styleRenderedParagraphs(root: HTMLElement, doc: PMNode, styles: DocxStyles): void {
  const els = Array.from(root.querySelectorAll<HTMLElement>("p, h1, h2, h3, h4, h5, h6"));
  let i = 0;
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.paragraph) return true;
    els[i++]?.setAttribute("style", paragraphCss(styles, doc, node, pos));
    return false;
  });
}

function decorations(styles: DocxStyles, doc: PMNode): DecorationSet {
  const marks = styles.listMarks(doc);
  const decos: Decoration[] = [];
  let note = 0;
  doc.descendants((node, pos) => {
    if (node.type === schema.nodes.opaque && FOOTNOTE_RE.test(node.attrs.xml as string)) decos.push(Decoration.node(pos, pos + 1, { "data-note": String(++note) }));
    return true;
  });
  for (const [id, r] of commentRanges(doc)) {
    if (r.to > r.from) decos.push(Decoration.inline(r.from, r.to, { class: "docx-comment-range", "data-comment": id }));
  }
  doc.descendants((node, pos) => {
    if (node.type === schema.nodes.table) {
      tableDecorations(node, pos, decos);
      return true;
    }
    if (node.type !== schema.nodes.paragraph) return true;
    const css = [paragraphCss(styles, doc, node, pos)];
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

/** Row spans for vertical merges, hidden covered cells, column widths and border presets. */
function tableDecorations(table: PMNode, tablePos: number, decos: Decoration[]): void {
  const t = attrsOf<TableAttrs>(table);
  const rows = rowsOf(table);
  let rowPos = tablePos + 1;
  rows.forEach((cells, r) => {
    let cellPos = rowPos + 1;
    cells.forEach((cell, c) => {
      const a = attrsOf<CellAttrs>(cell);
      const g = gridStart(cells, c);
      const attrs: Record<string, string> = {};
      if (a.vMerge === "continue") attrs.class = "docx-merged";
      else if (a.vMerge === "restart") {
        let span = 1;
        while (rows[r + span] && attrsOf<CellAttrs>(rows[r + span][cellAtGrid(rows[r + span], g)] ?? cell).vMerge === "continue" && cellAtGrid(rows[r + span], g) >= 0) span++;
        attrs.rowspan = String(span);
      }
      if (t.widths && a.vMerge !== "continue") {
        const w = t.widths.slice(g, g + a.colspan).reduce((n, x) => n + x, 0);
        if (w) attrs.style = `width: ${Math.round(w / 15)}px`;
      }
      if (Object.keys(attrs).length) decos.push(Decoration.node(cellPos, cellPos + cell.nodeSize, attrs));
      cellPos += cell.nodeSize;
    });
    rowPos += table.child(r).nodeSize;
  });
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
    findPlugin(),
  ];
}


// -----------------------------------------------------------------------------
// Page view
// -----------------------------------------------------------------------------

export interface PagesOptions {
  /** Current page geometry, or null for the continuous (pageless) view. */
  setup(): PageSetup | null;
  /** Header / footer content for a 1-based page number out of `total`. */
  header(page: number, total: number): Node | null;
  footer(page: number, total: number): Node | null;
}

interface PagesState {
  starts: number[];
  pages: number;
  decorations: DecorationSet;
  /** Header / footer band heights the layout assumed. */
  bands: [number, number];
}

export const pagesKey = new PluginKey<PagesState>("docx-pages");
const schedulers = new WeakMap<EditorView, () => void>();

/** Re-measure the pages (after a resize or a header / footer change). */
export function repaginate(view: EditorView): void {
  schedulers.get(view)?.();
}

function hasPageBreak(node: PMNode): boolean {
  let found = false;
  node.descendants((n) => {
    if (n.type === schema.nodes.page_break) found = true;
    return !found;
  });
  return found;
}

/** Bumped on every measure so header / footer widgets are rebuilt with fresh content. */
let pageGeneration = 0;

/** Heights of the drawn header and footer bands: a header or footer taller than its margin pushes the text in, as in Word. */
function bandHeights(view: EditorView, setup: PageSetup): [number, number] {
  const head = view.dom.querySelector<HTMLElement>(".docx-page-first")?.offsetHeight ?? 0;
  const foot = view.dom.querySelector<HTMLElement>(".docx-page-last .docx-page-foot")?.offsetHeight ?? 0;
  return [Math.max(setup.top, head), Math.max(setup.bottom, foot)];
}

function pageDecorations(view: EditorView, opts: PagesOptions): PagesState {
  const gen = ++pageGeneration;
  const setup = opts.setup();
  const doc = view.state.doc;
  if (!setup) return { starts: [], pages: 1, decorations: DecorationSet.empty, bands: [0, 0] };
  const blocks: PageBlock[] = [];
  const offsets: number[] = [];
  doc.forEach((node, offset) => {
    const dom = view.nodeDOM(offset) as HTMLElement | null;
    const cs = dom instanceof HTMLElement ? getComputedStyle(dom) : null;
    const height = dom instanceof HTMLElement ? dom.offsetHeight + (parseFloat(cs!.marginTop) || 0) + (parseFloat(cs!.marginBottom) || 0) : 0;
    blocks.push({ height, breakAfter: hasPageBreak(node), splittable: node.type === schema.nodes.table });
    offsets.push(offset);
  });
  const bands = bandHeights(view, setup);
  const textHeight = setup.height - bands[0] - bands[1];
  const { starts, pageNumbers, fills, pages } = layoutPages(blocks, textHeight);
  // A drawn page can hold several real pages (a table running on); its footer shows the last of them.
  const lastOf = (k: number) => (k + 1 < pageNumbers.length ? pageNumbers[k + 1] - 1 : pages);
  const part = (cls: string, content: Node | null) => {
    const el = createDiv({ cls, attr: { contenteditable: "false" } });
    if (content) el.appendChild(content);
    return el;
  };
  // Blank space that pads a page to its full height.
  const fill = (page: number) => {
    const el = createDiv({ cls: "docx-page-fill" });
    el.setCssProps({ "--docx-fill": `${fills[page]}px` });
    return el;
  };
  const decos = [
    Decoration.widget(0, () => part("docx-page-first", opts.header(1, pages)), { side: -1, key: `first-${gen}`, ignoreSelection: true }),
    ...starts.map((i, k) =>
      Decoration.widget(
        offsets[i],
        () => {
          const gap = createDiv({ cls: "docx-page-gap", attr: { contenteditable: "false" } });
          gap.appendChild(fill(k));
          gap.appendChild(part("docx-page-foot", opts.footer(lastOf(k), pages)));
          gap.createDiv({ cls: "docx-page-space" });
          gap.appendChild(part("docx-page-head", opts.header(pageNumbers[k + 1], pages)));
          return gap;
        },
        { side: -1, key: `gap-${gen}-${k}`, ignoreSelection: true }
      )
    ),
    Decoration.widget(
      doc.content.size,
      () => {
        const last = createDiv({ cls: "docx-page-last", attr: { contenteditable: "false" } });
        last.appendChild(fill(fills.length - 1));
        last.appendChild(part("docx-page-foot", opts.footer(pages, pages)));
        return last;
      },
      { side: 1, key: `last-${gen}`, ignoreSelection: true }
    ),
  ];
  return { starts, pages, decorations: DecorationSet.create(doc, decos), bands };
}

/** Lays the document out on pages: gaps with footers and headers where a page fills up (blocks never split). */
export function pagesPlugin(opts: PagesOptions): Plugin<PagesState> {
  return new Plugin<PagesState>({
    key: pagesKey,
    state: {
      init: () => ({ starts: [], pages: 1, decorations: DecorationSet.empty, bands: [0, 0] }),
      apply(tr, old) {
        const meta = tr.getMeta(pagesKey) as PagesState | undefined;
        return meta ?? { ...old, decorations: old.decorations.map(tr.mapping, tr.doc) };
      },
    },
    props: {
      decorations: (state) => pagesKey.getState(state)?.decorations,
    },
    view(view) {
      let frame = 0;
      const measure = () => {
        frame = 0;
        if (!view.dom.isConnected) return;
        const next = pageDecorations(view, opts);
        view.dispatch(view.state.tr.setMeta(pagesKey, next).setMeta("addToHistory", false));
        // The header / footer just drawn may be taller than the margins assumed: lay out again once with their real size.
        const setup = opts.setup();
        if (setup && bandHeights(view, setup).some((h, i) => h !== next.bands[i])) schedule();
      };
      // A timer rather than requestAnimationFrame: frames stop while the window is in the background.
      const schedule = () => {
        if (!frame) frame = window.setTimeout(measure, 60);
      };
      schedulers.set(view, schedule);
      schedule();
      return {
        update: (v, prev) => {
          if (v.state.doc !== prev.doc) schedule();
        },
        destroy: () => window.clearTimeout(frame),
      };
    },
  });
}
