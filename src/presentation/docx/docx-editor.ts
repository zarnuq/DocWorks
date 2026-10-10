/*
 * ProseMirror commands, keymap, input rules and decorations for the docx
 * editor. No Obsidian imports, so everything here runs under jsdom tests.
 */
import { Attrs, MarkType, Node as PMNode, ResolvedPos } from "prosemirror-model";
import { Command, EditorState, Plugin, TextSelection, Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { baseKeymap, chainCommands, toggleMark } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { InputRule, inputRules } from "prosemirror-inputrules";
import { attrsOf, docxSchema as schema, HEADING_LEVEL, Align, CellAttrs, ParagraphAttrs, TableAttrs, TableBorders } from "../../domain/docx/schema";
import { DocxStyles, lineHeight } from "../../domain/docx/docx-styles";
import { fontStack } from "../../core/utils/fonts";
import { letterSpacing } from "../pptx/text-fit";
import { pageAt } from "./docx-pages";
import { Columns, DEFAULT_SECTPR, ONE_COLUMN, regroup, sections, sectionType, withColumns, withType } from "../../domain/docx/sections";
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
    // After the block the cursor is in; inside a column section, after the block in that section.
    const depth = $from.depth > 1 && $from.node(1).type === schema.nodes.columns ? 2 : 1;
    const pos = $from.depth >= depth ? $from.after(depth) : $from.pos;
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
export function tocEntries(doc: PMNode): { level: number; text: string; pos: number }[] {
  const out: { level: number; text: string; pos: number }[] = [];
  doc.descendants((node, pos) => {
    const style = node.type === schema.nodes.paragraph ? attrsOf<ParagraphAttrs>(node).style : null;
    const level = style && style !== "Title" ? HEADING_LEVEL[style] : undefined;
    if (level && level <= 3 && node.textContent.trim()) out.push({ level, text: node.textContent.trim(), pos });
    return node.type !== schema.nodes.paragraph;
  });
  return out;
}

/**
 * A Word TOC field (\o "1-3") in a content control, with the current headings as its cached result; Word refreshes
 * it on open. With the page view on, each entry ends with its heading's page after a dotted right tab.
 */
function tocNode(state: EditorState): PMNode {
  const entries = tocEntries(state.doc);
  const rows = entries.length ? entries : [{ level: 1, text: "No headings yet", pos: -1 }];
  const run = (inner: string) => opaqueXml(`<w:r xmlns:w="${NS_W}">${inner}</w:r>`);
  const begin = [run(`<w:fldChar w:fldCharType="begin"/>`), run(`<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>`), run(`<w:fldChar w:fldCharType="separate"/>`)];
  const para = (style: string | null, content: PMNode[], tabLeader: string | null = null) => schema.nodes.paragraph.create({ style, tabLeader }, content);
  const page = (pos: number) => (pos < 0 ? null : pageAt(state, pos));
  const xml = `<w:sdt xmlns:w="${NS_W}"><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr></w:sdt>`;
  return schema.nodes.sdt.create({ kind: "toc", xml }, [
    para("TOCHeading", [schema.text("Contents")]),
    ...rows.map((e, i) => {
      const n = page(e.pos);
      return para(`TOC${e.level}`, [...(i === 0 ? begin : []), schema.text(n === null ? e.text : `${e.text}\t${n}`)], n === null ? null : "dot");
    }),
    para(null, [run(`<w:fldChar w:fldCharType="end"/>`)]),
  ]);
}

// -----------------------------------------------------------------------------
// Sections: columns and section breaks
// -----------------------------------------------------------------------------

/** Blocks with column sections opened up, and where each starts. */
function unitsOf(doc: PMNode): { node: PMNode; pos: number }[] {
  const out: { node: PMNode; pos: number }[] = [];
  doc.forEach((n, pos) => (n.type === schema.nodes.columns ? n.forEach((c, o) => out.push({ node: c, pos: pos + 1 + o })) : out.push({ node: n, pos })));
  return out;
}

const unitAt = (units: { pos: number }[], pos: number) => {
  let i = 0;
  while (i + 1 < units.length && units[i + 1].pos <= pos) i++;
  return i;
};

const withSectPr = (node: PMNode, sectPr: string | null) => node.type.create({ ...node.attrs, sectPr }, node.content, node.marks);

/** Replace the document's blocks with `units` regrouped into column sections, keeping the cursor in unit `at.unit`. */
function rebuild(state: EditorState, units: PMNode[], finalSectPr: string | null, at: { unit: number; offset: number }): Transaction {
  const tr = state.tr.replaceWith(0, state.doc.content.size, regroup(units, finalSectPr));
  if (finalSectPr !== state.doc.attrs.sectPr) tr.setDocAttribute("sectPr", finalSectPr);
  const placed = unitsOf(tr.doc);
  const unit = placed[Math.min(at.unit, placed.length - 1)];
  tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(unit.pos + at.offset, unit.pos + unit.node.nodeSize - 1))));
  return tr.scrollIntoView();
}

/** Columns of the section the cursor is in. */
export function columnsAt(state: EditorState): Columns {
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type === schema.nodes.columns) return { count: n.attrs.count as number, space: n.attrs.space as number, sep: !!n.attrs.sep };
  }
  return ONE_COLUMN;
}

/**
 * Format › Columns, as in Google Docs: with nothing selected the cursor's whole section gets `cols`; a selection
 * inside a section gets them alone, between continuous section breaks.
 */
function setColumns(cols: Columns): Command {
  return (state, dispatch) => {
    const placed = unitsOf(state.doc);
    const nodes = placed.map((u) => u.node);
    const { from, to, empty } = state.selection;
    const a = unitAt(placed, from);
    const b = unitAt(placed, to);
    const s = sections(nodes).find((x) => a >= x.start && a <= x.end)!;
    if (!dispatch) return true;
    const governing = (s.at === null ? (state.doc.attrs.sectPr as string | null) : (nodes[s.at].attrs.sectPr as string)) ?? DEFAULT_SECTPR;
    let start = a;
    let end = b;
    if (empty || b > s.end || (a <= s.start && b >= s.end)) [start, end] = [s.start, s.end];
    // Section breaks sit on paragraphs: a range next to a table takes in the rest of the section.
    if (start > s.start && nodes[start - 1].type !== schema.nodes.paragraph) start = s.start;
    if (end < s.end && nodes[end].type !== schema.nodes.paragraph) end = s.end;
    const units = nodes.slice();
    let final = state.doc.attrs.sectPr as string | null;
    const set = (i: number | null, xml: string) => (i === null ? (final = xml) : (units[i] = withSectPr(units[i], xml)));
    if (start > s.start) set(start - 1, governing);
    if (end < s.end) {
      set(end, withType(withColumns(governing, cols), start > s.start ? "continuous" : sectionType(governing)));
      set(s.at, withType(governing, "continuous"));
    } else {
      set(s.at, start > s.start ? withType(withColumns(governing, cols), "continuous") : withColumns(governing, cols));
    }
    dispatch(rebuild(state, units, final, { unit: a, offset: from - placed[a].pos }));
    return true;
  };
}

/** Insert › Break › Section break: the section ends at the cursor (splitting the paragraph), the next one starts as `type`. */
function insertSectionBreak(type: "nextPage" | "continuous"): Command {
  return (state, dispatch) => {
    const $from = state.selection.$from;
    const placed = unitsOf(state.doc);
    const k = unitAt(placed, $from.pos);
    const para = placed[k].node;
    if (para.type !== schema.nodes.paragraph || $from.parent !== para) return false;
    if (!dispatch) return true;
    let units = placed.map((u) => u.node);
    const off = $from.parentOffset;
    let end: number;
    let cursor: { unit: number; offset: number };
    if (off === para.content.size && k < units.length - 1) {
      end = k;
      cursor = { unit: k + 1, offset: 1 };
    } else if (off === 0 && k > 0 && units[k - 1].type === schema.nodes.paragraph && !units[k - 1].attrs.sectPr) {
      end = k - 1;
      cursor = { unit: k, offset: 1 };
    } else {
      const first = para.type.create({ ...para.attrs, sectPr: null }, para.content.cut(0, off), para.marks);
      const second = para.type.create(para.attrs, para.content.cut(off), para.marks);
      units = [...units.slice(0, k), first, second, ...units.slice(k + 1)];
      end = k;
      cursor = { unit: k + 1, offset: 1 };
    }
    const secs = sections(units);
    let final = state.doc.attrs.sectPr as string | null;
    const governingOf = (i: number | null) => (i === null ? final : (units[i].attrs.sectPr as string)) ?? DEFAULT_SECTPR;
    const set = (i: number | null, xml: string) => (i === null ? (final = xml) : (units[i] = withSectPr(units[i], xml)));
    if (units[end].attrs.sectPr) {
      // Already the end of a section: the next section starts as asked.
      const next = secs[secs.findIndex((x) => x.at === end) + 1];
      if (next) set(next.at, withType(governingOf(next.at), type));
    } else {
      const s = secs.find((x) => end >= x.start && end <= x.end)!;
      const governing = governingOf(s.at);
      set(end, governing);
      set(s.at, withType(governing, type));
    }
    dispatch(rebuild(state, units, final, cursor));
    return true;
  };
}

/** Split copies of a paragraph that ends a section both carry its w:sectPr; only the last one ends the section. */
function sectionDedupe(): Plugin {
  return new Plugin({
    appendTransaction(trs, _old, state) {
      if (!trs.some((t) => t.docChanged)) return null;
      const units = unitsOf(state.doc);
      let tr: Transaction | null = null;
      for (let i = 0; i + 1 < units.length; i++) {
        const [x, y] = [units[i].node, units[i + 1].node];
        if (x.attrs.sectPr && x.attrs.sectPr === y.attrs.sectPr && x.attrs.key === y.attrs.key) {
          tr ??= state.tr;
          tr.setNodeMarkup(units[i].pos, undefined, { ...x.attrs, sectPr: null });
        }
      }
      return tr;
    },
  });
}

/** Address of the link at the cursor or selection start, or null. */
export function linkHref(state: EditorState): string | null {
  const { $from } = state.selection;
  const link = [...$from.marks(), ...($from.nodeAfter?.marks ?? [])].find((x) => x.type === schema.marks.link);
  return (link?.attrs.href as string | null | undefined) ?? null;
}

export const cmd = {
  setColumns,
  columnsAt,
  insertSectionBreak,
  insertColumnBreak: ((state, dispatch) => {
    if (dispatch) dispatch(state.tr.replaceSelectionWith(opaqueXml(`<w:r xmlns:w="${NS_W}"><w:br w:type="column"/></w:r>`)).scrollIntoView());
    return true;
  }) as Command,

  insertToc: ((state, dispatch) => insertBlockAfter(tocNode(state))(state, dispatch)) as Command,

  /** Rebuild every table of contents from the current headings. */
  updateToc: ((state, dispatch) => {
    const at: { pos: number; size: number }[] = [];
    state.doc.forEach((node, pos) => {
      if (node.type === schema.nodes.sdt && node.attrs.kind === "toc") at.push({ pos, size: node.nodeSize });
    });
    if (!at.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { pos, size } of at.reverse()) tr.replaceWith(pos, pos + size, tocNode(state));
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

  /** Whether rows of the table at the cursor may break across pages (none is marked "can't split"). */
  rowsCanBreak(state: EditorState): boolean {
    const info = cellAround(state.selection.$from);
    let can = true;
    info?.table.forEach((row) => (can &&= !row.attrs.cantSplit));
    return can;
  },

  /** Let every row of the table at the cursor break across pages, or move whole (Google Docs' "Allow row to overflow across pages"). */
  setRowsCanBreak(allow: boolean): Command {
    return (state, dispatch) => {
      const info = cellAround(state.selection.$from);
      if (!info) return false;
      if (dispatch) {
        const tr = state.tr;
        let pos = info.tablePos + 1;
        info.table.forEach((row) => {
          tr.setNodeMarkup(pos, undefined, { ...row.attrs, cantSplit: !allow });
          pos += row.nodeSize;
        });
        dispatch(tr);
      }
      return true;
    };
  },

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
  // Word sizes lines by the text's font, not the style's: 10pt text in a 12pt style gets 10pt lines.
  const text = textFont(node, a);
  if (text.size) css.push(`font-size: ${text.size}pt`);
  if (text.font) css.push(`font-family: ${fontStack(text.font)}`);
  const lines = a.lineSpacing ? Number(a.lineSpacing) : text.size || text.font ? styles.lines(a.style, tableStyle) : null;
  if (lines) css.push(`line-height: ${lineHeight(lines, text.font ?? styles.fontFor(a.style, tableStyle))}`);
  if (a.spaceBefore !== null) css.push(`margin-top: ${a.spaceBefore}pt`);
  if (a.spaceAfter !== null) css.push(`margin-bottom: ${a.spaceAfter}pt`);
  if (a.firstLine) css.push(`text-indent: ${a.firstLine}pt`);
  return css.filter(Boolean).join("; ");
}

/**
 * Size and font Word sizes a paragraph's lines by: the text's when every run sets a size (an empty paragraph: its
 * mark's). Otherwise the style's, since text without a size of its own takes the style's, never the mark's.
 */
function textFont(node: PMNode, a: ParagraphAttrs): { size: number | null; font: string | null } {
  let size = 0;
  let font: string | null = null;
  let all = true;
  let any = false;
  node.forEach((child) => {
    if (!child.isText) return;
    any = true;
    const pt = child.marks.find((m) => m.type === schema.marks.size)?.attrs.pt as number | undefined;
    if (!pt) all = false;
    else if (pt > size) {
      size = pt;
      font = (child.marks.find((m) => m.type === schema.marks.font)?.attrs.font as string | undefined) ?? null;
    }
  });
  if (!any) return { size: a.markSize, font: a.markFont };
  return all ? { size, font: font ?? a.markFont } : { size: null, font: null };
}

const fitCache = new Map<string, number | null>();

/**
 * Letter-spacing that makes each run as wide as in its Word font (from font-metrics), so lines wrap where Word and
 * Google Docs wrap them even when the font isn't installed. With the real font installed it comes out ~0.
 */
function fitText(styles: DocxStyles, doc: PMNode, node: PMNode, pos: number, decos: Decoration[]): void {
  const a = attrsOf<ParagraphAttrs>(node);
  let table: string | null = null;
  const $p = doc.resolve(pos);
  for (let d = $p.depth; d > 0 && !table; d--) if ($p.node(d).type === schema.nodes.table) table = attrsOf<TableAttrs>($p.node(d)).style;
  const base = { font: styles.fontFor(a.style, table), size: styles.sizeFor(a.style, table) ?? 10, ...styles.runFlags(a.style, table) };
  node.forEach((child, offset) => {
    // ponytail: runs with tabs are left alone (a tab's width isn't a character's); they're short in practice.
    if (!child.isText || child.text!.includes("\t")) return;
    const mark = (type: MarkType) => child.marks.find((m) => m.type === type);
    const font = (mark(schema.marks.font)?.attrs.font as string | undefined) ?? base.font;
    if (!font) return;
    const pt = (mark(schema.marks.size)?.attrs.pt as number | undefined) ?? base.size;
    const b = !!mark(schema.marks.bold) || base.b;
    const i = !!mark(schema.marks.italic) || base.i;
    const text = base.caps ? child.text!.toUpperCase() : child.text!;
    const key = `${font}|${b}|${i}|${pt}|${text}`;
    if (!fitCache.has(key)) {
      if (fitCache.size > 20000) fitCache.clear();
      fitCache.set(key, letterSpacing(text, { font, b, i }, (pt * 96) / 72));
    }
    const ls = fitCache.get(key);
    if (ls !== null && ls !== undefined && Math.abs(ls) > 0.01) {
      decos.push(Decoration.inline(pos + 1 + offset, pos + 1 + offset + child.nodeSize, { style: `letter-spacing: ${Math.round(ls * 1000) / 1000}px` }));
    }
  });
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
      // The list level's indent replaces the style's.
      css.push("margin-left: 0", `padding-left: ${mark.indentPt}pt`, `--docx-hanging: ${mark.hangingPt}pt`);
      attrs.class = "docx-list";
      attrs["data-list"] = mark.text;
    } else if (node.attrs.indent) {
      css.push(`margin-left: ${node.attrs.indent * 36}pt`);
    }
    // The last tab runs to a right-aligned tab stop (table of contents lines): the line becomes a flex row
    // where that tab stretches, drawing its leader.
    // ponytail: in that row a text node of only spaces between two styled spans collapses; real TOC lines don't have one.
    const leader = attrsOf<ParagraphAttrs>(node).tabLeader;
    const tabs: number[] = [];
    if (leader)
      node.forEach((child, offset) => {
        for (let i = child.isText ? child.text!.indexOf("\t") : -1; i >= 0; i = child.text!.indexOf("\t", i + 1)) tabs.push(pos + 1 + offset + i);
      });
    if (tabs.length) {
      attrs.class = [attrs.class, "docx-tab-right"].filter(Boolean).join(" ");
      tabs.forEach((at, i) => decos.push(Decoration.inline(at, at + 1, { class: i === tabs.length - 1 ? `docx-tab-fill docx-leader-${leader}` : "docx-tab" })));
    }
    attrs.style = css.filter(Boolean).join("; ");
    decos.push(Decoration.node(pos, pos + node.nodeSize, attrs));
    fitText(styles, doc, node, pos, decos);
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
    sectionDedupe(),
  ];
}
