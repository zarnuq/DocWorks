/*
 * Find and replace for the docx editor: matches within runs of text (images and
 * other atoms break a match), highlighted with decorations.
 */
import { Node as PMNode } from "prosemirror-model";
import { Command, Plugin, PluginKey, Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { DocxDocument, DocxComment } from "../../infrastructure/docx/docx-document";
import { expandReplacement, SearchOptions, textMatches, TextMatch } from "../../core/utils/text-matches";

export type CommentMatch = TextMatch & { comment: DocxComment };
export function findCommentMatches(doc: DocxDocument, query: string, opts: SearchOptions | boolean): CommentMatch[] {
  if (!query) return [];
  return doc.comments().flatMap(comment => textMatches(comment.text, query, opts).map(match => ({ ...match, comment })));
}

export interface Range {
  from: number;
  to: number;
  /** A regular expression search's groups, for $1 in the replacement. */
  groups?: string[];
}

export function findRanges(doc: PMNode, query: string, opts: SearchOptions | boolean): Range[] {
  if (!query) return [];
  const out: Range[] = [];
  doc.descendants((block, pos) => {
    if (!block.isTextblock) return true;
    // Contiguous text pieces of the block, split at atoms.
    let start = -1;
    let text = "";
    const flush = () => {
      for (const m of textMatches(text, query, opts)) out.push({ from: start + m.start, to: start + m.end, ...(m.groups && { groups: m.groups }) });
      text = "";
      start = -1;
    };
    block.forEach((child, offset) => {
      if (child.isText) {
        if (start < 0) start = pos + 1 + offset;
        text += child.text;
      } else flush();
    });
    flush();
    return false;
  });
  return out;
}

/** Replace a range; the new text takes the marks at its start. */
export function replaceRangeCmd(range: Range, replacement: string): Command {
  replacement = expandReplacement(replacement, { start: 0, end: 0, groups: range.groups });
  return (state, dispatch) => {
    if (dispatch) {
      const marks = state.doc.resolve(range.from + 1).marks();
      const tr = state.tr;
      if (replacement) tr.replaceWith(range.from, range.to, state.schema.text(replacement, marks));
      else tr.delete(range.from, range.to);
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function replaceAllCmd(query: string, replacement: string, opts: SearchOptions | boolean): Command {
  return (state, dispatch) => {
    const ranges = findRanges(state.doc, query, opts);
    if (!ranges.length) return false;
    if (dispatch) {
      const tr = state.tr;
      // Back to front so earlier positions stay valid.
      for (const r of [...ranges].reverse()) {
        const marks = state.doc.resolve(r.from + 1).marks();
        const text = expandReplacement(replacement, { start: 0, end: 0, groups: r.groups });
        if (text) tr.replaceWith(r.from, r.to, state.schema.text(text, marks));
        else tr.delete(r.from, r.to);
      }
      dispatch(tr);
    }
    return true;
  };
}

interface FindState {
  query: string;
  opts: SearchOptions;
  matches: Range[];
  current: number;
  decorations: DecorationSet;
}

export const findKey = new PluginKey<FindState>("docx-find");

/** Set the search; `current` is the index of the match to emphasize (-1 for none). */
export function setFind(tr: Transaction, query: string, opts: SearchOptions | boolean, current: number): Transaction {
  return tr.setMeta(findKey, { query, opts: typeof opts === "boolean" ? { matchCase: opts } : opts, current });
}

function build(doc: PMNode, query: string, opts: SearchOptions, current: number): FindState {
  const matches = findRanges(doc, query, opts);
  const cur = matches.length && current >= 0 ? Math.min(current, matches.length - 1) : -1;
  const decos = matches.map((r, i) =>
    Decoration.inline(r.from, r.to, { class: i === cur ? "docx-find-hit is-current" : "docx-find-hit" }, { current: i === cur })
  );
  return { query, opts, matches, current: cur, decorations: DecorationSet.create(doc, decos) };
}

export function findPlugin(): Plugin<FindState> {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: (_, state) => build(state.doc, "", {}, -1),
      apply(tr, old) {
        const meta = tr.getMeta(findKey) as { query: string; opts: SearchOptions; current: number } | undefined;
        if (meta) return build(tr.doc, meta.query, meta.opts, meta.current);
        return tr.docChanged && old.query ? build(tr.doc, old.query, old.opts, old.current) : old;
      },
    },
    props: {
      decorations(state) {
        return findKey.getState(state)?.decorations;
      },
    },
  });
}

/** Words, characters (with spaces, without paragraph breaks) and characters excluding spaces, like Google Docs' count. */
export function countWords(doc: PMNode, from = 0, to = doc.content.size): { words: number; chars: number; letters: number } {
  const text = doc.textBetween(from, to, "\n", "\n");
  return { words: text.split(/\s+/).filter(Boolean).length, chars: text.replace(/\n/g, "").length, letters: text.replace(/\s/g, "").length };
}
