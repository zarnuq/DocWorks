/*
 * Find and replace over the slides' text (shapes, groups and table cells).
 * Matches are paragraph offsets; a:br counts as one "\n" character.
 */
import { NS } from "../../core/utils/xml";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { textMatches, TextMatch } from "../../core/utils/text-matches";
import { slideComments } from "./comments";

export interface FindOptions {
  matchCase: boolean;
  notes?: boolean;
  comments?: boolean;
}

export type SlideSearchMatch = ({ kind: "text" } & Match) | (TextMatch & {
  kind: "notes" | "comment";
  slide: number;
  text: string;
  author?: string;
  messageId?: string;
  resolved?: boolean;
});

/** Visible slide text plus opt-in speaker notes and all comment messages, in slide order. */
export function findSlideMatches(doc: PptxDocument, query: string, opts: FindOptions): SlideSearchMatch[] {
  if (!query) return [];
  const text = findAll(doc, query, opts);
  const bySlide = new Map<number, Match[]>();
  for (const match of text) {
    const matches = bySlide.get(match.slide) ?? [];
    matches.push(match);
    bySlide.set(match.slide, matches);
  }
  const out: SlideSearchMatch[] = [];
  doc.slides().forEach((ref, slide) => {
    out.push(...(bySlide.get(slide) ?? []).map(m => ({ ...m, kind: "text" as const })));
    if (opts.notes) {
      const notes = doc.notesText(slide);
      out.push(...textMatches(notes, query, opts.matchCase).map(m => ({ ...m, kind: "notes" as const, slide, text: notes })));
    }
    if (opts.comments) for (const thread of slideComments(doc, ref.path)) for (const message of thread.messages) {
      out.push(...textMatches(message.text, query, opts.matchCase).map(m => ({ ...m, kind: "comment" as const, slide, text: message.text, author: message.author, messageId: message.id, resolved: thread.resolved })));
    }
  });
  return out;
}

export interface Match {
  /** Slide index. */
  slide: number;
  /** The p:sp or a:tc holding the text. */
  owner: Element;
  /** The a:p, and the match's character range in it. */
  para: Element;
  start: number;
  end: number;
}

interface Piece {
  /** a:r, a:fld or a:br. */
  node: Element;
  /** Its a:t (none for a:br). */
  t: Element | null;
  text: string;
}

function pieces(p: Element): Piece[] {
  return Array.from(p.children).flatMap((node): Piece[] => {
    if (node.namespaceURI !== NS.a) return [];
    if (node.localName === "br") return [{ node, t: null, text: "\n" }];
    if (node.localName !== "r" && node.localName !== "fld") return [];
    const t = Array.from(node.children).find((c) => c.localName === "t") ?? null;
    return [{ node, t, text: t?.textContent ?? "" }];
  });
}

/** Text bodies of a slide in document order (shapes, group children, table cells). */
function textBodies(slide: Document): Element[] {
  const tree = slide.getElementsByTagNameNS(NS.p, "spTree")[0];
  if (!tree) return [];
  const out: Element[] = [];
  const walker = slide.createTreeWalker(tree, 1 /* NodeFilter.SHOW_ELEMENT */);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) if ((n as Element).localName === "txBody") out.push(n as Element);
  return out;
}

export function findAll(doc: PptxDocument, query: string, opts: FindOptions): Match[] {
  if (!query) return [];
  const norm = (s: string) => (opts.matchCase ? s : s.toLowerCase());
  const q = norm(query);
  const out: Match[] = [];
  doc.slides().forEach((ref, slide) => {
    const xml = doc.xml(ref.path);
    if (!xml) return;
    for (const body of textBodies(xml)) {
      const owner = body.parentElement!;
      for (const para of Array.from(body.children).filter((c) => c.localName === "p")) {
        const text = norm(pieces(para).map((pc) => pc.text).join(""));
        for (let i = text.indexOf(q); i >= 0; i = text.indexOf(q, i + q.length)) out.push({ slide, owner, para, start: i, end: i + q.length });
      }
    }
  });
  return out;
}

/** Replace one match; the replacement takes the formatting of the run the match starts in. */
export function replaceMatch(m: Match, replacement: string): void {
  let at = 0;
  let placed = false;
  for (const pc of pieces(m.para)) {
    const from = at;
    at += pc.text.length;
    if (at <= m.start || from >= m.end) continue;
    const before = pc.text.slice(0, Math.max(0, m.start - from));
    const after = pc.text.slice(Math.max(0, m.end - from));
    if (!placed && pc.t) {
      pc.t.textContent = before + replacement + after;
      placed = true;
    } else if (pc.t && before + after) {
      pc.t.textContent = before + after;
    } else {
      pc.node.remove();
    }
  }
  if (!placed && replacement) {
    // The match was only line breaks: put the text in a plain run.
    const r = m.para.ownerDocument.createElementNS(NS.a, "a:r");
    r.appendChild(m.para.ownerDocument.createElementNS(NS.a, "a:t")).textContent = replacement;
    m.para.insertBefore(r, Array.from(m.para.children).find((c) => c.localName === "endParaRPr") ?? null);
  }
}

/** Replace every match; returns how many were replaced. */
export function replaceAll(doc: PptxDocument, query: string, replacement: string, opts: FindOptions): number {
  const matches = findAll(doc, query, opts);
  // Back to front, so earlier offsets in a paragraph stay valid.
  for (const m of [...matches].reverse()) replaceMatch(m, replacement);
  return matches.length;
}
