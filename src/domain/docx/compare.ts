/*
 * Google Docs' Tools › Compare documents: the base document's body rewritten with tracked changes (w:ins / w:del)
 * that turn it into the other document, attributed to one author. Blocks are aligned by their text (paragraphs
 * and tables), a changed paragraph that still shares enough with its counterpart is diffed by words, and the rest
 * is deleted / inserted whole. Text kept is the base's (with its formatting); inserted text is the other file's.
 * ponytail: formatting-only changes aren't tracked, and inserted content loses what points into the other file
 * (pictures, links, notes, list numbering); `dropped` counts it.
 */
import { NS, insertInOrder } from "../../core/utils/xml";
import { diff, diffWords } from "../../core/utils/diff";
import { ORDER } from "./docx-writer";

export interface CompareOptions {
  author: string;
  /** ISO date for every revision. */
  date: string;
}
export interface CompareResult {
  /** Revisions written (insertions and deletions, paragraph marks included). */
  changes: number;
  /** Pieces of inserted content left out because they point into the other file. */
  dropped: number;
}

/** Google's name for the result: "Comparison of Draft and Final.docx". */
export const comparisonName = (base: string, other: string) => `Comparison of ${base.replace(/\.[^.]+$/, "")} and ${other.replace(/\.[^.]+$/, "")}.docx`;

const W = NS.w;
const isW = (el: Element, ...locals: string[]) => el.namespaceURI === W && locals.includes(el.localName);
const kidsOf = (el: Element) => Array.from(el.children);
/** Inside a tracked deletion. */
function inDel(el: Element): boolean {
  for (let e = el.parentElement; e; e = e.parentElement) if (isW(e, "del", "moveFrom")) return true;
  return false;
}

/** What a run says: text, tabs as \t, breaks as \n (deleted text left out). */
function runText(r: Element): string {
  return kidsOf(r).map((c) => (isW(c, "t") ? c.textContent ?? "" : isW(c, "tab") ? "\t" : isW(c, "br", "cr") ? "\n" : "")).join("");
}
/** A block's text, for aligning blocks. */
function blockText(el: Element): string {
  if (isW(el, "p")) return Array.from(el.getElementsByTagNameNS(W, "r")).filter((r) => !inDel(r)).map(runText).join("");
  if (isW(el, "tbl")) return Array.from(el.getElementsByTagNameNS(W, "p")).map(blockText).join("\u0001");
  return el.textContent ?? "";
}

/** A paragraph of plain runs only (text, tabs, breaks): one we can diff by words. */
function plain(p: Element): boolean {
  return isW(p, "p") && kidsOf(p).every((c) => isW(c, "pPr", "bookmarkStart", "bookmarkEnd", "proofErr") || (isW(c, "r") && kidsOf(c).every((k) => isW(k, "rPr", "t", "tab", "br", "lastRenderedPageBreak"))));
}

interface Segment {
  text: string;
  rPr: Element | null;
}
const segments = (p: Element): Segment[] => kidsOf(p).filter((c) => isW(c, "r")).map((r) => ({ text: runText(r), rPr: kidsOf(r).find((k) => isW(k, "rPr")) ?? null }));

class Comparer {
  changes = 0;
  dropped = 0;
  private id: number;

  constructor(private doc: Document, private opts: CompareOptions, base: Element) {
    // Revision ids above every id the document already uses.
    this.id = Array.from(base.ownerDocument.getElementsByTagNameNS(W, "*")).reduce((m, e) => Math.max(m, Number(e.getAttributeNS(W, "id")) || 0), 0) + 1;
  }

  private el(local: string): Element {
    return this.doc.createElementNS(W, `w:${local}`);
  }

  /** An empty w:ins / w:del with id, author and date. */
  mark(kind: "ins" | "del"): Element {
    const m = this.el(kind);
    m.setAttributeNS(W, "w:id", String(this.id++));
    m.setAttributeNS(W, "w:author", this.opts.author);
    m.setAttributeNS(W, "w:date", this.opts.date);
    this.changes++;
    return m;
  }

  /** Runs saying `text` formatted as `rPr` (deleted text as w:delText). */
  private runs(text: string, rPr: Element | null, deleted: boolean): Element[] {
    return text.split(/(\t|\n)/).filter(Boolean).map((piece) => {
      const r = this.el("r");
      if (rPr) r.appendChild(this.doc.importNode(rPr, true));
      if (piece === "\t") r.appendChild(this.el("tab"));
      else if (piece === "\n") r.appendChild(this.el("br"));
      else {
        const t = r.appendChild(this.el(deleted ? "delText" : "t"));
        t.setAttributeNS(NS.xml, "xml:space", "preserve");
        t.textContent = piece;
      }
      return r;
    });
  }

  /** Runs for [from, to) of `segs`, each piece with its own formatting. */
  private slice(segs: Segment[], from: number, to: number, deleted: boolean): Element[] {
    const out: Element[] = [];
    let at = 0;
    for (const s of segs) {
      const a = Math.max(from, at), b = Math.min(to, at + s.text.length);
      if (a < b) out.push(...this.runs(s.text.slice(a - at, b - at), s.rPr, deleted));
      at += s.text.length;
    }
    return out;
  }

  /** `base` rewritten word by word into `other`'s text. */
  diffParagraph(base: Element, other: Element): void {
    const a = segments(base), b = segments(other);
    const pPr = kidsOf(base).find((c) => isW(c, "pPr"));
    const out: Element[] = [];
    let ai = 0, bi = 0;
    for (const op of diffWords(a.map((s) => s.text).join(""), b.map((s) => s.text).join(""))) {
      const n = op.text.length;
      if (op.kind === "same") {
        out.push(...this.slice(a, ai, ai + n, false));
        ai += n;
        bi += n;
      } else if (op.kind === "removed") {
        const del = this.mark("del");
        del.append(...this.slice(a, ai, ai + n, true));
        out.push(del);
        ai += n;
      } else {
        const ins = this.mark("ins");
        ins.append(...this.slice(b, bi, bi + n, false));
        out.push(ins);
        bi += n;
      }
    }
    base.replaceChildren(...(pPr ? [pPr] : []), ...out);
  }

  /** Mark a paragraph's mark inserted / deleted (w:pPr/w:rPr/w:ins|del), in schema order. */
  private markParagraph(p: Element, kind: "ins" | "del"): void {
    let pPr = kidsOf(p).find((c) => isW(c, "pPr"));
    if (!pPr) pPr = p.insertBefore(this.el("pPr"), p.firstChild);
    // A section break's paragraph keeps its mark: joining it would move the break.
    if (kidsOf(pPr).some((c) => isW(c, "sectPr"))) return;
    const rPr = kidsOf(pPr).find((c) => isW(c, "rPr")) ?? insertInOrder(pPr, this.el("rPr"), ORDER.pPr);
    rPr.insertBefore(this.mark(kind), rPr.firstChild);
  }

  /** Every run of `p` deleted (text as w:delText), and its mark too unless it is in a table cell. */
  deleteParagraph(p: Element, inCell = false): void {
    for (const r of Array.from(p.getElementsByTagNameNS(W, "r"))) {
      if (inDel(r)) continue;
      for (const t of Array.from(r.getElementsByTagNameNS(W, "t"))) {
        const d = this.el("delText");
        for (const at of Array.from(t.attributes)) d.setAttributeNS(at.namespaceURI, at.name, at.value);
        d.textContent = t.textContent;
        t.replaceWith(d);
      }
      for (const t of Array.from(r.getElementsByTagNameNS(W, "instrText"))) {
        const d = this.el("delInstrText");
        d.textContent = t.textContent;
        t.replaceWith(d);
      }
      const del = this.mark("del");
      r.replaceWith(del);
      del.appendChild(r);
    }
    if (!inCell) this.markParagraph(p, "del");
  }

  /** A paragraph from the other file, cleaned of what points into it, its runs and mark inserted. */
  insertParagraph(p: Element, inCell = false): Element {
    const copy = this.doc.importNode(p, true);
    this.clean(copy);
    for (const r of Array.from(copy.getElementsByTagNameNS(W, "r"))) {
      const ins = this.mark("ins");
      r.replaceWith(ins);
      ins.appendChild(r);
    }
    if (!inCell) this.markParagraph(copy, "ins");
    return copy;
  }

  /** Drop what refers to the other file's parts or ids; keep the text. */
  private clean(el: Element): void {
    for (const h of Array.from(el.getElementsByTagNameNS(W, "hyperlink"))) h.replaceWith(...Array.from(h.childNodes));
    const drop = (e: Element) => {
      e.remove();
      this.dropped++;
    };
    for (const local of ["numPr", "sectPr", "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd"]) {
      for (const e of Array.from(el.getElementsByTagNameNS(W, local))) drop(e);
    }
    for (const r of Array.from(el.getElementsByTagNameNS(W, "r"))) {
      const refers = Array.from(r.getElementsByTagNameNS("*", "*")).some((c) =>
        ["drawing", "pict", "object", "footnoteReference", "endnoteReference", "commentReference"].includes(c.localName) || c.hasAttributeNS(NS.r, "id") || c.hasAttributeNS(NS.r, "embed"));
      if (refers) drop(r);
    }
  }

  /** Mark every row of a table deleted, its runs with it. */
  deleteTable(tbl: Element): void {
    for (const tr of kidsOf(tbl).filter((c) => isW(c, "tr"))) {
      const trPr = kidsOf(tr).find((c) => isW(c, "trPr")) ?? tr.insertBefore(this.el("trPr"), kidsOf(tr).find((c) => isW(c, "tblPrEx"))?.nextSibling ?? tr.firstChild);
      insertInOrder(trPr, this.mark("del"), ORDER.trPr);
      for (const p of Array.from(tr.getElementsByTagNameNS(W, "p"))) this.deleteParagraph(p, true);
    }
  }

  /** A block deleted: paragraphs and tables, and those inside content controls (the control stays). */
  deleteBlock(el: Element): void {
    if (isW(el, "p")) this.deleteParagraph(el);
    else if (isW(el, "tbl")) this.deleteTable(el);
    else if (isW(el, "sdt")) for (const inner of kidsOf(kidsOf(el).find((c) => isW(c, "sdtContent")) ?? el)) this.deleteBlock(inner);
  }

  /** A block of the other file inserted; a content control brings its contents, not itself. */
  insertBlock(el: Element): Element[] {
    if (isW(el, "p")) return [this.insertParagraph(el)];
    if (isW(el, "tbl")) return [this.insertTable(el)];
    if (isW(el, "sdt")) return kidsOf(kidsOf(el).find((c) => isW(c, "sdtContent")) ?? el).flatMap((inner) => this.insertBlock(inner));
    this.dropped++;
    return [];
  }

  insertTable(tbl: Element): Element {
    const copy = this.doc.importNode(tbl, true);
    this.clean(copy);
    // Its paragraphs are already this document's: mark them in place.
    for (const tr of kidsOf(copy).filter((c) => isW(c, "tr"))) {
      const trPr = kidsOf(tr).find((c) => isW(c, "trPr")) ?? tr.insertBefore(this.el("trPr"), tr.firstChild);
      insertInOrder(trPr, this.mark("ins"), ORDER.trPr);
      for (const r of Array.from(tr.getElementsByTagNameNS(W, "r"))) {
        const ins = this.mark("ins");
        r.replaceWith(ins);
        ins.appendChild(r);
      }
    }
    return copy;
  }
}

/** How alike two paragraph texts are: the share of characters a word diff keeps. */
function likeness(a: string, b: string): number {
  if (!a && !b) return 1;
  const same = diffWords(a, b).filter((c) => c.kind === "same").reduce((n, c) => n + c.text.length, 0);
  return same / Math.max(a.length, b.length);
}

/** Rewrite `base` (a w:body) with tracked changes that turn it into `other`. */
export function compareBodies(base: Element, other: Element, opts: CompareOptions): CompareResult {
  const c = new Comparer(base.ownerDocument, opts, base);
  const blocks = (b: Element) => kidsOf(b).filter((e) => !isW(e, "sectPr"));
  const a = blocks(base), b = blocks(other);
  const ops = diff(a, b, (x, y) => x.localName === y.localName && blockText(x) === blockText(y));
  const finalSect = kidsOf(base).find((e) => isW(e, "sectPr")) ?? null;
  // New blocks go before the next kept / deleted base block, else before the final section break.
  const out: Element[] = [];
  let i = 0;
  while (i < ops.length) {
    const op = ops[i];
    if (op.kind === "same") {
      out.push(op.a);
      i++;
      continue;
    }
    // A run of removed blocks then added ones: pair alike plain paragraphs, the rest deleted / inserted whole.
    const removed: Element[] = [], added: Element[] = [];
    while (i < ops.length && ops[i].kind !== "same") {
      const o = ops[i++];
      if (o.kind === "removed") removed.push(o.a);
      else if (o.kind === "added") added.push(o.b);
    }
    let r = 0, d = 0;
    while (r < removed.length || d < added.length) {
      const x = removed[r], y = added[d];
      // As many paragraphs changed as came in: each was edited in place. Otherwise only alike ones pair.
      if (x && y && plain(x) && plain(y) && (removed.length === added.length || likeness(blockText(x), blockText(y)) >= 0.4)) {
        c.diffParagraph(x, y);
        out.push(x);
        r++;
        d++;
      } else if (x) {
        c.deleteBlock(x);
        out.push(x);
        r++;
      } else {
        out.push(...c.insertBlock(y));
        d++;
      }
    }
  }
  base.replaceChildren(...out, ...(finalSect ? [finalSect] : []));
  return { changes: c.changes, dropped: c.dropped };
}
