import { Fragment, Mark, Node as PMNode, Slice } from "prosemirror-model";
import { TextSelection, Transaction } from "prosemirror-state";
import { isHistoryTransaction } from "prosemirror-history";
import { NS } from "../../core/utils/xml";
import { Origins, ReadContext, readDocx, w, wVal } from "./docx-reader";
import { WriteContext, writeBody } from "./docx-writer";
import { attrsOf, XmlAttrs } from "./schema";

/** XML for comparing, without namespace declarations: Chromium writes those a node carries even where redundant, so the same run read from a file and rewritten from the editor would differ. */
const xml = (el: Element) => new XMLSerializer().serializeToString(el).replace(/\sxmlns(?::[\w.-]+)?="[^"]*"/g, "");
const copy = (el: Element) => el.cloneNode(true) as Element;
const element = (doc: Document, name: string) => doc.createElementNS(NS.w, `w:${name}`);

/** Align only the changed span. Unchanged prefixes/suffixes also bound work on long documents. */
function align<T>(a: T[], b: T[], same: (a: T, b: T) => boolean, compatible: (a: T, b: T) => boolean): [T | null, T | null][] {
  // The common end first: a letter typed (or deleted) before the same letter then counts as the edit, where the cursor is.
  let start = 0, end = 0;
  while (end < Math.min(a.length, b.length) && same(a[a.length - end - 1], b[b.length - end - 1])) end++;
  while (start < Math.min(a.length, b.length) - end && same(a[start], b[start])) start++;
  const left = a.slice(start, a.length - end), right = b.slice(start, b.length - end);
  const out: [T | null, T | null][] = a.slice(0, start).map((v, i) => [v, b[i]]);
  // Avoid quadratic allocations on unusually large pasted ranges; preserve their exact before/after versions.
  if (left.length * right.length > 1_000_000) {
    out.push(...left.map(v => [v, null] as [T, null]), ...right.map(v => [null, v] as [null, T]));
  } else {
    const width = right.length + 1;
    const costs = new Uint32Array((left.length + 1) * width);
    for (let i = left.length; i >= 0; i--) for (let j = right.length; j >= 0; j--) {
      const at = i * width + j;
      if (i === left.length) costs[at] = right.length - j;
      else if (j === right.length) costs[at] = left.length - i;
      else costs[at] = Math.min(1 + costs[at + width], 1 + costs[at + 1],
        (same(left[i], right[j]) ? 0 : compatible(left[i], right[j]) ? 1 : 3) + costs[at + width + 1]);
    }
    let i = 0, j = 0;
    while (i < left.length || j < right.length) {
      const at = i * width + j;
      // An equal pair that could as well come after an insertion or deletion (a letter typed or deleted before the
      // same letter) goes after it: the edit sits where the cursor was, next to what was typed before.
      const equal = i < left.length && j < right.length && same(left[i], right[j]);
      if (equal && j < right.length && costs[at] === 1 + costs[at + 1]) out.push([null, right[j++]]);
      else if (equal && costs[at] === 1 + costs[at + width]) out.push([left[i++], null]);
      else if (i < left.length && j < right.length && costs[at] === (same(left[i], right[j]) ? 0 : compatible(left[i], right[j]) ? 1 : 3) + costs[at + width + 1]) out.push([left[i++], right[j++]]);
      else if (i < left.length && costs[at] === 1 + costs[at + width]) out.push([left[i++], null]);
      else out.push([null, right[j++]]);
    }
  }
  out.push(...a.slice(a.length - end).map((v, i) => [v, b[b.length - end + i]] as [T, T]));
  return out;
}

/**
 * One character of a text run, or one other run-level element. A text token shares its run (`el`) with the run's
 * other characters; `tokenEl` makes the one-character copy only when it is written out. `sig` (properties and
 * wrappers) is serialized once per run, so comparing tokens doesn't serialize XML on every keystroke.
 */
interface Token { el: Element; wrappers: Element[]; text: string | null; key: string; sig: string; }
function tokens(parent: Element, wrappers: Element[] = [], wrapSig = ""): Token[] {
  return Array.from(parent.children).flatMap(el => {
    if (el.namespaceURI === NS.w && ["ins", "hyperlink"].includes(el.localName)) {
      const wrapper = el.cloneNode(false) as Element;
      return tokens(el, [...wrappers, wrapper], wrapSig + xml(wrapper));
    }
    const t = w(el, "t");
    if (el.localName === "r" && t && Array.from(el.children).every(c => c === t || c.localName === "rPr")) {
      const bare = el.cloneNode(true) as Element;
      w(bare, "t")!.remove();
      const sig = xml(bare) + wrapSig;
      return Array.from(t.textContent ?? "").map(text => ({ el, wrappers, text, key: text, sig }));
    }
    // An empty run (Word leaves them in blank paragraphs) holds nothing to record.
    if (el.localName === "pPr" || (el.localName === "r" && Array.from(el.children).every(c => c.localName === "rPr"))) return [];
    const key = xml(el);
    return [{ el: copy(el), wrappers, text: null, key, sig: key + wrapSig }];
  });
}
const tokenKey = (t: Token) => t.key;
const tokenSame = (a: Token, b: Token) => a.text === b.text && a.sig === b.sig;
/** A copy of the token's element; for a character, its run holding just that character (not a copy of the whole run). */
function tokenEl(t: Token): Element {
  if (t.text === null) return copy(t.el);
  const run = t.el.cloneNode(false) as Element;
  for (const child of Array.from(t.el.children)) {
    if (child.localName !== "t") run.appendChild(copy(child));
    else {
      const text = run.appendChild(child.cloneNode(false)) as Element;
      text.textContent = t.text;
      text.setAttributeNS(NS.xml, "xml:space", "preserve");
    }
  }
  return run;
}
const xmlCache = new WeakMap<Element, string>();
/** XML of an element that is never changed after it is made (token wrappers), serialized once. */
function cachedXml(el: Element): string {
  let out = xmlCache.get(el);
  if (out === undefined) xmlCache.set(el, (out = xml(el)));
  return out;
}

class Recorder {
  private formatChanges = new Map<string, Element>();
  private date = new Date().toISOString();
  constructor(public id: number, private author: string) {}
  private revision(doc: Document, name: string): Element {
    const el = element(doc, name);
    for (const [key, value] of Object.entries({ id: String(this.id++), author: this.author, date: this.date })) el.setAttributeNS(NS.w, `w:${key}`, value);
    return el;
  }
  private properties(old: Element | null, next: Element, name: string): void {
    const now = w(next, name);
    // Compare the written XML: rebuilt properties can differ from parsed ones only in how attributes were created.
    if ((!old && !now) || (old && now && xml(old) === xml(now))) return;
    const props = now ?? next.insertBefore(element(next.ownerDocument, name), next.firstChild);
    if (w(props, `${name}Change`)) return; // retain the original baseline of an existing suggestion
    const signature = `${name}:${old ? xml(old) : ""}:${xml(props)}`;
    const previous = this.formatChanges.get(signature);
    if (previous) { props.appendChild(copy(previous)); return; }
    const change = this.revision(next.ownerDocument, `${name}Change`);
    const baseline = old ? copy(old) : element(next.ownerDocument, name);
    w(baseline, `${name}Change`)?.remove();
    if (name === "pPr") {
      w(baseline, "rPr")?.remove();
      w(baseline, "sectPr")?.remove();
    }
    change.appendChild(baseline);
    props.appendChild(change);
    this.formatChanges.set(signature, change);
  }
  /** Whether this author's insertion / deletion sits at pair `p` (for keeping one edit together). */
  private mine(p: [Token | null, Token | null] | undefined, kind: "ins" | "del"): boolean {
    if (!p) return false;
    const token = p[1] ?? p[0]!;
    if (kind === "del") return token.el.localName === "del" && wVal(token.el, "author") === this.author;
    return !!p[1] && token.wrappers.some((w) => w.localName === "ins" && wVal(w, "author") === this.author);
  }
  private inline(old: Element, next: Element): void {
    let pairs = align(tokens(old), tokens(next), tokenSame, (a, b) => tokenKey(a) === tokenKey(b));
    // A letter added or removed beside the same letter ("ll") could be either one: take the one next to this
    // author's edit there, so typing or deleting a key at a time stays one suggestion.
    const equal = (p: [Token | null, Token | null] | undefined) => !!p && !!p[0] && !!p[1] && tokenSame(p[0], p[1]);
    for (let k = 0; k < pairs.length; k++) {
      const [a, b] = pairs[k];
      const edit = a && !b ? a : b && !a ? b : null;
      if (!edit || edit.text === null) continue;
      const kind = a ? "del" : "ins";
      let lo = k, hi = k;
      while (equal(pairs[lo - 1]) && pairs[lo - 1][0]!.text === edit.text) lo--;
      while (equal(pairs[hi + 1]) && pairs[hi + 1][0]!.text === edit.text) hi++;
      if (lo === hi || this.mine(pairs[k - 1], kind) || this.mine(pairs[k + 1], kind)) continue;
      const to = this.mine(pairs[hi + 1], kind) ? hi : this.mine(pairs[lo - 1], kind) ? lo : k;
      if (to === k) continue;
      const [moved] = pairs.splice(k, 1);
      pairs.splice(to, 0, moved);
    }
    // A replaced word is one deletion plus one insertion; shared letters within it are not separate edits.
    const first = pairs.findIndex(([a, b]) => !a || !b);
    let last = pairs.length - 1;
    while (last >= 0 && pairs[last][0] && pairs[last][1]) last--;
    if (first >= 0) {
      const span = pairs.slice(first, last + 1);
      let streak = 0, longest = 0;
      for (const [a, b] of span) { streak = a && b ? streak + 1 : 0; longest = Math.max(streak, longest); }
      if (longest < 8 && span.every(([a, b]) => (a ?? b)!.text !== null)) {
        pairs = [...pairs.slice(0, first), ...span.flatMap(([a]) => a ? [[a, null] as [Token, null]] : []),
          ...span.flatMap(([, b]) => b ? [[null, b] as [null, Token]] : []), ...pairs.slice(last + 1)];
      }
    }
    const children: Element[] = [];
    let lastSignature = "", target: Element | null = null;
    let active: { kind: string; el: Element } | null = null;
    /** This author's insertion / deletion right beside pair `i`, to extend instead of starting another (typing goes a key at a time). */
    const neighbour = (i: number, kind: string): Element | null => {
      for (const j of [i - 1, i + 1]) {
        const p = pairs[j];
        if (!p) continue;
        const token = p[1] ?? p[0]!;
        const el = kind === "del" ? (token.el.localName === "del" ? token.el : null) : p[1] ? token.wrappers.find((w) => w.localName === "ins") ?? null : null;
        if (el && wVal(el, "author") === this.author) return el;
      }
      return null;
    };
    // Per character: wrappers stay the shared (never mutated) elements, compared by cached XML and copied only when a
    // new wrapper starts; a run joins the previous one when their properties match (compared by key, not re-serialized).
    let lastKey: string | null = null;
    for (const [index, [a, b]] of pairs.entries()) {
      const token = b ?? a!;
      let run = tokenEl(token);
      let runKey = token.text !== null ? token.sig : null;
      const wrappers = token.wrappers.slice();
      if (a && b) {
        active = null;
        const insertion = a.wrappers.find(e => e.localName === "ins");
        if (insertion && !wrappers.some(e => e.localName === "ins")) wrappers.unshift(insertion);
        if (!tokenSame(a, b) && a.text !== null && b.text !== null && !wrappers.some(e => e.localName === "ins")) {
          this.properties(w(a.el, "rPr"), run, "rPr");
          runKey = xml(w(run, "rPr") ?? element(run.ownerDocument, "rPr"));
        }
      } else {
        const kind = b ? "ins" : "del";
        if (!b && wrappers.some(e => e.localName === "ins" && wVal(e, "author") === this.author)) continue;
        // Existing review anchors/deletions must not become nested revisions when deleting a selection.
        if (!b && ["del", "commentRangeStart", "commentRangeEnd"].includes(run.localName)) {
          if (run.localName === "del") children.push(run);
          continue;
        }
        if (!active || active.kind !== kind) {
          const near = neighbour(index, kind);
          active = { kind, el: near ? (near.cloneNode(false) as Element) : this.revision(next.ownerDocument, kind) };
        }
        for (let i = wrappers.length - 1; i >= 0; i--) if (wrappers[i].localName === "ins") wrappers.splice(i, 1);
        wrappers.unshift(active.el);
        if (!b) for (const t of Array.from(run.getElementsByTagNameNS(NS.w, "t"))) {
          const deleted = element(next.ownerDocument, "delText");
          deleted.setAttributeNS(NS.xml, "xml:space", "preserve");
          deleted.textContent = t.textContent;
          t.replaceWith(deleted);
        }
        if (runKey !== null) runKey = `${kind}:${runKey}`;
      }
      const signature = wrappers.map(cachedXml).join("");
      if (signature !== lastSignature) lastKey = null;
      if (signature !== lastSignature || !target) {
        target = null;
        for (const wrapper of wrappers) {
          const placed = wrapper.cloneNode(false) as Element;
          if (target) target.appendChild(placed); else children.push(placed);
          target = placed;
        }
        lastSignature = signature;
      }
      const container = target;
      const previous = container?.lastElementChild ?? (!wrappers.length ? children.at(-1) : null);
      const t = w(run, "t") ?? w(run, "delText");
      const previousText = previous && (w(previous, "t") ?? w(previous, "delText"));
      if (t && previousText && runKey !== null && runKey === lastKey && previous.localName === "r") previousText.textContent += t.textContent ?? "";
      else if (container) container.appendChild(run); else children.push(run);
      lastKey = t ? runKey : null;
    }
    for (const el of Array.from(next.children)) if (el.localName !== "pPr") el.remove();
    next.append(...children);
    // One insertion / deletion that grew a key at a time is one element.
    for (const el of Array.from(next.children)) {
      const prev = el.previousElementSibling;
      if (prev && (el.localName === "ins" || el.localName === "del") && prev.localName === el.localName && wVal(prev, "id") === wVal(el, "id")) {
        prev.append(...Array.from(el.childNodes));
        el.remove();
      }
    }
  }
  private block(el: Element, kind: "ins" | "del"): Element {
    const out = copy(el);
    if (out.localName === "p") {
      const empty = out.cloneNode(false) as Element;
      if (kind === "ins") this.inline(empty, out);
      else {
        const properties = w(out, "pPr");
        this.inline(out, empty);
        out.replaceChildren(...(properties ? [copy(properties)] : []), ...Array.from(empty.children));
      }
      const pPr = w(out, "pPr") ?? out.insertBefore(element(out.ownerDocument, "pPr"), out.firstChild);
      const rPr = w(pPr, "rPr") ?? pPr.insertBefore(element(out.ownerDocument, "rPr"), w(pPr, "sectPr"));
      rPr.appendChild(this.revision(out.ownerDocument, kind));
    } else if (out.localName === "tr") {
      // Word strikes through (or underlines) a deleted (inserted) row's text as well.
      for (const cell of Array.from(out.children).filter(c => c.localName === "tc")) {
        for (const p of Array.from(cell.children).filter(c => c.localName === "p")) p.replaceWith(this.block(p, kind));
      }
      const props = w(out, "trPr") ?? out.insertBefore(element(out.ownerDocument, "trPr"), out.firstChild);
      props.insertBefore(this.revision(out.ownerDocument, kind), w(props, "trPrChange"));
    } else if (out.localName === "tbl") {
      for (const row of Array.from(out.children).filter(c => c.localName === "tr")) row.replaceWith(this.block(row, kind));
    } else if (out.localName === "tc") {
      // An inserted / deleted cell (a table column): w:cellIns / w:cellDel, with its text marked as well.
      for (const p of Array.from(out.children).filter(c => c.localName === "p")) p.replaceWith(this.block(p, kind));
      const props = w(out, "tcPr") ?? out.insertBefore(element(out.ownerDocument, "tcPr"), out.firstChild);
      props.insertBefore(this.revision(out.ownerDocument, kind === "ins" ? "cellIns" : "cellDel"), w(props, "tcPrChange"));
    }
    return out;
  }
  /** Whether `whole` holds the tokens of `parts` back to back (a paragraph split in two, or two joined). */
  private static joined(whole: Element, parts: Element[]): boolean {
    const a = tokens(whole), b = parts.flatMap(p => tokens(p));
    return a.length === b.length && a.every((t, i) => tokenSame(t, b[i]));
  }
  /** Mark a paragraph's mark as inserted or deleted, leaving its text alone (Enter / joining paragraphs). */
  private mark(p: Element, kind: "ins" | "del"): Element {
    const out = copy(p);
    const pPr = w(out, "pPr") ?? out.insertBefore(element(out.ownerDocument, "pPr"), out.firstChild);
    const rPr = w(pPr, "rPr") ?? pPr.insertBefore(element(out.ownerDocument, "rPr"), w(pPr, "sectPr") ?? w(pPr, "pPrChange"));
    rPr.appendChild(this.revision(out.ownerDocument, kind));
    return out;
  }
  /**
   * Paragraph splits and joins among aligned block pairs: one old paragraph against itself plus inserted ones
   * (Enter), or old paragraphs against one new paragraph plus deleted ones (Backspace / Delete at a boundary).
   * Returns the recorded blocks, or null when `pairs[i]` doesn't start one.
   */
  private splitOrJoin(pairs: [Element | null, Element | null][], i: number): { out: Element[]; used: number } | null {
    const [a, b] = pairs[i];
    if (!a || !b || a.localName !== "p" || b.localName !== "p") return null;
    let n = 1;
    while (i + n < pairs.length && !pairs[i + n][0] && pairs[i + n][1]?.localName === "p") n++;
    if (n > 1) {
      const parts = pairs.slice(i, i + n).map(([, x]) => x!);
      if (!Recorder.joined(a, parts)) return null;
      // The new paragraphs end in inserted marks; the last one keeps the original mark.
      const last = copy(parts[n - 1]);
      w(last, "pPr")?.replaceWith(...(w(a, "pPr") ? [copy(w(a, "pPr")!)] : []));
      return { out: [...parts.slice(0, -1).map(x => this.mark(x, "ins")), last], used: n };
    }
    while (i + n < pairs.length && pairs[i + n][0]?.localName === "p" && !pairs[i + n][1]) n++;
    if (n > 1) {
      const olds = pairs.slice(i, i + n).map(([x]) => x!);
      if (!Recorder.joined(b, olds)) return null;
      return { out: [...olds.slice(0, -1).map(x => this.mark(x, "del")), copy(olds[n - 1])], used: n };
    }
    return null;
  }
  /** A paragraph like `p` (its properties) holding `parts`, each wrapped back in its insertion / link. */
  private static paragraph(p: Element, parts: Token[]): Element {
    const out = p.cloneNode(false) as Element;
    const props = w(p, "pPr");
    if (props) out.appendChild(copy(props));
    for (const t of parts) out.appendChild(t.wrappers.reduceRight<Element>((inner, wrapper) => {
      const outer = wrapper.cloneNode(false) as Element;
      outer.appendChild(inner);
      return outer;
    }, tokenEl(t)));
    return out;
  }
  /**
   * A selection deleted (or typed over) across paragraphs: old paragraphs against the one they became. Each old
   * paragraph keeps its own text with the cut part deleted, the typed text goes where the cut starts, and every
   * paragraph mark but the last is deleted, as Word records it. Returns null when `pairs[i]` doesn't start one.
   */
  private partialJoin(pairs: [Element | null, Element | null][], i: number): { out: Element[]; used: number } | null {
    // Old blocks (tables in between included) against one new paragraph, which must come from old paragraphs at both ends.
    let n = 0, at = -1;
    for (; i + n < pairs.length; n++) {
      const [a, b] = pairs[i + n];
      if (!a || (b && (b.localName !== "p" || a.localName !== "p" || at >= 0))) break;
      if (b) at = n;
    }
    while (n > 0 && pairs[i + n - 1][0]!.localName !== "p") n--;
    const olds = pairs.slice(i, i + n).map(([a]) => a!);
    if (n < 2 || at < 0 || at >= n || olds[0].localName !== "p") return null;
    const merged = pairs[i + at][1]!;
    const parts = olds.map(p => (p.localName === "p" ? tokens(p) : [])), all = parts.flat(), now = tokens(merged);
    let head = 0, tail = 0;
    while (head < Math.min(now.length, all.length) && tokenKey(now[head]) === tokenKey(all[head])) head++;
    while (tail < Math.min(now.length, all.length) - head && tokenKey(now.at(-1 - tail)!) === tokenKey(all.at(-1 - tail)!)) tail++;
    const typed = now.slice(head, now.length - tail);
    // Blocks other than paragraphs can only be deleted whole: the cut has to cover them.
    let offset = 0;
    for (const [j, old] of olds.entries()) {
      if (old.localName !== "p" && (head > offset || all.length - tail < offset)) return null;
      offset += parts[j].length;
    }
    let start = 0, placed = false;
    const out = olds.map((old, j) => {
      if (old.localName !== "p") return this.block(old, "del");
      const end = start + parts[j].length;
      const kept = now.slice(start, Math.max(start, Math.min(end, head)));
      if (!placed && head <= end) { kept.push(...typed); placed = true; }
      for (let k = Math.max(start, all.length - tail); k < end; k++) kept.push(now[k - all.length + now.length]);
      start = end;
      const next = Recorder.paragraph(old, kept);
      this.inline(old, next);
      return j < n - 1 ? this.mark(next, "del") : next;
    });
    return { out, used: n };
  }
  patch(old: Element, next: Element): Element {
    const out = copy(next);
    if (old.isEqualNode(next)) return out;
    if (out.localName === "p") {
      const oldSection = w(w(old, "pPr"), "sectPr"), nextSection = w(w(out, "pPr"), "sectPr");
      if (oldSection && nextSection && !oldSection.isEqualNode(nextSection)) {
        const holder = element(out.ownerDocument, "holder");
        holder.appendChild(copy(nextSection));
        this.properties(oldSection, holder, "sectPr");
        nextSection.replaceWith(holder.firstElementChild!);
      }
      this.properties(w(old, "pPr"), out, "pPr");
      this.inline(old, out);
      return out;
    }
    if (["sectPr", "tblPr", "trPr", "tcPr"].includes(out.localName)) {
      const holder = element(out.ownerDocument, "holder");
      holder.appendChild(out);
      this.properties(old, holder, out.localName);
      return out;
    }
    if (out.localName === "tblGrid") {
      // Column widths before the change, so a rejected column insertion / deletion gets its grid back.
      const change = this.revision(out.ownerDocument, "tblGridChange");
      const grid = copy(old);
      w(grid, "tblGridChange")?.remove();
      change.appendChild(grid);
      w(out, "tblGridChange")?.remove();
      out.appendChild(change);
      return out;
    }
    if (["body", "tbl", "tr", "tc", "sdt", "sdtContent"].includes(out.localName)) {
      const pairs = align(Array.from(old.children), Array.from(next.children), (a, b) => a.isEqualNode(b),
        (a, b) => a.namespaceURI === b.namespaceURI && a.localName === b.localName && ["p", "tbl", "tr", "tc", "sectPr", "tblPr", "tblGrid", "trPr", "tcPr", "sdt", "sdtContent"].includes(a.localName));
      const children: Element[] = [];
      for (let i = 0; i < pairs.length; i++) {
        const joined = this.splitOrJoin(pairs, i) ?? this.partialJoin(pairs, i);
        if (joined) {
          children.push(...joined.out);
          i += joined.used - 1;
          continue;
        }
        const [a, b] = pairs[i];
        const lone = (a ?? b)!;
        // Properties added or removed whole (a cell's first gridSpan) are a change from / to empty properties.
        if ((!a || !b) && ["tblPr", "trPr", "tcPr"].includes(lone.localName)) children.push(this.patch(a ?? element(lone.ownerDocument, lone.localName), b ?? element(lone.ownerDocument, lone.localName)));
        else children.push(a && b ? this.patch(a, b) : this.block(lone, b ? "ins" : "del"));
      }
      out.replaceChildren(...children);
    }
    return out;
  }
}

let generation = 0;
const originIndexes = new WeakMap<Origins, Map<string, string[]>>();
/** Preserve origin versions referenced by undo history instead of replacing them after each edit. */
export function retainOrigins(doc: PMNode, readOrigins: Origins, origins: Origins): PMNode {
  const prefix = `suggestion${++generation}:`;
  let index = originIndexes.get(origins);
  if (!index) {
    index = new Map();
    for (const [key, origin] of origins) {
      const signature = xml(origin.el);
      index.set(signature, [...index.get(signature) ?? [], key]);
    }
    originIndexes.set(origins, index);
  }
  const used = new Set<string>();
  const visit = (node: PMNode): PMNode => {
    if (node.isText) return node;
    const children: PMNode[] = [];
    node.forEach(child => children.push(visit(child)));
    const key = attrsOf<{ key?: string }>(node).key;
    const origin = key ? readOrigins.get(key)! : null;
    const signature = origin ? xml(origin.el) : "";
    const candidates = index.get(signature) ?? [];
    const retainedKey = key ? candidates.find(candidate => !used.has(candidate)) ?? prefix + key : null;
    const result = node.type.create(key ? { ...node.attrs, key: retainedKey } : node.attrs, Fragment.from(children), node.marks);
    if (retainedKey) {
      used.add(retainedKey);
      if (!origins.has(retainedKey)) {
        origins.set(retainedKey, { el: origin!.el, node: result });
        index.set(signature, [...candidates, retainedKey]);
      }
    }
    return result;
  };
  return visit(doc);
}

/** Map the proposed cursor past the deleted atoms which suggestions retain in the editor. */
function cursorAt(source: PMNode, target: PMNode, pos: number, origins: Origins, forward = false): number {
  const at = cursorBefore(source, target, pos, origins);
  if (!forward) return at;
  // After the Delete key the cursor moves past what it just struck through, so the next Delete reaches new text.
  let out = at;
  for (let node = target.resolve(out).nodeAfter; node && node.type.name === "opaque" && /<(?:\w+:)?del\b/.test(attrsOf<XmlAttrs>(node).xml); node = target.resolve(out).nodeAfter) out += node.nodeSize;
  return out;
}

function cursorBefore(source: PMNode, target: PMNode, pos: number, origins: Origins): number {
  const deleted = (node: PMNode) => node.type.name === "opaque" && /<(?:\w+:)?(?:del|moveFrom)\b/.test(attrsOf<XmlAttrs>(node).xml);
  /** Paragraphs as the reader sees them: one whose mark is deleted runs on into the next. */
  const paragraphs = (doc: PMNode) => {
    const out: { node: PMNode; pos: number }[][] = [];
    let joining = false;
    doc.descendants((node, at) => {
      if (node.type.name !== "paragraph") return true;
      const key = attrsOf<{ key: string | null }>(node).key;
      const original = key ? origins.get(key)?.el : null;
      if (joining) out[out.length - 1].push({ node, pos: at });
      else out.push([{ node, pos: at }]);
      joining = !!w(w(w(original, "pPr"), "rPr"), "del");
      return false;
    });
    return out;
  };
  const from = paragraphs(source), to = paragraphs(target);
  const index = from.findIndex(group => group.some(p => pos > p.pos && pos < p.pos + p.node.nodeSize));
  const old = from[index], next = to[index];
  if (!old || !next) return Math.min(pos, target.content.size);
  let remaining = 0;
  for (const p of old) p.node.forEach((child, offset) => {
    const distance = pos - p.pos - 1 - offset;
    if (distance > 0 && !deleted(child)) remaining += Math.min(child.nodeSize, distance);
  });
  // Walk the target group's text; a position at a paragraph's end belongs to the next paragraph of the group.
  for (const [k, p] of next.entries()) {
    let result = p.pos + 1;
    let done = false;
    p.node.forEach((child, offset) => {
      if (done) return;
      result = p.pos + 1 + offset;
      if (deleted(child)) { result += child.nodeSize; return; }
      if (remaining < child.nodeSize || (remaining === child.nodeSize && k === next.length - 1)) { result += remaining; done = true; }
      else { remaining -= child.nodeSize; result += child.nodeSize; }
    });
    if (done || k === next.length - 1) return Math.min(done ? result : p.pos + p.node.nodeSize - 1, target.content.size);
  }
  return Math.min(pos, target.content.size);
}

/**
 * Apply `next` over `old` (the same blocks, starting at `pos`) as small steps: changed attributes in place and only
 * the differing text inside a paragraph replaced. Decorations elsewhere (page gaps, spacers) then stay put instead of
 * vanishing with a replaced block. Returns false, changing nothing, when the block structure differs.
 */
function patchInPlace(tr: Transaction, pos: number, old: Fragment, next: Fragment): boolean {
  const ops: (() => void)[] = [];
  const walk = (at: number, a: Fragment, b: Fragment): boolean => {
    if (a.childCount !== b.childCount) return false;
    for (let i = 0; i < a.childCount; at += a.child(i).nodeSize, i++) {
      const x = a.child(i), y = b.child(i), here = at;
      if (x.type !== y.type) return false;
      if (x !== y && !x.eq(y)) {
        if (x.isTextblock) {
          const start = x.content.findDiffStart(y.content);
          const end = x.content.findDiffEnd(y.content);
          if (start !== null && end) {
            const [endA, endB] = [Math.max(end.a, start), Math.max(end.b, start)];
            ops.push(() => tr.replace(here + 1 + start, here + 1 + endA, new Slice(y.content.cut(start, endB), 0, 0)));
          }
        } else if (x.isLeaf || !walk(here + 1, x.content, y.content)) return false;
        // Attribute steps (unlike setNodeMarkup) don't touch positions, so nothing mapped across them is lost.
        if (!Mark.sameSet(x.marks, y.marks)) ops.push(() => tr.setNodeMarkup(here, undefined, y.attrs, y.marks));
        else for (const [name, value] of Object.entries(y.attrs)) if (x.attrs[name] !== value) ops.push(() => tr.setNodeAttribute(here, name, value));
      }
    }
    return true;
  };
  if (!walk(pos, old, next)) return false;
  // Last first, so earlier positions stay valid.
  for (const op of ops.reverse()) op();
  return true;
}

const maxId = (root: Element) => Array.from(root.getElementsByTagNameNS(NS.w, "*")).reduce((n, el) => Math.max(n, Number(wVal(el, "id")) || 0), 0);
/** The next free revision id per document (its origins), found from the whole body once, then counted up. */
const nextIds = new WeakMap<Origins, number>();

/**
 * Turn the user's transaction into one undoable Word-compatible suggestion. Does not mutate the package.
 * Only the top-level blocks the edit changed (and one neighbour each side) are written, compared and read back,
 * so a keystroke costs the same in a long document as in a short one.
 */
export function suggestTransaction(tr: Transaction, origins: Origins, ctx: WriteContext, readContext: ReadContext, author: string, opts: { forward?: boolean } = {}): Transaction {
  if (!tr.docChanged || isHistoryTransaction(tr) || tr.getMeta("addToHistory") === false || tr.getMeta("suggestionsRecorded")) return tr;
  const a = tr.before, b = tr.doc;
  let start = 0, endA = a.childCount, endB = b.childCount;
  // A change to the document's own attributes (the final section) goes through the whole body.
  if (a.sameMarkup(b)) {
    const same = (x: PMNode, y: PMNode) => x === y || x.eq(y);
    while (start < endA && start < endB && same(a.child(start), b.child(start))) start++;
    while (endA > start && endB > start && same(a.child(endA - 1), b.child(endB - 1))) { endA--; endB--; }
    start = Math.max(0, start - 1);
    const after = Math.min(1, a.childCount - endA);
    endA += after;
    endB += after;
  }
  const at = (doc: PMNode, index: number) => { let pos = 0; for (let k = 0; k < index; k++) pos += doc.child(k).nodeSize; return pos; };
  const fromA = at(a, start), toA = at(a, endA), fromB = at(b, start), toB = at(b, endB);
  const owner = document.implementation.createDocument(NS.w, "w:body");
  const before = owner.documentElement;
  writeBody(before, a.copy(a.content.cut(fromA, toA)), origins, ctx);
  const afterBody = copy(before);
  writeBody(afterBody, b.copy(b.content.cut(fromB, toB)), origins, ctx);
  let first = nextIds.get(origins);
  if (first === undefined) {
    const whole = owner.createElementNS(NS.w, "w:body");
    writeBody(whole, a, origins, ctx);
    first = maxId(whole) + 1;
  }
  const recorder = new Recorder(Math.max(first, maxId(before) + 1), author);
  const recorded = recorder.patch(before, afterBody);
  nextIds.set(origins, recorder.id);
  const read = readDocx(recorded, readContext);
  const suggested = retainOrigins(read.doc, read.origins, origins);
  // The cursor, mapped within the recorded blocks only (outside them nothing moved but the size of what's between).
  const source = b.copy(b.content.cut(fromB, toB)), grown = suggested.content.size - (toB - fromB);
  const place = (pos: number, forward?: boolean) => pos < fromB ? pos : pos > toB ? pos + grown : fromB + cursorAt(source, suggested, pos - fromB, origins, forward);
  const selection = tr.selection;
  const from = place(selection.from, opts.forward && selection.empty), to = selection.empty ? from : place(selection.to);
  if (!patchInPlace(tr, fromB, b.content.cut(fromB, toB), suggested.content)) tr.replaceWith(fromB, toB, suggested.content);
  if (!a.sameMarkup(b)) tr.setDocAttribute("sectPr", suggested.attrs.sectPr);
  tr.setSelection(TextSelection.near(tr.doc.resolve(from)));
  if (from !== to) tr.setSelection(TextSelection.create(tr.doc, from, to));
  return tr.setMeta("suggestionsRecorded", true);
}
