import { Fragment, Node as PMNode } from "prosemirror-model";
import { TextSelection, Transaction } from "prosemirror-state";
import { isHistoryTransaction } from "prosemirror-history";
import { NS } from "../../core/utils/xml";
import { Origins, ReadContext, readDocx, w, wVal } from "./docx-reader";
import { WriteContext, writeBody } from "./docx-writer";
import { attrsOf, XmlAttrs } from "./schema";

const xml = (el: Element) => new XMLSerializer().serializeToString(el);
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

interface Token { el: Element; wrappers: Element[]; text: string | null; }
function tokens(parent: Element, wrappers: Element[] = []): Token[] {
  return Array.from(parent.children).flatMap(el => {
    if (el.namespaceURI === NS.w && ["ins", "hyperlink"].includes(el.localName)) return tokens(el, [...wrappers, el.cloneNode(false) as Element]);
    const t = w(el, "t");
    if (el.localName === "r" && t && Array.from(el.children).every(c => c === t || c.localName === "rPr")) {
      return Array.from(t.textContent ?? "").map(text => {
        const run = copy(el);
        w(run, "t")!.textContent = text;
        w(run, "t")!.setAttributeNS(NS.xml, "xml:space", "preserve");
        return { el: run, wrappers, text };
      });
    }
    return el.localName === "pPr" ? [] : [{ el: copy(el), wrappers, text: null }];
  });
}
const tokenKey = (t: Token) => t.text ?? xml(t.el);
const tokenSame = (a: Token, b: Token) => tokenKey(a) === tokenKey(b) && xml(a.el) === xml(b.el) && a.wrappers.map(xml).join("") === b.wrappers.map(xml).join("");

class Recorder {
  private id: number;
  private formatChanges = new Map<string, Element>();
  private date = new Date().toISOString();
  constructor(before: Element, private author: string) {
    this.id = Array.from(before.getElementsByTagNameNS(NS.w, "*")).reduce((n, el) => Math.max(n, Number(wVal(el, "id")) || 0), 0) + 1;
  }
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
    for (const [index, [a, b]] of pairs.entries()) {
      const token = b ?? a!;
      let run = copy(token.el);
      const wrappers = token.wrappers.map(copy);
      if (a && b) {
        active = null;
        const insertion = a.wrappers.find(e => e.localName === "ins");
        if (insertion && !wrappers.some(e => e.localName === "ins")) wrappers.unshift(copy(insertion));
        if (!tokenSame(a, b) && a.text !== null && b.text !== null && !wrappers.some(e => e.localName === "ins")) this.properties(w(a.el, "rPr"), run, "rPr");
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
        wrappers.unshift(active.el.cloneNode(false) as Element);
        if (!b) for (const t of Array.from(run.getElementsByTagNameNS(NS.w, "t"))) {
          const deleted = element(next.ownerDocument, "delText");
          deleted.setAttributeNS(NS.xml, "xml:space", "preserve");
          deleted.textContent = t.textContent;
          t.replaceWith(deleted);
        }
      }
      const signature = wrappers.map(xml).join("");
      if (signature !== lastSignature || !target) {
        target = null;
        for (const wrapper of wrappers) {
          if (target) target.appendChild(wrapper); else children.push(wrapper);
          target = wrapper;
        }
        lastSignature = signature;
      }
      const container = target;
      const previous = container?.lastElementChild ?? (!wrappers.length ? children.at(-1) : null);
      const t = w(run, "t") ?? w(run, "delText");
      const previousText = previous && (w(previous, "t") ?? w(previous, "delText"));
      if (t && previousText && t.localName === previousText.localName && previous.localName === "r" && xml(w(run, "rPr") ?? element(run.ownerDocument, "rPr")) === xml(w(previous, "rPr") ?? element(run.ownerDocument, "rPr"))) previousText.textContent += t.textContent ?? "";
      else if (container) container.appendChild(run); else children.push(run);
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
        const joined = this.splitOrJoin(pairs, i);
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
function retainOrigins(doc: PMNode, readOrigins: Origins, origins: Origins): PMNode {
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

/** Turn the user's transaction into one undoable Word-compatible suggestion. Does not mutate the package. */
export function suggestTransaction(tr: Transaction, origins: Origins, ctx: WriteContext, readContext: ReadContext, author: string, opts: { forward?: boolean } = {}): Transaction {
  if (!tr.docChanged || isHistoryTransaction(tr) || tr.getMeta("addToHistory") === false || tr.getMeta("suggestionsRecorded")) return tr;
  const owner = document.implementation.createDocument(NS.w, "w:body");
  const before = owner.documentElement;
  writeBody(before, tr.before, origins, ctx);
  const after = copy(before);
  writeBody(after, tr.doc, origins, ctx);
  const recorded = new Recorder(before, author).patch(before, after);
  const read = readDocx(recorded, readContext);
  const suggested = retainOrigins(read.doc, read.origins, origins);
  const selection = tr.selection;
  const from = cursorAt(tr.doc, suggested, selection.from, origins, opts.forward && selection.empty), to = selection.empty ? from : cursorAt(tr.doc, suggested, selection.to, origins);
  tr.replaceWith(0, tr.doc.content.size, suggested.content);
  tr.setDocAttribute("sectPr", suggested.attrs.sectPr);
  tr.setSelection(TextSelection.near(tr.doc.resolve(from)));
  if (from !== to) tr.setSelection(TextSelection.create(tr.doc, from, to));
  return tr.setMeta("suggestionsRecorded", true);
}
