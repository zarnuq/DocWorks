/*
 * Tracked changes (revisions) in WordprocessingML: listing them, and accepting or
 * rejecting all of them or a chosen set by w:id. Works on the XML directly.
 */
import { NS } from "../../core/utils/xml";

export interface Revision {
  id: string;
  kind: "insert" | "delete" | "format";
  author: string;
  date: string;
  /** Affected text; "¶" for a paragraph mark. */
  text: string;
}

const W = NS.w;
const INS = new Set(["ins", "moveTo"]);
const DEL = new Set(["del", "moveFrom"]);
const PROP_CHANGES = new Set(["rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange", "tblGridChange", "tblPrExChange", "numberingChange"]);
const RANGE_MARKS = new Set(["moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"]);

const attr = (el: Element, name: string) => el.getAttributeNS(W, name) ?? "";
const all = (root: Element) => Array.from(root.getElementsByTagNameNS(W, "*"));
const textOf = (el: Element) =>
  Array.from(el.getElementsByTagNameNS(W, "*"))
    .filter((t) => t.localName === "t" || t.localName === "delText")
    .map((t) => t.textContent)
    .join("");

/** Paragraph or row mark revisions live inside w:rPr of w:pPr, or w:trPr. */
const markOwner = (el: Element) => {
  const parent = el.parentElement;
  if (parent?.localName === "rPr" && parent.parentElement?.localName === "pPr") return parent.parentElement.parentElement;
  if (parent?.localName === "trPr") return parent.parentElement;
  return null;
};

export function revisions(root: Element): Revision[] {
  const out: Revision[] = [];
  for (const el of all(root)) {
    const kind = INS.has(el.localName) ? "insert" : DEL.has(el.localName) ? "delete" : PROP_CHANGES.has(el.localName) ? "format" : null;
    if (!kind) continue;
    const owner = markOwner(el);
    const text = owner ? (owner.localName === "tr" ? "row" : "¶") : kind === "format" ? textOf(el.parentElement!.parentElement!) : textOf(el);
    out.push({ id: attr(el, "id"), kind, author: attr(el, "author"), date: attr(el, "date"), text });
  }
  return out;
}

function unwrap(el: Element): void {
  el.replaceWith(...Array.from(el.childNodes));
}

/** Join paragraph `p` with the next one (its mark was deleted). */
function joinNext(p: Element): void {
  const next = p.nextElementSibling;
  if (!next || next.localName !== "p") {
    // Removing a proposed trailing paragraph must not leave its empty shell behind.
    if (p.previousElementSibling && !Array.from(p.children).some(c => c.localName !== "pPr")) p.remove();
    return;
  }
  for (const child of Array.from(next.childNodes)) if ((child as Element).localName !== "pPr") p.appendChild(child);
  next.remove();
}

/** Make deleted text normal text again. */
function undelete(el: Element): void {
  for (const t of Array.from(el.getElementsByTagNameNS(W, "*"))) {
    if (t.localName !== "delText" && t.localName !== "delInstrText") continue;
    const n = el.ownerDocument.createElementNS(W, t.localName === "delText" ? "w:t" : "w:instrText");
    for (const a of Array.from(t.attributes)) n.setAttributeNS(a.namespaceURI, a.name, a.value);
    n.append(...Array.from(t.childNodes));
    t.replaceWith(n);
  }
  unwrap(el);
}

/** Put the properties recorded in a w:*PrChange back into its parent. */
function restoreProps(change: Element): void {
  const parent = change.parentElement!;
  const old = Array.from(change.children).find((c) => c.localName.endsWith("Pr") || c.localName === "tblGrid") ?? null;
  // Keep what the change record doesn't cover (the paragraph's run props, section props).
  const keep = new Set(["rPr", "sectPr", "pPrChange"]);
  for (const c of Array.from(parent.children)) {
    if (c === change || (parent.localName === "pPr" && keep.has(c.localName))) continue;
    c.remove();
  }
  change.remove();
  // What's left in a pPr (rPr, sectPr) comes last in schema order.
  const ref = Array.from(parent.children).find((c) => keep.has(c.localName)) ?? null;
  if (old) for (const c of Array.from(old.children)) parent.insertBefore(c, ref);
}

function apply(root: Element, ids: Set<string> | undefined, accept: boolean): number {
  let n = 0;
  const chosen = (el: Element) => !ids || ids.has(attr(el, "id"));
  // Innermost first, so nested changes are handled before their containers.
  for (const el of all(root).reverse()) {
    if (!el.isConnected && !root.contains(el)) continue;
    const local = el.localName;
    if (RANGE_MARKS.has(local)) {
      if (!ids) el.remove();
      continue;
    }
    const isIns = INS.has(local);
    const isDel = DEL.has(local);
    const isProp = PROP_CHANGES.has(local);
    if ((!isIns && !isDel && !isProp) || !chosen(el)) continue;
    n++;
    const owner = markOwner(el);
    if (owner) {
      // Paragraph / row mark: deleting it joins the paragraph with the next, or removes the row.
      el.remove();
      const removeOwner = accept ? isDel : isIns;
      if (removeOwner) {
        if (owner.localName === "p") joinNext(owner);
        else {
          const table = owner.parentElement;
          owner.remove();
          if (table?.localName === "tbl" && !Array.from(table.children).some(c => c.localName === "tr")) table.remove();
        }
      }
    } else if (isProp) {
      if (accept) el.remove();
      else restoreProps(el);
    } else if (isIns) {
      if (accept) unwrap(el);
      else el.remove();
    } else if (accept) {
      el.remove();
    } else {
      undelete(el);
    }
  }
  return n;
}

/** Accept tracked changes (all, or those with the given ids). Returns how many were accepted. */
export function acceptRevisions(root: Element, ids?: Set<string>): number {
  return apply(root, ids, true);
}

export function rejectRevisions(root: Element, ids?: Set<string>): number {
  return apply(root, ids, false);
}
