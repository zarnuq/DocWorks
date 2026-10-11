/**
 * Cell notes (Google Sheets' notes; Excel's comments): kept in the sheet's comments part. The VML drawing Excel
 * shows them with is rebuilt from that part when the workbook is saved (xlsx-workbook).
 */
import type { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";

const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export interface Note {
  text: string;
  author: string;
}

const kids = (el: Element | null, name: string) => Array.from(el?.children ?? []).filter((c) => c.namespaceURI === S && c.localName === name);

/** A comment's text: its runs (or plain text) joined. */
const textOf = (comment: Element) => Array.from(comment.getElementsByTagNameNS(S, "t")).map((t) => t.textContent ?? "").join("");

/** The sheet's notes by cell reference. */
export function notesOf(pkg: XlsxWorkbook, sheetPath: string): Map<string, Note> {
  const path = pkg.commentsPath(sheetPath);
  const doc = path ? pkg.xml(path) : null;
  const out = new Map<string, Note>();
  if (!doc) return out;
  const authors = kids(kids(doc.documentElement, "authors")[0] ?? null, "author").map((a) => a.textContent ?? "");
  for (const c of Array.from(doc.getElementsByTagNameNS(S, "comment"))) {
    out.set(c.getAttribute("ref") ?? "", { text: textOf(c), author: authors[Number(c.getAttribute("authorId"))] ?? "" });
  }
  return out;
}

/** Set (or with null, remove) the note on cell `ref`. Removing the sheet's last note removes its parts. */
export function setNote(pkg: XlsxWorkbook, sheetPath: string, ref: string, text: string | null, author: string): void {
  const existing = pkg.commentsPath(sheetPath);
  if (!text && !existing) return;
  const doc = text ? pkg.ensureComments(sheetPath) : pkg.xml(existing!)!;
  const list = kids(doc.documentElement, "commentList")[0] ?? doc.documentElement.appendChild(doc.createElementNS(S, "commentList"));
  for (const c of kids(list, "comment")) if (c.getAttribute("ref") === ref) c.remove();
  if (text) {
    const authors = kids(doc.documentElement, "authors")[0] ?? doc.documentElement.insertBefore(doc.createElementNS(S, "authors"), list);
    let id = kids(authors, "author").findIndex((a) => a.textContent === author);
    if (id < 0) {
      authors.appendChild(doc.createElementNS(S, "author")).textContent = author;
      id = kids(authors, "author").length - 1;
    }
    const comment = list.appendChild(doc.createElementNS(S, "comment"));
    comment.setAttribute("ref", ref);
    comment.setAttribute("authorId", String(id));
    const t = comment.appendChild(doc.createElementNS(S, "text")).appendChild(doc.createElementNS(S, "r")).appendChild(doc.createElementNS(S, "t"));
    t.setAttribute("xml:space", "preserve");
    t.textContent = text;
  }
  pkg.touch(pkg.commentsPath(sheetPath)!);
  if (!kids(list, "comment").length) pkg.removeComments(sheetPath);
}
