/** Excel's threaded comments, with authors shared by the workbook and messages stored per worksheet. */
import type { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import type { Edits } from "./commands";
import { assertAction } from "./protection";
import type { Sheet } from "./workbook-model";

export const THREAD_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
const guid = () => `{${crypto.randomUUID().toUpperCase()}}`;
export interface CommentMessage { id: string; text: string; author: string; date: string; }
export interface CommentThread { id: string; resolved: boolean; messages: CommentMessage[]; }
const messages = (doc: Document | null) => Array.from(doc?.getElementsByTagNameNS(THREAD_NS, "threadedComment") ?? []);
const threadPath = (pkg: XlsxWorkbook, sheetPath: string) => pkg.rels(sheetPath).find(r => r.type === "threadedComment" && !r.external)?.target;

export function commentsOf(pkg: XlsxWorkbook, sheetPath: string): Map<string, CommentThread[]> {
  const out = new Map<string, CommentThread[]>();
  const people = new Map<string, string>();
  for (const rel of pkg.rels(pkg.workbookPath)) if (rel.type === "person" && !rel.external) {
    for (const p of Array.from(pkg.xml(rel.target)?.getElementsByTagNameNS(THREAD_NS, "person") ?? [])) people.set(p.getAttribute("id") ?? "", p.getAttribute("displayName") ?? "");
  }
  const path = threadPath(pkg, sheetPath), all = messages(path ? pkg.xml(path) : null);
  const read = (el: Element): CommentMessage => ({
    id: el.getAttribute("id") ?? "",
    text: Array.from(el.children).find(c => c.localName === "text")?.textContent ?? "",
    author: people.get(el.getAttribute("personId") ?? "") ?? "Unknown author",
    date: el.getAttribute("dT") ?? "",
  });
  for (const root of all.filter(c => !c.getAttribute("parentId"))) {
    const ref = root.getAttribute("ref") ?? "", id = root.getAttribute("id") ?? "";
    const thread: CommentThread = { id, resolved: ["1", "true"].includes(root.getAttribute("done") ?? ""), messages: [read(root), ...all.filter(c => c.getAttribute("parentId") === id).map(read)] };
    out.set(ref, [...(out.get(ref) ?? []), thread]);
  }
  return out;
}

/** Post a root comment or a reply as one undo step; adding a reply reopens its thread. */
export function addComment(e: Edits, sheet: Sheet, ref: string, text: string, author: string, parentId?: string): string {
  assertAction(sheet, "objects");
  if (!text.trim()) throw new Error("Enter a comment before posting.");
  const id = guid();
  e.structural(() => {
    const pkg = e.wb.pkg;
    const path = pkg.ensureThreadedComments(sheet.ref.path), doc = pkg.xml(path)!;
    const parent = parentId ? messages(doc).find(c => c.getAttribute("id") === parentId && !c.hasAttribute("parentId")) : null;
    if (parentId && (!parent || parent.getAttribute("ref") !== ref)) throw new Error("This comment thread no longer exists.");
    const persons = pkg.ensurePersons(), people = pkg.xml(persons)!;
    let person = Array.from(people.documentElement.children).find(p => p.getAttribute("displayName") === author && p.getAttribute("providerId") === "None");
    if (!person) {
      person = people.createElementNS(THREAD_NS, "person");
      person.setAttribute("id", guid()); person.setAttribute("displayName", author);
      person.setAttribute("userId", person.getAttribute("id")!); person.setAttribute("providerId", "None");
      people.documentElement.insertBefore(person, Array.from(people.documentElement.children).find(c => c.localName === "extLst") ?? null);
      pkg.touch(persons);
    }
    const message = doc.createElementNS(THREAD_NS, "threadedComment");
    for (const [k, v] of Object.entries({ id, ref, personId: person.getAttribute("id")!, dT: new Date().toISOString() })) message.setAttribute(k, v);
    if (parent) { message.setAttribute("parentId", parentId!); parent.removeAttribute("done"); }
    message.appendChild(doc.createElementNS(THREAD_NS, "text")).textContent = text;
    doc.documentElement.insertBefore(message, Array.from(doc.documentElement.children).find(c => c.localName === "extLst") ?? null);
    pkg.touch(path);
  });
  return id;
}

function changeMessage(e: Edits, sheet: Sheet, id: string, fn: (message: Element, doc: Document) => void): void {
  assertAction(sheet, "objects");
  const path = threadPath(e.wb.pkg, sheet.ref.path);
  if (!path) throw new Error("This comment no longer exists.");
  e.structural(() => {
    const doc = e.wb.pkg.xml(path)!;
    const message = messages(doc).find(c => c.getAttribute("id") === id);
    if (!message) throw new Error("This comment no longer exists.");
    fn(message, doc); e.wb.pkg.touch(path);
  });
}
export function editComment(e: Edits, sheet: Sheet, id: string, text: string): void {
  if (!text.trim()) throw new Error("Enter a comment before saving.");
  changeMessage(e, sheet, id, (message, doc) => {
    const old = Array.from(message.children).find(c => c.localName === "text");
    const next = doc.createElementNS(THREAD_NS, "text"); next.textContent = text;
    if (old) old.replaceWith(next); else message.insertBefore(next, message.firstChild);
    // Mention offsets refer to the old text; preserve mentions only on untouched comments.
    Array.from(message.children).find(c => c.localName === "mentions")?.remove();
  });
}
export function resolveThread(e: Edits, sheet: Sheet, id: string, resolved: boolean): void {
  changeMessage(e, sheet, id, message => {
    if (message.hasAttribute("parentId")) throw new Error("Choose the thread's first comment.");
    if (resolved) message.setAttribute("done", "1"); else message.removeAttribute("done");
  });
}
export function removeThread(e: Edits, sheet: Sheet, id: string): void {
  changeMessage(e, sheet, id, (message, doc) => {
    for (const reply of messages(doc)) if (reply.getAttribute("parentId") === id) reply.remove();
    message.remove();
  });
}
