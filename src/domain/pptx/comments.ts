/*
 * Slide comments. New threads are PowerPoint's modern comments (p188: a comment part per slide, authors in
 * ppt/authors.xml), which hold replies and a resolved status as Google Slides' threads do. Legacy comments
 * (p:cmLst, PowerPoint 2013 reply threading) are read and can be deleted.
 */
import { NS } from "../../core/utils/xml";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";

const P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
const PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const LEGACY_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

export interface SlideMessage {
  id: string;
  text: string;
  author: string;
  date: string;
}
export interface SlideThread {
  id: string;
  resolved: boolean;
  /** A legacy comment: shown and deletable, but replies and status need PowerPoint's modern comments. */
  legacy: boolean;
  messages: SlideMessage[];
}

const kids = (el: Element | null | undefined, ns: string, name: string) => Array.from(el?.children ?? []).filter((c) => c.namespaceURI === ns && c.localName === name);
const kid = (el: Element | null | undefined, ns: string, name: string) => kids(el, ns, name)[0] ?? null;
const textOf = (body: Element | null) => kids(body, NS.a, "p").map((p) => Array.from(p.getElementsByTagNameNS(NS.a, "t")).map((t) => t.textContent).join("")).join("\n");
const guid = () => `{${crypto.randomUUID().toUpperCase()}}`;
/** Local time without a zone ("2026-10-12T10:56:59.123"), as PowerPoint stamps comments. */
const now = () => {
  const d = new Date();
  return new Date(d.getTime() - d.getTimezoneOffset() * 60_000).toISOString().replace("Z", "");
};

function authorNames(doc: PptxDocument): Map<string, string> {
  const path = doc.commentAuthorsPart(false);
  const list = path ? doc.xml(path) : null;
  return new Map(Array.from(list?.getElementsByTagNameNS(P188, "author") ?? []).map((a) => [a.getAttribute("id") ?? "", a.getAttribute("name") ?? ""]));
}

const legacyPart = (doc: PptxDocument, slide: string) => doc.rels(slide).find((r) => r.uri === LEGACY_REL)?.target ?? null;
const legacyKey = (cm: Element) => `legacy:${cm.getAttribute("authorId")}:${cm.getAttribute("idx")}`;
const legacyParent = (cm: Element) => {
  const parent = cm.getElementsByTagNameNS(P15, "parentCm")[0];
  return parent ? `legacy:${parent.getAttribute("authorId")}:${parent.getAttribute("idx")}` : null;
};

/** The slide's comment threads, legacy ones first (they are older), each with its replies in order. */
export function slideComments(doc: PptxDocument, slide: string): SlideThread[] {
  const out: SlideThread[] = [];
  const legacy = legacyPart(doc, slide);
  const legacyDoc = legacy ? doc.xml(legacy) : null;
  if (legacyDoc) {
    const authorsPath = doc.rels("ppt/presentation.xml").find((r) => r.type === "commentAuthors")?.target;
    const authorsDoc = authorsPath ? doc.xml(authorsPath) : null;
    const names = new Map(Array.from(authorsDoc?.getElementsByTagNameNS(NS.p, "cmAuthor") ?? []).map((a) => [a.getAttribute("id") ?? "", a.getAttribute("name") ?? ""]));
    const byKey = new Map<string, SlideThread>();
    for (const cm of Array.from(legacyDoc.getElementsByTagNameNS(NS.p, "cm"))) {
      const message = { id: legacyKey(cm), text: kid(cm, NS.p, "text")?.textContent ?? "", author: names.get(cm.getAttribute("authorId") ?? "") ?? "", date: cm.getAttribute("dt") ?? "" };
      const parent = legacyParent(cm);
      const thread = parent ? byKey.get(parent) : undefined;
      if (thread) thread.messages.push(message);
      else {
        const t: SlideThread = { id: message.id, resolved: false, legacy: true, messages: [message] };
        byKey.set(message.id, t);
        out.push(t);
      }
    }
  }
  const part = doc.slideCommentPart(slide, false);
  const list = part ? doc.xml(part) : null;
  if (list) {
    const names = authorNames(doc);
    const message = (el: Element): SlideMessage => ({ id: el.getAttribute("id") ?? "", text: textOf(kid(el, P188, "txBody")), author: names.get(el.getAttribute("authorId") ?? "") ?? "", date: el.getAttribute("created") ?? "" });
    for (const cm of kids(list.documentElement, P188, "cm")) {
      const status = cm.getAttribute("status");
      out.push({ id: cm.getAttribute("id") ?? "", resolved: status === "resolved" || status === "closed", legacy: false, messages: [message(cm), ...kids(kid(cm, P188, "replyLst"), P188, "reply").map(message)] });
    }
  }
  return out;
}

/** The modern author id for `name`, adding the author (with initials, as PowerPoint lists them) if needed. */
function authorId(doc: PptxDocument, name: string): string {
  const path = doc.commentAuthorsPart(true)!;
  const list = doc.xml(path)!;
  const found = Array.from(list.getElementsByTagNameNS(P188, "author")).find((a) => a.getAttribute("name") === name);
  if (found) return found.getAttribute("id")!;
  const id = guid();
  const initials = name.split(/\s+/).filter(Boolean).map((w) => w[0].toUpperCase()).join("").slice(0, 3) || "?";
  const a = list.documentElement.appendChild(list.createElementNS(P188, "p188:author"));
  // userId and providerId keep PowerPoint from asking to repair the file.
  for (const [k, v] of [["id", id], ["name", name], ["initials", initials], ["userId", name], ["providerId", "None"]]) a.setAttribute(k, v);
  doc.touch(path);
  return id;
}

function body(list: Document, text: string): Element {
  const tx = list.createElementNS(P188, "p188:txBody");
  tx.appendChild(list.createElementNS(NS.a, "a:bodyPr"));
  tx.appendChild(list.createElementNS(NS.a, "a:lstStyle"));
  for (const line of text.split("\n")) {
    const p = tx.appendChild(list.createElementNS(NS.a, "a:p"));
    const r = p.appendChild(list.createElementNS(NS.a, "a:r"));
    r.appendChild(list.createElementNS(NS.a, "a:rPr")).setAttribute("lang", "en-US");
    r.appendChild(list.createElementNS(NS.a, "a:t")).textContent = line;
  }
  return tx;
}

const findCm = (list: Document, id: string) => Array.from(list.getElementsByTagNameNS(P188, "cm")).find((c) => c.getAttribute("id") === id) ?? null;

/** Add a comment (or, with `parentId`, a reply to that thread); returns its id. */
export function addSlideComment(doc: PptxDocument, slide: string, text: string, author: string, parentId?: string): string {
  const who = authorId(doc, author);
  const part = doc.slideCommentPart(slide, true)!;
  const list = doc.xml(part)!;
  const id = guid();
  if (parentId) {
    const cm = findCm(list, parentId);
    if (!cm) throw new Error("Replies to this comment need PowerPoint.");
    let replies = kid(cm, P188, "replyLst");
    if (!replies) replies = cm.insertBefore(list.createElementNS(P188, "p188:replyLst"), kid(cm, P188, "txBody"));
    const reply = replies.appendChild(list.createElementNS(P188, "p188:reply"));
    for (const [k, v] of [["id", id], ["authorId", who], ["created", now()]]) reply.setAttribute(k, v);
    reply.appendChild(body(list, text));
  } else {
    const cm = list.documentElement.appendChild(list.createElementNS(P188, "p188:cm"));
    for (const [k, v] of [["id", id], ["authorId", who], ["created", now()]]) cm.setAttribute(k, v);
    // Anchored to the slide: its id in the presentation and its creation id (or a fresh one, as Office's samples).
    const ref = doc.slides().find((s) => s.path === slide);
    // The slide's own creation id lives in its root extension list (shapes carry other creation ids).
    const root = doc.xml(slide)?.documentElement;
    const creation = kids(kid(root, NS.p, "extLst"), NS.p, "ext").flatMap((e) => kids(e, P14, "creationId"))[0]?.getAttribute("val");
    const marks = cm.appendChild(list.createElementNS(PC, "pc:sldMkLst"));
    marks.appendChild(list.createElementNS(PC, "pc:docMk"));
    const mark = marks.appendChild(list.createElementNS(PC, "pc:sldMk"));
    mark.setAttribute("cId", creation ?? String(100000000 + Math.floor(Math.random() * 899999999)));
    mark.setAttribute("sldId", ref?.id ?? "256");
    cm.appendChild(body(list, text));
  }
  doc.touch(part);
  return id;
}

/** Change a comment's or reply's text. */
export function editSlideComment(doc: PptxDocument, slide: string, id: string, text: string): void {
  const part = doc.slideCommentPart(slide, false);
  const list = part ? doc.xml(part) : null;
  const el = list && [...Array.from(list.getElementsByTagNameNS(P188, "cm")), ...Array.from(list.getElementsByTagNameNS(P188, "reply"))].find((c) => c.getAttribute("id") === id);
  if (!list || !el) return;
  kid(el, P188, "txBody")?.replaceWith(body(list, text));
  doc.touch(part!);
}

export function resolveSlideThread(doc: PptxDocument, slide: string, id: string, resolved: boolean): void {
  const part = doc.slideCommentPart(slide, false);
  const cm = part ? findCm(doc.xml(part)!, id) : null;
  if (!cm) return;
  if (resolved) cm.setAttribute("status", "resolved");
  else cm.removeAttribute("status");
  doc.touch(part!);
}

/** Delete a thread with its replies (legacy threads too). */
export function removeSlideThread(doc: PptxDocument, slide: string, id: string): void {
  if (id.startsWith("legacy:")) {
    const part = legacyPart(doc, slide);
    const list = part ? doc.xml(part) : null;
    for (const cm of Array.from(list?.getElementsByTagNameNS(NS.p, "cm") ?? [])) if (legacyKey(cm) === id || legacyParent(cm) === id) cm.remove();
    if (part) doc.touch(part);
    return;
  }
  const part = doc.slideCommentPart(slide, false);
  const cm = part ? findCm(doc.xml(part)!, id) : null;
  if (!cm) return;
  cm.remove();
  doc.touch(part!);
}
