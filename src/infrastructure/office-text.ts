/**
 * The text of a .docx / .pptx / .xlsx for searching across the vault (Obsidian's own search skips them): only the
 * parts holding text are unzipped and read, nothing is laid out. Each piece is labelled with where it is
 * ("Slide 3", "Budget!B4"; documents have no label).
 */
import JSZip from "jszip";
import { NS } from "../core/utils/xml";
import { SearchOptions, textMatches } from "../core/utils/text-matches";
import { parseXml, resolvePath } from "./ooxml/ooxml-package";

export interface TextPart {
  label: string;
  text: string;
}

const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function xml(zip: JSZip, path: string): Promise<Document | null> {
  const file = zip.file(path.replace(/^\//, ""));
  return file ? parseXml(await file.async("text")) : null;
}

/** Targets of a part's relationships by id. */
async function rels(zip: JSZip, part: string): Promise<Map<string, string>> {
  const dir = part.slice(0, part.lastIndexOf("/"));
  const doc = await xml(zip, `${dir}/_rels/${part.slice(dir.length + 1)}.rels`);
  const out = new Map<string, string>();
  for (const r of Array.from(doc?.getElementsByTagName("Relationship") ?? [])) {
    const target = r.getAttribute("Target") ?? "";
    out.set(r.getAttribute("Id") ?? "", target.startsWith("/") ? target.slice(1) : resolvePath(dir, target));
  }
  return out;
}

/** Paragraph texts: each a:t / w:t goes to its nearest paragraph, so text boxes inside a paragraph aren't read twice. */
function paragraphs(doc: Document, ns: string): string[] {
  const texts = new Map<Element, string>();
  for (const t of Array.from(doc.getElementsByTagNameNS(ns, "t"))) {
    let p: Element | null = t.parentElement;
    while (p && !(p.localName === "p" && p.namespaceURI === ns)) p = p.parentElement;
    if (p) texts.set(p, (texts.get(p) ?? "") + (t.textContent ?? ""));
  }
  return [...texts.values()].filter((t) => t.trim());
}

export async function officeText(ext: "docx" | "pptx" | "xlsx", bytes: ArrayBuffer | Uint8Array): Promise<TextPart[]> {
  const zip = await JSZip.loadAsync(bytes);
  if (ext === "docx") {
    const doc = await xml(zip, "word/document.xml");
    return doc ? paragraphs(doc, NS.w).map((text) => ({ label: "", text })) : [];
  }
  if (ext === "pptx") {
    const pres = await xml(zip, "ppt/presentation.xml");
    const targets = await rels(zip, "ppt/presentation.xml");
    const out: TextPart[] = [];
    let n = 0;
    for (const id of Array.from(pres?.getElementsByTagNameNS(NS.p, "sldId") ?? [])) {
      const slide = await xml(zip, targets.get(id.getAttributeNS(R, "id") ?? "") ?? "");
      n++;
      if (slide) out.push(...paragraphs(slide, NS.a).map((text) => ({ label: `Slide ${n}`, text })));
    }
    return out;
  }
  const book = await xml(zip, "xl/workbook.xml");
  const targets = await rels(zip, "xl/workbook.xml");
  const sst = await xml(zip, "xl/sharedStrings.xml");
  const shared = Array.from(sst?.getElementsByTagName("si") ?? [], (si) => Array.from(si.getElementsByTagName("t"), (t) => t.textContent ?? "").join(""));
  const out: TextPart[] = [];
  for (const sheet of Array.from(book?.getElementsByTagName("sheet") ?? [])) {
    const doc = await xml(zip, targets.get(sheet.getAttributeNS(R, "id") ?? "") ?? "");
    const name = sheet.getAttribute("name") ?? "";
    for (const c of Array.from(doc?.getElementsByTagName("c") ?? [])) {
      const v = c.getElementsByTagName("v")[0]?.textContent ?? "";
      const type = c.getAttribute("t");
      const text = type === "s" ? shared[Number(v)] ?? "" : type === "inlineStr" ? Array.from(c.getElementsByTagName("t"), (t) => t.textContent ?? "").join("") : v;
      if (text.trim()) out.push({ label: `${name}!${c.getAttribute("r") ?? ""}`, text });
    }
  }
  return out;
}

export interface TextHit {
  label: string;
  before: string;
  match: string;
  after: string;
}

/** Every match in `parts`, with up to 40 characters before and 60 after (cut with "…"). */
export function searchText(parts: TextPart[], query: string, opts: SearchOptions): TextHit[] {
  if (!query) return [];
  return parts.flatMap(({ label, text }) =>
    textMatches(text, query, opts).map((m) => ({
      label,
      before: m.start > 40 ? `…${text.slice(m.start - 40, m.start)}` : text.slice(0, m.start),
      match: text.slice(m.start, m.end),
      after: text.length - m.end > 60 ? `${text.slice(m.end, m.end + 60)}…` : text.slice(m.end),
    }))
  );
}
