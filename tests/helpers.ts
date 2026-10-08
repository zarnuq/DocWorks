import { readFileSync } from "fs";
import { join } from "path";
import { PptxDocument } from "../src/infrastructure/pptx/pptx-document";

export const fixture = (name: string) => new Uint8Array(readFileSync(join(__dirname, "fixtures", name)));

export const loadSample = () => PptxDocument.load(fixture("sample.pptx"));

/** Save and re-open, proving the edit survives the zip round trip. */
export async function roundTrip(doc: PptxDocument): Promise<PptxDocument> {
  return PptxDocument.load(await doc.save());
}

export const texts = (el: Element | Document) =>
  Array.from(el.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "t")).map((t) => t.textContent);

// ---------------------------------------------------------------------------
// docx
// ---------------------------------------------------------------------------
import { DocxDocument } from "../src/infrastructure/docx/docx-document";
import { readDocx } from "../src/domain/docx/docx-reader";
import type { Node as PMNode } from "prosemirror-model";

export const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const serializeXml = (n: Node) => new XMLSerializer().serializeToString(n);

export async function readSample(name = "sample.docx") {
  const docx = await DocxDocument.load(fixture(name));
  const urls = await docx.imageUrls();
  const read = readDocx(docx.body, { imageSrc: (id) => urls.get(id) ?? null, hyperlinkHref: (id) => docx.hyperlinkHref(id) });
  return { docx, ...read };
}

/** Document position just before top-level block `i`. */
export function blockPos(doc: PMNode, i: number): number {
  let pos = 0;
  for (let k = 0; k < i; k++) pos += doc.child(k).nodeSize;
  return pos;
}

/** Direct w:* children of the body, excluding the final sectPr. */
export const bodyBlocks = (body: Element) => Array.from(body.children).filter((c) => c.localName !== "sectPr");
