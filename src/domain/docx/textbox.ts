/** A Word text box's content (w:txbxContent) as an editor document, and edits written back into its run XML. */
import { Node as PMNode } from "prosemirror-model";
import { NS } from "../../core/utils/xml";
import { Origins, ReadContext, readDocx } from "./docx-reader";
import { WriteContext, writeBody } from "./docx-writer";

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml").documentElement;

export function readTextBox(xml: string, ctx: ReadContext): { doc: PMNode; origins: Origins } | null {
  if (!xml.includes("txbxContent")) return null;
  const content = parse(xml).getElementsByTagNameNS(NS.w, "txbxContent")[0];
  return content ? readDocx(content, ctx) : null;
}

/** The run XML with the text box's content replaced by `doc` (in the drawing and in its VML fallback). */
export function writeTextBox(xml: string, doc: PMNode, origins: Origins, ctx: WriteContext): string {
  const root = parse(xml);
  const [main, ...copies] = Array.from(root.getElementsByTagNameNS(NS.w, "txbxContent"));
  writeBody(main, doc, origins, ctx);
  for (const copy of copies) copy.replaceChildren(...Array.from(main.children).map((c) => c.cloneNode(true)));
  return new XMLSerializer().serializeToString(root);
}
