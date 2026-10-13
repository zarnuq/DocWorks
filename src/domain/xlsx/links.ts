/**
 * Cell links (Google Sheets' Insert › Link): native hyperlinks in the sheet's <hyperlinks>. A web address is an
 * external relationship of the sheet; a place in the workbook ("#'Sheet 2'!C3") is the hyperlink's location.
 */
import type { XlsxWorkbook } from "../../infrastructure/xlsx/xlsx-workbook";
import { WORKSHEET_ORDER } from "./sheet-edit";

const S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const kids = (el: Element | null, name: string) => Array.from(el?.children ?? []).filter((c) => c.namespaceURI === S && c.localName === name);

/** The sheet's links by cell reference (a range's links are listed at its first cell): a web address, or "#" + location. */
export function linksOf(pkg: XlsxWorkbook, sheetPath: string): Map<string, string> {
  const out = new Map<string, string>();
  const root = pkg.xml(sheetPath)?.documentElement ?? null;
  const rels = new Map(pkg.rels(sheetPath).map((r) => [r.id, r.target]));
  for (const h of kids(kids(root, "hyperlinks")[0] ?? null, "hyperlink")) {
    const ref = (h.getAttribute("ref") ?? "").split(":")[0];
    const id = h.getAttributeNS(R, "id");
    const target = id ? rels.get(id) : h.hasAttribute("location") ? `#${h.getAttribute("location")}` : undefined;
    if (ref && target !== undefined) out.set(ref, target);
  }
  return out;
}

/** Link cell `ref` to `url` (a web address, or "#" + a place in the workbook); null removes its link. */
export function setLink(pkg: XlsxWorkbook, sheetPath: string, ref: string, url: string | null): void {
  const doc = pkg.xml(sheetPath)!;
  const root = doc.documentElement;
  let list = kids(root, "hyperlinks")[0] ?? null;
  for (const h of kids(list, "hyperlink")) {
    if ((h.getAttribute("ref") ?? "").split(":")[0] !== ref) continue;
    const id = h.getAttributeNS(R, "id");
    h.remove();
    if (id && !doc.querySelector(`[*|id="${id}"]`)) pkg.dropRel(sheetPath, id);
  }
  if (url) {
    if (!list) {
      list = doc.createElementNS(S, "hyperlinks");
      const after = WORKSHEET_ORDER.slice(WORKSHEET_ORDER.indexOf("hyperlinks") + 1);
      root.insertBefore(list, Array.from(root.children).find((c) => after.includes(c.localName)) ?? null);
    }
    const h = list.appendChild(doc.createElementNS(S, "hyperlink"));
    h.setAttribute("ref", ref);
    if (url.startsWith("#")) h.setAttribute("location", url.slice(1));
    else h.setAttributeNS(R, "r:id", pkg.addLinkRel(sheetPath, url));
  }
  if (list && !kids(list, "hyperlink").length) list.remove();
  pkg.touch(sheetPath);
}
