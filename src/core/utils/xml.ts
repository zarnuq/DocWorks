export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  xml: "http://www.w3.org/XML/1998/namespace",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  dgm: "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  dsp: "http://schemas.microsoft.com/office/drawing/2008/diagram",
  pkgRels: "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

export const REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

export function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/** Child elements named `local` (all of them without it), in any namespace or only `ns`. */
export const kids = (el: Element | null | undefined, local?: string, ns?: string): Element[] =>
  el ? Array.from(el.children).filter((c) => (!local || c.localName === local) && (!ns || c.namespaceURI === ns)) : [];

/** The first child element named `local`, in any namespace or only `ns`. */
export function kid(el: Element | null | undefined, local: string, ns?: string): Element | null {
  if (!el) return null;
  for (let i = 0; i < el.children.length; i++) {
    const c = el.children[i];
    if (c.localName === local && (!ns || c.namespaceURI === ns)) return c;
  }
  return null;
}

/** English Metric Units: OOXML's lengths. */
export const EMU_PER_PX = 9525;
export const EMU_PER_PT = 12700;

/** A new GUID in braces, as Office writes ids. */
export const guid = () => `{${crypto.randomUUID().toUpperCase()}}`;

/** Parse an XML string and return its root element. */
export const parseElement = (xml: string): Element => new DOMParser().parseFromString(xml, "application/xml").documentElement;

/** Insert `el` into `parent` before the first child that comes later in schema `order` (ranked by `key`, default the local name). */
export function insertInOrder(parent: Element, el: Element, order: readonly string[], key: (e: Element) => string = (e) => e.localName): Element {
  const at = order.indexOf(key(el));
  return parent.insertBefore(el, Array.from(parent.children).find((c) => order.indexOf(key(c)) > at) ?? null);
}
