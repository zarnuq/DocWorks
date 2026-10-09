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

export function childNS(el: Element, ns: string, local: string): Element | null {
  for (let i = 0; i < el.children.length; i++) {
    if (el.children[i].localName === local && el.children[i].namespaceURI === ns) {
      return el.children[i];
    }
  }
  return null;
}
