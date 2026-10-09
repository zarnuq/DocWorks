import type { DOMOutputSpec } from "prosemirror-model";

/**
 * A read-only rendering of an Office Math (OMML) equation: text with sub/superscripts,
 * fractions, n-ary operators, delimiters and roots. The XML itself is kept untouched.
 */
// ponytail: approximate layout (no stacked fractions or limits); a MathML conversion if equations matter more.
export function mathSpec(xml: string): DOMOutputSpec {
  const root = new DOMParser().parseFromString(xml, "application/xml").documentElement;
  return ["span", { class: "docx-math" }, ...convert(root)];
}

type Spec = DOMOutputSpec | string;

const kid = (el: Element, local: string) => Array.from(el.children).find((c) => c.localName === local) ?? null;
const prop = (el: Element, pr: string, local: string) => {
  const p = kid(el, pr);
  const v = p && kid(p, local);
  return v?.getAttributeNS(v.namespaceURI, "val") ?? v?.getAttribute("m:val") ?? null;
};

function convert(el: Element | null): Spec[] {
  if (!el) return [];
  const all = (e: Element | null) => (e ? Array.from(e.children).flatMap((c) => convert(c)) : []);
  switch (el.localName) {
    case "t":
      return [el.textContent ?? ""];
    case "sSub":
      return [...all(kid(el, "e")), ["sub", ...all(kid(el, "sub"))]];
    case "sSup":
      return [...all(kid(el, "e")), ["sup", ...all(kid(el, "sup"))]];
    case "sSubSup":
      return [...all(kid(el, "e")), ["sub", ...all(kid(el, "sub"))], ["sup", ...all(kid(el, "sup"))]];
    case "f":
      return ["(", ...all(kid(el, "num")), ")/(", ...all(kid(el, "den")), ")"];
    case "nary":
      return [prop(el, "naryPr", "chr") ?? "∫", ["sub", ...all(kid(el, "sub"))], ["sup", ...all(kid(el, "sup"))], " ", ...all(kid(el, "e"))];
    case "d":
      return [prop(el, "dPr", "begChr") ?? "(", ...Array.from(el.children).filter((c) => c.localName === "e").flatMap((e) => all(e)), prop(el, "dPr", "endChr") ?? ")"];
    case "rad":
      return ["√(", ...all(kid(el, "e")), ")"];
    case "rPr":
    case "ctrlPr":
    case "sSubPr":
    case "sSupPr":
    case "naryPr":
    case "dPr":
    case "fPr":
    case "radPr":
    case "oMathParaPr":
      return [];
    default:
      return all(el);
  }
}
