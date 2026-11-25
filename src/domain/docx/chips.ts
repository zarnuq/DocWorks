/*
 * Smart chips (Google Docs' @ menu) as Word's own inline content controls, which is also what Google writes for them:
 * a date chip is a date picker (w:sdtPr/w:date), a dropdown chip a drop-down list (w:dropDownList, or a combo box
 * read the same way). File and people chips are plain links (see docx-view). The control's XML is kept: an
 * unchanged chip is written back as it was, a changed one only gets its value and its shown text rewritten.
 */
import { NS, parseElement } from "../../core/utils/xml";
import { citationText, type CitationStyle, type Source } from "./citations";

export type ChipKind = "date" | "dropdown" | "citation";
export interface ChipOption {
  text: string;
  value: string;
}
export interface ChipAttrs {
  kind: ChipKind;
  /** The original w:sdt ("" for a chip made here). */
  xml: string;
  /** What the chip shows. */
  text: string;
  /** Date chips: yyyy-mm-dd, or null while the control shows its placeholder. */
  date: string | null;
  /** Date chips: Word's date format (MMM d, yyyy...). */
  format: string;
  /** Dropdown chips: the items. */
  options: ChipOption[];
  /** Citation chips: the source's tag (its CITATION field names it). */
  tag: string;
}

/** Google's date formats, as Word format strings. */
export const DATE_FORMATS: [label: string, format: string][] = [
  ["Nov 21, 2026", "MMM d, yyyy"],
  ["November 21, 2026", "MMMM d, yyyy"],
  ["Friday, November 21, 2026", "dddd, MMMM d, yyyy"],
  ["11/21/2026", "M/d/yyyy"],
  ["21/11/2026", "d/M/yyyy"],
  ["2026-11-21", "yyyy-MM-dd"],
];
export const DEFAULT_DATE_FORMAT = "MMM d, yyyy";

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/** `iso` (yyyy-mm-dd) in a Word date format: d dd ddd dddd, M MM MMM MMMM, yy yyyy, 'quoted text'. */
export function formatDate(iso: string, format: string): string {
  const [y, m, d] = iso.split("-").map(Number);
  const day = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  const pad = (n: number) => String(n).padStart(2, "0");
  return format.replace(/'([^']*)'|d{1,4}|M{1,4}|y{2,4}/g, (tok, quoted?: string) => {
    if (quoted !== undefined) return quoted;
    switch (tok) {
      case "d": return String(d);
      case "dd": return pad(d);
      case "ddd": return DAYS[day].slice(0, 3);
      case "dddd": return DAYS[day];
      case "M": return String(m);
      case "MM": return pad(m);
      case "MMM": return MONTHS[m - 1].slice(0, 3);
      case "MMMM": return MONTHS[m - 1];
      case "yy": return pad(y % 100);
      default: return String(y);
    }
  });
}

/** Local date as yyyy-mm-dd, `days` from today. */
export function isoDay(days = 0, now = new Date()): string {
  const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() + days);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

const wKid = (el: Element | null | undefined, local: string) => (el ? Array.from(el.children).find((c) => c.namespaceURI === NS.w && c.localName === local) ?? null : null);
const wAttr = (el: Element | null, attr: string) => el?.getAttributeNS(NS.w, attr) ?? null;

/** A chip for an inline w:sdt holding a date picker or a dropdown with only runs inside it; null for anything else. */
export function chipAttrs(sdt: Element): ChipAttrs | null {
  const pr = wKid(sdt, "sdtPr"), content = wKid(sdt, "sdtContent");
  if (!pr || !content) return null;
  const date = wKid(pr, "date"), list = wKid(pr, "dropDownList") ?? wKid(pr, "comboBox"), citation = wKid(pr, "citation");
  if (!date && !list && !citation) return null;
  if (!Array.from(content.children).every((c) => c.namespaceURI === NS.w && ["r", "proofErr"].includes(c.localName))) return null;
  const text = Array.from(content.getElementsByTagNameNS(NS.w, "t")).map((t) => t.textContent).join("");
  const xml = new XMLSerializer().serializeToString(sdt);
  if (citation) {
    const instr = Array.from(content.getElementsByTagNameNS(NS.w, "instrText")).map((t) => t.textContent).join("");
    const tag = /^\s*CITATION\s+(\S+)/.exec(instr)?.[1];
    return tag ? { kind: "citation", xml, text, date: null, format: DEFAULT_DATE_FORMAT, options: [], tag } : null;
  }
  if (date) {
    const full = wAttr(date, "fullDate");
    return { kind: "date", xml, text, date: full && /^\d{4}-\d\d-\d\d/.test(full) ? full.slice(0, 10) : null, format: wAttr(wKid(date, "dateFormat"), "val") ?? DEFAULT_DATE_FORMAT, options: [], tag: "" };
  }
  const options = Array.from(list!.children)
    .filter((c) => c.localName === "listItem")
    .map((c) => {
      const shown = wAttr(c, "displayText") ?? wAttr(c, "value") ?? "";
      return { text: shown, value: wAttr(c, "value") ?? shown };
    });
  return { kind: "dropdown", xml, text, date: null, format: DEFAULT_DATE_FORMAT, options, tag: "" };
}

/** Attributes of a chip made here (`xml` empty: the writer builds the control). */
export function newChip(kind: "date" | "dropdown", opts: { date?: string; format?: string; options?: string[] } = {}): ChipAttrs {
  if (kind === "date") {
    const date = opts.date ?? isoDay();
    const format = opts.format ?? DEFAULT_DATE_FORMAT;
    return { kind, xml: "", text: formatDate(date, format), date, format, options: [], tag: "" };
  }
  const options = dropdownOptions(opts.options ?? []);
  return { kind: "dropdown", xml: "", text: options[0]?.text ?? "", date: null, format: DEFAULT_DATE_FORMAT, options, tag: "" };
}

/** An in-text citation of `source`. */
export function citationChip(source: Source, style: CitationStyle): ChipAttrs {
  return { kind: "citation", xml: "", text: citationText(source, style), date: null, format: DEFAULT_DATE_FORMAT, options: [], tag: source.tag };
}

/** Item names, trimmed, blanks and repeats dropped; each is its own value. */
export function dropdownOptions(names: string[]): ChipOption[] {
  const seen = new Set<string>();
  return names.map((n) => n.trim()).filter((n) => n && !seen.has(n) && seen.add(n)).map((n) => ({ text: n, value: n }));
}

/** Whether `a` still says what its original XML says (then the XML is written back as it was). */
export function chipUnchanged(a: ChipAttrs): boolean {
  if (!a.xml) return false;
  const was = chipAttrs(parseElement(a.xml));
  return !!was && was.text === a.text && was.tag === a.tag && was.date === a.date && was.format === a.format && JSON.stringify(was.options) === JSON.stringify(a.options);
}

/** sdtPr children before the control type (ECMA-376 CT_SdtPr). */
const BEFORE_TYPE = new Set(["rPr", "alias", "tag", "id", "lock", "placeholder", "temporary", "showingPlcHdr", "dataBinding", "label", "tabIndex"]);

/**
 * The w:sdt for chip `a`, in `owner`: the original control with its value and shown text rewritten, or a new one.
 * `run` makes the run holding the text (the writer gives it the chip's formatting).
 */
export function chipElement(owner: Document, a: ChipAttrs, run: (text: string) => Element): Element {
  const el = (local: string) => owner.createElementNS(NS.w, `w:${local}`);
  const set = (e: Element, attr: string, v: string) => e.setAttributeNS(NS.w, `w:${attr}`, v);
  const sdt = a.xml
    ? owner.importNode(parseElement(a.xml), true)
    : owner.importNode(parseElement(`<w:sdt xmlns:w="${NS.w}"><w:sdtPr><w:id w:val="${1 + Math.floor(Math.random() * 0x7ffffffe)}"/></w:sdtPr><w:sdtContent/></w:sdt>`), true);
  const pr = wKid(sdt, "sdtPr")!;
  /** The control type element, made in place (after the common properties) when missing. */
  const control = (local: string): Element => {
    const found = wKid(pr, local);
    if (found) return found;
    const made = el(local);
    pr.insertBefore(made, Array.from(pr.children).find((c) => !BEFORE_TYPE.has(c.localName)) ?? null);
    return made;
  };
  const content = wKid(sdt, "sdtContent") ?? sdt.appendChild(el("sdtContent"));
  if (a.kind === "citation") {
    control("citation");
    // The CITATION field, its cached result the citation's text; every run formatted as the text.
    const result = run(a.text);
    const rPr = wKid(result, "rPr");
    const fieldRun = (child: Element) => {
      const r = el("r");
      if (rPr) r.appendChild(rPr.cloneNode(true));
      r.appendChild(child);
      return r;
    };
    const fld = (type: string) => {
      const f = el("fldChar");
      set(f, "fldCharType", type);
      return fieldRun(f);
    };
    const instr = el("instrText");
    instr.setAttributeNS(NS.xml, "xml:space", "preserve");
    const was = Array.from(content.getElementsByTagNameNS(NS.w, "instrText")).map((t) => t.textContent).join("");
    instr.textContent = new RegExp(`^\\s*CITATION\\s+${a.tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(\\s|$)`).test(was) ? was : ` CITATION ${a.tag} \\l 1033 `;
    content.replaceChildren(fld("begin"), fieldRun(instr), fld("separate"), result, fld("end"));
    return sdt;
  }
  if (a.kind === "date") {
    const date = control("date");
    if (a.date) {
      set(date, "fullDate", `${a.date}T00:00:00Z`);
      wKid(pr, "showingPlcHdr")?.remove();
    } else date.removeAttributeNS(NS.w, "fullDate");
    const format = wKid(date, "dateFormat");
    if (format) set(format, "val", a.format);
    if (!date.children.length) {
      // Word's own new date picker.
      for (const [local, v] of [["dateFormat", a.format], ["lid", "en-US"], ["storeMappedDataAs", "dateTime"], ["calendar", "gregorian"]]) set(date.appendChild(el(local)), "val", v);
    }
  } else {
    const list = wKid(pr, "dropDownList") ?? wKid(pr, "comboBox") ?? control("dropDownList");
    for (const item of Array.from(list.children).filter((c) => c.localName === "listItem")) item.remove();
    for (const o of a.options) {
      const item = list.appendChild(el("listItem"));
      set(item, "displayText", o.text);
      set(item, "value", o.value);
    }
    const chosen = a.options.find((o) => o.text === a.text);
    if (chosen) set(list, "lastValue", chosen.value);
    else list.removeAttributeNS(NS.w, "lastValue");
    if (chosen) wKid(pr, "showingPlcHdr")?.remove();
  }
  content.replaceChildren(run(a.text));
  return sdt;
}
