/*
 * Google Docs' citations (Tools › Citations) kept the way Word keeps them, so Word's References tab sees the same
 * sources: the sources live in the bibliography customXml part (b:Sources), each in-text citation is a CITATION
 * field in a citation content control (a "citation" chip, see chips.ts), and the list is a BIBLIOGRAPHY field in a
 * Bibliographies content control. Formatting follows APA 7, MLA 9 and Chicago author-date, Google's three styles.
 */

export const BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

export type CitationStyle = "apa" | "mla" | "chicago";
export type SourceType = "book" | "website" | "journal";
export interface Person {
  first: string;
  last: string;
}
export interface Source {
  /** Word's b:Tag, which CITATION fields name. */
  tag: string;
  type: SourceType;
  authors: Person[];
  title: string;
  year: string;
  /** Books. */
  publisher: string;
  /** Journal name (journal articles) or website name. */
  container: string;
  volume: string;
  issue: string;
  pages: string;
  url: string;
}

export const CITATION_STYLES: [CitationStyle, string][] = [["mla", "MLA (9th edition)"], ["apa", "APA (7th edition)"], ["chicago", "Chicago (author-date)"]];
export const SOURCE_TYPES: [SourceType, string][] = [["book", "Book"], ["website", "Website"], ["journal", "Journal article"]];

/** Word's names for our styles and source types. */
const WORD_STYLE: Record<CitationStyle, [file: string, name: string]> = {
  apa: ["\\APASixthEdition.xsl", "APA"],
  mla: ["\\MLASeventhEdition.xsl", "MLA"],
  chicago: ["\\Chicago.xsl", "Chicago"],
};
const WORD_TYPE: Record<SourceType, string> = { book: "Book", website: "InternetSite", journal: "JournalArticle" };
const FROM_WORD_TYPE: Record<string, SourceType> = { InternetSite: "website", DocumentFromInternetSite: "website", ElectronicSource: "website", JournalArticle: "journal", ArticleInAPeriodical: "journal" };

export const emptySource = (): Source => ({ tag: "", type: "book", authors: [], title: "", year: "", publisher: "", container: "", volume: "", issue: "", pages: "", url: "" });

// -----------------------------------------------------------------------------
// Formatting
// -----------------------------------------------------------------------------

export interface Segment {
  text: string;
  italic?: boolean;
}

const initials = (first: string) => first.split(/[\s-]+/).filter(Boolean).map((n) => `${n[0].toUpperCase()}.`).join(" ");
const lastFirst = (p: Person) => (p.first ? `${p.last}, ${p.first}` : p.last);
const firstLast = (p: Person) => (p.first ? `${p.first} ${p.last}` : p.last);
/** "a, b, and c" (`and` = "and" or "&"), APA's "a, b, & c" with the serial comma. */
function list(names: string[], and: string): string {
  if (names.length <= 1) return names.join("");
  if (names.length === 2) return `${names[0]}${and === "&" ? ", &" : ` ${and}`} ${names[1]}`;
  return `${names.slice(0, -1).join(", ")}, ${and} ${names[names.length - 1]}`;
}
/** Ends with a full stop unless it already ends with punctuation. */
const stop = (s: string) => (/[.?!]$/.test(s) ? s : `${s}.`);

/** The in-text citation: (Smith, 2020), (Smith), (Smith 2020). */
export function citationText(s: Source, style: CitationStyle): string {
  const last = s.authors.map((a) => a.last);
  const year = s.year || "n.d.";
  const quoted = s.type === "book" ? s.title : `"${s.title}"`;
  if (style === "apa") {
    const who = !last.length ? s.title : last.length === 1 ? last[0] : last.length === 2 ? `${last[0]} & ${last[1]}` : `${last[0]} et al.`;
    return `(${who}, ${year})`;
  }
  if (style === "mla") {
    const who = !last.length ? quoted : last.length === 1 ? last[0] : last.length === 2 ? `${last[0]} and ${last[1]}` : `${last[0]} et al.`;
    return `(${who})`;
  }
  const who = !last.length ? quoted : last.length <= 3 ? list(last, "and") : `${last[0]} et al.`;
  return `(${who} ${year})`;
}

/** The list's heading: References (APA), Works Cited (MLA), Bibliography (Chicago). */
export const bibliographyTitle = (style: CitationStyle) => ({ apa: "References", mla: "Works Cited", chicago: "Bibliography" })[style];

/** One entry of the list, titles of whole works in italics. */
export function bibliographyEntry(s: Source, style: CitationStyle): Segment[] {
  const out: Segment[] = [];
  const add = (text: string, italic = false) => text && out.push(italic ? { text, italic } : { text });
  const year = s.year || "n.d.";
  const whole = s.type === "book" || (style === "apa" && s.type === "website");
  if (style === "apa") {
    const names = s.authors.map((a) => (a.first ? `${a.last}, ${initials(a.first)}` : a.last));
    const authors = names.length === 1 ? names[0] : names.length > 1 ? `${names.slice(0, -1).join(", ")}, & ${names[names.length - 1]}` : "";
    const title = () => {
      if (whole) {
        add(s.title, true);
        add(". ");
      } else add(`${stop(s.title)} `);
    };
    if (authors) {
      add(`${authors} (${year}). `);
      title();
    } else {
      title();
      add(`(${year}). `);
    }
    if (s.type === "journal") {
      add(s.container, true);
      if (s.volume) {
        add(", ");
        add(s.volume, true);
      }
      if (s.issue) add(`(${s.issue})`);
      if (s.pages) add(`, ${s.pages}`);
      add(".");
    } else if (s.type === "book") add(s.publisher ? `${stop(s.publisher)}` : "");
    else add(s.container ? stop(s.container) : "");
    if (s.url) add(` ${s.url}`);
    return trim(out);
  }
  const [lead, ...others] = s.authors;
  const authors = !lead
    ? ""
    : style === "mla" && others.length > 1
      ? `${lastFirst(lead)}, et al.`
      : others.length === 1 && lead.first
        ? `${lastFirst(lead)}, and ${firstLast(others[0])}`
        : list([lastFirst(lead), ...others.map(firstLast)], "and");
  if (authors) add(`${stop(authors)} `);
  const title = () => {
    if (whole) {
      add(s.title, true);
      add(". ");
    } else add(`"${stop(s.title)}" `);
  };
  if (style === "chicago") {
    if (authors) {
      add(`${year}. `);
      title();
    } else {
      title();
      add(`${year}. `);
    }
    if (s.type === "journal") {
      add(s.container, true);
      if (s.volume) add(` ${s.volume}`);
      if (s.issue) add(` (${s.issue})`);
      add(s.pages ? `: ${s.pages}.` : ".");
    } else if (s.type === "book") add(s.publisher ? stop(s.publisher) : "");
    else {
      if (s.container) add(`${stop(s.container)}`);
      if (s.url) add(`${s.container ? " " : ""}${stop(s.url)}`);
    }
    return trim(out);
  }
  // MLA: author. title. container, details, year, location.
  title();
  if (s.type === "book") {
    const tail = [s.publisher, s.year].filter(Boolean).join(", ");
    if (tail) add(`${tail}.`);
  } else {
    const details: string[] = [];
    if (s.type === "journal" && s.volume) details.push(`vol. ${s.volume}`);
    if (s.type === "journal" && s.issue) details.push(`no. ${s.issue}`);
    if (s.year) details.push(s.year);
    if (s.type === "journal" && s.pages) details.push(`pp. ${s.pages}`);
    if (s.url && s.type === "website") details.push(s.url);
    if (s.container) {
      add(s.container, true);
      add(details.length ? `, ${details.join(", ")}.` : ".");
    } else if (details.length) add(`${details.join(", ")}.`);
  }
  return trim(out);
}

function trim(out: Segment[]): Segment[] {
  if (out.length) out[out.length - 1] = { ...out[out.length - 1], text: out[out.length - 1].text.replace(/\s+$/, "") };
  return out.filter((s) => s.text);
}

/** By first author's last name (or the title when there's none), then title, then year. */
export function sortSources(sources: Source[]): Source[] {
  const key = (s: Source) => `${(s.authors[0]?.last ?? s.title).toLowerCase()}\u0000${s.title.toLowerCase()}\u0000${s.year}`;
  return [...sources].sort((a, b) => key(a).localeCompare(key(b)));
}

/** One author per line, "First Last" or "Last, First"; one word is a last name (or an organisation). */
export function parseAuthors(text: string): Person[] {
  return text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const comma = line.indexOf(",");
      if (comma >= 0) return { first: line.slice(comma + 1).trim(), last: line.slice(0, comma).trim() };
      const space = line.lastIndexOf(" ");
      return space < 0 ? { first: "", last: line } : { first: line.slice(0, space).trim(), last: line.slice(space + 1) };
    });
}

/** Word's tag for a new source: three letters of the first author (or title) and the year's last two digits, unique. */
export function newTag(s: Source, taken: string[]): string {
  const name = (s.authors[0]?.last ?? s.title).replace(/[^\p{L}\p{N}]/gu, "").slice(0, 3) || "Src";
  const base = `${name}${s.year.slice(-2)}`;
  let tag = base;
  for (let n = 1; taken.includes(tag); n++) tag = `${base}${n}`;
  return tag;
}

// -----------------------------------------------------------------------------
// Word's sources part
// -----------------------------------------------------------------------------

const kid = (el: Element | null | undefined, local: string) => (el ? Array.from(el.children).find((c) => c.namespaceURI === BIB_NS && c.localName === local) ?? null : null);
const text = (el: Element | null | undefined, local: string) => kid(el, local)?.textContent?.trim() ?? "";

function readSource(el: Element): Source {
  const type = FROM_WORD_TYPE[text(el, "SourceType")] ?? "book";
  const author = kid(kid(el, "Author"), "Author");
  const corporate = text(author, "Corporate");
  const people = Array.from(kid(author, "NameList")?.children ?? [])
    .filter((p) => p.localName === "Person")
    .map((p) => ({ first: [text(p, "First"), text(p, "Middle")].filter(Boolean).join(" "), last: text(p, "Last") }));
  return {
    tag: text(el, "Tag"),
    type,
    authors: corporate ? [{ first: "", last: corporate }] : people,
    title: text(el, "Title"),
    year: text(el, "Year"),
    publisher: text(el, "Publisher"),
    container: type === "journal" ? text(el, "JournalName") || text(el, "PeriodicalTitle") : type === "website" ? text(el, "InternetSiteTitle") : "",
    volume: text(el, "Volume"),
    issue: text(el, "Issue"),
    pages: text(el, "Pages"),
    url: text(el, "URL"),
  };
}

/** The style and sources in Word's bibliography part (APA and none without one). */
export function readSources(xml: Document | null): { style: CitationStyle; sources: Source[] } {
  const root = xml?.documentElement;
  if (!root || root.namespaceURI !== BIB_NS) return { style: "apa", sources: [] };
  const name = (root.getAttribute("StyleName") ?? root.getAttribute("SelectedStyle") ?? "").toLowerCase();
  const style: CitationStyle = name.includes("mla") ? "mla" : name.includes("chicago") ? "chicago" : "apa";
  return { style, sources: Array.from(root.children).filter((c) => c.localName === "Source").map(readSource) };
}

/**
 * The sources part for `sources`: each source keeps its original element (fields we don't model, its Guid) with
 * the modelled fields rewritten; Word's element order inside b:Source is free.
 */
export function writeSources(original: Document | null, style: CitationStyle, sources: Source[]): string {
  const doc = new DOMParser().parseFromString(`<b:Sources xmlns:b="${BIB_NS}"/>`, "application/xml");
  const root = doc.documentElement;
  const [file, name] = WORD_STYLE[style];
  root.setAttribute("SelectedStyle", file);
  root.setAttribute("StyleName", name);
  root.setAttribute("Version", style === "apa" ? "6" : "7");
  const before = new Map(Array.from(original?.documentElement.children ?? []).filter((c) => c.localName === "Source").map((c) => [text(c, "Tag"), c]));
  const el = (local: string, value?: string) => {
    const e = doc.createElementNS(BIB_NS, `b:${local}`);
    if (value !== undefined) e.textContent = value;
    return e;
  };
  for (const s of sources) {
    const old = before.get(s.tag);
    const source = old ? (doc.importNode(old, true)) : el("Source");
    const modelled = ["Tag", "SourceType", "Author", "Title", "Year", "Publisher", "JournalName", "PeriodicalTitle", "InternetSiteTitle", "Volume", "Issue", "Pages", "URL"];
    for (const c of Array.from(source.children)) if (modelled.includes(c.localName)) c.remove();
    const fields: [string, string][] = [["Tag", s.tag], ["SourceType", WORD_TYPE[s.type]]];
    const head = fields.map(([k, v]) => el(k, v));
    if (s.authors.length) {
      const outer = el("Author");
      const inner = outer.appendChild(el("Author"));
      const names = inner.appendChild(el("NameList"));
      for (const p of s.authors) {
        const person = names.appendChild(el("Person"));
        person.appendChild(el("Last", p.last));
        if (p.first) person.appendChild(el("First", p.first));
      }
      head.push(outer);
    }
    const rest: [string, string][] = [["Title", s.title], ["Year", s.year], ["Publisher", s.type === "book" ? s.publisher : ""],
      [s.type === "journal" ? "JournalName" : "InternetSiteTitle", s.type === "book" ? "" : s.container],
      ["Volume", s.volume], ["Issue", s.issue], ["Pages", s.pages], ["URL", s.url]];
    for (const [k, v] of rest) if (v) head.push(el(k, v));
    source.prepend(...head);
    root.appendChild(source);
  }
  return new XMLSerializer().serializeToString(doc);
}
