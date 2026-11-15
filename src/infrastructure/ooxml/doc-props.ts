/*
 * Document properties (docProps/core.xml) shared by the three formats: title, subject, author, keywords,
 * description and category, plus the dates and last editor Office records. The part's children are an
 * unordered xsd:all, so edited fields are set in place and new ones appended.
 */
import JSZip from "jszip";
import { OoxmlPackage, parseXml } from "./ooxml-package";

const CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC = "http://purl.org/dc/elements/1.1/";
const DCTERMS = "http://purl.org/dc/terms/";
const XSI = "http://www.w3.org/2001/XMLSchema-instance";
const CORE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const CORE_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";

export const EDITABLE_PROPS = ["title", "subject", "creator", "keywords", "description", "category"] as const;
export type EditableProp = (typeof EDITABLE_PROPS)[number];
export type DocProps = Record<EditableProp | "lastModifiedBy" | "created" | "modified", string>;

const FIELDS: Record<keyof DocProps, [ns: string, qname: string]> = {
  title: [DC, "dc:title"],
  subject: [DC, "dc:subject"],
  creator: [DC, "dc:creator"],
  keywords: [CP, "cp:keywords"],
  description: [DC, "dc:description"],
  category: [CP, "cp:category"],
  lastModifiedBy: [CP, "cp:lastModifiedBy"],
  created: [DCTERMS, "dcterms:created"],
  modified: [DCTERMS, "dcterms:modified"],
};

const field = (doc: Document | null, key: keyof DocProps) => {
  const [ns, qname] = FIELDS[key];
  return doc?.getElementsByTagNameNS(ns, qname.split(":")[1])[0] ?? null;
};

function propsOf(doc: Document | null): DocProps {
  const out = {} as DocProps;
  for (const key of Object.keys(FIELDS) as (keyof DocProps)[]) out[key] = field(doc, key)?.textContent?.trim() ?? "";
  return out;
}

/** The properties in a core.xml text (for search, which reads parts without loading the package). */
export const corePropsOf = (xml: string): DocProps => propsOf(parseXml(xml));

/** The properties of a file's bytes, unzipping only core.xml (for search and companion notes); null without the part. */
export async function filePropsOf(bytes: ArrayBuffer | Uint8Array | JSZip): Promise<DocProps | null> {
  const zip = bytes instanceof JSZip ? bytes : await JSZip.loadAsync(bytes);
  const rels = await zip.file("_rels/.rels")?.async("text");
  const target = rels && /<Relationship\b[^>]*Type="[^"]*\/core-properties"[^>]*>/.exec(rels)?.[0].match(/Target="\/?([^"]+)"/)?.[1];
  const core = zip.file(target || "docProps/core.xml");
  return core ? corePropsOf(await core.async("text")) : null;
}

const corePath = (pkg: OoxmlPackage) => pkg.relOfType("", "core-properties") ?? "docProps/core.xml";

export const readProps = (pkg: OoxmlPackage): DocProps => propsOf(pkg.xml(corePath(pkg)));

function setField(doc: Document, key: keyof DocProps, value: string): boolean {
  const el = field(doc, key);
  const text = value.trim();
  if ((el?.textContent?.trim() ?? "") === text) return false;
  if (!text) {
    el?.remove();
    return true;
  }
  if (el) el.textContent = text;
  else {
    const [ns, qname] = FIELDS[key];
    const added = doc.createElementNS(ns, qname);
    if (ns === DCTERMS) added.setAttributeNS(XSI, "xsi:type", "dcterms:W3CDTF");
    added.textContent = text;
    doc.documentElement.appendChild(added);
  }
  return true;
}

const emptyCore = () =>
  parseXml(`<cp:coreProperties xmlns:cp="${CP}" xmlns:dc="${DC}" xmlns:dcterms="${DCTERMS}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="${XSI}"/>`)!;

/** Set the edited fields (empty = remove); the part is created when missing. Returns whether anything changed. */
export function writeProps(pkg: OoxmlPackage, patch: Partial<Record<EditableProp, string>>): boolean {
  const path = corePath(pkg);
  const existing = pkg.xml(path);
  const doc = existing ?? emptyCore();
  let changed = false;
  for (const key of EDITABLE_PROPS) if (patch[key] !== undefined && setField(doc, key, patch[key])) changed = true;
  if (!changed) return false;
  pkg.putXml(path, doc, existing ? undefined : CORE_TYPE);
  if (!existing) pkg.relate("", CORE_REL, path);
  return true;
}

const TEMPLATE_TYPES: [from: string, to: string][] = [
  ["wordprocessingml.template.main+xml", "wordprocessingml.document.main+xml"],
  ["presentationml.template.main+xml", "presentationml.presentation.main+xml"],
  ["spreadsheetml.template.main+xml", "spreadsheetml.sheet.main+xml"],
];

/** A new file from a template file (.docx / .dotx, .pptx / .potx, .xlsx / .xltx): a document's content type, created and modified now. */
export async function fromTemplate(bytes: Uint8Array, now = new Date()): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const typesFile = zip.file("[Content_Types].xml");
  if (typesFile) {
    let types = await typesFile.async("text");
    for (const [from, to] of TEMPLATE_TYPES) types = types.replace(from, to);
    zip.file("[Content_Types].xml", types);
  }
  const core = zip.file("docProps/core.xml");
  const doc = core && parseXml(await core.async("text"));
  if (doc) {
    const stamp = now.toISOString().replace(/\.\d+Z$/, "Z");
    setField(doc, "created", stamp);
    setField(doc, "modified", stamp);
    zip.file("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new XMLSerializer().serializeToString(doc));
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}
