import JSZip from "jszip";
import { NS, REL_TYPE } from "../../core/utils/xml";

export interface Rel {
  id: string;
  /** Last segment of the relationship type URL, e.g. "slideLayout", "image". */
  type: string;
  /** Absolute zip path for internal targets, raw value for external ones. */
  target: string;
  external: boolean;
}

export type PackageSnapshot = Map<string, string>;

export const CONTENT_TYPES = "[Content_Types].xml";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const IMAGE_CT: Record<string, string> = { jpg: "image/jpeg", svg: "image/svg+xml" };

const isXmlPart = (path: string) => path.endsWith(".xml") || path.endsWith(".rels");
export const dirOf = (path: string) => path.slice(0, path.lastIndexOf("/") + 1);
export const relsPathOf = (path: string) => `${dirOf(path)}_rels/${path.slice(path.lastIndexOf("/") + 1)}.rels`;

/** Resolve "../media/x.png" against the directory of the source part. */
export function resolvePath(fromDir: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const out = fromDir.split("/").filter(Boolean);
  for (const seg of target.split("/")) {
    if (seg === "..") out.pop();
    else if (seg && seg !== ".") out.push(seg);
  }
  return out.join("/");
}

/** Inverse of resolvePath: target path written relative to fromDir. */
export function relativePath(fromDir: string, to: string): string {
  const from = fromDir.split("/").filter(Boolean);
  const dest = to.split("/");
  let i = 0;
  while (i < from.length && i < dest.length - 1 && from[i] === dest[i]) i++;
  return "../".repeat(from.length - i) + dest.slice(i).join("/");
}

export function parseXml(text: string): Document | null {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  return doc.getElementsByTagName("parsererror").length ? null : doc;
}

const serialize = (doc: Document) =>
  XML_DECL + new XMLSerializer().serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, "");

/**
 * An OPC zip (docx, pptx, ...) whose XML parts are all kept as live DOMs, so
 * edits pass through everything the editors don't understand. Media stays in
 * the zip and is read on demand.
 */
export class OoxmlPackage {
  protected parts = new Map<string, Document>();
  /** XML parts that came from the zip and parsed; removed from the zip on save if deleted. */
  private parsedPaths = new Set<string>();

  protected constructor(protected zip: JSZip) {}

  /** Parse every XML part. Unparseable parts stay in the zip and are written back verbatim. */
  protected async readParts(): Promise<void> {
    for (const file of Object.values(this.zip.files)) {
      if (file.dir || !isXmlPart(file.name)) continue;
      const parsed = parseXml(await file.async("text"));
      if (parsed) {
        this.parts.set(file.name, parsed);
        this.parsedPaths.add(file.name);
      }
    }
  }

  xml(path: string): Document | null {
    return this.parts.get(path) ?? null;
  }

  rels(path: string): Rel[] {
    const relsDoc = this.parts.get(relsPathOf(path));
    if (!relsDoc) return [];
    return Array.from(relsDoc.getElementsByTagNameNS(NS.pkgRels, "Relationship")).map((r) => {
      const external = r.getAttribute("TargetMode") === "External";
      const raw = r.getAttribute("Target") ?? "";
      const type = r.getAttribute("Type") ?? "";
      return {
        id: r.getAttribute("Id") ?? "",
        type: type.slice(type.lastIndexOf("/") + 1),
        target: external ? raw : resolvePath(dirOf(path), raw),
        external,
      };
    });
  }

  relTarget(path: string, rId: string): string | null {
    return this.rels(path).find((r) => r.id === rId)?.target ?? null;
  }

  relOfType(path: string, type: string): string | null {
    return this.rels(path).find((r) => r.type === type)?.target ?? null;
  }

  mediaPaths(): string[] {
    return Object.keys(this.zip.files)
      .filter((p) => /^[^/]+\/media\//.test(p) && !this.zip.files[p].dir)
      .sort();
  }

  async media(path: string): Promise<Uint8Array> {
    const file = this.zip.file(path);
    if (!file) throw new Error(`Missing media part: ${path}`);
    return file.async("uint8array");
  }

  // ponytail: snapshots serialize every XML part; switch to per-part diffs if huge files feel slow.
  snapshot(): PackageSnapshot {
    const snap: PackageSnapshot = new Map();
    for (const [path, doc] of this.parts) snap.set(path, new XMLSerializer().serializeToString(doc));
    return snap;
  }

  restore(snap: PackageSnapshot): void {
    this.parts = new Map(Array.from(snap, ([path, text]) => [path, parseXml(text)!]));
  }

  async save(): Promise<Uint8Array> {
    for (const path of this.parsedPaths) {
      if (!this.parts.has(path)) this.zip.remove(path);
    }
    for (const [path, doc] of this.parts) this.zip.file(path, serialize(doc));
    this.parsedPaths = new Set(this.parts.keys());
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  // ---------------------------------------------------------------------------
  // Helpers for subclasses
  // ---------------------------------------------------------------------------

  /** e.g. ("ppt/slides/slide", ".xml") -> "ppt/slides/slide4.xml", unique across parts and zip. */
  protected nextPartPath(prefix: string, suffix: string): string {
    let max = 0;
    const re = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(\\d+)\\.`);
    for (const p of [...this.parts.keys(), ...Object.keys(this.zip.files)]) {
      const m = re.exec(p);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return `${prefix}${max + 1}${suffix}`;
  }

  protected emptyRels(): Document {
    return parseXml(`<Relationships xmlns="${NS.pkgRels}"/>`)!;
  }

  protected addRel(fromPath: string, type: string, targetPath: string, external = false): string {
    if (!this.parts.has(relsPathOf(fromPath))) this.parts.set(relsPathOf(fromPath), this.emptyRels());
    const relsDoc = this.parts.get(relsPathOf(fromPath))!;
    const existing = Array.from(relsDoc.getElementsByTagNameNS(NS.pkgRels, "Relationship"));
    const max = existing.reduce((m, r) => Math.max(m, Number(/^rId(\d+)$/.exec(r.getAttribute("Id") ?? "")?.[1] ?? 0)), 0);
    const id = `rId${max + 1}`;
    const rel = relsDoc.createElementNS(NS.pkgRels, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", REL_TYPE + type);
    rel.setAttribute("Target", external ? targetPath : relativePath(dirOf(fromPath), targetPath));
    if (external) rel.setAttribute("TargetMode", "External");
    relsDoc.documentElement.appendChild(rel);
    return id;
  }

  protected removeRel(fromPath: string, rId: string): void {
    const relsDoc = this.parts.get(relsPathOf(fromPath));
    Array.from(relsDoc?.getElementsByTagNameNS(NS.pkgRels, "Relationship") ?? [])
      .find((r) => r.getAttribute("Id") === rId)
      ?.remove();
  }

  /** Register a new part's content type in [Content_Types].xml. */
  protected addOverride(path: string, contentType: string): void {
    const types = this.parts.get(CONTENT_TYPES)!;
    const override = types.createElementNS(NS.contentTypes, "Override");
    override.setAttribute("PartName", `/${path}`);
    override.setAttribute("ContentType", contentType);
    types.documentElement.appendChild(override);
  }

  protected removePart(path: string): void {
    this.parts.delete(path);
    this.parts.delete(relsPathOf(path));
    const types = this.parts.get(CONTENT_TYPES)!;
    Array.from(types.getElementsByTagNameNS(NS.contentTypes, "Override"))
      .find((o) => o.getAttribute("PartName") === `/${path}`)
      ?.remove();
  }

  /** Store image bytes as `<dir>/media/imageN.ext` with a content-type default. Returns the zip path. */
  protected storeImage(dir: string, bytes: Uint8Array, ext: string): string {
    ext = ext.toLowerCase();
    const path = this.nextPartPath(`${dir}/media/image`, `.${ext}`);
    this.zip.file(path, bytes);

    const types = this.parts.get(CONTENT_TYPES)!;
    const hasDefault = Array.from(types.getElementsByTagNameNS(NS.contentTypes, "Default")).some(
      (d) => d.getAttribute("Extension")?.toLowerCase() === ext
    );
    if (!hasDefault) {
      const d = types.createElementNS(NS.contentTypes, "Default");
      d.setAttribute("Extension", ext);
      d.setAttribute("ContentType", IMAGE_CT[ext] ?? `image/${ext}`);
      types.documentElement.insertBefore(d, types.documentElement.firstChild);
    }
    return path;
  }
}
