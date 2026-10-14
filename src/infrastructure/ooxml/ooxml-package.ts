import JSZip from "jszip";
import { NS, REL_TYPE } from "../../core/utils/xml";
import { IMAGE_MIME } from "../../core/utils/images";

export interface Rel {
  id: string;
  /** Last segment of the relationship type URL, e.g. "slideLayout", "image". */
  type: string;
  uri: string;
  /** Absolute zip path for internal targets, raw value for external ones. */
  target: string;
  external: boolean;
}

export interface ResourcePart { xml?: string; bytes?: Uint8Array; contentType: string }
export type ResourceBundle = Map<string, ResourcePart>;

export type PackageSnapshot = Map<string, string> & { binaries?: Map<string, Uint8Array> };

export const CONTENT_TYPES = "[Content_Types].xml";
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

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

  /**
   * Save only parts passed to `touch` (plus new and removed ones) instead of reserializing every part, so untouched
   * parts keep their original bytes. Formats whose editors mutate parts without calling `touch` leave this off.
   */
  protected dirtyOnly = false;
  private touched = new Set<string>();

  private binaryEdits = new Map<string, Uint8Array>();
  private managedBinary = new Set<string>();
  private originalBinary = new Map<string, JSZip.JSZipObject>();
  protected constructor(protected zip: JSZip) {
    for (const file of Object.values(zip.files)) if (!file.dir && !isXmlPart(file.name)) this.originalBinary.set(file.name, file);
  }

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

  /** Mark a part as changed, for packages that save only changed parts. */
  touch(path: string): void {
    this.touched.add(path);
  }

  /** Whether a part changed since the last save. */
  protected isTouched(path: string): boolean {
    return this.touched.has(path);
  }

  /** Remove a non-XML part (and its content type override) from the package. */
  protected dropBinary(path: string): void {
    this.binaryEdits.delete(path);
    this.managedBinary.delete(path);
    this.originalBinary.delete(path);
    this.zip.remove(path);
    this.removePart(path);
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
        uri: type,
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
    return [...new Set([...Object.keys(this.zip.files), ...this.binaryEdits.keys()])]
      .filter((p) => (!this.managedBinary.has(p) || this.binaryEdits.has(p) || this.originalBinary.has(p)) && /^[^/]+\/media\//.test(p) && (this.binaryEdits.has(p) || (!!this.zip.files[p] && !this.zip.files[p].dir)))
      .sort();
  }

  async media(path: string): Promise<Uint8Array> {
    if (this.binaryEdits.has(path)) return this.binaryEdits.get(path)!.slice();
    const file = this.managedBinary.has(path) ? this.originalBinary.get(path) : this.zip.file(path);
    if (!file) throw new Error(`Missing media part: ${path}`);
    return file.async("uint8array");
  }

  // ponytail: snapshots serialize every XML part; switch to per-part diffs if huge files feel slow.
  snapshot(): PackageSnapshot {
    const snap: PackageSnapshot = new Map();
    for (const [path, doc] of this.parts) snap.set(path, new XMLSerializer().serializeToString(doc));
    snap.binaries = new Map(this.binaryEdits);
    return snap;
  }

  restore(snap: PackageSnapshot): void {
    this.parts = new Map(Array.from(snap, ([path, text]) => [path, parseXml(text)!]));
    for (const path of this.parts.keys()) this.touch(path);
    this.binaryEdits = new Map(snap.binaries ?? []);
  }

  async save(): Promise<Uint8Array> {
    for (const path of this.parsedPaths) {
      if (!this.parts.has(path)) this.zip.remove(path);
    }
    for (const [path, doc] of this.parts) {
      if (!this.dirtyOnly || this.touched.has(path) || !this.parsedPaths.has(path)) this.zip.file(path, serialize(doc));
    }
    this.touched.clear();
    for (const path of this.managedBinary) {
      const edited = this.binaryEdits.get(path);
      const original = this.originalBinary.get(path);
      if (edited) this.zip.file(path, edited);
      else if (original) this.zip.file(path, await original.async("uint8array"));
      else this.zip.remove(path);
    }
    this.parsedPaths = new Set(this.parts.keys());
    return this.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }

  putXml(path: string, xml: Document, contentType?: string): void {
    this.parts.set(path, xml);
    this.touch(path);
    if (contentType) this.registerContentType(path, contentType);
  }
  putBinary(path: string, bytes: Uint8Array, contentType?: string): void {
    this.binaryEdits.set(path, bytes.slice());
    this.managedBinary.add(path);
    if (contentType) this.registerContentType(path, contentType);
  }
  uniquePath(prefix: string, suffix: string): string { return this.nextPartPath(prefix, suffix); }
  private contentType(path: string): string {
    const types = this.parts.get(CONTENT_TYPES)!;
    const override = Array.from(types.getElementsByTagNameNS(NS.contentTypes, "Override")).find(e => e.getAttribute("PartName") === "/" + path);
    const ext = path.slice(path.lastIndexOf(".") + 1);
    return override?.getAttribute("ContentType") ?? Array.from(types.getElementsByTagNameNS(NS.contentTypes, "Default")).find(e => e.getAttribute("Extension") === ext)?.getAttribute("ContentType") ?? "application/octet-stream";
  }
  private registerContentType(path: string, type: string): void {
    const existing = Array.from(this.parts.get(CONTENT_TYPES)!.getElementsByTagNameNS(NS.contentTypes, "Override")).find(e => e.getAttribute("PartName") === "/" + path);
    if (existing) existing.setAttribute("ContentType", type); else this.addOverride(path, type);
    this.touch(CONTENT_TYPES);
  }
  /** Capture the entire dependency graph now, so later source edits cannot alter a clipboard. */
  async captureResources(paths: string[]): Promise<ResourceBundle> {
    const bundle: ResourceBundle = new Map();
    const visit = async (path: string): Promise<void> => {
      if (bundle.has(path)) return;
      const xml = this.xml(path);
      const part: ResourcePart = { contentType: this.contentType(path) };
      if (xml) part.xml = new XMLSerializer().serializeToString(xml); else part.bytes = await this.media(path);
      bundle.set(path, part);
      const rels = this.xml(relsPathOf(path));
      if (rels) bundle.set(relsPathOf(path), { xml: new XMLSerializer().serializeToString(rels), contentType: "application/vnd.openxmlformats-package.relationships+xml" });
      for (const rel of this.rels(path)) if (!rel.external) await visit(rel.target);
    };
    for (const path of paths) await visit(path);
    return bundle;
  }
  importResources(bundle: ResourceBundle): Map<string, string> {
    const mapping = new Map<string, string>();
    for (const path of bundle.keys()) {
      if (path.endsWith(".rels")) continue;
      const dot = path.lastIndexOf(".");
      const prefix = path.slice(0, dot).replace(/\d+$/, "") + "Copy";
      const copied = this.nextPartPath(prefix, path.slice(dot));
      mapping.set(path, copied);
      const part = bundle.get(path)!;
      if (part.xml !== undefined) this.putXml(copied, parseXml(part.xml)!, part.contentType);
      else this.putBinary(copied, part.bytes!, part.contentType);
    }
    for (const [oldPath, newPath] of mapping) {
      const part = bundle.get(relsPathOf(oldPath));
      if (!part?.xml) continue;
      const rels = parseXml(part.xml)!;
      for (const rel of Array.from(rels.documentElement.children)) {
        if (rel.getAttribute("TargetMode") === "External") continue;
        const target = resolvePath(dirOf(oldPath), rel.getAttribute("Target")!);
        const copied = mapping.get(target);
        if (!copied) throw new Error("Missing copied resource: " + target);
        rel.setAttribute("Target", relativePath(dirOf(newPath), copied));
      }
      this.putXml(relsPathOf(newPath), rels);
    }
    return mapping;
  }

  // ---------------------------------------------------------------------------
  // Helpers for subclasses
  // ---------------------------------------------------------------------------

  /** e.g. ("ppt/slides/slide", ".xml") -> "ppt/slides/slide4.xml", unique across parts and zip. */
  protected nextPartPath(prefix: string, suffix: string): string {
    let max = 0;
    const re = new RegExp(`^${prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}(\\d+)\\.`);
    for (const p of [...this.parts.keys(), ...Object.keys(this.zip.files), ...this.binaryEdits.keys()]) {
      const m = re.exec(p);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return `${prefix}${max + 1}${suffix}`;
  }

  protected emptyRels(): Document {
    return parseXml(`<Relationships xmlns="${NS.pkgRels}"/>`)!;
  }

  /** A new relationship from `fromPath` (rels types without a scheme are officeDocument ones); returns its id. */
  relate(fromPath: string, type: string, targetPath: string): string {
    return this.addRel(fromPath, type, targetPath);
  }

  protected addRel(fromPath: string, type: string, targetPath: string, external = false): string {
    if (!this.parts.has(relsPathOf(fromPath))) this.parts.set(relsPathOf(fromPath), this.emptyRels());
    const relsDoc = this.parts.get(relsPathOf(fromPath))!;
    const existing = Array.from(relsDoc.getElementsByTagNameNS(NS.pkgRels, "Relationship"));
    const max = existing.reduce((m, r) => Math.max(m, Number(/^rId(\d+)$/.exec(r.getAttribute("Id") ?? "")?.[1] ?? 0)), 0);
    const id = `rId${max + 1}`;
    this.touch(relsPathOf(fromPath));
    const rel = relsDoc.createElementNS(NS.pkgRels, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", type.includes("://") ? type : REL_TYPE + type);
    rel.setAttribute("Target", external ? targetPath : relativePath(dirOf(fromPath), targetPath));
    if (external) rel.setAttribute("TargetMode", "External");
    relsDoc.documentElement.appendChild(rel);
    return id;
  }

  protected removeRel(fromPath: string, rId: string): void {
    const relsDoc = this.parts.get(relsPathOf(fromPath));
    this.touch(relsPathOf(fromPath));
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
    this.touch(CONTENT_TYPES);
  }

  protected removePart(path: string): void {
    this.parts.delete(path);
    this.parts.delete(relsPathOf(path));
    const types = this.parts.get(CONTENT_TYPES)!;
    Array.from(types.getElementsByTagNameNS(NS.contentTypes, "Override"))
      .find((o) => o.getAttribute("PartName") === `/${path}`)
      ?.remove();
    this.touch(CONTENT_TYPES);
  }

  /** Store image bytes as `<dir>/media/imageN.ext` with a content-type default. Returns the zip path. */
  protected storeImage(dir: string, bytes: Uint8Array, ext: string): string {
    ext = ext.toLowerCase();
    const path = this.nextPartPath(`${dir}/media/image`, `.${ext}`);
    this.putBinary(path, bytes, IMAGE_MIME[ext] ?? `image/${ext}`);

    const types = this.parts.get(CONTENT_TYPES)!;
    const hasDefault = Array.from(types.getElementsByTagNameNS(NS.contentTypes, "Default")).some(
      (d) => d.getAttribute("Extension")?.toLowerCase() === ext
    );
    if (!hasDefault) {
      const d = types.createElementNS(NS.contentTypes, "Default");
      d.setAttribute("Extension", ext);
      d.setAttribute("ContentType", IMAGE_MIME[ext] ?? `image/${ext}`);
      types.documentElement.insertBefore(d, types.documentElement.firstChild);
      this.touch(CONTENT_TYPES);
    }
    return path;
  }
}
