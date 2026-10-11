import { NS } from "../../core/utils/xml";
import { parseXml } from "../../infrastructure/ooxml/ooxml-package";
import type { PptxDocument } from "../../infrastructure/pptx/pptx-document";
import { addPicture, shapeId } from "./slide-edit";
export const MEDIA_MIME: Record<string, string> = { mp4: "video/mp4", webm: "video/webm", ogv: "video/ogg", mp3: "audio/mpeg", wav: "audio/wav", ogg: "audio/ogg", m4a: "audio/mp4", oga: "audio/ogg" };
const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
export interface SlideMedia {
  kind: "audio" | "video";
  path: string;
  autoPlay: boolean;
  loop: boolean;
  external: boolean;
}
const timingNode = (slide: Document, id: string) => Array.from(slide.getElementsByTagNameNS(NS.p, "cMediaNode")).find(e => e.getElementsByTagNameNS(NS.p, "spTgt")[0]?.getAttribute("spid") === id);
export function readMedia(doc: PptxDocument, part: string, shape: Element): SlideMedia | null {
  const file = shape.getElementsByTagNameNS(NS.a, "videoFile")[0] ?? shape.getElementsByTagNameNS(NS.a, "audioFile")[0];
  if (!file)
    return null;
  const rid = shape.getElementsByTagNameNS(P14, "media")[0]?.getAttributeNS(NS.r, "embed") ?? file.getAttributeNS(NS.r, "link");
  const rel = doc.rels(part).find(r => r.id === rid);
  if (!rel)
    return null;
  const node = timingNode(doc.xml(part)!, shapeId(shape) ?? "");
  const ct = node?.getElementsByTagNameNS(NS.p, "cTn")[0];
  const condition = ct?.getElementsByTagNameNS(NS.p, "cond")[0];
  return { kind: file.localName === "videoFile" ? "video" : "audio", path: rel.target, external: rel.external, autoPlay: !!condition && condition.getAttribute("delay") !== "indefinite" && condition.getAttribute("evt") !== "onClick", loop: ct?.getAttribute("repeatCount") === "indefinite" };
}
export function addMedia(doc: PptxDocument, part: string, bytes: Uint8Array, kind: "audio" | "video", ext: string, name: string, poster?: Uint8Array): Element {
  if (!MEDIA_MIME[ext] || !MEDIA_MIME[ext].startsWith(kind + "/"))
    throw new Error("Choose a supported audio or video file");
  const path = doc.uniquePath("ppt/media/media", "." + ext);
  doc.putBinary(path, bytes, MEDIA_MIME[ext]);
  const link = doc.addSlideRel(part, kind, path);
  const embedded = doc.addSlideRel(part, "http://schemas.microsoft.com/office/2007/relationships/media", path);
  const png = poster ?? Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="), c => c.charCodeAt(0));
  const image = doc.addImage(part, png, "png");
  const { cx, cy } = doc.slideSize;
  const rect = kind === "video" ? { x: cx * .2, y: cy * .2, w: cx * .6, h: cy * .6 } : { x: cx * .4, y: cy * .4, w: cx * .2, h: cy * .2 };
  const shape = addPicture(doc.xml(part)!, rect, image, name);
  shape.getElementsByTagNameNS(NS.p,"cNvPr")[0].setAttribute("name",name);
  const nv = shape.getElementsByTagNameNS(NS.p, "nvPr")[0];
  nv.append(doc.xml(part)!.importNode(parseXml(`<a:${kind}File xmlns:a="${NS.a}" xmlns:r="${NS.r}" r:link="${link}"/>`)!.documentElement, true));
  nv.append(doc.xml(part)!.importNode(parseXml(`<p:extLst xmlns:p="${NS.p}" xmlns:p14="${P14}" xmlns:r="${NS.r}"><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media r:embed="${embedded}"/></p:ext></p:extLst>`)!.documentElement, true));
  const hlink = doc.xml(part)!.createElementNS(NS.a, "a:hlinkClick");
  hlink.setAttributeNS(NS.r, "r:id", "");
  hlink.setAttribute("action", "ppaction://media");
  shape.getElementsByTagNameNS(NS.p, "cNvPr")[0].append(hlink);
  setMediaPlayback(doc.xml(part)!, shape, { autoPlay: false, loop: false });
  return shape;
}
export function setMediaPlayback(slide: Document, shape: Element, patch: Partial<Pick<SlideMedia, "autoPlay" | "loop">>): void {
  const id = shapeId(shape);
  let node = timingNode(slide, id);
  const make = (tag: string) => slide.createElementNS(NS.p, "p:" + tag);
  const nextId = () => String(Math.max(0, ...Array.from(slide.getElementsByTagNameNS(NS.p, "cTn")).map(e => Number(e.getAttribute("id")))) + 1);
  if (!node) {
    let timing = slide.getElementsByTagNameNS(NS.p, "timing")[0];
    if (!timing) {
      timing = make("timing");
      slide.documentElement.insertBefore(timing, Array.from(slide.documentElement.children).find(e => e.localName === "extLst") ?? null);
    }
    let list = Array.from(timing.children).find(e => e.localName === "tnLst");
    if (!list) {
      list = make("tnLst");
      timing.insertBefore(list, timing.firstChild);
    }
    let root = Array.from(list.getElementsByTagNameNS(NS.p, "cTn")).find(e => e.getAttribute("nodeType") === "tmRoot");
    if (!root) {
      const par = make("par");
      root = make("cTn");
      root.setAttribute("id", nextId());
      root.setAttribute("dur", "indefinite");
      root.setAttribute("restart", "never");
      root.setAttribute("nodeType", "tmRoot");
      par.append(root);
      list.append(par);
    }
    let children = Array.from(root.children).find(e => e.localName === "childTnLst");
    if (!children) {
      children = make("childTnLst");
      root.append(children);
    }
    const kind = shape.getElementsByTagNameNS(NS.a, "videoFile").length ? "video" : "audio";
    const media = make(kind);
    node = make("cMediaNode");
    node.setAttribute("vol", "100000");
    const ct = make("cTn");
    ct.setAttribute("id", nextId());
    ct.setAttribute("dur", "media");
    ct.setAttribute("fill", "hold");
    ct.append(make("stCondLst"));
    const target = make("tgtEl");
    const sp = make("spTgt");
    sp.setAttribute("spid", id);
    target.append(sp);
    node.append(ct, target);
    media.append(node);
    children.append(media);
  }
  const ct = node.getElementsByTagNameNS(NS.p, "cTn")[0];
  if (patch.loop !== undefined) {
    if (patch.loop)
      ct.setAttribute("repeatCount", "indefinite");
    else
      ct.removeAttribute("repeatCount");
  }
  if (patch.autoPlay !== undefined) {
    let conditions = Array.from(ct.children).find(e => e.localName === "stCondLst");
    if (!conditions) {
      conditions = make("stCondLst");
      ct.insertBefore(conditions, ct.firstChild);
    }
    const cond = make("cond");
    cond.setAttribute("delay", "0");
    if (!patch.autoPlay) {
      cond.setAttribute("evt", "onClick");
      const target = make("tgtEl");
      const sp = make("spTgt");
      sp.setAttribute("spid", id);
      target.append(sp);
      cond.append(target);
    }
    conditions.replaceChildren(cond);
  }
}

/**
 * Vault paths to try for media the deck links to instead of embedding (`r:link` with TargetMode="External"), best
 * first. Only audio / video files inside the vault: a relative link is resolved from the deck's folder, and any
 * link (absolute paths from the author's machine, or one leaving the vault) falls back to a same-named file beside
 * the deck, as PowerPoint does. Web links are refused, so nothing is ever fetched.
 */
export function linkedMediaCandidates(target: string, deckPath: string): string[] {
  let t = target.trim();
  if (/^file:/i.test(t)) t = t.replace(/^file:\/*/i, "/");
  else if (/^[a-z][a-z0-9+.-]+:/i.test(t) && !/^[a-z]:[\\/]/i.test(t)) return [];
  try {
    t = decodeURIComponent(t);
  } catch {
    // keep it as written
  }
  t = t.replace(/\\/g, "/");
  const name = t.slice(t.lastIndexOf("/") + 1);
  if (!MEDIA_MIME[name.slice(name.lastIndexOf(".") + 1).toLowerCase()]) return [];
  const dir = deckPath.includes("/") ? deckPath.slice(0, deckPath.lastIndexOf("/")) : "";
  const beside = dir ? `${dir}/${name}` : name;
  const absolute = t.startsWith("/") || /^[a-z]:\//i.test(t);
  const out: string[] = [];
  if (!absolute) {
    const parts = dir ? dir.split("/") : [];
    let inside = true;
    for (const seg of t.split("/")) {
      if (seg === "..") {
        if (!parts.length) inside = false;
        parts.pop();
      } else if (seg && seg !== ".") parts.push(seg);
    }
    if (inside) out.push(parts.join("/"));
  }
  if (!out.includes(beside)) out.push(beside);
  return out;
}
