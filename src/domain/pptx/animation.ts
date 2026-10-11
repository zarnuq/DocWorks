/**
 * Slide transitions (p:transition) and animations (p:timing) for the slideshow. Animations are read from their
 * behaviours (show / hide, filter effects, property keyframes, scale, rotation, motion paths) rather than from preset
 * ids, so effects from PowerPoint, Google Slides and LibreOffice all play without a preset table.
 */
import { NS } from "../../core/utils/xml";

const P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

export interface Transition {
  /** Element name: fade, push, wipe, cover, pull, split, zoom, cut, p14 vortex / flip / gallery... */
  type: string;
  dir: string | null;
  /** ms */
  dur: number;
  /** Advance to the next slide after this many ms. */
  advanceAfter: number | null;
}

interface Timed {
  /** ms from the effect's start */
  start: number;
  dur: number;
  autoRev?: true;
}

export type Behavior = Timed &
  (
    | { type: "visibility"; visible: boolean }
    | { type: "effect"; filter: string; in: boolean }
    | { type: "prop"; attr: string; keys: { t: number; v: string }[] }
    | { type: "scale"; by: { x: number; y: number } }
    | { type: "rotate"; by: number }
    | { type: "motion"; points: { x: number; y: number }[] }
  );

export interface Effect {
  spid: string;
  /** Paragraph range [first, last] for text animated by paragraph. */
  para: [number, number] | null;
  /** entr, exit, emph, path... */
  cls: string;
  /** ms from the step's start */
  start: number;
  dur: number;
  behaviors: Behavior[];
}

export interface Step {
  /** Starts without a click (with the slide, or after the previous step). */
  auto: boolean;
  effects: Effect[];
}

const kids = (el: Element | null | undefined, local?: string) => (el ? Array.from(el.children).filter((c) => !local || c.localName === local) : []);
const kid = (el: Element | null | undefined, local: string) => kids(el, local)[0] ?? null;
const num = (v: string | null | undefined) => (v && /^-?\d+(\.\d+)?$/.test(v) ? Number(v) : 0);

/** The first delay of a cTn's start conditions. */
const delayOf = (cTn: Element | null) => num(kid(kid(cTn, "stCondLst"), "cond")?.getAttribute("delay"));

export function readTransition(slide: Document): Transition | null {
  const t = slide.getElementsByTagNameNS(NS.p, "transition")[0];
  if (!t) return null;
  const kind = kids(t).find((c) => c.localName !== "sndAc" && c.localName !== "extLst") ?? null;
  const type = kind?.localName ?? "cut";
  const spd = { fast: 500, med: 750, slow: 1000 }[t.getAttribute("spd") ?? "fast"] ?? 500;
  const p14dur = t.getAttributeNS(P14, "dur");
  const adv = t.getAttribute("advTm");
  return {
    type,
    dir: kind?.getAttribute("dir") ?? null,
    dur: type === "cut" ? 0 : p14dur ? num(p14dur) : spd,
    advanceAfter: adv ? num(adv) : null,
  };
}

/** Effects the editor can author; other imported transitions remain readable and keep their XML. */
export const TRANSITIONS: [string, string][] = [["None", "cut"], ["Fade", "fade"], ["Slide", "push"], ["Wipe", "wipe"], ["Cover", "cover"], ["Uncover", "pull"], ["Zoom", "zoom"]];

export type TransitionPatch = Partial<Pick<Transition, "type" | "dir" | "dur">>;

/** Patch the visual transition only. Slide advance, sounds, extensions and animation timing are preserved. */
export function writeTransition(slide: Document, patch: TransitionPatch): void {
  if (patch.type !== undefined && !TRANSITIONS.some(([, type]) => type === patch.type)) throw new Error("Unsupported transition");
  if (patch.dur !== undefined && (!Number.isFinite(patch.dur) || patch.dur < 1)) throw new Error("Duration must be positive");
  const root = slide.documentElement;
  const transitions = Array.from(slide.getElementsByTagNameNS(NS.p, "transition"));
  if (!transitions.length) {
    if (!patch.type || patch.type === "cut") return;
    const t = slide.createElementNS(NS.p, "p:transition");
    root.insertBefore(t, kids(root).find((c) => c.localName === "timing" || c.localName === "extLst") ?? null);
    transitions.push(t);
  }
  for (const t of transitions) {
    if (patch.type !== undefined) {
      for (const c of kids(t)) if (c.localName !== "sndAc" && c.localName !== "extLst") c.remove();
      if (patch.type !== "cut") {
        const kind = slide.createElementNS(NS.p, `p:${patch.type}`);
        if (["push", "wipe", "cover", "pull"].includes(patch.type)) kind.setAttribute("dir", patch.dir ?? "l");
        else if (patch.type === "zoom") kind.setAttribute("dir", "in");
        t.insertBefore(kind, t.firstChild);
      } else {
        t.removeAttribute("spd");
        t.removeAttributeNS(P14, "dur");
      }
    } else if (patch.dir !== undefined) {
      const kind = kids(t).find((c) => ["push", "wipe", "cover", "pull"].includes(c.localName));
      if (kind) kind.setAttribute("dir", patch.dir ?? "l");
    }
    if (patch.dur !== undefined && patch.type !== "cut") {
      const dur = Math.round(patch.dur);
      t.setAttribute("spd", dur <= 500 ? "fast" : dur <= 750 ? "med" : "slow");
      // Exact durations use Office's alternate representation, with a speed-only fallback.
      if (t.parentElement === root) {
        const ac = slide.createElementNS(NS.mc, "mc:AlternateContent");
        ac.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:p14", P14);
        const choice = slide.createElementNS(NS.mc, "mc:Choice");
        choice.setAttribute("Requires", "p14");
        const fallback = slide.createElementNS(NS.mc, "mc:Fallback");
        const copy = t.cloneNode(true) as Element;
        copy.removeAttributeNS(P14, "dur");
        fallback.append(copy);
        root.replaceChild(ac, t);
        choice.append(t);
        ac.append(choice, fallback);
        t.setAttributeNS(P14, "p14:dur", String(dur));
      } else if (t.parentElement?.localName === "Choice") {
        const choice = t.parentElement;
        const prefix = choice.lookupPrefix(P14) ?? "dw14";
        choice.setAttributeNS("http://www.w3.org/2000/xmlns/", `xmlns:${prefix}`, P14);
        const requires = new Set((choice.getAttribute("Requires") ?? "").split(/\s+/).filter(Boolean));
        requires.add(prefix);
        choice.setAttribute("Requires", Array.from(requires).join(" "));
        t.setAttributeNS(P14, `${prefix}:dur`, String(dur));
      } else t.removeAttributeNS(P14, "dur");
    }
  }
}

function behavior(el: Element): Behavior | null {
  const cBhvr = kid(el, "cBhvr");
  const cTn = kid(cBhvr, "cTn");
  const timed: Timed = { start: delayOf(cTn), dur: num(cTn?.getAttribute("dur")) };
  if (cTn?.getAttribute("autoRev") === "1") timed.autoRev = true;
  const attr = kid(kid(cBhvr, "attrNameLst"), "attrName")?.textContent?.trim() ?? "";
  const xy = (e: Element | null) => (e ? { x: num(e.getAttribute("x")) / 100000, y: num(e.getAttribute("y")) / 100000 } : null);
  switch (el.localName) {
    case "set": {
      const val = kid(kid(el, "to"), "strVal")?.getAttribute("val");
      return attr === "style.visibility" && val ? { type: "visibility", visible: val !== "hidden", ...timed } : null;
    }
    case "animEffect":
      return { type: "effect", filter: el.getAttribute("filter") ?? "", in: el.getAttribute("transition") !== "out", ...timed };
    case "anim": {
      const tavs = kids(kid(el, "tavLst"), "tav");
      const keys = tavs.length
        ? tavs.map((tav) => ({ t: num(tav.getAttribute("tm")) / 100000, v: kid(kid(tav, "val"), "strVal")?.getAttribute("val") ?? kid(kid(tav, "val"), "fltVal")?.getAttribute("val") ?? "" }))
        : el.hasAttribute("from") && el.hasAttribute("to")
          ? [
              { t: 0, v: el.getAttribute("from")! },
              { t: 1, v: el.getAttribute("to")! },
            ]
          : [];
      return attr && keys.length ? { type: "prop", attr, keys, ...timed } : null;
    }
    case "animScale": {
      const by = xy(kid(el, "by")) ?? xy(kid(el, "to"));
      return by ? { type: "scale", by, ...timed } : null;
    }
    case "animRot":
      return { type: "rotate", by: num(el.getAttribute("by")) / 60000, ...timed };
    case "animMotion": {
      // ponytail: only each segment's end point is kept; curves play as straight lines.
      const segs: string[] = (el.getAttribute("path") ?? "").match(/[MLC][^MLCEZ]*/gi) ?? [];
      const points = segs.map((s) => {
        const n = s.match(/-?\d*\.?\d+(e-?\d+)?/gi)?.map(Number) ?? [];
        return { x: n[n.length - 2] ?? 0, y: n[n.length - 1] ?? 0 };
      });
      return points.length ? { type: "motion", points, ...timed } : null;
    }
  }
  return null;
}

function effectOf(par: Element, groupStart: number): Effect | null {
  const cTn = kid(par, "cTn");
  if (!cTn?.hasAttribute("presetClass")) return null;
  const behaviors = kids(kid(cTn, "childTnLst"))
    .map(behavior)
    .filter((b): b is Behavior => !!b);
  const spTgt = cTn.getElementsByTagNameNS(NS.p, "spTgt")[0];
  if (!spTgt) return null;
  const pRg = spTgt.getElementsByTagNameNS(NS.p, "pRg")[0];
  return {
    spid: spTgt.getAttribute("spid") ?? "",
    para: pRg ? [num(pRg.getAttribute("st")), num(pRg.getAttribute("end"))] : null,
    cls: cTn.getAttribute("presetClass") ?? "",
    start: groupStart + delayOf(cTn),
    dur: Math.max(0, ...behaviors.map((b) => b.start + b.dur * (b.autoRev ? 2 : 1))),
    behaviors,
  };
}

/** The slide's main animation sequence, one step per click. */
export function readAnimations(slide: Document): Step[] {
  const main = Array.from(slide.getElementsByTagNameNS(NS.p, "cTn")).find((c) => c.getAttribute("nodeType") === "mainSeq");
  return kids(kid(main, "childTnLst"), "par").map((click) => {
    const cTn = kid(click, "cTn");
    const conds = kids(kid(cTn, "stCondLst"), "cond");
    const auto = conds.some((c) => c.getAttribute("evt") === "onBegin") || !conds.some((c) => c.getAttribute("delay") === "indefinite");
    const effects = kids(kid(cTn, "childTnLst"), "par").flatMap((group) => {
      const g = kid(group, "cTn");
      return kids(kid(g, "childTnLst"), "par")
        .map((par) => effectOf(par, delayOf(g)))
        .filter((e): e is Effect => !!e);
    });
    return { auto, effects };
  });
}

/** Evaluate a PowerPoint animation value: numbers, #ppt_x / #ppt_y / #ppt_w / #ppt_h, + - * / and brackets. */
export function evalFormula(src: string, vars: Record<string, number>): number | null {
  const tokens = src.match(/#\w+|\d*\.?\d+(?:e-?\d+)?|[-+*/()]|\S/gi) ?? [];
  let i = 0;
  const fail = () => {
    throw new Error("formula");
  };
  const atom = (): number => {
    const t = tokens[i++];
    if (t === undefined) return fail();
    if (t === "(") {
      const v = sum();
      if (tokens[i++] !== ")") fail();
      return v;
    }
    if (t === "-") return -atom();
    if (t === "+") return atom();
    if (t.startsWith("#")) return vars[t.slice(1)] ?? fail();
    return /^\d*\.?\d+(e-?\d+)?$/i.test(t) ? Number(t) : fail();
  };
  const product = (): number => {
    let v = atom();
    while (tokens[i] === "*" || tokens[i] === "/") v = tokens[i++] === "*" ? v * atom() : v / atom();
    return v;
  };
  const sum = (): number => {
    let v = product();
    while (tokens[i] === "+" || tokens[i] === "-") v = tokens[i++] === "+" ? v + product() : v - product();
    return v;
  };
  try {
    const v = sum();
    return i === tokens.length && Number.isFinite(v) ? v : null;
  } catch {
    return null;
  }
}

/** A shape's own cNvPr id, then the ids of the groups it sits in (an animation on a group moves its shapes). */
export function shapeIds(el: Element): string[] {
  const idOf = (e: Element) => kid(kids(e).find((c) => c.localName.startsWith("nv")), "cNvPr")?.getAttribute("id");
  const out: string[] = [];
  for (let e: Element | null = el; e; e = e.parentElement) {
    if (e !== el && e.localName !== "grpSp") continue;
    const id = idOf(e);
    if (id) out.push(id);
  }
  return out;
}
