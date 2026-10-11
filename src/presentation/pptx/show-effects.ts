/*
 * Plays slide animations and transitions on rendered slides with the Web Animations API. Shapes are found by
 * their `data-spids` (own id + group ids); paragraphs by their index in the shape's text.
 */
import { Behavior, Effect, evalFormula, Step, Transition } from "../../domain/pptx/animation";

type Frames = { frames: Keyframe[]; composite?: CompositeOperation };

/** Elements an effect animates: the shape(s), or a range of their paragraphs. */
export function targets(slide: HTMLElement, e: Effect): HTMLElement[] {
  if (!/^\w+$/.test(e.spid)) return [];
  const shapes = Array.from(slide.querySelectorAll<HTMLElement>(`.pptx-shape[data-spids~="${e.spid}"]`));
  if (!e.para) return shapes;
  const [st, end] = e.para;
  return shapes.flatMap((s) => Array.from(s.querySelectorAll<HTMLElement>(".pptx-text > p")).slice(st, end + 1));
}

const setVisible = (els: HTMLElement[], on: boolean) => els.forEach((el) => (el.style.visibility = on ? "visible" : "hidden"));

/** The visibility an effect leaves its target with, if it sets one. */
function finalVisibility(e: Effect): boolean | null {
  let v: boolean | null = null;
  for (const b of e.behaviors) if (b.type === "visibility") v = b.visible;
  return v;
}

/** Draw the slide as it looks after its first `done` steps: later entrances hidden, earlier exits gone. */
export function showState(slide: HTMLElement, steps: Step[], done: number): void {
  const seen = new Set<string>();
  for (const s of steps)
    for (const e of s.effects) {
      const key = `${e.spid}:${e.para?.join("-") ?? ""}`;
      if (seen.has(key)) continue;
      seen.add(key);
      if (e.cls === "entr") setVisible(targets(slide, e), false);
    }
  for (const s of steps.slice(0, done))
    for (const e of s.effects) {
      const v = finalVisibility(e);
      if (v !== null) setVisible(targets(slide, e), v);
    }
}

export const stepDuration = (s: Step) => Math.max(0, ...s.effects.map((e) => e.start + e.dur));

/** Keyframes for a filter effect (fade, wipe, split, circle...) as it plays in; others fade. */
function filterFrames(filter: string): Keyframe[] {
  const [, type, sub = ""] = /^(\w+)(?:\((\w+)\))?/.exec(filter) ?? [];
  const clip = (from: string, to = "inset(0 0 0 0)") => [{ clipPath: from }, { clipPath: to }];
  switch (type) {
    case "wipe":
      return clip(`inset(${{ down: "0 0 100% 0", up: "100% 0 0 0", left: "0 0 0 100%", right: "0 100% 0 0" }[sub] ?? "0 100% 0 0"})`);
    case "barn":
      return clip(sub.endsWith("Horizontal") ? "inset(50% 0 50% 0)" : "inset(0 50% 0 50%)");
    case "circle":
    case "diamond":
    case "plus":
    case "wheel":
    case "wedge":
    case "box":
      return clip("circle(0% at 50% 50%)", "circle(75% at 50% 50%)");
    default:
      return [{ opacity: 0 }, { opacity: 1 }];
  }
}

function framesOf(b: Behavior, el: HTMLElement, box: DOMRect): Frames | null {
  const W = box.width || 1;
  const H = box.height || 1;
  const r = el.getBoundingClientRect();
  const vars = { ppt_x: (r.left - box.left + r.width / 2) / W, ppt_y: (r.top - box.top + r.height / 2) / H, ppt_w: r.width / W, ppt_h: r.height / H };
  const add = (frames: Keyframe[]): Frames => ({ frames, composite: "add" });
  switch (b.type) {
    case "effect": {
      const f = filterFrames(b.filter);
      return { frames: b.in ? f : f.reverse() };
    }
    case "prop": {
      const vals = b.keys.map((k) => evalFormula(k.v, vars));
      if (vals.some((v) => v === null)) return null;
      const at = (i: number, kf: Keyframe): Keyframe => ({ offset: b.keys[i].t, ...kf });
      const v = vals as number[];
      switch (b.attr) {
        case "ppt_x":
          return add(v.map((x, i) => at(i, { transform: `translateX(${(x - vars.ppt_x) * W}px)` })));
        case "ppt_y":
          return add(v.map((y, i) => at(i, { transform: `translateY(${(y - vars.ppt_y) * H}px)` })));
        case "ppt_w":
          return vars.ppt_w ? add(v.map((w, i) => at(i, { transform: `scaleX(${w / vars.ppt_w})` }))) : null;
        case "ppt_h":
          return vars.ppt_h ? add(v.map((h, i) => at(i, { transform: `scaleY(${h / vars.ppt_h})` }))) : null;
        case "style.opacity":
          return { frames: v.map((o, i) => at(i, { opacity: o })) };
        case "style.rotation":
        case "r":
          return add(v.map((d, i) => at(i, { transform: `rotate(${d}deg)` })));
      }
      return null;
    }
    case "scale":
      return add([{ transform: "scale(1)" }, { transform: `scale(${b.by.x}, ${b.by.y})` }]);
    case "rotate":
      return add([{ transform: "rotate(0deg)" }, { transform: `rotate(${b.by}deg)` }]);
    case "motion": {
      const pts = b.points.length > 1 ? b.points : [b.points[0], b.points[0]];
      return add(pts.map((p, i) => ({ offset: i / (pts.length - 1), transform: `translate(${p.x * W}px, ${p.y * H}px)` })));
    }
  }
  return null;
}

/** Play one step on a drawn slide. Returns a function that stops it. */
export function playStep(slide: HTMLElement, step: Step): () => void {
  const timers: number[] = [];
  const anims: Animation[] = [];
  const box = slide.getBoundingClientRect();
  for (const e of step.effects)
    for (const el of targets(slide, e))
      for (const b of e.behaviors) {
        const t = e.start + b.start;
        if (b.type === "visibility") {
          const show = () => setVisible([el], b.visible);
          if (t <= 0) show();
          else timers.push(window.setTimeout(show, t));
          continue;
        }
        const f = typeof el.animate === "function" ? framesOf(b, el, box) : null;
        if (!f) continue;
        anims.push(
          el.animate(f.frames, {
            duration: Math.max(1, b.dur),
            delay: t,
            fill: "forwards",
            composite: f.composite ?? "replace",
            iterations: b.autoRev ? 2 : 1,
            direction: b.autoRev ? "alternate" : "normal",
          })
        );
      }
  return () => {
    timers.forEach((t) => window.clearTimeout(t));
    anims.forEach((a) => a.cancel());
  };
}

/** Movement of a transition with direction `dir` (l = towards the left...), as a unit vector. */
const DIRS: Record<string, [number, number]> = { l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1], lu: [-1, -1], ru: [1, -1], ld: [-1, 1], rd: [1, 1] };
const PUSHES = new Set(["push", "gallery", "conveyor", "pan", "ferris", "prism", "cube", "switch", "flythrough"]);

/** Animate from the old slide (kept on screen meanwhile) to the new one, both inside `stage`. */
export function playTransition(stage: HTMLElement, from: HTMLElement, to: HTMLElement, t: Transition): void {
  if (t.type === "cut" || t.dur <= 0 || typeof to.animate !== "function") return;
  const [dx, dy] = DIRS[t.dir ?? "l"] ?? DIRS.l;
  const off = (k: number) => `translate(${dx * k * 100}%, ${dy * k * 100}%)`;
  const opts: KeyframeAnimationOptions = { duration: t.dur, easing: "ease-in-out" };
  const over = t.type === "pull" || t.type === "uncover";
  from.classList.add("dw-show-leaving");
  from.classList.toggle("is-over", over);
  if (over) stage.append(from);
  else stage.prepend(from);

  const anims: Animation[] = [];
  const play = (el: HTMLElement, frames: Keyframe[]) => anims.push(el.animate(frames, opts));
  if (PUSHES.has(t.type)) {
    play(to, [{ transform: off(-1) }, { transform: off(0) }]);
    play(from, [{ transform: off(0) }, { transform: off(1) }]);
  } else if (t.type === "cover") play(to, [{ transform: off(-1) }, { transform: off(0) }]);
  else if (over) play(from, [{ transform: off(0) }, { transform: off(1) }]);
  else if (t.type === "wipe") {
    const hidden = { l: "0 0 0 100%", r: "0 100% 0 0", u: "100% 0 0 0", d: "0 0 100% 0" }[t.dir ?? "l"] ?? "0 0 0 100%";
    play(to, [{ clipPath: `inset(${hidden})` }, { clipPath: "inset(0 0 0 0)" }]);
  } else if (t.type === "split") play(to, [{ clipPath: "inset(0 50% 0 50%)" }, { clipPath: "inset(0 0 0 0)" }]);
  else if (["circle", "diamond", "plus", "wheel", "wedge", "wheelReverse"].includes(t.type)) play(to, [{ clipPath: "circle(0% at 50% 50%)" }, { clipPath: "circle(75% at 50% 50%)" }]);
  else if (t.type === "zoom" || t.type === "newsflash") play(to, [{ transform: "scale(0.3)", opacity: 0 }, { transform: "scale(1)", opacity: 1 }]);
  else if (t.type === "flip") play(to, [{ transform: "perspective(1200px) rotateY(90deg)" }, { transform: "perspective(1200px) rotateY(0deg)" }]);
  // ponytail: other transitions (dissolve, blinds, vortex, shred...) fade.
  else play(to, [{ opacity: 0 }, { opacity: 1 }]);
  void Promise.all(anims.map((a) => a.finished)).then(
    () => from.remove(),
    () => from.remove()
  );
}
