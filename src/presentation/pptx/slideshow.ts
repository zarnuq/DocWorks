/*
 * Full-window slideshow with an optional presenter view (current slide, next
 * slide, notes, timer). Plain DOM: the deck is supplied by the caller.
 */

import { linkTarget } from "../../domain/pptx/links";
import type { Step, Transition } from "../../domain/pptx/animation";
import { playStep, playTransition, showState, stepDuration } from "./show-effects";

export interface SlideEffects {
  /** How this slide comes in. */
  transition: Transition | null;
  /** Animation steps, one per click. */
  steps: Step[];
}

export interface ShowDeck {
  size: { cx: number; cy: number };
  count: number;
  isHidden(i: number): boolean;
  notes(i: number): string;
  /** Draw slide `i` into `host` at `width` px. */
  render(host: HTMLElement, i: number, width: number): void;
  effects?(i: number): SlideEffects;
  /** Slide part paths in order, for links to slides. */
  slidePaths?: string[];
}

/** Next slide from `from` in direction `dir` that isn't hidden, or null past the end. */
export function nextVisible(hidden: boolean[], from: number, dir: 1 | -1): number | null {
  for (let i = from + dir; i >= 0 && i < hidden.length; i += dir) if (!hidden[i]) return i;
  return null;
}

const NEXT = new Set(["ArrowRight", "ArrowDown", "PageDown", " ", "Enter", "n", "N"]);
const PREV = new Set(["ArrowLeft", "ArrowUp", "PageUp", "Backspace"]);
const pad = (n: number) => String(n).padStart(2, "0");
const SVG = "http://www.w3.org/2000/svg";

/** A pen stroke, as fractions of the slide's width and height (so it survives a resize or presenter view). */
type Stroke = [x: number, y: number][];
export type ShowTool = "none" | "laser" | "pen";

/** SVG path data of a stroke drawn on a w x h slide. */
export function strokePath(stroke: Stroke, w: number, h: number): string {
  return stroke.map(([x, y], i) => `${i ? "L" : "M"}${(x * w).toFixed(1)} ${(y * h).toFixed(1)}`).join(" ");
}

export class Slideshow {
  private overlay: HTMLElement | null = null;
  private presenter = false;
  private ended = false;
  private typed = "";
  private started = Date.now();
  private timer: number | null = null;
  private wasFullscreen = false;
  /** Animation steps played on the current slide. */
  private step = 0;
  /** Stops the step that is playing. */
  private stopStep: (() => void) | null = null;
  private stepTimer: number | null = null;
  private advanceTimer: number | null = null;
  /** The current slide's host. */
  private stage: HTMLElement | null = null;
  private mediaIndex = -1;
  private effectsCache = new Map<number, SlideEffects>();
  /** Laser pointer (Google's L) or pen (PowerPoint's Ctrl+P). */
  tool: ShowTool = "none";
  /** Pen strokes per slide, kept until the show ends. */
  private ink = new Map<number, Stroke[]>();
  private drawing: Stroke | null = null;
  private pointingTimer: number | null = null;

  constructor(private deck: ShowDeck, public index: number, private onExit: (index: number) => void) {}

  open(previewTransition = false): void {
    const o = (this.overlay = document.body.createDiv({ cls: "dw-show", attr: { tabindex: "-1" } }));
    o.addEventListener("keydown", this.onKey);
    o.addEventListener("click", (e) => {
      const target = e.target as HTMLElement;
      const link = target.closest<HTMLElement>("[data-link]")?.dataset.link;
      if (link?.startsWith("#")) {
        const to = linkTarget(link, this.index, this.deck.slidePaths ?? []);
        if (to !== null) this.goTo(to);
      } else if (link) window.open(link);
      else if (this.tool !== "pen" && !target.closest("button, audio, video, .dw-show-notes, .dw-show-tools")) this.next();
    });
    o.addEventListener("pointermove", this.onPointerMove);
    o.addEventListener("pointerdown", this.onPointerDown);
    o.addEventListener("pointerup", this.onPointerUp);
    o.addEventListener("wheel", (e) => (e.deltaY > 0 ? this.next() : this.prev()), { passive: true });
    document.addEventListener("fullscreenchange", this.onFullscreenChange);
    window.addEventListener("resize", this.redraw);
    this.timer = window.setInterval(this.tick, 1000);
    this.draw();
    const t = previewTransition ? this.effects().transition : null;
    const fresh = this.stage?.querySelector<HTMLElement>(".pptx-slide");
    if (t && fresh && this.stage) {
      const blank = fresh.cloneNode(false) as HTMLElement;
      blank.addClass("dw-show-blank");
      playTransition(this.stage, blank, fresh, t);
    }
    if (t?.dur) this.stepTimer = window.setTimeout(() => {
      this.stepTimer = null;
      this.autoPlay();
    }, t.dur);
    else this.autoPlay();
    o.focus();
    o.requestFullscreen?.().catch(() => {
      // Fullscreen can be refused; the overlay still covers the window.
    });
  }

  close(): void {
    if (!this.overlay) return;
    this.stop();
    if (this.timer) window.clearInterval(this.timer);
    if (this.pointingTimer) window.clearTimeout(this.pointingTimer);
    document.removeEventListener("fullscreenchange", this.onFullscreenChange);
    window.removeEventListener("resize", this.redraw);
    if (document.fullscreenElement === this.overlay) void document.exitFullscreen?.().catch(() => undefined);
    this.overlay.querySelectorAll("audio,video").forEach(el => (el as HTMLMediaElement).pause());
    this.overlay.remove();
    this.overlay = null;
    this.onExit(this.index);
  }

  private hidden(): boolean[] {
    return Array.from({ length: this.deck.count }, (_, i) => this.deck.isHidden(i));
  }

  private effects(i = this.index): SlideEffects {
    let e = this.effectsCache.get(i);
    if (!e) this.effectsCache.set(i, (e = this.deck.effects?.(i) ?? { transition: null, steps: [] }));
    return e;
  }

  /** Stop the playing step (leaving it built) and the automatic advance. */
  private stop(): void {
    this.stopStep?.();
    this.stopStep = null;
    if (this.stepTimer !== null) window.clearTimeout(this.stepTimer);
    if (this.advanceTimer !== null) window.clearTimeout(this.advanceTimer);
    this.stepTimer = this.advanceTimer = null;
  }

  private play(): void {
    const step = this.effects().steps[this.step++];
    if (this.stage) this.stopStep = playStep(this.stage, step);
    this.stepTimer = window.setTimeout(() => {
      this.stepTimer = null;
      this.autoPlay();
    }, stepDuration(step));
  }

  /** Play the next step if it starts by itself, else wait for the slide's automatic advance. */
  private autoPlay(): void {
    if (this.ended) return;
    const { steps, transition } = this.effects();
    if (this.step < steps.length && steps[this.step].auto) this.play();
    else if (this.step >= steps.length && transition?.advanceAfter != null) this.advanceTimer = window.setTimeout(() => this.next(), transition.advanceAfter);
  }

  next(): void {
    if (this.ended) return this.close();
    // A click while a step plays finishes it.
    if (this.stepTimer !== null) return this.redraw();
    this.stop();
    if (this.step < this.effects().steps.length) return this.play();
    const n = nextVisible(this.hidden(), this.index, 1);
    if (n === null) {
      this.ended = true;
      return this.draw();
    }
    this.index = n;
    this.step = 0;
    const t = this.effects().transition;
    this.draw(t);
    // Animations start once the transition is over.
    if (t?.dur) this.stepTimer = window.setTimeout(() => {
      this.stepTimer = null;
      this.autoPlay();
    }, t.dur);
    else this.autoPlay();
  }

  prev(): void {
    if (this.ended) this.ended = false;
    else if (this.step > 0) this.step--;
    else {
      const p = nextVisible(this.hidden(), this.index, -1);
      if (p !== null) {
        this.index = p;
        this.step = this.effects().steps.length;
      }
    }
    this.draw();
  }

  goTo(i: number): void {
    if (i < 0 || i >= this.deck.count) return;
    this.ended = false;
    this.index = i;
    this.step = 0;
    this.draw();
    this.autoPlay();
  }

  private onKey = (e: KeyboardEvent) => {
    if ((e.target as HTMLElement).closest("audio,video") && e.key !== "Escape") return;
    const k = e.key;
    if (/^\d$/.test(k)) {
      this.typed += k;
      return;
    }
    if (k === "Enter" && this.typed) {
      this.goTo(Number(this.typed) - 1);
      this.typed = "";
    } else if (k === "Escape") {
      // Escape puts a pen or pointer down first, as in PowerPoint.
      if (this.tool !== "none") this.setTool("none");
      else this.close();
    } else if ((k === "l" || k === "L") && !e.altKey) this.setTool(this.tool === "laser" ? "none" : "laser");
    else if ((k === "p" || k === "P") && (e.ctrlKey || e.metaKey)) {
      // In Obsidian a scope the editor pushes may have taken it already (it's the command palette key there).
      if (e.defaultPrevented) return;
      e.stopPropagation();
      this.setTool(this.tool === "pen" ? "none" : "pen");
    }
    else if (k === "e" || k === "E") this.eraseInk();
    else if (k === "p" || k === "P") {
      this.presenter = !this.presenter;
      this.redraw();
    } else if (k === "Home") this.goTo(nextVisible(this.hidden(), -1, 1) ?? 0);
    else if (k === "End") this.goTo(nextVisible(this.hidden(), this.deck.count, -1) ?? this.deck.count - 1);
    else if (NEXT.has(k)) this.next();
    else if (PREV.has(k)) this.prev();
    else return;
    e.preventDefault();
    this.typed = "";
  };

  setTool(tool: ShowTool): void {
    this.tool = tool;
    this.drawing = null;
    const o = this.overlay;
    if (!o) return;
    o.classList.toggle("is-laser", tool === "laser");
    o.classList.toggle("is-pen", tool === "pen");
    o.querySelectorAll<HTMLElement>(".dw-show-tools button[data-tool]").forEach((b) => b.classList.toggle("is-active", b.dataset.tool === tool));
  }

  /** E: the current slide's ink goes. */
  eraseInk(): void {
    this.ink.delete(this.index);
    this.drawInk();
  }

  /** The slide as drawn in the current stage (null on the end screen). */
  private slideBox(): DOMRect | null {
    const slide = this.stage?.querySelector<HTMLElement>(":scope > .pptx-slide:not(.dw-show-leaving)") ?? this.stage;
    const r = slide?.getBoundingClientRect();
    return r && r.width && r.height ? r : null;
  }

  private onPointerMove = (e: PointerEvent) => {
    const o = this.overlay;
    if (!o) return;
    // Moving the mouse shows the pointer and the tools for a moment.
    o.classList.add("is-pointing");
    if (this.pointingTimer) window.clearTimeout(this.pointingTimer);
    this.pointingTimer = window.setTimeout(() => o.classList.remove("is-pointing"), 2500);
    const dot = o.querySelector<HTMLElement>(".dw-show-laser");
    dot?.setCssProps({ "--dw-laser-x": `${e.clientX}px`, "--dw-laser-y": `${e.clientY}px` });
    if (this.drawing) {
      const r = this.slideBox();
      if (r) this.drawing.push([(e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height]);
      this.drawInk();
    }
  };

  private onPointerDown = (e: PointerEvent) => {
    if (this.tool !== "pen" || e.button !== 0 || (e.target as HTMLElement).closest(".dw-show-tools, button, .dw-show-notes")) return;
    const r = this.slideBox();
    if (!r || e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom) return;
    this.drawing = [[(e.clientX - r.left) / r.width, (e.clientY - r.top) / r.height]];
    const strokes = this.ink.get(this.index) ?? [];
    strokes.push(this.drawing);
    this.ink.set(this.index, strokes);
    (e.target as Element).setPointerCapture?.(e.pointerId);
    e.preventDefault();
    this.drawInk();
  };

  private onPointerUp = () => {
    this.drawing = null;
  };

  /** Draw the current slide's ink over it. */
  private drawInk(): void {
    const slide = this.stage?.querySelector<HTMLElement>(":scope > .pptx-slide:not(.dw-show-leaving)") ?? this.stage;
    if (!slide || !this.stage) return;
    let svg = this.stage.querySelector<SVGSVGElement>(":scope > .dw-show-ink");
    const strokes = this.ink.get(this.index) ?? [];
    if (!strokes.length) return void svg?.remove();
    const w = slide.offsetWidth || this.stage.offsetWidth;
    const h = slide.offsetHeight || this.stage.offsetHeight;
    if (!svg) {
      svg = document.createElementNS(SVG, "svg");
      svg.classList.add("dw-show-ink");
      this.stage.appendChild(svg);
    }
    svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
    svg.replaceChildren(...strokes.map((stroke) => {
      const path = document.createElementNS(SVG, "path");
      path.setAttribute("d", strokePath(stroke.length > 1 ? stroke : [stroke[0], stroke[0]], w, h));
      path.setAttribute("stroke-width", String(Math.max(2, w / 300)));
      return path;
    }));
  }

  /** Laser dot and the tools bar, over whatever is drawn. */
  private drawTools(o: HTMLElement): void {
    o.createDiv({ cls: "dw-show-laser" });
    const bar = o.createDiv({ cls: "dw-show-tools" });
    const button = (text: string, title: string, run: () => void, tool?: ShowTool) => {
      const b = bar.createEl("button", { text, attr: { "aria-label": title, title } });
      if (tool) b.dataset.tool = tool;
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        run();
        o.focus();
      });
    };
    button("Laser", "Laser pointer (L)", () => this.setTool(this.tool === "laser" ? "none" : "laser"), "laser");
    button("Pen", "Pen (Ctrl+P)", () => this.setTool(this.tool === "pen" ? "none" : "pen"), "pen");
    button("Erase", "Erase ink on this slide (E)", () => this.eraseInk());
    this.setTool(this.tool);
  }

  private onFullscreenChange = () => {
    // The browser eats Escape to leave fullscreen; treat that as ending the show.
    // With a pen or pointer in hand that Escape puts it down instead; the show goes on in the window.
    if (document.fullscreenElement === this.overlay) this.wasFullscreen = true;
    else if (this.wasFullscreen) {
      this.wasFullscreen = false;
      if (this.tool !== "none") this.setTool("none");
      else this.close();
    }
  };

  private tick = () => {
    const s = Math.floor((Date.now() - this.started) / 1000);
    const now = new Date();
    const set = (cls: string, text: string) => {
      const el = this.overlay?.querySelector(cls);
      if (el) el.textContent = text;
    };
    set(".dw-show-timer", `${pad(Math.floor(s / 60))}:${pad(s % 60)}`);
    set(".dw-show-clock", `${pad(now.getHours())}:${pad(now.getMinutes())}`);
  };

  /** Largest slide width that fits a w x h box. */
  private fit(w: number, h: number): number {
    return Math.max(1, Math.min(w, (h * this.deck.size.cx) / this.deck.size.cy));
  }

  private redraw = () => {
    this.draw();
    this.autoPlay();
  };

  /** Draw the slide (built up to the current step), with the transition into it when given. */
  private draw = (transition: Transition | null = null) => {
    const o = this.overlay;
    if (!o) return;
    this.stop();
    const mediaState = new Map<string, {time:number; paused:boolean}>();
    o.querySelectorAll<HTMLMediaElement>("audio,video").forEach(media=>{
      if(this.mediaIndex === this.index) mediaState.set(media.closest<HTMLElement>("[data-spids]")?.dataset.spids ?? media.src,{time:media.currentTime,paused:media.paused});
      media.pause();
    });
    this.mediaIndex=this.index;
    const restoreMedia = () => o.querySelectorAll<HTMLMediaElement>("audio,video").forEach(media=>{
      const state=mediaState.get(media.closest<HTMLElement>("[data-spids]")?.dataset.spids ?? media.src);
      if(state) { media.autoplay=false; const restore=()=>{media.currentTime=state.time;if(!state.paused)void media.play().catch(()=>undefined);}; if(media.readyState)restore();else media.addEventListener("loadedmetadata",restore,{once:true}); }
    });
    const old = transition ? this.stage?.querySelector<HTMLElement>(".pptx-slide") : null;
    this.stage = null;
    o.replaceChildren();
    o.classList.toggle("is-presenter", this.presenter);
    const W = window.innerWidth;
    const H = window.innerHeight;

    if (!this.presenter) {
      if (this.ended) o.createDiv({ cls: "dw-show-end", text: "End of slide show. Click to exit." });
      else { this.drawStage(o.createDiv({ cls: "dw-show-slide" }), this.fit(W, H), old); restoreMedia(); }
      this.drawTools(o);
      return;
    }

    const bar = o.createDiv({ cls: "dw-show-bar" });
    bar.createSpan({ cls: "dw-show-count", text: `${this.index + 1} / ${this.deck.count}` });
    bar.createSpan({ cls: "dw-show-timer", text: "00:00" });
    bar.createSpan({ cls: "dw-show-clock" });
    const restart = bar.createEl("button", { text: "Restart timer" });
    restart.addEventListener("click", () => {
      this.started = Date.now();
      this.tick();
    });
    const exit = bar.createEl("button", { text: "End show" });
    exit.addEventListener("click", () => this.close());

    const main = o.createDiv({ cls: "dw-show-current" });
    if (this.ended) main.createDiv({ cls: "dw-show-end", text: "End of slide show." });
    else this.drawStage(main.createDiv({ cls: "dw-show-slide" }), this.fit(W * 0.62, H - 80), old);

    const side = o.createDiv({ cls: "dw-show-side" });
    const next = side.createDiv({ cls: "dw-show-next" });
    next.createDiv({ cls: "dw-show-label", text: "Next" });
    const n = this.ended ? null : nextVisible(this.hidden(), this.index, 1);
    if (n === null) next.createDiv({ cls: "dw-show-end", text: "End of slide show" });
    else this.deck.render(next.createDiv({ cls: "dw-show-slide" }), n, this.fit(W * 0.3, H * 0.3));
    const notes = this.ended ? "" : this.deck.notes(this.index);
    side.createDiv({ cls: `dw-show-notes${notes ? "" : " is-empty"}`, text: notes || "No notes for this slide." });
    restoreMedia();
    next.querySelectorAll<HTMLMediaElement>("audio,video").forEach(media=>{media.autoplay=false;media.pause();media.controls=false;});
    this.drawTools(o);
    this.tick();
  };

  private drawStage(stage: HTMLElement, width: number, old: HTMLElement | null | undefined): void {
    this.stage = stage;
    this.deck.render(stage, this.index, width);
    showState(stage, this.effects().steps, this.step);
    const fresh = stage.querySelector<HTMLElement>(".pptx-slide");
    const t = this.effects().transition;
    if (old && fresh && t) playTransition(stage, old, fresh, t);
    this.drawInk();
  }
}
