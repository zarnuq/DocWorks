/*
 * Full-window slideshow with an optional presenter view (current slide, next
 * slide, notes, timer). Plain DOM: the deck is supplied by the caller.
 */

export interface ShowDeck {
  size: { cx: number; cy: number };
  count: number;
  isHidden(i: number): boolean;
  notes(i: number): string;
  /** Draw slide `i` into `host` at `width` px. */
  render(host: HTMLElement, i: number, width: number): void;
}

/** Next slide from `from` in direction `dir` that isn't hidden, or null past the end. */
export function nextVisible(hidden: boolean[], from: number, dir: 1 | -1): number | null {
  for (let i = from + dir; i >= 0 && i < hidden.length; i += dir) if (!hidden[i]) return i;
  return null;
}

const NEXT = new Set(["ArrowRight", "ArrowDown", "PageDown", " ", "Enter", "n", "N"]);
const PREV = new Set(["ArrowLeft", "ArrowUp", "PageUp", "Backspace"]);
const pad = (n: number) => String(n).padStart(2, "0");

export class Slideshow {
  private overlay: HTMLElement | null = null;
  private presenter = false;
  private ended = false;
  private typed = "";
  private started = Date.now();
  private timer: number | null = null;
  private wasFullscreen = false;

  constructor(private deck: ShowDeck, public index: number, private onExit: (index: number) => void) {}

  open(): void {
    const o = (this.overlay = document.body.createDiv({ cls: "dw-show", attr: { tabindex: "-1" } }));
    o.addEventListener("keydown", this.onKey);
    o.addEventListener("click", (e) => {
      if (!(e.target as HTMLElement).closest("button, .dw-show-notes")) this.next();
    });
    o.addEventListener("wheel", (e) => (e.deltaY > 0 ? this.next() : this.prev()), { passive: true });
    document.addEventListener("fullscreenchange", this.onFullscreenChange);
    window.addEventListener("resize", this.draw);
    this.timer = window.setInterval(this.tick, 1000);
    this.draw();
    o.focus();
    o.requestFullscreen?.().catch(() => {
      // Fullscreen can be refused; the overlay still covers the window.
    });
  }

  close(): void {
    if (!this.overlay) return;
    if (this.timer) window.clearInterval(this.timer);
    document.removeEventListener("fullscreenchange", this.onFullscreenChange);
    window.removeEventListener("resize", this.draw);
    if (document.fullscreenElement === this.overlay) void document.exitFullscreen?.().catch(() => undefined);
    this.overlay.remove();
    this.overlay = null;
    this.onExit(this.index);
  }

  private hidden(): boolean[] {
    return Array.from({ length: this.deck.count }, (_, i) => this.deck.isHidden(i));
  }

  next(): void {
    if (this.ended) return this.close();
    const n = nextVisible(this.hidden(), this.index, 1);
    if (n === null) this.ended = true;
    else this.index = n;
    this.draw();
  }

  prev(): void {
    if (this.ended) this.ended = false;
    else this.index = nextVisible(this.hidden(), this.index, -1) ?? this.index;
    this.draw();
  }

  goTo(i: number): void {
    if (i < 0 || i >= this.deck.count) return;
    this.ended = false;
    this.index = i;
    this.draw();
  }

  private onKey = (e: KeyboardEvent) => {
    const k = e.key;
    if (/^\d$/.test(k)) {
      this.typed += k;
      return;
    }
    if (k === "Enter" && this.typed) {
      this.goTo(Number(this.typed) - 1);
      this.typed = "";
    } else if (k === "Escape") this.close();
    else if (k === "p" || k === "P") {
      this.presenter = !this.presenter;
      this.draw();
    } else if (k === "Home") this.goTo(nextVisible(this.hidden(), -1, 1) ?? 0);
    else if (k === "End") this.goTo(nextVisible(this.hidden(), this.deck.count, -1) ?? this.deck.count - 1);
    else if (NEXT.has(k)) this.next();
    else if (PREV.has(k)) this.prev();
    else return;
    e.preventDefault();
    this.typed = "";
  };

  private onFullscreenChange = () => {
    // The browser eats Escape to leave fullscreen; treat that as ending the show.
    if (document.fullscreenElement === this.overlay) this.wasFullscreen = true;
    else if (this.wasFullscreen) this.close();
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

  private draw = () => {
    const o = this.overlay;
    if (!o) return;
    o.replaceChildren();
    o.classList.toggle("is-presenter", this.presenter);
    const W = window.innerWidth;
    const H = window.innerHeight;

    if (!this.presenter) {
      if (this.ended) o.createDiv({ cls: "dw-show-end", text: "End of slide show. Click to exit." });
      else this.deck.render(o.createDiv({ cls: "dw-show-slide" }), this.index, this.fit(W, H));
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
    else this.deck.render(main.createDiv({ cls: "dw-show-slide" }), this.index, this.fit(W * 0.62, H - 80));

    const side = o.createDiv({ cls: "dw-show-side" });
    const next = side.createDiv({ cls: "dw-show-next" });
    next.createDiv({ cls: "dw-show-label", text: "Next" });
    const n = this.ended ? null : nextVisible(this.hidden(), this.index, 1);
    if (n === null) next.createDiv({ cls: "dw-show-end", text: "End of slide show" });
    else this.deck.render(next.createDiv({ cls: "dw-show-slide" }), n, this.fit(W * 0.3, H * 0.3));
    const notes = this.ended ? "" : this.deck.notes(this.index);
    side.createDiv({ cls: `dw-show-notes${notes ? "" : " is-empty"}`, text: notes || "No notes for this slide." });
    this.tick();
  };
}
