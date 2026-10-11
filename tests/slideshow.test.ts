import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { nextVisible, ShowDeck, SlideEffects, Slideshow } from "../src/presentation/pptx/slideshow";
import type { Behavior, Effect } from "../src/domain/pptx/animation";

const deck = (hidden: boolean[] = [false, false, false, false]): ShowDeck => ({
  size: { cx: 1600, cy: 900 },
  count: hidden.length,
  isHidden: (i) => hidden[i],
  notes: (i) => `notes ${i}`,
  render: (host, i) => {
    host.textContent = `slide ${i}`;
  },
});

const key = (el: HTMLElement, k: string) => el.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true }));

describe("nextVisible", () => {
  const hidden = [false, true, false, true];
  it("skips hidden slides in both directions", () => {
    expect(nextVisible(hidden, 0, 1)).toBe(2);
    expect(nextVisible(hidden, 2, -1)).toBe(0);
  });
  it("returns null past either end", () => {
    expect(nextVisible(hidden, 2, 1)).toBeNull();
    expect(nextVisible(hidden, 0, -1)).toBeNull();
  });
});

describe("Slideshow", () => {
  let exit: ReturnType<typeof vi.fn<(index: number) => void>>;
  let show: Slideshow;
  const overlay = () => document.querySelector<HTMLElement>(".dw-show")!;

  beforeEach(() => {
    document.body.innerHTML = "";
    exit = vi.fn<(index: number) => void>();
  });
  afterEach(() => show?.close());

  it("opens full-window on the start slide", () => {
    show = new Slideshow(deck(), 1, exit);
    show.open();
    expect(overlay().textContent).toContain("slide 1");
  });

  it("advances and goes back with the keyboard, skipping hidden slides", () => {
    show = new Slideshow(deck([false, true, false, false]), 0, exit);
    show.open();
    key(overlay(), "ArrowRight");
    expect(show.index).toBe(2);
    key(overlay(), " ");
    expect(show.index).toBe(3);
    key(overlay(), "PageUp");
    expect(show.index).toBe(2);
    key(overlay(), "Home");
    expect(show.index).toBe(0);
    key(overlay(), "End");
    expect(show.index).toBe(3);
  });

  it("shows an end screen after the last slide, then exits", () => {
    show = new Slideshow(deck([false, false]), 1, exit);
    show.open();
    key(overlay(), "ArrowRight");
    expect(overlay().textContent).toMatch(/end of slide show/i);
    key(overlay(), "ArrowRight");
    expect(document.querySelector(".dw-show")).toBeNull();
    expect(exit).toHaveBeenCalledWith(1);
  });

  it("jumps to a typed slide number on Enter", () => {
    show = new Slideshow(deck(), 0, exit);
    show.open();
    key(overlay(), "3");
    key(overlay(), "Enter");
    expect(show.index).toBe(2);
  });

  it("exits on Escape and reports the current slide", () => {
    show = new Slideshow(deck(), 2, exit);
    show.open();
    key(overlay(), "Escape");
    expect(document.querySelector(".dw-show")).toBeNull();
    expect(exit).toHaveBeenCalledWith(2);
  });

  it("presenter view shows notes, the next slide, a counter and a timer", () => {
    show = new Slideshow(deck([false, true, false, false]), 0, exit);
    show.open();
    key(overlay(), "p");
    const o = overlay();
    expect(o.classList.contains("is-presenter")).toBe(true);
    expect(o.querySelector(".dw-show-notes")!.textContent).toBe("notes 0");
    expect(o.querySelector(".dw-show-next")!.textContent).toContain("slide 2");
    expect(o.querySelector(".dw-show-count")!.textContent).toBe("1 / 4");
    expect(o.querySelector(".dw-show-timer")!.textContent).toBe("00:00");
    key(o, "p");
    expect(overlay().classList.contains("is-presenter")).toBe(false);
  });

  it("opens links instead of advancing", () => {
    const open = vi.spyOn(window, "open").mockReturnValue(null);
    show = new Slideshow({ ...deck(), render: (host) => (host.innerHTML = '<span data-link="https://x.y/">link</span>') }, 0, exit);
    show.open();
    overlay().querySelector<HTMLElement>("[data-link]")!.click();
    expect(open).toHaveBeenCalledWith("https://x.y/");
    expect(show.index).toBe(0);
    open.mockRestore();
  });

  it("says so when a slide has no notes", () => {
    show = new Slideshow({ ...deck(), notes: () => "" }, 0, exit);
    show.open();
    key(overlay(), "p");
    expect(overlay().querySelector(".dw-show-notes")!.textContent).toBe("No notes for this slide.");
  });
});

describe("Slideshow animations", () => {
  let show: Slideshow;
  const overlay = () => document.querySelector<HTMLElement>(".dw-show")!;
  const shape = (id: string) => overlay().querySelector<HTMLElement>(`[data-spids~="${id}"]`)!;
  const vis = (b: Behavior[] = []) => [{ type: "visibility" as const, visible: true, start: 0, dur: 1 }, ...b];
  const enter = (spid: string, para: [number, number] | null = null, start = 0): Effect => ({ spid, para, cls: "entr", start, dur: 500, behaviors: vis() });
  const exitFx = (spid: string): Effect => ({ spid, para: null, cls: "exit", start: 0, dur: 500, behaviors: [{ type: "visibility", visible: false, start: 499, dur: 1 }] });
  const animDeck = (effects: (i: number) => SlideEffects, count = 2): ShowDeck => ({
    ...deck(Array(count).fill(false)),
    render: (host, i) => {
      host.innerHTML = `<div class="pptx-slide"><div class="pptx-shape" data-spids="3"><div class="pptx-text"><p>a</p><p>b</p></div></div><div class="pptx-shape" data-spids="4 9">s${i}</div></div>`;
    },
    effects,
  });

  beforeEach(() => {
    document.body.innerHTML = "";
    vi.useFakeTimers();
  });
  afterEach(() => {
    show?.close();
    vi.useRealTimers();
  });

  it("hides shapes that enter later and shows them one click at a time", () => {
    show = new Slideshow(animDeck((i) => ({ transition: null, steps: i ? [] : [{ auto: false, effects: [enter("3")] }, { auto: false, effects: [enter("4")] }] })), 0, vi.fn());
    show.open();
    expect(shape("3").style.visibility).toBe("hidden");
    expect(shape("4").style.visibility).toBe("hidden");
    key(overlay(), "ArrowRight");
    expect(show.index).toBe(0);
    expect(shape("3").style.visibility).toBe("visible");
    expect(shape("4").style.visibility).toBe("hidden");
    vi.advanceTimersByTime(1000);
    key(overlay(), "ArrowRight");
    expect(shape("4").style.visibility).toBe("visible");
    vi.advanceTimersByTime(1000);
    key(overlay(), "ArrowRight");
    expect(show.index).toBe(1);
  });

  it("animates the shapes inside an animated group", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("9")] }] })), 0, vi.fn());
    show.open();
    expect(shape("4").style.visibility).toBe("hidden");
  });

  it("builds text paragraph by paragraph", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("3", [1, 1])] }] })), 0, vi.fn());
    show.open();
    const [a, b] = Array.from(shape("3").querySelectorAll("p"));
    expect(a.style.visibility).toBe("");
    expect(b.style.visibility).toBe("hidden");
    key(overlay(), "ArrowRight");
    expect(b.style.visibility).toBe("visible");
  });

  it("plays delayed and exit effects on time", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("4", null, 1000), exitFx("3")] }] })), 0, vi.fn());
    show.open();
    expect(shape("3").style.visibility).toBe("");
    key(overlay(), "ArrowRight");
    expect(shape("4").style.visibility).toBe("hidden");
    vi.advanceTimersByTime(600);
    expect(shape("3").style.visibility).toBe("hidden");
    vi.advanceTimersByTime(500);
    expect(shape("4").style.visibility).toBe("visible");
  });

  it("finishes a playing step on click instead of skipping ahead", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("4", null, 1000)] }, { auto: false, effects: [enter("3")] }] })), 0, vi.fn());
    show.open();
    key(overlay(), "ArrowRight");
    key(overlay(), "ArrowRight");
    expect(shape("4").style.visibility).toBe("visible");
    expect(shape("3").style.visibility).toBe("hidden");
  });

  it("plays automatic steps when the slide starts and after the previous step", () => {
    show = new Slideshow(animDeck((i) => ({ transition: null, steps: i ? [] : [{ auto: true, effects: [enter("3")] }, { auto: true, effects: [enter("4")] }] })), 0, vi.fn());
    show.open();
    expect(shape("3").style.visibility).toBe("visible");
    vi.advanceTimersByTime(500);
    expect(shape("4").style.visibility).toBe("visible");
  });

  it("continues automatic effects when fullscreen resizing redraws a playing slide", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("3")] }, { auto: true, effects: [exitFx("3")] }] })), 0, vi.fn());
    show.open();
    key(overlay(), "ArrowRight");
    window.dispatchEvent(new Event("resize"));
    vi.advanceTimersByTime(500);
    expect(shape("3").style.visibility).toBe("hidden");
  });

  it("continues automatic effects after a click finishes the playing step", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("3")] }, { auto: true, effects: [exitFx("3")] }] })), 0, vi.fn());
    show.open();
    key(overlay(), "ArrowRight");
    key(overlay(), "ArrowRight");
    vi.advanceTimersByTime(500);
    expect(shape("3").style.visibility).toBe("hidden");
  });

  it("goes back one step, then to the previous slide fully built", () => {
    show = new Slideshow(animDeck(() => ({ transition: null, steps: [{ auto: false, effects: [enter("3")] }] })), 0, vi.fn());
    show.open();
    key(overlay(), "ArrowRight");
    key(overlay(), "ArrowLeft");
    expect(shape("3").style.visibility).toBe("hidden");
    key(overlay(), "ArrowRight");
    vi.advanceTimersByTime(1000);
    key(overlay(), "ArrowRight");
    expect(show.index).toBe(1);
    key(overlay(), "ArrowLeft");
    expect(show.index).toBe(0);
    expect(shape("3").style.visibility).toBe("visible");
  });

  it("advances by itself after the slide's timing", () => {
    show = new Slideshow(animDeck((i) => ({ transition: i ? null : { type: "fade", dir: null, dur: 500, advanceAfter: 2000 }, steps: [] })), 0, vi.fn());
    show.open();
    vi.advanceTimersByTime(1999);
    expect(show.index).toBe(0);
    vi.advanceTimersByTime(1);
    expect(show.index).toBe(1);
  });

  it("keeps the previous slide on screen while the next slide's transition plays", () => {
    const anims: { finish: () => void }[] = [];
    const proto = HTMLElement.prototype as unknown as { animate?: unknown };
    proto.animate = function () {
      let done = () => {};
      const finished = new Promise<void>((r) => (done = r));
      const a = { finished, cancel: () => {}, finish: () => done() };
      anims.push(a);
      return a;
    };
    try {
      show = new Slideshow(animDeck((i) => ({ transition: i ? { type: "push", dir: "l", dur: 500, advanceAfter: null } : null, steps: [] })), 0, vi.fn());
      show.open();
      key(overlay(), "ArrowRight");
      expect(overlay().querySelectorAll(".pptx-slide")).toHaveLength(2);
      expect(anims.length).toBeGreaterThan(0);
    } finally {
      delete proto.animate;
    }
  });

  it("previews the current slide's entry transition before playing automatic effects", () => {
    const animate = vi.fn(() => ({ finished: Promise.resolve(), cancel: vi.fn() }));
    Object.defineProperty(HTMLElement.prototype, "animate", { configurable: true, value: animate });
    try {
      show = new Slideshow(animDeck(() => ({ transition: { type: "fade", dir: null, dur: 800, advanceAfter: null }, steps: [{ auto: true, effects: [enter("3")] }] })), 0, vi.fn());
      show.open(true);
      expect(animate).toHaveBeenCalledWith([{ opacity: 0 }, { opacity: 1 }], expect.objectContaining({ duration: 800 }));
      expect(shape("3").style.visibility).toBe("hidden");
      vi.advanceTimersByTime(800);
      expect(shape("3").style.visibility).toBe("visible");
    } finally {
      delete (HTMLElement.prototype as unknown as { animate?: unknown }).animate;
    }
  });
});
