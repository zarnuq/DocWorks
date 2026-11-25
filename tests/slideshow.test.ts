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

describe("Slideshow links to slides", () => {
  afterEach(() => document.body.empty());

  it("jumps to the linked slide instead of opening a page or advancing", () => {
    const d: ShowDeck = {
      ...deck(),
      slidePaths: ["s1", "s2", "s3", "s4"],
      render: (host, i) => {
        host.textContent = `slide ${i}`;
        host.createSpan({ cls: "go", attr: { "data-link": i === 0 ? "#slide:s3" : "#jump:firstslide" } });
      },
    };
    const open = vi.spyOn(window, "open").mockImplementation(() => null);
    const show = new Slideshow(d, 0, vi.fn());
    show.open();
    document.querySelector<HTMLElement>(".dw-show .go")!.click();
    expect(show.index).toBe(2);
    document.querySelector<HTMLElement>(".dw-show .go")!.click();
    expect(show.index).toBe(0);
    expect(open).not.toHaveBeenCalled();
    show.close();
  });
});

describe("laser pointer and pen", () => {
  let show: Slideshow;
  const overlay = () => document.querySelector<HTMLElement>(".dw-show")!;
  const stage = () => overlay().querySelector<HTMLElement>(".dw-show-slide")!;
  const keyWith = (k: string, init: KeyboardEventInit = {}) => overlay().dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true, ...init }));
  const pointer = (type: string, x: number, y: number) => stage().dispatchEvent(new PointerEvent(type, { clientX: x, clientY: y, button: 0, bubbles: true }));
  beforeEach(() => {
    document.body.innerHTML = "";
    show = new Slideshow(deck(), 0, () => undefined);
    show.open();
    // jsdom lays nothing out: the slide is 800 x 450 at the window's top left.
    vi.spyOn(HTMLElement.prototype, "getBoundingClientRect").mockReturnValue({ left: 0, top: 0, right: 800, bottom: 450, width: 800, height: 450, x: 0, y: 0, toJSON: () => ({}) });
  });
  afterEach(() => {
    vi.restoreAllMocks();
    show?.close();
  });

  it("L turns the laser on and off; the dot follows the mouse", () => {
    keyWith("l");
    expect(show.tool).toBe("laser");
    expect(overlay().classList.contains("is-laser")).toBe(true);
    overlay().dispatchEvent(new PointerEvent("pointermove", { clientX: 120, clientY: 80, bubbles: true }));
    const dot = overlay().querySelector<HTMLElement>(".dw-show-laser")!;
    expect(dot.style.getPropertyValue("--dw-laser-x")).toBe("120px");
    expect(dot.style.getPropertyValue("--dw-laser-y")).toBe("80px");
    keyWith("L");
    expect(show.tool).toBe("none");
  });

  it("with the laser on a click still advances", () => {
    keyWith("l");
    stage().click();
    expect(show.index).toBe(1);
    expect(show.tool).toBe("laser");
    expect(overlay().classList.contains("is-laser")).toBe(true);
  });

  it("Ctrl+P is the pen (plain P stays presenter view); a drag draws a stroke and doesn't advance", () => {
    keyWith("p", { ctrlKey: true });
    expect(show.tool).toBe("pen");
    expect(overlay().classList.contains("is-presenter")).toBe(false);
    pointer("pointerdown", 80, 45);
    pointer("pointermove", 400, 225);
    pointer("pointerup", 400, 225);
    stage().click();
    expect(show.index).toBe(0);
    const path = stage().querySelector(".dw-show-ink path")!;
    expect(path.getAttribute("d")).toMatch(/^M\d/);
  });

  it("ink belongs to its slide and comes back when the show returns to it", () => {
    keyWith("p", { ctrlKey: true });
    pointer("pointerdown", 10, 10);
    pointer("pointermove", 50, 50);
    pointer("pointerup", 50, 50);
    keyWith("ArrowRight");
    expect(stage().querySelector(".dw-show-ink")).toBeNull();
    keyWith("ArrowLeft");
    expect(stage().querySelectorAll(".dw-show-ink path")).toHaveLength(1);
  });

  it("ink survives a switch to presenter view (it's stored relative to the slide)", () => {
    keyWith("p", { ctrlKey: true });
    pointer("pointerdown", 400, 225);
    pointer("pointermove", 800, 450);
    pointer("pointerup", 800, 450);
    keyWith("p");
    expect(overlay().classList.contains("is-presenter")).toBe(true);
    expect(overlay().querySelectorAll(".dw-show-current .dw-show-ink path")).toHaveLength(1);
  });

  it("E erases the slide's ink", () => {
    keyWith("p", { ctrlKey: true });
    pointer("pointerdown", 10, 10);
    pointer("pointerup", 10, 10);
    expect(stage().querySelector(".dw-show-ink")).not.toBeNull();
    keyWith("e");
    expect(stage().querySelector(".dw-show-ink")).toBeNull();
  });

  it("a stroke outside the slide isn't started", () => {
    keyWith("p", { ctrlKey: true });
    overlay().dispatchEvent(new PointerEvent("pointerdown", { clientX: 900, clientY: 500, button: 0, bubbles: true }));
    expect(stage().querySelector(".dw-show-ink")).toBeNull();
  });

  it("Escape puts the tool down first, then ends the show", () => {
    keyWith("l");
    keyWith("Escape");
    expect(show.tool).toBe("none");
    expect(overlay()).not.toBeNull();
    keyWith("Escape");
    expect(document.querySelector(".dw-show")).toBeNull();
  });

  it("the tools bar has Laser, Pen and Erase; the active tool is marked", () => {
    const buttons = () => Array.from(overlay().querySelectorAll<HTMLButtonElement>(".dw-show-tools button"));
    expect(buttons().map((b) => b.textContent)).toEqual(["Laser", "Pen", "Erase"]);
    buttons()[1].click();
    expect(show.tool).toBe("pen");
    expect(show.index).toBe(0);
    expect(buttons()[1].classList.contains("is-active")).toBe(true);
    keyWith("ArrowRight");
    expect(buttons()[1].classList.contains("is-active")).toBe(true);
  });
});

describe("strokePath", () => {
  it("scales fractions to the slide", async () => {
    const { strokePath } = await import("../src/presentation/pptx/slideshow");
    expect(strokePath([[0, 0], [0.5, 1]], 200, 100)).toBe("M0.0 0.0 L100.0 100.0");
  });
});

describe("leaving fullscreen", () => {
  it("with a tool in hand puts it down and keeps the show; without one ends it", () => {
    document.body.innerHTML = "";
    const show = new Slideshow(deck(), 0, () => undefined);
    show.open();
    const o = document.querySelector<HTMLElement>(".dw-show")!;
    const fullscreen = (el: Element | null) => {
      Object.defineProperty(document, "fullscreenElement", { configurable: true, get: () => el });
      document.dispatchEvent(new Event("fullscreenchange"));
    };
    fullscreen(o);
    show.setTool("pen");
    fullscreen(null);
    expect(show.tool).toBe("none");
    expect(document.querySelector(".dw-show")).not.toBeNull();
    fullscreen(o);
    fullscreen(null);
    expect(document.querySelector(".dw-show")).toBeNull();
    Object.defineProperty(document, "fullscreenElement", { configurable: true, get: () => null });
  });
});
