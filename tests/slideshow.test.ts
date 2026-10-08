import { describe, it, expect, beforeEach, afterEach, vi } from "vitest";
import { nextVisible, ShowDeck, Slideshow } from "../src/presentation/pptx/slideshow";

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

  it("says so when a slide has no notes", () => {
    show = new Slideshow({ ...deck(), notes: () => "" }, 0, exit);
    show.open();
    key(overlay(), "p");
    expect(overlay().querySelector(".dw-show-notes")!.textContent).toBe("No notes for this slide.");
  });
});
