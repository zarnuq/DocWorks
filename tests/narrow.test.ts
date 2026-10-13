import { afterEach, describe, expect, it } from "vitest";
import { openDrawer, watchNarrow } from "../src/presentation/common/narrow";

afterEach(() => document.body.empty());

describe("watchNarrow", () => {
  it("toggles is-narrow below 700 px and reports only changes", () => {
    const root = document.body.createDiv();
    let w = 1000;
    const seen: boolean[] = [];
    const stop = watchNarrow(root, (n) => seen.push(n), () => w);
    stop.check();
    w = 500;
    stop.check();
    stop.check();
    expect(root.hasClass("is-narrow")).toBe(true);
    w = 900;
    stop.check();
    expect(seen).toEqual([true, false]);
    expect(root.hasClass("is-narrow")).toBe(false);
    stop();
  });

  it("starts from the class the pane already has (a file reopened in a narrow pane that is now wide)", () => {
    const root = document.body.createDiv({ cls: "is-narrow" });
    const seen: boolean[] = [];
    watchNarrow(root, (n) => seen.push(n), () => 1000).check();
    expect(seen).toEqual([false]);
    expect(root.hasClass("is-narrow")).toBe(false);
  });

  it("treats an unlaid-out pane (width 0) as wide", () => {
    const root = document.body.createDiv();
    const seen: boolean[] = [];
    watchNarrow(root, (n) => seen.push(n), () => 0).check();
    expect(seen).toEqual([]);
  });
});

describe("openDrawer", () => {
  function setup() {
    const root = document.body.createDiv();
    const panel = root.createDiv();
    panel.hide();
    const toggle = root.createEl("button");
    const outside = root.createDiv();
    let closed = 0;
    const close = openDrawer(panel, { root, toggle, onClose: () => closed++ });
    return { root, panel, toggle, outside, close, closed: () => closed };
  }

  it("shows the panel as a drawer and closes on a click outside", () => {
    const { panel, outside, closed } = setup();
    expect(panel.hasClass("is-drawer")).toBe(true);
    expect(panel.style.display).not.toBe("none");
    outside.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(panel.hasClass("is-drawer")).toBe(false);
    expect(closed()).toBe(1);
  });

  it("closes on Escape inside it, letting the Escape through to the panel's own handler", () => {
    const { panel, closed } = setup();
    const button = panel.createEl("button");
    let reached = false;
    panel.addEventListener("keydown", (e) => (reached = !e.defaultPrevented));
    button.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect([closed(), reached]).toEqual([1, true]);
  });

  it("ignores a pointerdown on its toggle or inside it, and closes once", () => {
    const { panel, toggle, close, closed } = setup();
    toggle.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    panel.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(closed()).toBe(0);
    close();
    close();
    expect(closed()).toBe(1);
  });
});
