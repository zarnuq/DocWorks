import { afterEach, describe, expect, it } from "vitest";
import { Scope } from "./obsidian-stub";
import { escapeToEditor, focusRegions, returnFocusAfterOverlays, rovingToolbar } from "../src/presentation/common/keyboard";

afterEach(() => document.body.empty());

const key = (el: Element, k: string, opts: KeyboardEventInit = {}) => {
  const e = new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...opts });
  el.dispatchEvent(e);
  return e;
};

function bar() {
  const el = document.body.createDiv({ cls: "dw-toolbar" });
  const g1 = el.createDiv({ cls: "dw-group" });
  const a = g1.createEl("button", { text: "A" });
  const sel = g1.createEl("select");
  sel.createEl("option", { text: "x" });
  const g2 = el.createDiv({ cls: "dw-group" });
  const color = g2.createEl("label").createEl("input", { type: "color" });
  const hidden = g2.createEl("button", { text: "hidden" });
  hidden.style.display = "none";
  const b = g2.createEl("button", { text: "B" });
  rovingToolbar(el);
  return { el, a, sel, color, hidden, b };
}

describe("rovingToolbar", () => {
  it("is one Tab stop: the first control, marked as a toolbar", () => {
    const { el, a, sel, color, b } = bar();
    expect(el.getAttribute("role")).toBe("toolbar");
    expect([a, sel, color, b].map((x) => x.tabIndex)).toEqual([0, -1, -1, -1]);
  });

  it("moves with Left / Right / Home / End, skipping hidden controls, and keeps the last one as the Tab stop", () => {
    const { a, sel, color, b } = bar();
    a.focus();
    key(a, "ArrowRight");
    expect(document.activeElement).toBe(sel);
    key(sel, "ArrowRight");
    expect(document.activeElement).toBe(color);
    key(color, "ArrowRight");
    expect(document.activeElement).toBe(b);
    key(b, "ArrowRight");
    expect(document.activeElement).toBe(a);
    key(a, "End");
    expect(document.activeElement).toBe(b);
    key(b, "Home");
    expect(document.activeElement).toBe(a);
    key(a, "ArrowLeft");
    expect(document.activeElement).toBe(b);
    expect([a.tabIndex, b.tabIndex]).toEqual([-1, 0]);
  });

  it("leaves arrow keys alone in text fields", () => {
    const { el, a } = bar();
    const input = el.createEl("input", { type: "text" });
    input.focus();
    const e = key(input, "ArrowLeft");
    expect(e.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(input);
    void a;
  });
});

describe("focusRegions", () => {
  it("cycles F6 / Shift+F6 through visible regions, landing on each one's focus target", () => {
    const scope = new Scope();
    const { el: toolbar, a } = bar();
    const editor = document.body.createDiv({ attr: { tabindex: "0" } });
    const panel = document.body.createDiv();
    const inPanel = panel.createEl("button", { text: "p" });
    const hiddenPanel = document.body.createDiv({ attr: { tabindex: "0" } });
    hiddenPanel.style.display = "none";
    focusRegions(scope as never, document.body, () => [toolbar, editor, hiddenPanel, panel]);
    const f6 = scope.keys.find((k) => k.key === "F6" && !k.mods.length)!;
    const shiftF6 = scope.keys.find((k) => k.key === "F6" && k.mods.join() === "Shift")!;
    editor.focus();
    f6.fn();
    expect(document.activeElement).toBe(inPanel);
    f6.fn();
    expect(document.activeElement).toBe(a);
    shiftF6.fn();
    expect(document.activeElement).toBe(inPanel);
    shiftF6.fn();
    expect(document.activeElement).toBe(editor);
  });
});

describe("focusRegions with a menu row", () => {
  it("treats a menu row like a toolbar: F6 lands on its current item", () => {
    const scope = new Scope();
    const menubar = document.body.createDiv();
    menubar.createEl("button", { text: "File" });
    const edit = menubar.createEl("button", { text: "Edit" });
    rovingToolbar(menubar);
    menubar.setAttribute("role", "menubar");
    const editor = document.body.createDiv({ attr: { tabindex: "0" } });
    focusRegions(scope as never, document.body, () => [menubar, editor]);
    edit.focus();
    editor.focus();
    scope.keys.find((k) => k.key === "F6" && !k.mods.length)!.fn();
    expect(document.activeElement).toBe(edit);
  });
});

describe("escapeToEditor", () => {
  it("sends Escape in a panel back to the editor, unless something else handled it", () => {
    const panel = document.body.createDiv();
    const button = panel.createEl("button");
    let focused = 0;
    escapeToEditor(panel, () => focused++);
    button.focus();
    key(button, "Escape");
    expect(focused).toBe(1);
    button.addEventListener("keydown", (e) => e.preventDefault(), { once: true });
    key(button, "Escape");
    expect(focused).toBe(1);
  });
});

describe("returnFocusAfterOverlays", () => {
  it("gives the editor focus back when a dialog closes and focus is lost", async () => {
    let focused = 0;
    const stop = returnFocusAfterOverlays(() => true, () => focused++);
    const modal = document.body.createDiv({ cls: "modal-container" });
    await new Promise((r) => setTimeout(r, 0));
    modal.remove();
    await new Promise((r) => setTimeout(r, 0));
    await new Promise((r) => setTimeout(r, 0));
    expect(focused).toBe(1);
    // Focus somewhere real (a field the dialog's action opened) is left alone.
    const other = document.body.createDiv({ cls: "modal-container" });
    const field = document.body.createEl("input");
    other.remove();
    field.focus();
    await new Promise((r) => setTimeout(r, 0));
    await new Promise((r) => setTimeout(r, 0));
    expect(focused).toBe(1);
    stop();
  });

  it("does nothing for views that aren't active", async () => {
    let focused = 0;
    const stop = returnFocusAfterOverlays(() => false, () => focused++);
    const modal = document.body.createDiv({ cls: "modal-container" });
    await new Promise((r) => setTimeout(r, 0));
    modal.remove();
    await new Promise((r) => setTimeout(r, 0));
    await new Promise((r) => setTimeout(r, 0));
    expect(focused).toBe(0);
    stop();
  });
});
