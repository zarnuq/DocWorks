import { describe, expect, it, vi } from "vitest";
import { findBar } from "../src/presentation/common/find-bar";

function setup(key?: string) {
  const host = createDiv();
  const body = host.createDiv({ cls: "body" });
  const calls = { step: vi.fn(), changed: vi.fn(), replace: vi.fn(), replaceAll: vi.fn(), close: vi.fn() };
  const bar = findBar(host, body, {
    key,
    options: [["Comments", "Search comments", true]],
    replaceTitles: ["Replace this match", "Replace every match"],
    ...calls,
  });
  return { host, bar, calls };
}

describe("findBar", () => {
  it("puts the bar above the body with find, replace, options, count and buttons", () => {
    const { host, bar } = setup();
    expect(host.firstElementChild).toBe(bar.el);
    expect(bar.input.getAttribute("aria-label")).toBe("Find");
    expect(bar.replaceInput.getAttribute("aria-label")).toBe("Replace with");
    expect(Array.from(bar.el.querySelectorAll<HTMLInputElement>("input[type=checkbox]")).map((b) => [b.getAttribute("aria-label"), b.checked])).toEqual([
      ["Match case", false], ["Match whole words", false], ["Match using regular expressions", false], ["Ignore accents", false], ["Search comments", true],
    ]);
    expect(bar.options.map((b) => b.getAttribute("aria-label"))).toEqual(["Search comments"]);
    expect(Array.from(bar.el.querySelectorAll("button")).map((b) => b.getAttribute("aria-label"))).toEqual(["Previous match (Shift+Enter)", "Next match (Enter)", "Replace this match", "Replace every match", "Close (Esc)"]);
    expect(bar.el.lastElementChild).toBe(bar.preview);
  });

  it("wires typing, options, keys and buttons to the view", () => {
    const { bar, calls } = setup();
    bar.input.dispatchEvent(new Event("input"));
    bar.options[0].dispatchEvent(new Event("change"));
    expect(calls.changed.mock.calls).toEqual([[true], [false]]);
    bar.input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    bar.input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", shiftKey: true, bubbles: true }));
    expect(calls.step.mock.calls).toEqual([[1], [-1]]);
    bar.replaceButton.click();
    bar.replaceAllButton.click();
    expect(calls.replace).toHaveBeenCalledOnce();
    expect(calls.replaceAll).toHaveBeenCalledOnce();
    bar.replaceInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(calls.close).toHaveBeenCalledOnce();
  });

  it("shows the count", () => {
    const { bar } = setup();
    bar.setCount("", 0, 0);
    expect(bar.count.textContent).toBe("");
    bar.setCount("x", -1, 0);
    expect(bar.count.textContent).toBe("No results");
    bar.setCount("x", 1, 3);
    expect(bar.count.textContent).toBe("2 of 3");
  });

  it("reads Google's options", () => {
    const { bar } = setup();
    expect(bar.search()).toEqual({ matchCase: false, wholeWord: false, regex: false, ignoreAccents: false });
    for (const b of Array.from(bar.el.querySelectorAll<HTMLInputElement>("input[type=checkbox]")).slice(0, 4)) b.checked = true;
    expect(bar.search()).toEqual({ matchCase: true, wholeWord: true, regex: true, ignoreAccents: true });
  });

  it("remembers the last search, replacement and options per file", () => {
    const first = setup("a.docx").bar;
    first.input.value = "needle";
    first.input.dispatchEvent(new Event("input"));
    first.replaceInput.value = "pin";
    first.replaceInput.dispatchEvent(new Event("input"));
    const regex = first.el.querySelector<HTMLInputElement>('[aria-label="Match using regular expressions"]')!;
    regex.checked = true;
    regex.dispatchEvent(new Event("change"));
    first.options[0].checked = false;
    first.options[0].dispatchEvent(new Event("change"));
    const again = setup("a.docx").bar;
    expect([again.input.value, again.replaceInput.value, again.search().regex, again.options[0].checked]).toEqual(["needle", "pin", true, false]);
    expect(setup("b.docx").bar.input.value).toBe("");
  });
});
