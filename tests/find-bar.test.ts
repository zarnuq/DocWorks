import { describe, expect, it, vi } from "vitest";
import { findBar } from "../src/presentation/common/find-bar";

function setup() {
  const host = createDiv();
  const body = host.createDiv({ cls: "body" });
  const calls = { step: vi.fn(), changed: vi.fn(), replace: vi.fn(), replaceAll: vi.fn(), close: vi.fn() };
  const bar = findBar(host, body, {
    options: [["Aa", "Match case", false], ["Comments", "Search comments", true]],
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
    expect(bar.options.map((b) => [b.getAttribute("aria-label"), b.checked])).toEqual([["Match case", false], ["Search comments", true]]);
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
});
