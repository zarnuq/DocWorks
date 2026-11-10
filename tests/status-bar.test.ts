import { afterEach, describe, expect, it } from "vitest";
import { Menu } from "./obsidian-stub";
import { StatusBar, type StatusItem } from "../src/presentation/common/status-bar";

afterEach(() => document.body.empty());

function storage() {
  const data = new Map<string, unknown>();
  return { loadLocalStorage: (k: string) => data.get(k) ?? null, saveLocalStorage: (k: string, v: unknown) => void data.set(k, v), data };
}

function setup(store = storage()) {
  let page = 1;
  const items: StatusItem[] = [
    { id: "page", label: "Page Number", side: "left", render: (el) => el.setText(`Page ${page} of 3`) },
    { id: "words", label: "Word Count", side: "left", render: (el) => el.setText("12 words") },
    { id: "views", label: "View Shortcuts", side: "right", render: (el) => el.createEl("button", { text: "Print" }) },
  ];
  const el = document.body.createDiv();
  const bar = new StatusBar(el, items, store, "doc-works-sb-test");
  return { el, bar, store, setPage: (n: number) => (page = n) };
}

const shown = (el: HTMLElement) => Array.from(el.querySelectorAll<HTMLElement>(".dw-sb-item")).filter((i) => !i.hidden).map((i) => i.textContent);

describe("status bar", () => {
  it("shows left items, then right items, and redraws them on update", () => {
    const { el, bar, setPage } = setup();
    expect(Array.from(el.querySelectorAll(".dw-sb-left .dw-sb-item")).map((i) => i.textContent)).toEqual(["Page 1 of 3", "12 words"]);
    expect(el.querySelector(".dw-sb-right .dw-sb-item")!.textContent).toBe("Print");
    setPage(2);
    bar.update();
    expect(shown(el)[0]).toBe("Page 2 of 3");
  });

  it("right-click customizes which items show, and remembers it", () => {
    const { el, store } = setup();
    el.querySelector(".dw-statusbar")!.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    expect(Menu.last!.items.map((i) => [i.title, i.checked])).toEqual([["Page Number", true], ["Word Count", true], ["View Shortcuts", true]]);
    Menu.last!.items[1].run!();
    expect(shown(el)).toEqual(["Page 1 of 3", "Print"]);
    expect(store.data.get("doc-works-sb-test-status")).toEqual(["words"]);
    const again = setup(store);
    expect(shown(again.el).filter((t) => t === "12 words")).toEqual([]);
  });

  it("has a zoom slider with − and + stepping through the levels, and the percentage opens the zoom menu", () => {
    const { el, bar } = setup();
    let zoom = "1";
    bar.zoom({ get: () => zoom, set: (z) => (zoom = z), levels: ["fit", "0.5", "0.75", "1", "1.5", "2"] });
    const pct = el.querySelector<HTMLButtonElement>(".dw-sb-zoom-pct")!;
    expect(pct.textContent).toBe("100%");
    el.querySelector<HTMLButtonElement>('.dw-sb-zoom [aria-label="Zoom in"]')!.click();
    expect(zoom).toBe("1.5");
    expect(pct.textContent).toBe("150%");
    el.querySelector<HTMLButtonElement>('.dw-sb-zoom [aria-label="Zoom out"]')!.click();
    el.querySelector<HTMLButtonElement>('.dw-sb-zoom [aria-label="Zoom out"]')!.click();
    expect(zoom).toBe("0.75");
    const slider = el.querySelector<HTMLInputElement>('.dw-sb-zoom input[type="range"]')!;
    slider.value = "4";
    slider.dispatchEvent(new Event("input"));
    expect(zoom).toBe("2");
    pct.click();
    expect(Menu.last!.items.map((i) => [i.title, i.checked])).toEqual([["Fit", false], ["50%", false], ["75%", false], ["100%", false], ["150%", false], ["200%", true]]);
    Menu.last!.items[0].run!();
    expect(zoom).toBe("fit");
    expect(pct.textContent).toBe("Fit");
    // From "fit", the buttons step from 100%.
    el.querySelector<HTMLButtonElement>('.dw-sb-zoom [aria-label="Zoom in"]')!.click();
    expect(zoom).toBe("1.5");
  });

  it("has a place for the save status", () => {
    const { bar } = setup();
    expect(bar.save.classList.contains("dw-sb-save")).toBe(true);
  });
});
