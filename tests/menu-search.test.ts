import { afterEach, describe, expect, it } from "vitest";
import { Menu } from "./obsidian-stub";
import { menuItem, subMenu, Toolbar } from "../src/presentation/common/toolbar";
import { matchMenus, recordMenu, type MenuEntry } from "../src/presentation/common/menu-search";
import type { Menu as ObsidianMenu } from "obsidian";

afterEach(() => document.body.empty());

const entry = (title: string, path: string[] = [], more: Partial<MenuEntry> = {}): MenuEntry => ({ title, path, run: () => {}, ...more });
const titles = (hits: MenuEntry[]) => hits.map((h) => h.title);

describe("recordMenu", () => {
  it("lists every item with its menu path, skipping separators and keeping checked / disabled / icon", () => {
    const entries = recordMenu("Format", (m) => {
      menuItem(m, "Bold", () => {}, { checked: true, icon: "bold" });
      m.addSeparator();
      menuItem(m, "Italic", () => {}, { disabled: true });
    });
    expect(entries).toEqual([
      { title: "Bold", path: ["Format"], run: expect.any(Function), checked: true, icon: "bold" },
      { title: "Italic", path: ["Format"], run: expect.any(Function), disabled: true },
    ]);
  });

  it("lists submenu leaves under their parents, not the parents themselves", () => {
    const entries = recordMenu("View", (m) => {
      subMenu(m, "Zoom", (z) => {
        menuItem(z, "50%", () => {});
        subMenu(z, "More", (x) => menuItem(x, "200%", () => {}));
      });
    });
    expect(entries.map((e) => [e.title, e.path.join(" › ")])).toEqual([
      ["50%", "View › Zoom"],
      ["200%", "View › Zoom › More"],
    ]);
  });

  it("runs the item's own click handler", () => {
    let ran = 0;
    const [e] = recordMenu("Edit", (m) => menuItem(m, "Undo", () => ran++));
    e.run(new KeyboardEvent("keydown"));
    expect(ran).toBe(1);
  });

  it("reads titles given as a DocumentFragment and drops items with no title or no handler", () => {
    const entries = recordMenu("Insert", (m) => {
      const f = createFragment();
      f.append("Special characters");
      m.addItem((i) => i.setTitle(f).onClick(() => {}));
      m.addItem((i) => i.setTitle("Heading only"));
      m.addItem((i) => i.onClick(() => {}));
    });
    expect(titles(entries)).toEqual(["Special characters"]);
  });

  it("tolerates menu methods it doesn't model", () => {
    const entries = recordMenu("Tools", (m) => {
      (m as ObsidianMenu & { setNoIcon(): ObsidianMenu }).setNoIcon();
      m.onHide(() => {});
      m.addItem((i) => i.setTitle("Word count").setSection("info").setWarning(false).onClick(() => {}));
    });
    expect(titles(entries)).toEqual(["Word count"]);
  });

  it("a menu that throws while building gives what it had so far", () => {
    const entries = recordMenu("Data", (m) => {
      menuItem(m, "Sort A → Z", () => {});
      throw new Error("no selection");
    });
    expect(titles(entries)).toEqual(["Sort A → Z"]);
  });
});

describe("matchMenus", () => {
  const all = [
    entry("Bold", ["Format", "Text"]),
    entry("Bulleted list", ["Format"]),
    entry("Table of contents", ["Insert"]),
    entry("Table", ["Insert"]),
    entry("Delete table", ["Format", "Table"]),
    entry("Find and replace", ["Edit"]),
    entry("150%", ["View", "Zoom"]),
    entry("Word count", ["Tools"]),
    entry("Columns…", ["Format"]),
    entry("Header & footer", ["Insert"]),
  ];

  it("nothing for an empty or blank query", () => {
    expect(matchMenus(all, "")).toEqual([]);
    expect(matchMenus(all, "   ")).toEqual([]);
  });

  it("ignores case", () => {
    expect(titles(matchMenus(all, "BOLD"))).toEqual(["Bold"]);
  });

  it("exact title first, then titles starting with it, then words starting with it, then anywhere", () => {
    expect(titles(matchMenus(all, "table"))).toEqual(["Table", "Table of contents", "Delete table"]);
    expect(titles(matchMenus(all, "b"))).toEqual(["Bold", "Bulleted list"]);
  });

  it("matches inside words after word starts", () => {
    expect(titles(matchMenus(all, "ount"))).toEqual(["Word count"]);
  });

  it("finds items by their menu path, after title matches", () => {
    expect(titles(matchMenus(all, "zoom"))).toEqual(["150%"]);
    expect(titles(matchMenus(all, "text"))).toEqual(["Bold"]);
  });

  it("every word of the query must match, in the title or the path", () => {
    expect(titles(matchMenus(all, "zoom 150"))).toEqual(["150%"]);
    expect(titles(matchMenus(all, "format bold"))).toEqual(["Bold"]);
    expect(matchMenus(all, "zoom bold")).toEqual([]);
  });

  it("ignores ellipses and ampersands in titles", () => {
    expect(titles(matchMenus(all, "columns"))).toEqual(["Columns…"]);
    expect(titles(matchMenus(all, "header footer"))).toEqual(["Header & footer"]);
  });

  it("keeps menu order among equally good matches", () => {
    const same = [entry("Copy", ["Edit"]), entry("Copy link to heading", ["Edit"]), entry("Copy slide", ["Slide"])];
    expect(titles(matchMenus(same, "copy "))).toEqual(["Copy", "Copy link to heading", "Copy slide"]);
  });

  it("lists an item reachable twice once", () => {
    const twice = [entry("Undo", ["Edit"]), entry("Undo", ["Edit"])];
    expect(matchMenus(twice, "undo")).toHaveLength(1);
  });
});

function searchBar(build?: (tb: Toolbar) => void) {
  const host = document.body.createDiv();
  const tb = new Toolbar(host.createDiv({ cls: "dw-toolbar" }));
  const ran: string[] = [];
  let bold = false;
  tb.fileMenu({}, (m) => menuItem(m, "Export…", () => ran.push("export")));
  tb.menu("Format", (m) => {
    menuItem(m, "Bold", () => ((bold = !bold), ran.push("bold")), { checked: bold });
    menuItem(m, "Bulleted list", () => ran.push("bullets"));
    menuItem(m, "Clear formatting", () => ran.push("clear"), { disabled: true });
  });
  tb.menu("View", (m) => subMenu(m, "Zoom", (z) => menuItem(z, "150%", () => ran.push("150"))));
  build?.(tb);
  tb.end();
  const input = tb.el.querySelector<HTMLInputElement>(".dw-menu-search input")!;
  const list = () => document.body.querySelector<HTMLElement>(".dw-menu-search-results");
  const shown = () => Array.from(list()?.querySelectorAll<HTMLElement>(".suggestion-item") ?? []).map((el) => el.querySelector(".dw-menu-search-title")!.textContent);
  const type = (text: string) => {
    input.focus();
    input.value = text;
    input.dispatchEvent(new Event("input"));
  };
  const key = (k: string, init: KeyboardEventInit = {}) => {
    const e = new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...init });
    input.dispatchEvent(e);
    return e;
  };
  return { tb, input, list, shown, type, key, ran, isBold: () => bold };
}

describe("menu search box", () => {
  it("sits first in the toolbar row of a bar with menus, as Google's 'Menus' box", () => {
    const { tb, input } = searchBar();
    const box = tb.el.firstElementChild!;
    expect(box.classList.contains("dw-menu-search")).toBe(true);
    expect(input.placeholder).toBe("Menus");
    expect(input.getAttribute("aria-label")).toBe("Search the menus");
    expect(input.getAttribute("aria-keyshortcuts")).toBe("Alt+/");
    expect(input.getAttribute("role")).toBe("combobox");
  });

  it("isn't added to a bar without menus", () => {
    const tb = new Toolbar(document.body.createDiv());
    tb.button("B", "Bold", () => {});
    tb.end();
    expect(tb.el.querySelector(".dw-menu-search")).toBeNull();
    expect(tb.focusSearch()).toBe(false);
  });

  it("never folds into ⋮ (it isn't a group)", () => {
    const { tb } = searchBar((tb) => tb.button("B", "Bold", () => {}));
    tb.overflows = () => true;
    tb.fit();
    expect(tb.el.querySelector<HTMLElement>(".dw-menu-search")!.hidden).toBe(false);
  });

  it("lists matching items from every menu as you type, with their menu path", () => {
    const { type, shown, list } = searchBar();
    type("b");
    expect(shown()).toEqual(["Bold", "Bulleted list"]);
    expect(list()!.querySelector(".dw-menu-search-path")!.textContent).toBe("Format");
    type("zoom");
    expect(shown()).toEqual(["150%"]);
    expect(list()!.querySelector(".dw-menu-search-path")!.textContent).toBe("View › Zoom");
    type("export");
    expect(shown()).toEqual(["Export…"]);
  });

  it("says so when nothing matches, and closes for an empty query", () => {
    const { type, shown, list, input } = searchBar();
    type("qqq");
    expect(shown()).toEqual([]);
    expect(list()!.textContent).toContain("No results");
    type("");
    expect(list()).toBeNull();
    expect(input.getAttribute("aria-expanded")).toBe("false");
  });

  it("Enter runs the first match, then clears and closes", () => {
    const { type, key, ran, input, list } = searchBar();
    type("bold");
    expect(key("Enter").defaultPrevented).toBe(true);
    expect(ran).toEqual(["bold"]);
    expect(input.value).toBe("");
    expect(list()).toBeNull();
  });

  it("arrow keys move the highlighted item (wrapping), Enter runs it", () => {
    const { type, key, ran, list, input } = searchBar();
    type("b");
    const selected = () => list()!.querySelector(".suggestion-item.is-selected .dw-menu-search-title")!.textContent;
    expect(selected()).toBe("Bold");
    key("ArrowDown");
    expect(selected()).toBe("Bulleted list");
    expect(input.getAttribute("aria-activedescendant")).toBe(list()!.querySelector(".is-selected")!.id);
    key("ArrowDown");
    expect(selected()).toBe("Bold");
    key("ArrowUp");
    expect(selected()).toBe("Bulleted list");
    key("Enter");
    expect(ran).toEqual(["bullets"]);
  });

  it("a click on a result runs it", () => {
    const { type, list, ran } = searchBar();
    type("zoom");
    const item = list()!.querySelector<HTMLElement>(".suggestion-item")!;
    item.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    item.click();
    expect(ran).toEqual(["150"]);
    expect(list()).toBeNull();
  });

  it("shows disabled items greyed and never runs them", () => {
    const { type, key, list, ran } = searchBar();
    type("clear");
    const item = list()!.querySelector<HTMLElement>(".suggestion-item")!;
    expect(item.classList.contains("is-disabled")).toBe(true);
    expect(item.getAttribute("aria-disabled")).toBe("true");
    key("Enter");
    item.click();
    expect(ran).toEqual([]);
  });

  it("shows which items are ticked, as the menu would", () => {
    const { type, key, list } = searchBar();
    type("bold");
    expect(list()!.querySelector(".suggestion-item")!.classList.contains("is-checked")).toBe(false);
    key("Enter");
    type("bold");
    expect(list()!.querySelector(".suggestion-item")!.classList.contains("is-checked")).toBe(true);
  });

  it("menus are read again each time the box gets focus, so items show the current state", () => {
    let extra = false;
    const { type, shown, input } = searchBar((tb) => tb.menu("Slide", (m) => extra && menuItem(m, "Bookmark slide", () => {})));
    type("bookmark");
    expect(shown()).toEqual([]);
    input.blur();
    extra = true;
    type("bookmark");
    expect(shown()).toEqual(["Bookmark slide"]);
  });

  it("Escape clears, closes and gives focus back to where it came from", () => {
    const { tb, type, key, input, list } = searchBar();
    const editor = tb.el.parentElement!.createDiv({ attr: { tabindex: "0" } });
    editor.focus();
    type("bold");
    expect(key("Escape").defaultPrevented).toBe(true);
    expect(input.value).toBe("");
    expect(list()).toBeNull();
    expect(document.activeElement).toBe(editor);
  });

  it("running an item gives focus back first, so commands act on the editor", () => {
    const host = document.body.createDiv();
    const editor = host.createDiv({ attr: { tabindex: "0" } });
    const tb = new Toolbar(host.createDiv());
    let focusedWhenRun: Element | null = null;
    tb.menu("Format", (m) => menuItem(m, "Bold", () => (focusedWhenRun = document.activeElement)));
    tb.end();
    const input = tb.el.querySelector<HTMLInputElement>(".dw-menu-search input")!;
    editor.focus();
    input.focus();
    input.value = "bold";
    input.dispatchEvent(new Event("input"));
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    expect(focusedWhenRun).toBe(editor);
  });

  it("closes when focus leaves the box", () => {
    const { type, input, list } = searchBar();
    type("b");
    input.dispatchEvent(new FocusEvent("blur"));
    expect(list()).toBeNull();
  });

  it("keys typed in the box don't reach the editor or the toolbar's arrow-key handling", () => {
    const { tb, type, key } = searchBar();
    let reached = 0;
    tb.el.parentElement!.addEventListener("keydown", () => reached++);
    type("b");
    key("ArrowDown");
    key("ArrowLeft");
    key("x");
    expect(reached).toBe(0);
  });

  it("focusSearch (Alt+/) focuses the box and selects what was typed", () => {
    const { tb, input } = searchBar();
    input.value = "zo";
    expect(tb.focusSearch()).toBe(true);
    expect(document.activeElement).toBe(input);
    expect(input.selectionStart).toBe(0);
    expect(input.selectionEnd).toBe(2);
  });

  it("caps the list so a one-letter query stays readable", () => {
    const { type, shown } = searchBar((tb) => tb.menu("Insert", (m) => {
      for (let n = 0; n < 40; n++) menuItem(m, `Item ${n}`, () => {});
    }));
    type("item");
    expect(shown()).toHaveLength(15);
  });

  it("read-only files: only the items their menus keep", () => {
    const tb = new Toolbar(document.body.createDiv().createDiv(), true);
    tb.menu("Edit", (m) => {
      menuItem(m, "Copy", () => {});
      menuItem(m, "Cut", () => {});
    });
    tb.menu("Format", (m) => menuItem(m, "Bold", () => {}));
    tb.end();
    const input = tb.el.querySelector<HTMLInputElement>(".dw-menu-search input")!;
    input.focus();
    for (const q of ["copy", "cut", "bold"]) {
      input.value = q;
      input.dispatchEvent(new Event("input"));
      const n = document.body.querySelectorAll(".dw-menu-search-results .suggestion-item").length;
      expect([q, n]).toEqual([q, q === "copy" ? 1 : 0]);
    }
  });

  it("menu building for the search doesn't open a menu", () => {
    const { type } = searchBar();
    Menu.last = null;
    type("b");
    expect(Menu.last).toBeNull();
  });
});
