import { afterEach, describe, expect, it } from "vitest";
import { Menu } from "./obsidian-stub";
import { menuItem, subMenu, Toolbar } from "../src/presentation/common/toolbar";

afterEach(() => document.body.empty());
const titles = (m: Menu | null) => m!.items.map((i) => i.title);

describe("menu row", () => {
  it("puts a menu row before the toolbar and builds each menu when it opens", () => {
    const host = document.body.createDiv();
    const tb = new Toolbar(host.createDiv({ cls: "dw-toolbar" }));
    let bold = false;
    const btn = tb.menu("Format", (m) => menuItem(m, "Bold", () => (bold = !bold), { checked: bold }));
    tb.end();
    expect(tb.menubar!.nextElementSibling).toBe(tb.el);
    expect(tb.menubar!.getAttribute("role")).toBe("menubar");
    btn.click();
    expect(Menu.last!.items[0]).toMatchObject({ title: "Bold", checked: false });
    Menu.last!.items[0].run!();
    btn.click();
    expect(Menu.last!.items[0].checked).toBe(true);
  });

  it("File opens the view's own ⋮ (More options) menu when Obsidian has one, else the menu built here", () => {
    const tb = new Toolbar(document.body.createDiv().createDiv());
    const opened: Event[] = [];
    const view = { onMoreOptions: (e: Event) => opened.push(e) };
    const file = tb.fileMenu(view, (m) => menuItem(m, "Export…", () => {}));
    Menu.last = null;
    file.click();
    expect(opened).toHaveLength(1);
    expect(opened[0].target).toBe(file);
    expect(Menu.last).toBeNull();
    const fallback = tb.fileMenu({}, (m) => menuItem(m, "Export…", () => {}));
    fallback.click();
    expect(titles(Menu.last)).toEqual(["Export…"]);
  });

  it("uses Obsidian's own submenus (hover, arrow keys, parent stays open)", () => {
    const tb = new Toolbar(document.body.createDiv().createDiv());
    tb.menu("Insert", (m) => subMenu(m, "Break", (s) => menuItem(s, "Page break", () => {}))).click();
    const parent = Menu.last!;
    expect(titles(parent)).toEqual(["Break"]);
    expect(titles(parent.items[0].submenu!)).toEqual(["Page break"]);
  });

  it("falls back to a ▸ item that opens a second menu when Obsidian has no submenus", () => {
    Menu.noSubmenus = true;
    try {
      const tb = new Toolbar(document.body.createDiv().createDiv());
      tb.menu("Insert", (m) => subMenu(m, "Break", (s) => menuItem(s, "Page break", () => {}))).click();
      const parent = Menu.last!;
      expect(titles(parent)).toEqual(["Break ▸"]);
      parent.items[0].run!(new MouseEvent("click"));
      expect(titles(Menu.last)).toEqual(["Page break"]);
    } finally {
      Menu.noSubmenus = false;
    }
  });
});

function bar(groups = 4) {
  const tb = new Toolbar(document.body.createDiv());
  const clicks: string[] = [];
  tb.button("U", "Undo", () => clicks.push("undo"));
  tb.pin();
  for (let g = 1; g < groups; g++) {
    tb.gap();
    tb.button(`b${g}`, `Button ${g}`, () => clicks.push(`b${g}`));
  }
  tb.end();
  // Pretend the row has room for `room` shown groups.
  let room = 99;
  const shown = () => Array.from(tb.el.querySelectorAll<HTMLElement>(".dw-group")).filter((g) => !g.hidden && g.childElementCount).length;
  tb.overflows = () => shown() > room;
  return { tb, clicks, setRoom: (n: number) => { room = n; tb.fit(); }, shown };
}
const more = (tb: Toolbar) => tb.el.querySelector<HTMLButtonElement>(".dw-more")!;

describe("overflow", () => {
  it("folds unpinned groups from the right into ⋮, whose items click the real buttons", () => {
    const { tb, clicks, setRoom, shown } = bar();
    setRoom(99);
    expect(more(tb).hidden).toBe(true);
    setRoom(2);
    expect(shown()).toBe(2);
    expect(more(tb).hidden).toBe(false);
    more(tb).click();
    expect(titles(Menu.last)).toEqual(["Button 2", "---", "Button 3"]);
    Menu.last!.items[2].run!();
    expect(clicks).toEqual(["b3"]);
  });

  it("never folds pinned groups", () => {
    const { tb, setRoom } = bar();
    setRoom(0);
    expect(tb.el.querySelector<HTMLElement>(".dw-group")!.hidden).toBe(false);
  });

  it("unfolds when room returns", () => {
    const { tb, setRoom, shown } = bar();
    setRoom(1);
    setRoom(99);
    expect(shown()).toBe(4);
    expect(more(tb).hidden).toBe(true);
  });

  it("folded select: lists options as a submenu that sets the value and fires change", () => {
    const tb = new Toolbar(document.body.createDiv());
    const picked: string[] = [];
    const s = tb.select("Font", "Font", [["Arial", "Arial"], ["Calibri", "Calibri"]], (v) => picked.push(v));
    tb.end();
    tb.overflows = () => !s.closest<HTMLElement>(".dw-group")!.hidden;
    tb.fit();
    more(tb).click();
    expect(Menu.last!.items[0].title).toBe("Font");
    const sub = Menu.last!.items[0].submenu!;
    expect(titles(sub)).toEqual(["Arial", "Calibri"]);
    sub.items[1].run!();
    expect([s.value, picked]).toEqual(["Calibri", ["Calibri"]]);
  });

  it("folded icon select: uses the wrapper's label", () => {
    const tb = new Toolbar(document.body.createDiv());
    tb.iconSelect("unfold-vertical", "Line spacing", [["1.5", "1.5"]], () => {});
    tb.end();
    tb.overflows = () => !tb.el.querySelector<HTMLElement>(".dw-group")!.hidden;
    tb.fit();
    more(tb).click();
    expect(titles(Menu.last)).toEqual(["Line spacing"]);
  });

  it("folded colour: opens a temporary picker and forwards its value to the real input", () => {
    const tb = new Toolbar(document.body.createDiv());
    const picked: string[] = [];
    tb.color("A", "Text color", "#000000", (c) => picked.push(c));
    tb.end();
    tb.overflows = () => !tb.el.querySelector<HTMLElement>(".dw-group")!.hidden;
    tb.fit();
    more(tb).click();
    expect(titles(Menu.last)).toEqual(["Text color…"]);
    Menu.last!.items[0].run!();
    const temp = document.body.querySelector<HTMLInputElement>("input.dw-color-temp")!;
    temp.value = "#ff0000";
    temp.dispatchEvent(new Event("change"));
    expect(picked).toEqual(["#FF0000"]);
    expect(document.body.querySelector("input.dw-color-temp")).toBeNull();
  });

  it("context groups follow sync() and never show in ⋮ while hidden", () => {
    const tb = new Toolbar(document.body.createDiv());
    let picture = false;
    tb.button("B", "Bold", () => {});
    tb.pin();
    tb.gap();
    tb.button("Crop", "Crop", () => {});
    tb.context(() => picture);
    tb.end();
    tb.overflows = () => false;
    tb.sync();
    const crop = tb.el.querySelector<HTMLElement>('[aria-label="Crop"]')!.parentElement!;
    expect(crop.hidden).toBe(true);
    picture = true;
    tb.sync();
    expect(crop.hidden).toBe(false);
    // Full row: the shown context group folds like any other.
    tb.overflows = () => !crop.hidden;
    tb.fit();
    more(tb).click();
    expect(titles(Menu.last)).toEqual(["Crop"]);
    picture = false;
    tb.sync();
    expect(more(tb).hidden).toBe(true);
  });

  it("folds the selection's tools last: the other groups go into ⋮ first", () => {
    const tb = new Toolbar(document.body.createDiv());
    tb.button("B", "Bold", () => {});
    tb.gap();
    tb.button("L", "Link", () => {});
    tb.gap();
    tb.button("Crop", "Crop", () => {});
    tb.context(() => true);
    tb.end();
    const shown = () => Array.from(tb.el.querySelectorAll<HTMLElement>(".dw-group")).filter((g) => !g.hidden && g.childElementCount).length;
    tb.overflows = () => shown() > 1;
    tb.sync();
    tb.fit();
    expect(tb.el.querySelector<HTMLElement>('[aria-label="Crop"]')!.parentElement!.hidden).toBe(false);
    more(tb).click();
    expect(titles(Menu.last)).toEqual(["Bold", "---", "Link"]);
  });

  it("marks active buttons as checked in ⋮", () => {
    const { tb, setRoom } = bar();
    tb.el.querySelector<HTMLElement>('[aria-label="Button 3"]')!.addClass("is-active");
    setRoom(1);
    more(tb).click();
    expect(Menu.last!.items.find((i) => i.title === "Button 3")!.checked).toBe(true);
    expect(Menu.last!.items.find((i) => i.title === "Button 2")!.checked).toBeUndefined();
  });

  it("leaves the shortcut hint in parentheses out of ⋮ item titles", () => {
    const tb = new Toolbar(document.body.createDiv());
    tb.button("B", "Bold (Ctrl+B)", () => {});
    tb.end();
    tb.overflows = () => !tb.el.querySelector<HTMLElement>(".dw-group")!.hidden;
    tb.fit();
    more(tb).click();
    expect(titles(Menu.last)).toEqual(["Bold"]);
  });

  it("keeps the end area (save status) out of the fold", () => {
    const { tb, setRoom } = bar();
    setRoom(0);
    expect(tb.el.querySelector<HTMLElement>(".dw-end")!.hidden).toBe(false);
  });
});

describe("review fixes", () => {
  it("icon selects are actions: picking the same option twice runs it twice", () => {
    const tb = new Toolbar(document.body.createDiv());
    const picked: string[] = [];
    const s = tb.iconSelect("rotate-ccw", "Reset image", [["Reset crop", "crop"]], (v) => picked.push(v));
    for (let i = 0; i < 2; i++) {
      s.value = "crop";
      s.dispatchEvent(new Event("change"));
    }
    expect(picked).toEqual(["crop", "crop"]);
    expect(s.value).toBe("");
  });

  it("refits when the end area changes (e.g. Couldn't save + Retry appears)", async () => {
    const tb = new Toolbar(document.body.createDiv());
    tb.button("B", "Bold", () => {});
    const { status, el } = tb.end();
    tb.overflows = () => (el.textContent ?? "").includes("Couldn't") && !tb.el.querySelector<HTMLElement>(".dw-group")!.hidden;
    tb.fit();
    status.failed(new Error("x"), () => {}, "a");
    await new Promise((r) => setTimeout(r, 0));
    expect(tb.el.querySelector<HTMLElement>(".dw-group")!.hidden).toBe(true);
    status.destroy();
  });
});
