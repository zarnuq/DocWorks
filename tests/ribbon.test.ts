import { afterEach, describe, expect, it, vi } from "vitest";
import { Menu } from "./obsidian-stub";
import type { Commands } from "../src/presentation/common/commands";
import { Ribbon, type Tab } from "../src/presentation/common/ribbon";

afterEach(() => document.body.empty());

function storage() {
  const data = new Map<string, unknown>();
  return { loadLocalStorage: (k: string) => data.get(k) ?? null, saveLocalStorage: (k: string, v: unknown) => void data.set(k, v), data };
}

function setup(over: { tabs?: Tab[]; cmds?: Commands; readOnly?: boolean; store?: ReturnType<typeof storage> } = {}) {
  const ran: string[] = [];
  const run = (id: string) => () => void ran.push(id);
  const state = { bold: false, table: false };
  const cmds: Commands = over.cmds ?? {
    Paste: { label: "Paste", icon: "clipboard-paste", run: run("Paste"), menu: (m) => m.addItem((i) => i.setTitle("Keep Text Only").onClick(run("PasteTextOnly"))) },
    Cut: { label: "Cut", icon: "scissors", run: run("Cut") },
    Copy: { label: "Copy", icon: "copy", run: run("Copy"), view: true },
    Bold: { label: "Bold", icon: "bold", run: run("Bold"), active: () => state.bold, shortcut: "Ctrl+B" },
    ParagraphDialog: { label: "Paragraph Settings", icon: "pilcrow", run: run("ParagraphDialog") },
    PictureInsertFromFile: { label: "Pictures", icon: "image", run: run("Pictures") },
    TableShading: { label: "Shading", icon: "paint-bucket", run: run("Shading") },
    ViewRulerWord: { label: "Ruler", icon: "ruler", run: run("Ruler"), view: true, active: () => true },
  };
  const tabs: Tab[] = over.tabs ?? [
    { id: "home", label: "Home", groups: [
      { id: "clipboard", label: "Clipboard", icon: "clipboard", controls: [{ kind: "large", id: "Paste" }, { kind: "column", items: [{ kind: "small", id: "Cut" }, { kind: "small", id: "Copy" }] }] },
      { id: "font", label: "Font", icon: "type", controls: [{ kind: "icon", id: "Bold" }] },
      { id: "paragraph", label: "Paragraph", icon: "pilcrow", launcher: "ParagraphDialog", controls: [] },
    ] },
    { id: "insert", label: "Insert", groups: [{ id: "illustrations", label: "Illustrations", icon: "image", controls: [{ kind: "large", id: "PictureInsertFromFile" }] }] },
    { id: "view", label: "View", groups: [{ id: "show", label: "Show", icon: "eye", controls: [{ kind: "check", id: "ViewRulerWord" }] }] },
    { id: "tableDesign", label: "Table Design", context: () => state.table, groups: [{ id: "styles", label: "Table Styles", icon: "table", controls: [{ kind: "large", id: "TableShading" }] }] },
  ];
  const el = document.body.createDiv();
  const store = over.store ?? storage();
  const fileMenu = vi.fn((m: Menu) => m.addItem((i) => i.setTitle("Export…")));
  const ribbon = new Ribbon(el, cmds, tabs, { readOnly: over.readOnly ?? false, app: store, key: "doc-works-rb-test", fileMenu: fileMenu as never });
  return { ribbon, el, ran, state, store, fileMenu };
}

const tabLabels = (el: HTMLElement) => Array.from(el.querySelectorAll<HTMLElement>(".dw-rb-tab")).filter((t) => !t.hidden).map((t) => t.textContent);
const groupLabels = (el: HTMLElement) => Array.from(el.querySelectorAll(".dw-rb-caption-label")).map((c) => c.textContent);
const button = (el: HTMLElement, label: string) => el.querySelector<HTMLButtonElement>(`.dw-rb-panel [aria-label^="${label}"]`)!;

describe("ribbon tabs", () => {
  it("shows File then the tabs in order, Home's groups first", () => {
    const { el } = setup();
    expect(el.querySelector(".dw-rb-file")!.textContent).toBe("File");
    expect(tabLabels(el)).toEqual(["Home", "Insert", "View"]);
    expect(el.querySelectorAll('[role="tab"]').length).toBe(4);
    expect(el.querySelector('[role="tab"][aria-selected="true"]')!.textContent).toBe("Home");
    expect(groupLabels(el)).toEqual(["Clipboard", "Font", "Paragraph"]);
  });

  it("shows a tab's groups when it's clicked", () => {
    const { el } = setup();
    el.querySelectorAll<HTMLElement>(".dw-rb-tab")[1].click();
    expect(groupLabels(el)).toEqual(["Illustrations"]);
    expect(el.querySelector('[role="tab"][aria-selected="true"]')!.textContent).toBe("Insert");
  });

  it("opens the view's file menu under File", () => {
    const { el, fileMenu } = setup();
    el.querySelector<HTMLElement>(".dw-rb-file")!.click();
    expect(fileMenu).toHaveBeenCalled();
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Export…"]);
  });
});

describe("ribbon controls", () => {
  it("draws large buttons with a caption, small ones with a caption, icon buttons with a label only, and runs them", () => {
    const { el, ran } = setup();
    expect(button(el, "Pictures")).toBeNull();
    const cut = button(el, "Cut");
    expect(cut.classList.contains("is-small")).toBe(true);
    expect(cut.textContent).toBe("Cut");
    const bold = button(el, "Bold");
    expect(bold.classList.contains("is-icon")).toBe(true);
    expect(bold.textContent).toBe("");
    expect(bold.getAttribute("aria-label")).toBe("Bold (Ctrl+B)");
    cut.click();
    bold.click();
    expect(ran).toEqual(["Cut", "Bold"]);
  });

  it("splits a command that has both an action and a menu", () => {
    const { el, ran } = setup();
    const paste = button(el, "Paste");
    expect(paste.classList.contains("is-large")).toBe(true);
    expect(paste.textContent).toBe("Paste");
    paste.click();
    expect(ran).toEqual(["Paste"]);
    el.querySelector<HTMLButtonElement>('.dw-rb-panel [aria-label="Paste options"]')!.click();
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Keep Text Only"]);
    Menu.last!.items[0].run!();
    expect(ran).toEqual(["Paste", "PasteTextOnly"]);
  });

  it("marks toggles that are on after sync", () => {
    const { el, ribbon, state } = setup();
    expect(button(el, "Bold").getAttribute("aria-pressed")).toBe("false");
    state.bold = true;
    ribbon.sync();
    expect(button(el, "Bold").classList.contains("is-active")).toBe(true);
    expect(button(el, "Bold").getAttribute("aria-pressed")).toBe("true");
  });

  it("puts a group's dialog launcher in its caption", () => {
    const { el, ran } = setup();
    el.querySelector<HTMLButtonElement>('.dw-rb-group[data-group="paragraph"] .dw-rb-launcher')!.click();
    expect(ran).toEqual(["ParagraphDialog"]);
  });

  it("draws checkboxes from the command's state", () => {
    const { el, ran } = setup();
    el.querySelectorAll<HTMLElement>(".dw-rb-tab")[2].click();
    const box = el.querySelector<HTMLInputElement>('.dw-rb-check input[type="checkbox"]')!;
    expect(box.checked).toBe(true);
    expect(box.parentElement!.textContent).toBe("Ruler");
    box.click();
    expect(ran).toEqual(["Ruler"]);
  });
});

describe("contextual tabs", () => {
  it("appear while their test holds without taking over, as Word's; going away hands back the tab that was open", () => {
    const { el, ribbon, state } = setup();
    el.querySelectorAll<HTMLElement>(".dw-rb-tab")[1].click();
    state.table = true;
    ribbon.sync();
    expect(tabLabels(el)).toEqual(["Home", "Insert", "View", "Table Design"]);
    const tab = Array.from(el.querySelectorAll<HTMLElement>(".dw-rb-tab")).find((t) => t.textContent === "Table Design")!;
    expect(tab.classList.contains("is-contextual")).toBe(true);
    expect(ribbon.current).toBe("insert");
    tab.click();
    expect(ribbon.current).toBe("tableDesign");
    state.table = false;
    ribbon.sync();
    expect(tabLabels(el)).toEqual(["Home", "Insert", "View"]);
    expect(ribbon.current).toBe("insert");
  });
});

describe("ribbon keyboard", () => {
  it("the tab row is one Tab stop", () => {
    const { el, ribbon } = setup();
    ribbon.side("history", "Version history", () => {});
    const stops = Array.from(el.querySelectorAll<HTMLElement>(".dw-rb-tabs button")).filter((b) => b.tabIndex === 0);
    expect(stops.length).toBe(1);
  });
});

describe("read-only files", () => {
  it("keep only commands that look at the file, dropping empty groups and tabs", () => {
    const { el } = setup({ readOnly: true });
    expect(tabLabels(el)).toEqual(["Home", "View"]);
    expect(groupLabels(el)).toEqual(["Clipboard"]);
    expect(button(el, "Copy")).not.toBeNull();
    expect(button(el, "Cut")).toBeNull();
  });
});

describe("collapsing the ribbon", () => {
  it("double-clicking the open tab hides the groups, and it stays that way next time", () => {
    const { el, store } = setup();
    const home = el.querySelector<HTMLElement>(".dw-rb-tab")!;
    home.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(el.querySelector(".dw-rb")!.classList.contains("is-collapsed")).toBe(true);
    expect(store.data.get("doc-works-rb-test-collapsed")).toBe("collapsed");
    const again = setup({ store });
    expect(again.el.querySelector(".dw-rb")!.classList.contains("is-collapsed")).toBe(true);
    again.el.querySelector<HTMLElement>(".dw-rb-tab")!.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(again.el.querySelector(".dw-rb")!.classList.contains("is-collapsed")).toBe(false);
    expect(store.data.get("doc-works-rb-test-collapsed")).toBe("open");
  });

  it("a click on a tab while collapsed shows its groups over the page until a command runs", () => {
    const { el, ran } = setup();
    el.querySelector<HTMLElement>(".dw-rb-tab")!.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    el.querySelectorAll<HTMLElement>(".dw-rb-tab")[1].click();
    const rb = el.querySelector(".dw-rb")!;
    expect(rb.classList.contains("is-peek")).toBe(true);
    button(el, "Pictures").click();
    expect(ran).toEqual(["Pictures"]);
    expect(rb.classList.contains("is-peek")).toBe(false);
  });
});

describe("side buttons", () => {
  it("sit at the end of the tab row", () => {
    const { el, ribbon } = setup();
    const onClick = vi.fn();
    ribbon.side("messages-square", "Comments", onClick).click();
    expect(el.querySelector(".dw-rb-tabs .dw-rb-side button[aria-label='Comments']")).not.toBeNull();
    expect(onClick).toHaveBeenCalled();
  });
});

describe("value controls", () => {
  function valueSetup() {
    const picked: string[] = [];
    let font = "Calibri";
    const cmds: Commands = {
      Font: { label: "Font", icon: "type" },
      FontColorPicker: { label: "Font Color", icon: "baseline", menu: (m) => m.addItem((i) => i.setTitle("Automatic").onClick(() => picked.push("auto"))) },
      QuickStylesGallery: { label: "Styles", icon: "pilcrow", menu: (m) => m.addItem((i) => i.setTitle("Update Normal to Match").onClick(() => picked.push("update"))) },
    };
    const tabs: Tab[] = [{ id: "home", label: "Home", groups: [{ id: "font", label: "Font", icon: "type", controls: [
      { kind: "select", id: "Font", options: () => [["Calibri", "Calibri"], ["Arial", "Arial"]], value: () => font, set: (v) => { font = v; picked.push(v); } },
      { kind: "color", id: "FontColorPicker", initial: "#FF0000", set: (hex) => picked.push(hex) },
      { kind: "gallery", id: "QuickStylesGallery", items: () => [["Normal", "Normal"], ["Heading 1", "Heading1"]], value: () => "Heading1", set: (v) => picked.push(v) },
    ] }] }];
    return { ...setup({ cmds, tabs }), picked, setFont: (f: string) => (font = f) };
  }

  it("a select shows the current value and picks from a menu with it ticked", () => {
    const { el, picked, ribbon, setFont } = valueSetup();
    const select = el.querySelector<HTMLButtonElement>('.dw-rb-select[aria-label="Font"]')!;
    expect(select.querySelector(".dw-rb-select-label")!.textContent).toBe("Calibri");
    select.click();
    expect(Menu.last!.items.map((i) => [i.title, i.checked])).toEqual([["Calibri", true], ["Arial", false]]);
    Menu.last!.items[1].run!();
    expect(picked).toEqual(["Arial"]);
    setFont("Georgia");
    ribbon.sync();
    expect(select.querySelector(".dw-rb-select-label")!.textContent).toBe("Georgia");
  });

  it("a colour button applies the last colour; its arrow lists the command's items", () => {
    const { el, picked } = valueSetup();
    el.querySelector<HTMLButtonElement>('.dw-rb-panel [aria-label="Font Color"]')!.click();
    expect(picked).toEqual(["#FF0000"]);
    el.querySelector<HTMLButtonElement>('.dw-rb-panel [aria-label="Font Color options"]')!.click();
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Automatic", "More Colors…"]);
    Menu.last!.items[0].run!();
    expect(picked).toEqual(["#FF0000", "auto"]);
  });

  it("a gallery shows its items in the ribbon, marks the current one, and lists all plus extras under its arrow", () => {
    const { el, picked } = valueSetup();
    const items = Array.from(el.querySelectorAll<HTMLButtonElement>(".dw-rb-gallery-item"));
    expect(items.map((b) => [b.textContent, b.classList.contains("is-active")])).toEqual([["Normal", false], ["Heading 1", true]]);
    items[0].click();
    expect(picked).toEqual(["Normal"]);
    el.querySelector<HTMLButtonElement>('.dw-rb-panel [aria-label="Styles options"]')!.click();
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Normal", "Heading 1", "---", "Update Normal to Match"]);
  });
});

describe("layout helpers", () => {
  it("stacks columns vertically and rows horizontally", () => {
    const cmds: Commands = { Bold: { label: "Bold", icon: "bold" }, Italic: { label: "Italic", icon: "italic" }, Cut: { label: "Cut", icon: "scissors" } };
    const tabs: Tab[] = [{ id: "home", label: "Home", groups: [{ id: "font", label: "Font", icon: "type", controls: [
      { kind: "column", items: [{ kind: "row", items: [{ kind: "icon", id: "Bold" }, { kind: "icon", id: "Italic" }] }, { kind: "small", id: "Cut" }] },
    ] }] }];
    const { el } = setup({ cmds, tabs });
    const col = el.querySelector(".dw-rb-col")!;
    expect(Array.from(col.children).map((c) => c.className)).toEqual(["dw-rb-row", "dw-rb-btn is-small"]);
    expect(col.querySelectorAll(".dw-rb-row .dw-rb-btn").length).toBe(2);
  });
});

describe("folding when the ribbon is too narrow", () => {
  /** `overflows` answers true `n` times, then false. */
  const narrow = (ribbon: Ribbon, n: number) => {
    let left = n;
    ribbon.overflows = () => left-- > 0;
  };
  const sizes = (el: HTMLElement) => Array.from(el.querySelectorAll<HTMLElement>(".dw-rb-group")).map((g) => g.dataset.size);

  it("drops captions from the rightmost group first", () => {
    const { el, ribbon } = setup();
    narrow(ribbon, 1);
    ribbon.fit();
    expect(sizes(el)).toEqual(["full", "full", "compact"]);
    expect(el.querySelector('.dw-rb-group[data-group="paragraph"]')!.classList.contains("dw-rb-compact")).toBe(true);
  });

  it("then folds whole groups into one button, right to left, the first group last", () => {
    const { el, ribbon } = setup();
    narrow(ribbon, 5);
    ribbon.fit();
    expect(sizes(el)).toEqual(["compact", "collapsed", "collapsed"]);
    narrow(ribbon, 99);
    ribbon.fit();
    expect(sizes(el)).toEqual(["collapsed", "collapsed", "collapsed"]);
  });

  it("a folded group's button opens its controls, which work and close it", () => {
    const { el, ribbon, ran } = setup();
    narrow(ribbon, 99);
    ribbon.fit();
    const clip = el.querySelector<HTMLButtonElement>('.dw-rb-group[data-group="clipboard"] .dw-rb-group-btn')!;
    expect(clip.getAttribute("aria-label")).toBe("Clipboard");
    expect(el.querySelector(".dw-rb-flyout")).toBeNull();
    clip.click();
    const flyout = el.querySelector<HTMLElement>(".dw-rb-flyout")!;
    expect(flyout.querySelector('[aria-label="Cut"]')).not.toBeNull();
    flyout.querySelector<HTMLButtonElement>('[aria-label="Cut"]')!.click();
    expect(ran).toEqual(["Cut"]);
    expect(el.querySelector(".dw-rb-flyout")).toBeNull();
  });

  it("Escape closes an open group and focuses its button", () => {
    const { el, ribbon } = setup();
    narrow(ribbon, 99);
    ribbon.fit();
    const clip = el.querySelector<HTMLButtonElement>('.dw-rb-group[data-group="clipboard"] .dw-rb-group-btn')!;
    clip.click();
    el.querySelector(".dw-rb-flyout")!.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(el.querySelector(".dw-rb-flyout")).toBeNull();
    expect(document.activeElement).toBe(clip);
  });

  it("unfolds everything again when there's room", () => {
    const { el, ribbon } = setup();
    narrow(ribbon, 99);
    ribbon.fit();
    el.querySelector<HTMLButtonElement>(".dw-rb-group-btn")!.click();
    narrow(ribbon, 0);
    ribbon.fit();
    expect(sizes(el)).toEqual(["full", "full", "full"]);
    expect(el.querySelector(".dw-rb-group-btn")).toBeNull();
    expect(el.querySelector(".dw-rb-flyout")).toBeNull();
    expect(el.querySelector('.dw-rb-group[data-group="clipboard"] .dw-rb-controls [aria-label="Cut"]')).not.toBeNull();
  });
});

describe("Quick Access Toolbar", () => {
  const qat = { defaults: ["Copy", "Bold"], choices: ["Copy", "Cut", "Bold"] };
  function qatSetup(store = storage(), readOnly = false) {
    const s = setup({ store, readOnly });
    s.ribbon.destroy();
    s.el.empty();
    const ribbon = new Ribbon(s.el, (s.ribbon as unknown as { cmds: Commands }).cmds, (s.ribbon as unknown as { tabs: Tab[] }).tabs, { readOnly, app: store, key: "doc-works-rb-test", fileMenu: () => {}, qat });
    return { ...s, ribbon, store };
  }
  const qatLabels = (el: HTMLElement) => Array.from(el.querySelectorAll(".dw-rb-qat .dw-rb-btn")).map((b) => b.getAttribute("aria-label"));
  const rightClick = (target: Element) => target.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));

  it("starts with the defaults, left of File, and runs them", () => {
    const { el, ran } = qatSetup();
    expect(qatLabels(el)).toEqual(["Copy", "Bold (Ctrl+B)"]);
    const row = el.querySelector(".dw-rb-tabs")!;
    expect(row.firstElementChild!.classList.contains("dw-rb-qat")).toBe(true);
    el.querySelector<HTMLButtonElement>('.dw-rb-qat [aria-label="Copy"]')!.click();
    expect(ran).toEqual(["Copy"]);
  });

  it("adds a ribbon command from its right-click menu, and keeps the list", () => {
    const { el, store } = qatSetup();
    rightClick(button(el, "Cut"));
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Add to Quick Access Toolbar"]);
    Menu.last!.items[0].run!();
    expect(qatLabels(el)).toEqual(["Copy", "Bold (Ctrl+B)", "Cut"]);
    expect(store.data.get("doc-works-rb-test-qat")).toEqual(["Copy", "Bold", "Cut"]);
    rightClick(button(el, "Cut"));
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Remove from Quick Access Toolbar"]);
  });

  it("removes a command from its right-click menu on the toolbar", () => {
    const { el, store } = qatSetup();
    rightClick(el.querySelector('.dw-rb-qat [aria-label="Copy"]')!);
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Remove from Quick Access Toolbar"]);
    Menu.last!.items[0].run!();
    expect(qatLabels(el)).toEqual(["Bold (Ctrl+B)"]);
    expect(store.data.get("doc-works-rb-test-qat")).toEqual(["Bold"]);
  });

  it("customizes from its ▾ menu: common commands ticked, then every other command", () => {
    const { el } = qatSetup();
    el.querySelector<HTMLButtonElement>(".dw-rb-qat-more")!.click();
    const menu = Menu.last!;
    expect(menu.items.map((i) => [i.title, i.checked ?? null])).toEqual([["Copy", true], ["Cut", false], ["Bold", true], ["---", null], ["More Commands…", null]]);
    menu.items[1].run!();
    expect(qatLabels(el)).toEqual(["Copy", "Bold (Ctrl+B)", "Cut"]);
    const more = menu.items[4].submenu!;
    expect(more.items.map((i) => i.title)).toEqual(["Paragraph Settings", "Paste", "Pictures", "Ruler", "Shading"]);
    more.items[3].run!();
    expect(qatLabels(el)).toEqual(["Copy", "Bold (Ctrl+B)", "Cut", "Ruler"]);
  });

  it("ignores saved ids it doesn't know", () => {
    const store = storage();
    store.data.set("doc-works-rb-test-qat", ["Bold", "NoSuchId", 7]);
    const { el } = qatSetup(store);
    expect(qatLabels(el)).toEqual(["Bold (Ctrl+B)"]);
  });

  it("shows only commands that look at the file for read-only files", () => {
    const { el } = qatSetup(storage(), true);
    expect(qatLabels(el)).toEqual(["Copy"]);
  });
});

describe("ribbon review fixes", () => {
  const narrowBy = (ribbon: Ribbon, n: number) => {
    let left = n;
    ribbon.overflows = () => left-- > 0;
  };

  it("scrolls the groups sideways when even fully folded they don't fit", () => {
    const { el, ribbon } = setup();
    narrowBy(ribbon, 99);
    ribbon.fit();
    expect(el.querySelector(".dw-rb-panel")!.classList.contains("is-scroll")).toBe(true);
    narrowBy(ribbon, 0);
    ribbon.fit();
    expect(el.querySelector(".dw-rb-panel")!.classList.contains("is-scroll")).toBe(false);
  });

  it("keeps a folded group's controls inside the ribbon's width", () => {
    const { el, ribbon } = setup();
    narrowBy(ribbon, 99);
    ribbon.fit();
    const root = el.querySelector<HTMLElement>(".dw-rb")!;
    Object.defineProperty(root, "clientWidth", { value: 390, configurable: true });
    const btn = el.querySelector<HTMLButtonElement>('.dw-rb-group[data-group="font"] .dw-rb-group-btn')!;
    vi.spyOn(btn, "getBoundingClientRect").mockReturnValue({ left: 300, top: 0, width: 50, height: 80, right: 350, bottom: 80, x: 300, y: 0, toJSON: () => ({}) });
    vi.spyOn(HTMLElement.prototype, "offsetWidth", "get").mockReturnValue(330);
    btn.click();
    expect(el.querySelector<HTMLElement>(".dw-rb-flyout")!.style.getPropertyValue("--dw-rb-flyout-x")).toBe("60px");
  });

  it("read-only files can't change the Quick Access Toolbar", () => {
    const store = storage();
    store.data.set("doc-works-rb-test-qat", ["Copy", "Bold"]);
    const s = setup({ store, readOnly: true });
    s.ribbon.destroy();
    s.el.empty();
    new Ribbon(s.el, (s.ribbon as unknown as { cmds: Commands }).cmds, (s.ribbon as unknown as { tabs: Tab[] }).tabs, { readOnly: true, app: store, key: "doc-works-rb-test", fileMenu: () => {}, qat: { defaults: [], choices: [] } });
    Menu.last = null;
    button(s.el, "Copy").dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true }));
    expect(Menu.last).toBeNull();
    expect(store.data.get("doc-works-rb-test-qat")).toEqual(["Copy", "Bold"]);
  });

  it("offers only commands that do something under More Commands…, and shows QAT toggles as on", () => {
    const store = storage();
    const s = setup({ store });
    s.ribbon.destroy();
    s.el.empty();
    const cmds = { ...(s.ribbon as unknown as { cmds: Commands }).cmds, Font: { label: "Font", icon: "type" } };
    const r = new Ribbon(s.el, cmds, (s.ribbon as unknown as { tabs: Tab[] }).tabs, { readOnly: false, app: store, key: "doc-works-rb-test", fileMenu: () => {}, qat: { defaults: ["Bold"], choices: [] } });
    s.el.querySelector<HTMLButtonElement>(".dw-rb-qat-more")!.click();
    const more = Menu.last!.items.find((i) => i.title === "More Commands…")!.submenu!;
    expect(more.items.map((i) => i.title)).not.toContain("Font");
    s.state.bold = true;
    r.sync();
    expect(s.el.querySelector(".dw-rb-qat [data-cmd='Bold']")!.classList.contains("is-active")).toBe(true);
  });
});

describe("ribbon keyboard after QAT changes", () => {
  it("keeps one Tab stop in the tab row after the Quick Access Toolbar is redrawn", () => {
    const store = storage();
    const s = setup({ store });
    s.ribbon.destroy();
    s.el.empty();
    new Ribbon(s.el, (s.ribbon as unknown as { cmds: Commands }).cmds, (s.ribbon as unknown as { tabs: Tab[] }).tabs, { readOnly: false, app: store, key: "doc-works-rb-test", fileMenu: () => {}, qat: { defaults: ["Copy"], choices: ["Copy", "Cut"] } });
    s.el.querySelector<HTMLButtonElement>(".dw-rb-qat-more")!.click();
    Menu.last!.items[1].run!();
    expect(Array.from(s.el.querySelectorAll<HTMLElement>(".dw-rb-tabs button")).filter((b) => b.tabIndex === 0).length).toBe(1);
  });
});

describe("ribbon screenshot fixes", () => {
  it("draws drop-down arrows as icons, not a ▾ character (too small in some fonts)", () => {
    const { el } = setup();
    const carets = Array.from(el.querySelectorAll(".dw-rb-caret"));
    expect(carets.length).toBeGreaterThan(0);
    expect(carets.map((c) => c.textContent)).toEqual(carets.map(() => ""));
  });

  it("draws gallery tiles in their own look", () => {
    const cmds: Commands = { QuickStylesGallery: { label: "Styles", icon: "pilcrow" } };
    const tabs: Tab[] = [{ id: "home", label: "Home", groups: [{ id: "styles", label: "Styles", icon: "pilcrow", controls: [
      { kind: "gallery", id: "QuickStylesGallery", items: () => [["Normal", "Normal"], ["Heading 1", "Heading1"]], value: () => "Normal", set: () => {},
        preview: (v) => (v === "Heading1" ? { font: "Calibri Light", size: 16, bold: true, color: "#2F5496" } : { font: "Calibri", size: 11 }) },
    ] }] }];
    const { el } = setup({ cmds, tabs });
    const [normal, heading] = Array.from(el.querySelectorAll<HTMLElement>(".dw-rb-gallery-item"));
    expect(heading.style.fontWeight).toBe("bold");
    expect(heading.style.color).toBe("rgb(47, 84, 150)");
    expect(heading.style.fontFamily).toContain("Calibri Light");
    expect(parseFloat(heading.style.fontSize)).toBeGreaterThan(parseFloat(normal.style.fontSize));
  });

  it("runs a command's double-click action (Format Painter keeps painting)", () => {
    const ran: string[] = [];
    const cmds: Commands = { FormatPainter: { label: "Format Painter", icon: "paintbrush", run: () => ran.push("once"), dblclick: () => ran.push("sticky") } };
    const tabs: Tab[] = [{ id: "home", label: "Home", groups: [{ id: "c", label: "Clipboard", icon: "clipboard", controls: [{ kind: "small", id: "FormatPainter" }] }] }];
    const { el } = setup({ cmds, tabs });
    el.querySelector<HTMLElement>('[data-cmd="FormatPainter"]')!.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }));
    expect(ran).toContain("sticky");
  });
});
