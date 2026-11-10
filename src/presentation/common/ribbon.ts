/**
 * The Office-style ribbon (beta interface): File, the app's tabs and contextual tabs in one row, the open tab's
 * labelled groups below it. Controls are drawn from a view's command table, so they run the same code as the
 * Google-style toolbar's.
 */
import { Menu, setIcon, type MenuItem } from "obsidian";
import { AppStorage, Cmd, Commands, tooltip } from "./commands";
import { rovingToolbar } from "./keyboard";
import { pickColor } from "./toolbar";

export type Control =
  | { kind: "large" | "small" | "icon"; id: string }
  | { kind: "select"; id: string; options: () => [label: string, value: string][]; value: () => string | null; set: (v: string) => void; width?: number }
  | { kind: "color"; id: string; initial: string; set: (hex: string) => void }
  | { kind: "check"; id: string }
  | { kind: "gallery"; id: string; items: () => [label: string, value: string][]; value: () => string | null; set: (v: string) => void; preview?: (value: string) => Preview }
  /** Controls stacked top to bottom (up to three small ones) or side by side. */
  | { kind: "column" | "row"; items: Control[] };

/** How a gallery tile shows its item (a paragraph style's font, size, weight and colour). */
export interface Preview {
  font?: string | null;
  /** pt */
  size?: number | null;
  bold?: boolean;
  italic?: boolean;
  color?: string | null;
}

/** A ▾ as a small chevron icon: the character is a dot in some UI fonts. */
export function caret(parent: HTMLElement): HTMLElement {
  const c = parent.createSpan({ cls: "dw-rb-caret" });
  setIcon(c, "chevron-down");
  return c;
}

/** Relative luminance of #RRGGBB, 0 (black) to 1 (white). */
function luminance(hex: string): number {
  const n = parseInt(hex.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => v / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

export interface Group {
  id: string;
  label: string;
  /** Shown when the group is folded into one button. */
  icon: string;
  /** Command of the group's dialog launcher (the small corner button). */
  launcher?: string;
  controls: Control[];
}

export interface Tab {
  id: string;
  label: string;
  groups: Group[];
  /** A contextual tab (Table Design…): shown only while this holds. */
  context?: () => boolean;
}

export interface RibbonOptions {
  readOnly: boolean;
  app: AppStorage;
  /** Storage key prefix, e.g. "doc-works-rb-docx". */
  key: string;
  fileMenu: (m: Menu) => void;
  /** Quick Access Toolbar: the commands it starts with, and the ones its ▾ menu offers first. */
  qat?: { defaults: string[]; choices: string[] };
}

function showBelow(menu: Menu, anchor: HTMLElement): void {
  const r = anchor.getBoundingClientRect();
  menu.showAtPosition({ x: r.left, y: r.bottom });
}

type Stack = Extract<Control, { items: Control[] }>;
const isStack = (c: Control): c is Stack => c.kind === "column" || c.kind === "row";
/** Command ids a control uses. */
export const controlIds = (c: Control): string[] => (isStack(c) ? c.items.flatMap(controlIds) : [c.id]);

/** Read-only files: only commands that look at the file; groups and tabs left empty go. */
function keepViewOnly(tabs: Tab[], cmds: Commands): Tab[] {
  const keep = (c: Control): Control | null => {
    if (!isStack(c)) return cmds[c.id]?.view ? c : null;
    const items = c.items.map(keep).filter((x): x is Control => !!x);
    return items.length ? { ...c, items } : null;
  };
  return tabs
    .map((t) => ({
      ...t,
      groups: t.groups
        .map((g) => ({ ...g, launcher: g.launcher && cmds[g.launcher]?.view ? g.launcher : undefined, controls: g.controls.map(keep).filter((x): x is Control => !!x) }))
        .filter((g) => g.controls.length),
    }))
    .filter((t) => t.groups.length);
}

export class Ribbon {
  readonly root: HTMLElement;
  readonly tabRow: HTMLElement;
  readonly panel: HTMLElement;
  private readonly tabs: Tab[];
  private readonly tabEls = new Map<string, HTMLButtonElement>();
  private shownContext = new Set<string>();
  private currentTab = "";
  private lastCore = "";
  /** State updaters of the drawn controls, run by `sync`. */
  private refresh: (() => void)[] = [];
  private sideEl: HTMLElement | null = null;
  /** The drawn groups' definitions, for folding. */
  private groupInfo = new Map<HTMLElement, Group>();
  /** A folded group opened from its button. */
  private flyout: { el: HTMLElement; group: HTMLElement; controls: HTMLElement; button: HTMLButtonElement } | null = null;
  private readonly abort = new AbortController();
  /** Whether the panel is wider than the ribbon; replaced in tests (jsdom has no layout). */
  overflows = (): boolean => this.panel.scrollWidth > this.panel.clientWidth;

  constructor(
    el: HTMLElement,
    private readonly cmds: Commands,
    tabs: Tab[],
    private readonly opts: RibbonOptions
  ) {
    this.tabs = (opts.readOnly ? keepViewOnly(tabs, cmds) : tabs).map((t) => ({ ...t, groups: t.groups.map((g) => ({ ...g, controls: g.controls.filter((c) => controlIds(c).every((id) => cmds[id])) })) }));
    this.root = el.createDiv({ cls: "dw-rb" });
    this.tabRow = this.root.createDiv({ cls: "dw-rb-tabs", attr: { role: "tablist" } });
    if (opts.qat) {
      this.qatEl = this.tabRow.createDiv({ cls: "dw-rb-qat", attr: { role: "toolbar", "aria-label": "Quick Access Toolbar" } });
      this.renderQat();
    }
    const file = (this.fileBtn = this.tabRow.createEl("button", { cls: "dw-rb-file", text: "File", attr: { "aria-haspopup": "menu" } }));
    file.addEventListener("click", () => {
      const m = new Menu();
      opts.fileMenu(m);
      showBelow(m, file);
    });
    for (const tab of this.tabs) {
      const b = this.tabRow.createEl("button", { cls: "dw-rb-tab", text: tab.label, attr: { role: "tab", "aria-selected": "false" } });
      if (tab.context) {
        b.addClass("is-contextual");
        b.hidden = true;
      }
      b.addEventListener("click", () => this.clickTab(tab.id));
      b.addEventListener("dblclick", () => this.toggleCollapsed());
      this.tabEls.set(tab.id, b);
    }
    this.panel = this.root.createDiv({ cls: "dw-rb-panel", attr: { role: "tabpanel" } });
    rovingToolbar(this.panel);
    // The tab row: one Tab stop, arrow keys between QAT, File, tabs and side buttons (rovingToolbar names it a toolbar).
    rovingToolbar(this.tabRow);
    // Right-click a command: add it to (or remove it from) the Quick Access Toolbar, as Word.
    if (opts.qat && !opts.readOnly) this.root.addEventListener("contextmenu", (e) => {
      const id = (e.target as HTMLElement).closest<HTMLElement>("[data-cmd]")?.dataset.cmd;
      if (!id || !this.cmds[id]) return;
      e.preventDefault();
      const has = this.qatIds().includes(id);
      const m = new Menu();
      m.addItem((i) => i.setTitle(has ? "Remove from Quick Access Toolbar" : "Add to Quick Access Toolbar").onClick(() => this.toggleQat(id)));
      m.showAtMouseEvent(e);
    });
    // Buttons keep the document's selection.
    this.root.addEventListener("mousedown", (e) => {
      if ((e.target as HTMLElement).closest("button")) e.preventDefault();
    });
    this.root.ownerDocument.addEventListener("mousedown", (e) => {
      if (this.root.contains(e.target as Node)) return;
      this.root.removeClass("is-peek");
      this.closeFlyout();
    }, { signal: this.abort.signal });
    this.root.toggleClass("is-collapsed", opts.app.loadLocalStorage(`${opts.key}-collapsed`) === "collapsed");
    const first = this.tabs.find((t) => !t.context);
    if (first) this.select(first.id);
    if (typeof ResizeObserver !== "undefined") {
      // Next frame: fitting changes the ribbon's own size, which inside the callback reports a ResizeObserver loop.
      let frame = 0;
      const ro = new ResizeObserver(() => {
        cancelAnimationFrame(frame);
        frame = requestAnimationFrame(() => this.fit());
      });
      ro.observe(this.root);
      this.abort.signal.addEventListener("abort", () => ro.disconnect());
    }
  }

  /**
   * Fit the open tab's groups to the width, as Word does: from the right, groups first drop their small buttons'
   * captions, then fold into one button that opens them; the first group goes last.
   */
  fit(): void {
    this.closeFlyout();
    this.panel.removeClass("is-scroll");
    const groups = Array.from(this.panel.querySelectorAll<HTMLElement>(":scope > .dw-rb-group"));
    for (const g of groups) this.setSize(g, "full");
    if (!this.overflows()) return;
    for (const size of ["compact", "collapsed"] as const) {
      for (const g of [...groups].reverse()) {
        this.setSize(g, size);
        if (!this.overflows()) return;
      }
    }
    // Still too wide (a phone): the folded groups scroll sideways rather than being cut off.
    this.panel.addClass("is-scroll");
  }

  private setSize(g: HTMLElement, size: "full" | "compact" | "collapsed"): void {
    g.dataset.size = size;
    g.toggleClass("dw-rb-compact", size !== "full");
    const controls = g.querySelector<HTMLElement>(":scope > .dw-rb-controls");
    let button = g.querySelector<HTMLButtonElement>(":scope > .dw-rb-group-btn");
    if (controls) controls.hidden = size === "collapsed";
    if (size !== "collapsed") return button?.remove();
    const info = this.groupInfo.get(g);
    if (button || !info) return;
    button = createEl("button", { cls: "dw-rb-btn is-large dw-rb-group-btn", attr: { "aria-label": info.label, "aria-haspopup": "true" } });
    setIcon(button.createSpan({ cls: "dw-rb-icon" }), info.icon);
    button.createSpan({ cls: "dw-rb-label", text: info.label });
    caret(button);
    g.prepend(button);
    const btn = button;
    btn.addEventListener("click", () => (this.flyout?.button === btn ? this.closeFlyout() : this.openFlyout(g, btn)));
  }

  private openFlyout(group: HTMLElement, button: HTMLButtonElement): void {
    this.closeFlyout();
    const controls = group.querySelector<HTMLElement>(":scope > .dw-rb-controls");
    if (!controls) return;
    const el = this.root.createDiv({ cls: "dw-rb-flyout" });
    el.appendChild(controls);
    controls.hidden = false;
    const r = button.getBoundingClientRect(), root = this.root.getBoundingClientRect();
    // Under its button, kept inside the ribbon's width.
    const x = Math.max(0, Math.min(r.left - root.left, this.root.clientWidth - el.offsetWidth));
    el.setCssProps({ "--dw-rb-flyout-x": `${x}px` });
    el.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      e.preventDefault();
      e.stopPropagation();
      this.closeFlyout();
      button.focus();
    });
    this.flyout = { el, group, controls, button };
    controls.querySelector<HTMLElement>("button, input")?.focus();
  }

  private closeFlyout(): void {
    const f = this.flyout;
    if (!f) return;
    this.flyout = null;
    f.group.insertBefore(f.controls, f.group.querySelector(":scope > .dw-rb-caption"));
    f.controls.hidden = f.group.dataset.size === "collapsed";
    f.el.remove();
  }

  private qatEl: HTMLElement | null = null;
  private fileBtn: HTMLButtonElement | null = null;
  private qatRefresh: (() => void)[] = [];

  /** The Quick Access Toolbar's commands, as saved (unknown ids dropped) or the defaults. */
  qatIds(): string[] {
    const saved = this.opts.app.loadLocalStorage(`${this.opts.key}-qat`);
    const list = Array.isArray(saved) ? saved.filter((x): x is string => typeof x === "string") : this.opts.qat?.defaults ?? [];
    return list.filter((id) => this.cmds[id] && (!this.opts.readOnly || this.cmds[id].view));
  }

  private toggleQat(id: string): void {
    // The saved list as it is (ids of another file type's commands, or ones hidden here, stay).
    const saved = this.opts.app.loadLocalStorage(`${this.opts.key}-qat`);
    const list = Array.isArray(saved) ? saved.filter((x): x is string => typeof x === "string") : this.qatIds();
    this.opts.app.saveLocalStorage(`${this.opts.key}-qat`, list.includes(id) ? list.filter((x) => x !== id) : [...list, id]);
    this.renderQat();
  }

  private renderQat(): void {
    const el = this.qatEl;
    if (!el) return;
    el.empty();
    this.qatRefresh = [];
    for (const id of this.qatIds()) {
      const cmd = this.cmds[id];
      const b = el.createEl("button", { cls: "dw-rb-btn is-icon", attr: { "aria-label": tooltip(cmd), "data-cmd": id, tabindex: "-1" } });
      this.qatRefresh.push(() => {
        if (cmd.active) b.toggleClass("is-active", cmd.active());
        if (cmd.enabled) b.disabled = !cmd.enabled();
      });
      setIcon(b.createSpan({ cls: "dw-rb-icon" }), cmd.icon);
      b.addEventListener("click", () => {
        if (cmd.run) return this.invoke(cmd.run);
        const m = new Menu();
        cmd.menu?.(m);
        showBelow(m, b);
      });
    }
    if (!this.opts.readOnly) {
      const more = this.menuButton(el, "Customize Quick Access Toolbar", (m) => {
      const ids = this.qatIds();
      const choices = (this.opts.qat?.choices ?? []).filter((id) => this.cmds[id]);
      for (const id of choices) m.addItem((i) => i.setTitle(this.cmds[id].label).setChecked(ids.includes(id)).onClick(() => this.toggleQat(id)));
      m.addSeparator();
      m.addItem((i) => {
        i.setTitle("More Commands…");
        const sub = (i as MenuItem & { setSubmenu: () => Menu }).setSubmenu();
        const rest = Object.keys(this.cmds).filter((id) => !choices.includes(id) && (this.cmds[id].run || this.cmds[id].menu)).sort((a, b) => this.cmds[a].label.localeCompare(this.cmds[b].label));
        for (const id of rest) sub.addItem((j) => j.setTitle(this.cmds[id].label).setChecked(ids.includes(id)).onClick(() => this.toggleQat(id)));
      });
    }, "dw-rb-arrow dw-rb-qat-more");
      more.tabIndex = -1;
    }
    // Redrawn buttons drop the row's Tab stop if it was on one of them: File takes it.
    if (this.fileBtn && !this.tabRow.querySelector('button[tabindex="0"]')) this.fileBtn.tabIndex = 0;
  }

  get current(): string {
    return this.currentTab;
  }

  /** Stop listening outside the ribbon (the view is closing or redrawing). */
  destroy(): void {
    this.abort.abort();
  }

  select(id: string): void {
    const tab = this.tabs.find((t) => t.id === id);
    if (!tab) return;
    this.currentTab = id;
    if (!tab.context) this.lastCore = id;
    for (const [tid, b] of this.tabEls) {
      b.setAttribute("aria-selected", String(tid === id));
      b.toggleClass("is-active", tid === id);
    }
    this.render(tab);
  }

  /** Toggle states, values and contextual tabs again after the selection changed. */
  sync(): void {
    for (const tab of this.tabs) {
      if (!tab.context) continue;
      const on = tab.context();
      this.tabEls.get(tab.id)!.hidden = !on;
      // Word shows a contextual tab while the cursor is in its object, without switching to it.
      if (on) this.shownContext.add(tab.id);
      else if (this.shownContext.delete(tab.id) && this.currentTab === tab.id) this.select(this.lastCore);
    }
    for (const f of [...this.refresh, ...this.qatRefresh]) f();
  }

  /** An icon button at the right end of the tab row (version history, comments…). */
  side(icon: string, title: string, onClick: () => void): HTMLButtonElement {
    this.sideEl ??= this.tabRow.createDiv({ cls: "dw-rb-side" });
    const b = this.sideEl.createEl("button", { cls: "clickable-icon dw-rb-side-btn", attr: { "aria-label": title, tabindex: "-1" } });
    setIcon(b, icon);
    b.addEventListener("click", onClick);
    return b;
  }

  private clickTab(id: string): void {
    if (this.root.hasClass("is-collapsed")) this.root.toggleClass("is-peek", !(this.currentTab === id && this.root.hasClass("is-peek")));
    this.select(id);
  }

  /** Show or hide the groups (double-click a tab, Ctrl+F1). */
  toggleCollapsed(): void {
    const collapsed = !this.root.hasClass("is-collapsed");
    this.root.toggleClass("is-collapsed", collapsed);
    this.root.removeClass("is-peek");
    this.opts.app.saveLocalStorage(`${this.opts.key}-collapsed`, collapsed ? "collapsed" : "open");
  }

  /** Run a command from the ribbon: a peeking (collapsed) ribbon closes, states refresh. */
  private invoke(fn: (() => void) | undefined): void {
    this.root.removeClass("is-peek");
    this.closeFlyout();
    fn?.();
    this.sync();
  }

  private render(tab: Tab): void {
    this.closeFlyout();
    this.panel.empty();
    this.refresh = [];
    this.groupInfo.clear();
    for (const group of tab.groups) {
      const g = this.panel.createDiv({ cls: "dw-rb-group", attr: { "data-group": group.id, "data-size": "full" } });
      this.groupInfo.set(g, group);
      const controls = g.createDiv({ cls: "dw-rb-controls" });
      for (const c of group.controls) this.control(controls, c);
      const caption = g.createDiv({ cls: "dw-rb-caption" });
      caption.createSpan({ cls: "dw-rb-caption-label", text: group.label });
      const launcher = group.launcher ? this.cmds[group.launcher] : undefined;
      if (launcher) {
        const b = caption.createEl("button", { cls: "dw-rb-launcher", attr: { "aria-label": tooltip(launcher) } });
        setIcon(b, "square-arrow-out-down-right");
        b.addEventListener("click", () => this.invoke(launcher.run));
      }
    }
    // One Tab stop in the panel; arrow keys move inside (rovingToolbar's listeners stay on the panel).
    const items = Array.from(this.panel.querySelectorAll<HTMLElement>("button, input"));
    items.forEach((el, i) => (el.tabIndex = i ? -1 : 0));
    for (const f of this.refresh) f();
    this.fit();
  }

  private control(parent: HTMLElement, c: Control): void {
    if (isStack(c)) {
      const col = parent.createDiv({ cls: c.kind === "column" ? "dw-rb-col" : "dw-rb-row" });
      for (const item of c.items) this.control(col, item);
      return;
    }
    const cmd = this.cmds[c.id];
    switch (c.kind) {
      case "large":
      case "small":
      case "icon":
        return this.button(parent, c.kind, c.id, cmd);
      case "select":
        return this.selectControl(parent, c, cmd);
      case "color":
        return this.color(parent, c, cmd);
      case "check":
        return this.check(parent, c.id, cmd);
      case "gallery":
        return this.gallery(parent, c, cmd);
    }
  }

  private menuButton(parent: HTMLElement, label: string, build: (m: Menu) => void, cls = "dw-rb-arrow"): HTMLButtonElement {
    const b = parent.createEl("button", { cls, attr: { "aria-label": label, "aria-haspopup": "menu" } });
    caret(b);
    b.addEventListener("click", () => {
      const m = new Menu();
      build(m);
      showBelow(m, b);
    });
    return b;
  }

  private button(parent: HTMLElement, kind: "large" | "small" | "icon", id: string, cmd: Cmd): void {
    const split = cmd.run && cmd.menu;
    const host = split ? parent.createSpan({ cls: `dw-rb-split is-${kind}` }) : parent;
    const b = host.createEl("button", { cls: `dw-rb-btn is-${kind}`, attr: { "aria-label": tooltip(cmd), "data-cmd": id } });
    setIcon(b.createSpan({ cls: "dw-rb-icon" }), cmd.icon);
    if (kind !== "icon") b.createSpan({ cls: "dw-rb-label", text: cmd.label });
    if (cmd.run) b.addEventListener("click", () => this.invoke(cmd.run));
    if (cmd.dblclick) b.addEventListener("dblclick", () => this.invoke(cmd.dblclick));
    else if (cmd.menu) {
      b.setAttribute("aria-haspopup", "menu");
      caret(b);
      b.addEventListener("click", () => {
        const m = new Menu();
        cmd.menu!(m);
        showBelow(m, b);
      });
    }
    if (split) this.menuButton(host, `${cmd.label} options`, cmd.menu!);
    this.refresh.push(() => {
      if (cmd.active) {
        const on = cmd.active();
        b.toggleClass("is-active", on);
        b.setAttribute("aria-pressed", String(on));
      }
      if (cmd.enabled) b.disabled = !cmd.enabled();
    });
  }

  private selectControl(parent: HTMLElement, c: Extract<Control, { kind: "select" }>, cmd: Cmd): void {
    const b = parent.createEl("button", { cls: "dw-rb-select", attr: { "aria-label": cmd.label, "aria-haspopup": "menu", "data-cmd": c.id } });
    if (c.width) b.setCssProps({ "--dw-rb-select-w": `${c.width}em` });
    const label = b.createSpan({ cls: "dw-rb-select-label" });
    caret(b);
    b.addEventListener("click", () => {
      const m = new Menu();
      const value = c.value();
      for (const [text, v] of c.options()) m.addItem((i) => i.setTitle(text).setChecked(v === value).onClick(() => this.invoke(() => c.set(v))));
      showBelow(m, b);
    });
    this.refresh.push(() => {
      const value = c.value();
      label.setText(c.options().find(([, v]) => v === value)?.[0] ?? value ?? "");
    });
  }

  private color(parent: HTMLElement, c: Extract<Control, { kind: "color" }>, cmd: Cmd): void {
    let current = c.initial;
    const host = parent.createSpan({ cls: "dw-rb-split is-icon" });
    const b = host.createEl("button", { cls: "dw-rb-btn is-icon dw-rb-color", attr: { "aria-label": cmd.label, "data-cmd": c.id } });
    setIcon(b.createSpan({ cls: "dw-rb-icon" }), cmd.icon);
    const swatch = b.createSpan({ cls: "dw-rb-swatch" });
    swatch.setCssProps({ "--dw-swatch": current });
    b.addEventListener("click", () => this.invoke(() => c.set(current)));
    this.menuButton(host, `${cmd.label} options`, (m) => {
      cmd.menu?.(m);
      m.addItem((i) => i.setTitle("More Colors…").onClick(() => pickColor(current, (hex) => {
        current = hex;
        swatch.setCssProps({ "--dw-swatch": hex });
        this.invoke(() => c.set(hex));
      })));
    });
  }

  private check(parent: HTMLElement, id: string, cmd: Cmd): void {
    const label = parent.createEl("label", { cls: "dw-rb-check" });
    const box = label.createEl("input", { type: "checkbox", attr: { "data-cmd": id } });
    label.appendText(cmd.label);
    box.addEventListener("change", () => this.invoke(cmd.run));
    this.refresh.push(() => (box.checked = !!cmd.active?.()));
  }

  private gallery(parent: HTMLElement, c: Extract<Control, { kind: "gallery" }>, cmd: Cmd): void {
    const host = parent.createDiv({ cls: "dw-rb-gallery", attr: { "data-cmd": c.id, role: "group", "aria-label": cmd.label } });
    const list = host.createDiv({ cls: "dw-rb-gallery-items" });
    const buttons = c.items().map(([text, v]) => {
      const b = list.createEl("button", { cls: "dw-rb-gallery-item", text, attr: { "data-value": v } });
      const p = c.preview?.(v);
      if (p) {
        // As Word's gallery: the style's font, weight and colour; sizes squeezed into the tile (9–16 px).
        b.addClass("has-preview");
        const css: Record<string, string> = {};
        if (p.font) css["font-family"] = `"${p.font}", var(--font-interface)`;
        if (p.size) css["font-size"] = `${Math.max(9, Math.min(16, p.size * 1.1))}px`;
        if (p.bold) css["font-weight"] = "bold";
        if (p.italic) css["font-style"] = "italic";
        // Near-black or near-white text would vanish on one of the themes: those keep the theme's text colour.
        if (p.color && /^#[0-9a-f]{6}$/i.test(p.color) && luminance(p.color) > 0.08 && luminance(p.color) < 0.85) css.color = p.color;
        b.setCssProps(css);
      }
      b.addEventListener("click", () => this.invoke(() => c.set(v)));
      return b;
    });
    this.menuButton(host, `${cmd.label} options`, (m) => {
      const value = c.value();
      for (const [text, v] of c.items()) m.addItem((i) => i.setTitle(text).setChecked(v === value).onClick(() => this.invoke(() => c.set(v))));
      if (cmd.menu) {
        m.addSeparator();
        cmd.menu(m);
      }
    });
    this.refresh.push(() => {
      const value = c.value();
      for (const b of buttons) b.toggleClass("is-active", b.dataset.value === value);
    });
  }
}

/** Shorthands for laying out a tab. */
export const large = (id: string): Control => ({ kind: "large", id });
export const small = (id: string): Control => ({ kind: "small", id });
export const icon = (id: string): Control => ({ kind: "icon", id });
export const col = (...items: Control[]): Control => ({ kind: "column", items });
export const row = (...items: Control[]): Control => ({ kind: "row", items });
