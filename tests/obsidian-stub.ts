/*
 * Runtime stand-in for the `obsidian` package (types only on npm), aliased in
 * vitest.config.mjs. Covers just what the views touch on load / unload / find /
 * gallery; extend it as tests reach further.
 */
/* eslint-disable @typescript-eslint/no-explicit-any */
type Fn = (...args: any[]) => any;

export class Scope {
  keys: { mods: string[]; key: string; fn: Fn }[] = [];
  constructor(public parent?: Scope) {}
  register(mods: string[], key: string, fn: Fn) {
    const k = { mods, key, fn };
    this.keys.push(k);
    return k;
  }
}

export class TFile {
  constructor(public path: string) {}
  get name() { return this.path.slice(this.path.lastIndexOf("/") + 1); }
  get basename() { return this.name.replace(/\.[^.]*$/, ""); }
  get extension() { return this.name.slice(this.name.lastIndexOf(".") + 1); }
}
export class TFolder {}

/** The vault's local storage behind App.loadLocalStorage / saveLocalStorage, shared by every App. */
export const localData = new Map<string, unknown>();
/** Desktop by default; tests flip it to check the mobile app. */
export const Platform = { isDesktopApp: true, isMobile: false };

export class App {
  scope = new Scope();
  loadLocalStorage(key: string): unknown {
    return localData.get(key) ?? null;
  }
  saveLocalStorage(key: string, data: unknown): void {
    if (data === null || data === undefined) localData.delete(key);
    else localData.set(key, data);
  }
  vault = {
    readBinary: async (_f: TFile): Promise<ArrayBuffer> => new ArrayBuffer(0),
    modifyBinary: async (_f: TFile, _d: ArrayBuffer) => {},
    on: (_name: string, _cb: Fn) => ({}),
  };
}

export class WorkspaceLeaf {
  constructor(public app: App) {}
}

export class FileView {
  app: App;
  containerEl = createDiv({ cls: "workspace-leaf-content" });
  contentEl = this.containerEl.createDiv({ cls: "view-content" });
  scope: Scope | null = null;
  file: TFile | null = null;
  constructor(public leaf: WorkspaceLeaf) {
    this.app = leaf.app;
  }
  registerEvent(_ref: unknown) {}
  onPaneMenu(_menu: Menu, _source: string) {}
}

export class Notice {
  constructor(public message: string, _duration?: number) {}
  hide() {}
}

export class Menu {
  addItem(cb: Fn) {
    const item: any = { setTitle: () => item, setIcon: () => item, setSection: () => item, onClick: () => item, setChecked: () => item, setDisabled: () => item };
    cb(item);
    return this;
  }
  hideCallbacks: Fn[] = [];
  addSeparator() { return this; }
  onHide(cb: Fn) { this.hideCallbacks.push(cb); return this; }
  hide() { this.hideCallbacks.forEach((cb) => cb()); return this; }
  showAtPosition() { Menu.last = this; }
  showAtMouseEvent() { Menu.last = this; }
  static last: Menu | null = null;
}

export class Modal {
  modalEl = createDiv();
  titleEl = createDiv();
  contentEl = createDiv();
  onClose() {}
  constructor(public app: App) {}
  open() { this.modalEl.addClass("test-modal"); this.modalEl.append(this.titleEl,this.contentEl); document.body.append(this.modalEl); (this as any).onOpen?.(); }
  close() { this.modalEl.remove(); this.onClose(); }
}

export class Setting {
  settingEl: HTMLElement;
  constructor(public containerEl: HTMLElement) {
    this.settingEl = containerEl.createDiv({ cls: "setting-item" });
  }
  setName(name: string) { this.settingEl.createDiv({ cls: "setting-item-name", text: name }); return this; }
  setDesc() { return this; }
  setHeading() { return this; }
  private control(cb: Fn, kind = "text") {
    const input = kind === "select" ? this.settingEl.createEl("select") : this.settingEl.createEl("input", { type: kind });
    const c: any = {
      setValue: (value: any) => { if (kind === "checkbox") (input as HTMLInputElement).checked = value; else input.value = String(value); return c; },
      onChange: (fn: Fn) => { input.addEventListener("change", () => fn(kind === "checkbox" ? (input as HTMLInputElement).checked : kind === "range" ? Number(input.value) : input.value)); return c; },
      addOptions: (options: Record<string, string>) => { for (const [value, text] of Object.entries(options)) input.createEl("option", { value, text }); return c; },
      setPlaceholder: () => c,
      setLimits: (min: number, max: number, step: number) => { Object.assign(input, { min: String(min), max: String(max), step: String(step) }); return c; },
      setDynamicTooltip: () => c, inputEl: input,
    };
    cb(c);
    return this;
  }
  addToggle(cb: Fn) { return this.control(cb, "checkbox"); }
  addDropdown(cb: Fn) { return this.control(cb, "select"); }
  addSlider(cb: Fn) { return this.control(cb, "range"); }
  addText(cb: Fn) { return this.control(cb); }
  addButton(cb: Fn) {
    const el = this.settingEl.createEl("button");
    const b: any = { setButtonText: (text: string) => { el.textContent = text; return b; }, setCta: () => b, onClick: (fn: () => void) => { el.addEventListener("click", fn); return b; } };
    cb(b); return this;
  }
}

export class Plugin {}

export function setIcon(_el: HTMLElement, _icon: string) {}

export class Component {
  load() { this.onload(); }
  onload() {}
  unload() { this.onunload(); }
  onunload() {}
  register(_cb: () => unknown) {}
}

export const MarkdownRenderer = {
  async render(_app: App, _markdown: string, _el: HTMLElement, _source: string, _component: Component) {},
};
export async function finishRenderMath() {}
export function parseLinktext(link: string) {
  const hash = link.indexOf("#");
  return hash < 0 ? { path: link, subpath: "" } : { path: link.slice(0, hash), subpath: link.slice(hash) };
}
export function resolveSubpath(_cache: unknown, _subpath: string): any { return null; }
