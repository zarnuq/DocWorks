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
}

export class Notice {
  constructor(public message: string) {}
}

export class Menu {
  addItem(cb: Fn) {
    const item: any = { setTitle: () => item, setIcon: () => item, onClick: () => item, setChecked: () => item, setDisabled: () => item };
    cb(item);
    return this;
  }
  addSeparator() { return this; }
  showAtPosition() {}
  showAtMouseEvent() {}
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
  private control(cb: Fn) {
    const c: any = { setValue: () => c, onChange: () => c, addOptions: () => c, setPlaceholder: () => c, setLimits: () => c, inputEl: this.settingEl.createEl("input") };
    cb(c);
    return this;
  }
  addToggle(cb: Fn) { return this.control(cb); }
  addDropdown(cb: Fn) { return this.control(cb); }
  addText(cb: Fn) { return this.control(cb); }
  addButton(cb: Fn) { const b: any = { setButtonText: () => b, setCta: () => b, onClick: () => b }; cb(b); return this; }
}

export class Plugin {}

export function setIcon(_el: HTMLElement, _icon: string) {}
