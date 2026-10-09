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

export class App {
  scope = new Scope();
  vault = { readBinary: async (_f: TFile): Promise<ArrayBuffer> => new ArrayBuffer(0), modifyBinary: async (_f: TFile, _d: ArrayBuffer) => {} };
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
  open() {}
  close() { this.onClose(); }
}

export class Setting {
  constructor(public containerEl: HTMLElement) {}
}

export class Plugin {}

export function setIcon(_el: HTMLElement, _icon: string) {}
