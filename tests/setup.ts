/*
 * Minimal stand-ins for the DOM helpers Obsidian adds globally (createEl,
 * createDiv, createSpan and their Node methods), so plain-DOM modules can use
 * them and still run under jsdom.
 */
type Info = {
  cls?: string | string[];
  text?: string;
  attr?: Record<string, string | number | boolean | null>;
  value?: string;
  type?: string;
  placeholder?: string;
  title?: string;
};

function create<K extends keyof HTMLElementTagNameMap>(tag: K, o?: Info | string, parent?: Node): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  const info = typeof o === "string" ? { cls: o } : o ?? {};
  if (info.cls) el.className = ([] as string[]).concat(info.cls).join(" ");
  if (info.text) el.textContent = info.text;
  for (const k of ["type", "value", "placeholder", "title"] as const) if (info[k] !== undefined) el.setAttribute(k, info[k]);
  if (info.value !== undefined) (el as HTMLInputElement).value = info.value;
  for (const [k, v] of Object.entries(info.attr ?? {})) if (v !== null && v !== false) el.setAttribute(k, String(v));
  parent?.appendChild(el);
  return el;
}

const g = globalThis as Record<string, unknown>;
g.createEl = (tag: keyof HTMLElementTagNameMap, o?: Info | string) => create(tag, o);
g.createDiv = (o?: Info | string) => create("div", o);
g.createSpan = (o?: Info | string) => create("span", o);
g.createFragment = () => document.createDocumentFragment();
Object.assign(Node.prototype, {
  createEl(this: Node, tag: keyof HTMLElementTagNameMap, o?: Info | string) {
    return create(tag, o, this);
  },
  createDiv(this: Node, o?: Info | string) {
    return create("div", o, this);
  },
  createSpan(this: Node, o?: Info | string) {
    return create("span", o, this);
  },
});
