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
Object.assign(HTMLElement.prototype, {
  empty(this: HTMLElement) {
    this.replaceChildren();
  },
  addClass(this: HTMLElement, ...c: string[]) {
    this.classList.add(...c);
  },
  removeClass(this: HTMLElement, ...c: string[]) {
    this.classList.remove(...c);
  },
  toggleClass(this: HTMLElement, c: string, on: boolean) {
    this.classList.toggle(c, on);
  },
  setText(this: HTMLElement, t: string) {
    this.textContent = t;
  },
  appendText(this: HTMLElement, t: string) {
    this.append(t);
  },
  setCssProps(this: HTMLElement, props: Record<string, string>) {
    for (const [k, v] of Object.entries(props)) this.style.setProperty(k, v);
  },
  setCssStyles(this: HTMLElement, styles: Partial<CSSStyleDeclaration>) {
    Object.assign(this.style, styles);
  },
  hide(this: HTMLElement) {
    this.style.display = "none";
  },
  show(this: HTMLElement) {
    this.style.display = "";
  },
  toggle(this: HTMLElement, on: boolean) {
    this.style.display = on ? "" : "none";
  },
});
// jsdom lacks these; the views only construct / observe / disconnect.
g.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
};
g.CSS ??= {};
Element.prototype.scrollIntoView ??= () => {};

// jsdom does no layout: ranges have no boxes.
if (!Range.prototype.getClientRects) {
  Range.prototype.getClientRects = () => [] as unknown as DOMRectList;
  Range.prototype.getBoundingClientRect = () => new DOMRect();
}
