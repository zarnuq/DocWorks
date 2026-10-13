# Toolbar Menus + Narrow Layout Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace each editor's wrapping toolbar with Google's layout (menu row + one compact row that folds into ⋮), add selection-dependent toolbar groups, and turn side panels into drawers in narrow panes.

**Architecture:** `presentation/common/toolbar.ts` grows a menu row (`menu()`), pinned / contextual groups (`pin()`, `context()`, `sync()`) and an overflow fitter that hides whole groups and lists their real controls in a ⋮ menu. A new `presentation/common/narrow.ts` toggles `is-narrow` and opens panels as drawers. The three views reorder their existing build code into the new menus; no command is reimplemented.

**Tech Stack:** TypeScript, Obsidian API (`Menu`, `setIcon`), vitest + jsdom with `tests/obsidian-stub.ts`.

**Spec:** `docs/superpowers/specs/2026-10-12-toolbar-menus-design.md`

## Global Constraints

- No new runtime dependencies; nothing touches the network.
- Only `main.ts`, `*-view.ts` and `presentation/common/` import `obsidian`.
- Use `createEl` / `createDiv` / `createSpan`, `.instanceOf()`, `setCssProps`; no `document.createElement`, no `instanceof` on DOM nodes.
- styles.css: no `!important`, no `text-indent`, no multi-column properties.
- View preferences via `app.loadLocalStorage` / `saveLocalStorage` only.
- Toolbar controls keep their existing `aria-label`s (tests and screen readers find them by label).
- Owner runs git: no `git` commands; "Checkpoint" steps list the files changed instead of committing.
- Done = `npm test`, `npm run typecheck`, `npm run lint` (0 errors), `npm run build`, then live check in Obsidian on scratch copies.

## Review Focus

1. A select folded into ⋮ is changed from the menu: its `onChange` must run and `showValue` state must survive (Task 2 test "folded select").
2. A colour control folded into ⋮: Chromium won't open the picker of a `display: none` input, so the menu opens a temporary input and forwards the value (Task 2 test "folded colour").
3. Pane widened again: folded groups must come back (Task 2 test "unfolds when room returns").
4. A contextual group hidden by `sync()` must never appear in ⋮, and one shown while the row is full folds like any other (Task 2 test "context groups").
5. Drawer toggle button clicked while its drawer is open closes it rather than reopening; going wide with a drawer open restores the docked panel state (Task 4 tests).

---

### Task 1: Menu row and menu helpers

**Files:**
- Modify: `src/presentation/common/toolbar.ts`
- Modify: `tests/obsidian-stub.ts` (Menu records checked / disabled / separators)
- Create: `tests/toolbar.test.ts`

**Interfaces:**
- Produces:
  - `Toolbar.menu(name: string, build: (m: Menu) => void): HTMLButtonElement` — button in `.dw-menubar` (created on first call, inserted before the toolbar element); click builds a fresh `Menu` and shows it below the button.
  - `Toolbar.menubar: HTMLElement | null`
  - `menuItem(m: Menu, title: string, run: () => void, opts?: { checked?: boolean; disabled?: boolean; icon?: string }): void`
  - `subMenu(m: Menu, title: string, build: (m: Menu) => void): void` — item titled `${title} ▸` whose click opens a second menu at the click (or the parent's) position.
  - `showBelow(menu: Menu, anchor: HTMLElement): void` (moved here from docx-view)

- [ ] **Step 1: Extend the Menu stub**

In `tests/obsidian-stub.ts` replace the `Menu` class:

```ts
export class Menu {
  /** Items added, by title, so tests can run one. Separators are `{ title: "---" }`. */
  items: { title: string; run?: Fn; checked?: boolean | null; disabled?: boolean }[] = [];
  addItem(cb: Fn) {
    const entry: { title: string; run?: Fn; checked?: boolean | null; disabled?: boolean } = { title: "" };
    const item: any = {
      setTitle: (t: string) => ((entry.title = t), item), setIcon: () => item, setSection: () => item,
      onClick: (fn: Fn) => ((entry.run = fn), item),
      setChecked: (c: boolean | null) => ((entry.checked = c), item),
      setDisabled: (d: boolean) => ((entry.disabled = d), item),
    };
    cb(item);
    this.items.push(entry);
    return this;
  }
  hideCallbacks: Fn[] = [];
  addSeparator() { this.items.push({ title: "---" }); return this; }
  onHide(cb: Fn) { this.hideCallbacks.push(cb); return this; }
  hide() { this.hideCallbacks.forEach((cb) => cb()); return this; }
  showAtPosition() { Menu.last = this; }
  showAtMouseEvent() { Menu.last = this; }
  static last: Menu | null = null;
}
```

Then run `npm test` and fix any existing test that compared `menu.items` titles and now sees `"---"` (filter with `.filter((i) => i.title !== "---")`).

- [ ] **Step 2: Write the failing tests**

`tests/toolbar.test.ts`:

```ts
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
    expect(tb.menubar!.nextElementSibling).toBe(tb.el);
    expect(tb.menubar!.getAttribute("role")).toBe("menubar");
    btn.click();
    expect(Menu.last!.items[0]).toMatchObject({ title: "Bold", checked: false });
    Menu.last!.items[0].run!();
    btn.click();
    expect(Menu.last!.items[0].checked).toBe(true);
  });

  it("opens a submenu from a ▸ item", () => {
    const tb = new Toolbar(document.body.createDiv());
    tb.menu("Insert", (m) => subMenu(m, "Break", (s) => menuItem(s, "Page break", () => {})))
      .click();
    const parent = Menu.last!;
    expect(titles(parent)).toEqual(["Break ▸"]);
    parent.items[0].run!(new MouseEvent("click"));
    expect(titles(Menu.last)).toEqual(["Page break"]);
  });
});
```

- [ ] **Step 3: Run to see it fail**

Run: `npx vitest run tests/toolbar.test.ts`
Expected: FAIL — `menuItem` / `subMenu` not exported, `tb.menu` not a function.

- [ ] **Step 4: Implement**

In `toolbar.ts` add `Menu` to the obsidian import, and:

```ts
/** Open a toolbar menu under its button. */
export function showBelow(menu: Menu, anchor: HTMLElement): void {
  const r = anchor.getBoundingClientRect();
  menu.showAtPosition({ x: r.left, y: r.bottom });
}

export function menuItem(m: Menu, title: string, run: () => void, opts: { checked?: boolean; disabled?: boolean; icon?: string } = {}): void {
  m.addItem((i) => {
    i.setTitle(title).onClick(run);
    if (opts.checked !== undefined) i.setChecked(opts.checked);
    if (opts.disabled) i.setDisabled(true);
    if (opts.icon) i.setIcon(opts.icon);
  });
}

// Where the last menu opened, for submenus opened from the keyboard.
let lastAt = { x: 0, y: 0 };

/** Obsidian's typed Menu has no submenus: a ▸ item opens a second menu where it was clicked. */
export function subMenu(m: Menu, title: string, build: (m: Menu) => void): void {
  m.addItem((i) => i.setTitle(`${title} ▸`).onClick((e) => {
    const sub = new Menu();
    build(sub);
    const at = e instanceof MouseEvent && e.clientX ? { x: e.clientX, y: e.clientY } : lastAt;
    sub.showAtPosition(at);
  }));
}
```

(`e instanceof MouseEvent` is on an event, not a DOM node, so the `.instanceOf` rule doesn't apply; if lint flags it, use `"clientX" in e`.)

In `Toolbar`:

```ts
  menubar: HTMLElement | null = null;

  /** A menu in the menu row above the toolbar; `build` runs each time it opens, so items show the current state. */
  menu(name: string, build: (m: Menu) => void): HTMLButtonElement {
    if (!this.menubar) {
      this.menubar = createDiv({ cls: "dw-menubar" });
      this.el.parentElement?.insertBefore(this.menubar, this.el);
    }
    const b = this.menubar.createEl("button", { text: name, cls: "dw-menu-btn", attr: { "aria-haspopup": "menu" } });
    b.addEventListener("mousedown", (e) => e.preventDefault());
    b.addEventListener("click", () => {
      const menu = new Menu();
      build(menu);
      const r = b.getBoundingClientRect();
      lastAt = { x: r.left, y: r.bottom };
      showBelow(menu, b);
    });
    b.addEventListener("keydown", (e) => {
      if (e.key !== "ArrowDown") return;
      e.preventDefault();
      b.click();
    });
    return b;
  }
```

At the end of `end()` add, next to `rovingToolbar(this.el)`:

```ts
    if (this.menubar) {
      rovingToolbar(this.menubar);
      this.menubar.setAttribute("role", "menubar");
    }
```

Delete `showBelow` from `docx-view.ts` (line ~63) and import it from `../common/toolbar`.

- [ ] **Step 5: Run tests**

Run: `npx vitest run tests/toolbar.test.ts && npm test`
Expected: PASS.

- [ ] **Step 6: Checkpoint** — changed: `toolbar.ts`, `docx-view.ts`, `tests/obsidian-stub.ts`, `tests/toolbar.test.ts`.

---

### Task 2: Pinned / contextual groups and the ⋮ overflow

**Files:**
- Modify: `src/presentation/common/toolbar.ts`
- Modify: `styles.css` (toolbar block near line 16)
- Test: `tests/toolbar.test.ts`

**Interfaces:**
- Produces:
  - `Toolbar.pin(): void` — current group never folds.
  - `Toolbar.context(test: () => boolean): void` — current group shows only while `test()` holds.
  - `Toolbar.sync(): void` — re-evaluate contextual groups, then `fit()`.
  - `Toolbar.fit(): void` — fold unpinned groups from the right while the row overflows.
  - `Toolbar.overflows: () => boolean` — default `() => this.el.scrollWidth > this.el.clientWidth`; tests replace it.
  - `Toolbar.overflowItems(m: Menu): void` — fills a menu with the folded groups' controls (used by ⋮).

- [ ] **Step 1: Write the failing tests**

Append to `tests/toolbar.test.ts`:

```ts
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
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Button 2", "---", "Button 3"]);
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
    expect(Menu.last!.items[0].title).toBe("Font ▸");
    Menu.last!.items[0].run!(new MouseEvent("click"));
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Arial", "Calibri"]);
    Menu.last!.items[1].run!();
    expect([s.value, picked]).toEqual(["Calibri", ["Calibri"]]);
  });

  it("folded colour: opens a temporary picker and forwards its value to the real input", () => {
    const tb = new Toolbar(document.body.createDiv());
    const picked: string[] = [];
    tb.color("A", "Text color", "#000000", (c) => picked.push(c));
    tb.end();
    tb.overflows = () => !tb.el.querySelector<HTMLElement>(".dw-group")!.hidden;
    tb.fit();
    more(tb).click();
    expect(Menu.last!.items[0].title).toBe("Text color…");
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
    expect(Menu.last!.items.map((i) => i.title)).toEqual(["Crop"]);
    picture = false;
    tb.sync();
    expect(more(tb).hidden).toBe(true);
  });

  it("marks active buttons as checked in ⋮", () => {
    const { tb, setRoom } = bar();
    tb.el.querySelector<HTMLElement>('[aria-label="Button 3"]')!.addClass("is-active");
    setRoom(1);
    more(tb).click();
    expect(Menu.last!.items.find((i) => i.title === "Button 3")!.checked).toBe(true);
  });
});
```

- [ ] **Step 2: Run to see it fail**

Run: `npx vitest run tests/toolbar.test.ts`
Expected: FAIL — `pin` / `fit` not functions; `tb.end()` signature (fixed in Task 3, so for now call `tb.end()` and accept a typecheck error until Task 3 — or do Task 3 Step 3 first; the executor may merge Tasks 2 and 3).

- [ ] **Step 3: Implement**

In `Toolbar`:

```ts
  private more: HTMLButtonElement | null = null;
  private tests = new Map<HTMLElement, () => boolean>();
  /** Whether the row is wider than the bar; replaced in tests (jsdom has no layout). */
  overflows = (): boolean => this.el.scrollWidth > this.el.clientWidth;

  /** The current group never folds into ⋮. */
  pin(): void {
    this.group.dataset.pinned = "true";
  }

  /** The current group shows only while `test()` holds (re-checked by `sync`). */
  context(test: () => boolean): void {
    this.group.addClass("dw-context");
    this.tests.set(this.group, test);
  }

  /** Re-check contextual groups after the selection changed, then refit. */
  sync(): void {
    for (const [group, test] of this.tests) group.toggleClass("is-off", !test());
    this.fit();
  }

  private groups(): HTMLElement[] {
    return Array.from(this.el.querySelectorAll<HTMLElement>(":scope > .dw-group")).filter((g) => g.childElementCount && !g.hasClass("is-off"));
  }

  /** Fold unpinned groups from the right into ⋮ until the row fits on one line. */
  fit(): void {
    const all = Array.from(this.el.querySelectorAll<HTMLElement>(":scope > .dw-group"));
    for (const g of all) {
      g.hidden = g.hasClass("is-off");
      delete g.dataset.folded;
    }
    if (this.more) this.more.hidden = true;
    const foldable = this.groups().filter((g) => !g.dataset.pinned);
    if (!this.overflows()) return;
    if (this.more) this.more.hidden = false;
    for (let i = foldable.length - 1; i >= 0 && this.overflows(); i--) {
      foldable[i].hidden = true;
      foldable[i].dataset.folded = "true";
    }
  }

  /** The folded groups' controls as menu items; each one works the real (hidden) control. */
  overflowItems(m: Menu): void {
    const folded = this.groups().filter((g) => g.dataset.folded);
    folded.forEach((g, n) => {
      if (n) m.addSeparator();
      for (const el of Array.from(g.children) as HTMLElement[]) {
        const select = el.matches("select") ? (el as HTMLSelectElement) : el.querySelector("select");
        const color = el.querySelector<HTMLInputElement>('input[type="color"]');
        const label = el.getAttribute("aria-label") ?? select?.getAttribute("aria-label") ?? el.textContent ?? "";
        if (select) {
          subMenu(m, label, (s) => {
            for (const o of Array.from(select.options).filter((o) => o.value !== "")) {
              menuItem(s, o.text, () => {
                select.value = o.value;
                select.dispatchEvent(new Event("change"));
              }, { checked: select.value === o.value });
            }
          });
        } else if (color) {
          menuItem(m, `${label}…`, () => pickColor(color));
        } else if (el.matches("button")) {
          const b = el as HTMLButtonElement;
          menuItem(m, label, () => b.click(), { checked: b.hasClass("is-active") || undefined, disabled: b.disabled });
        }
      }
    });
  }
```

Note: `checked: undefined` must not call `setChecked` (keeps plain items plain); `menuItem` already skips `undefined`. But the "marks active" test expects `checked === true` for the active one only — fine.

Module-level helper:

```ts
/** Chromium won't open the picker of a hidden input: open a temporary one and forward its value. */
function pickColor(real: HTMLInputElement): void {
  const temp = document.body.createEl("input", { type: "color", cls: "dw-color-temp", value: real.value });
  temp.addEventListener("change", () => {
    real.value = temp.value;
    real.dispatchEvent(new Event("change"));
    temp.remove();
  });
  temp.addEventListener("blur", () => window.setTimeout(() => temp.remove(), 1000));
  if (typeof temp.showPicker === "function") {
    try { temp.showPicker(); } catch { temp.click(); }
  } else temp.click();
}
```

`showPicker` doesn't exist in jsdom; the `typeof` guard falls to `click()`, which is harmless.

In `end()`, before creating `.dw-end`, add the ⋮ button and the resize observer:

```ts
    this.more = this.el.createEl("button", { cls: "dw-btn dw-icon dw-more", attr: { "aria-label": "More tools" } });
    setIcon(this.more, "more-vertical");
    this.more.hidden = true;
    this.more.addEventListener("click", () => {
      const menu = new Menu();
      this.overflowItems(menu);
      showBelow(menu, this.more!);
    });
    if (typeof ResizeObserver !== "undefined") new ResizeObserver(() => this.fit()).observe(this.el);
```

The observer is garbage-collected with the element when the view closes (it observes only `this.el`).

- [ ] **Step 4: CSS**

Replace the `.dw-toolbar` rule's `flex-wrap: wrap;` with `flex-wrap: nowrap; overflow: hidden;` and add:

```css
.dw-toolbar > .dw-group { flex-shrink: 0; }
.dw-group[hidden], .dw-more[hidden] { display: none; }

.dw-color-temp {
  position: fixed;
  left: 50%;
  top: 30%;
  width: 0;
  height: 0;
  opacity: 0;
}
```

- [ ] **Step 5: Run tests**

Run: `npx vitest run tests/toolbar.test.ts`
Expected: PASS.

- [ ] **Step 6: Checkpoint** — `toolbar.ts`, `styles.css`, `tests/toolbar.test.ts`.

---

### Task 3: `end()` without the theme switch; menubar in F6; menu-row CSS

**Files:**
- Modify: `src/presentation/common/toolbar.ts` (`end`)
- Modify: `src/presentation/common/keyboard.ts` (`focusTarget`)
- Modify: `styles.css`
- Test: `tests/views.test.ts` ("Toolbar groups"), `tests/keyboard.test.ts`

**Interfaces:**
- Produces: `Toolbar.end(): { status: SaveStatus; el: HTMLElement }` (no arguments).
- Consumes: Task 1's `menubar`.

- [ ] **Step 1: Failing tests**

In `tests/views.test.ts` "puts the controls between gaps…": replace `tb.end("light", () => {}, "paper");` with `tb.end();`, and add after the existing assertion:

```ts
    expect(tb.el.querySelector(".dw-segmented")).toBeNull();
```

In `tests/keyboard.test.ts` `focusRegions` describe, add:

```ts
  it("treats a menu row like a toolbar: F6 lands on its current item", () => {
    const scope = new Scope();
    const menubar = document.body.createDiv({ attr: { role: "menubar" } });
    const file = menubar.createEl("button", { text: "File" });
    const edit = menubar.createEl("button", { text: "Edit" });
    rovingToolbar(menubar);
    menubar.setAttribute("role", "menubar");
    const editor = document.body.createDiv({ attr: { tabindex: "0" } });
    focusRegions(scope as never, document.body, () => [menubar, editor]);
    edit.focus();
    editor.focus();
    scope.keys.find((k) => k.key === "F6" && !k.mods.length)!.fn();
    expect(document.activeElement).toBe(edit);
    void file;
  });
```

- [ ] **Step 2: Run to see failures**

Run: `npx vitest run tests/keyboard.test.ts tests/views.test.ts -t "menu row|between gaps"`
Expected: FAIL (focus lands on `file` because `focusTarget` only knows `role="toolbar"`; `.dw-segmented` still exists / typecheck of `end()` arguments).

- [ ] **Step 3: Implement**

`keyboard.ts` `focusTarget`, first line:

```ts
  if (["toolbar", "menubar"].includes(region.getAttribute("role") ?? "")) return toolbarItems(region).find((el) => el.tabIndex === 0) ?? toolbarItems(region)[0] ?? null;
```

`toolbar.ts` — `end()` becomes:

```ts
  /** Save status (and whatever the view adds) pushed to the end of the bar, after ⋮. */
  end(): { status: SaveStatus; el: HTMLElement } {
    // ⋮ button and ResizeObserver from Task 2 go here.
    const end = this.el.createDiv({ cls: "dw-end" });
    end.dataset.pinned = "true";
    rovingToolbar(this.el);
    if (this.menubar) {
      rovingToolbar(this.menubar);
      this.menubar.setAttribute("role", "menubar");
    }
    return { status: new SaveStatus(end), el: end };
  }
```

Remove the `PaperTheme` import use from `end` (keep the exported type; views still use it).

Each view's F6 region list gets the menu row first: `this.contentEl.querySelector<HTMLElement>(".dw-menubar")` before `.dw-toolbar` (docx-view.ts:152, xlsx-view.ts:137, pptx-view.ts:198).

- [ ] **Step 4: CSS**

```css
.dw-menubar {
  display: flex;
  gap: var(--size-2-1);
  padding: var(--size-2-1) var(--size-4-2) 0;
  background: var(--background-secondary);
  user-select: none;
}

.dw-menu-btn {
  height: auto;
  padding: var(--size-2-1) var(--size-4-2);
  font-size: var(--font-ui-small);
  background: transparent;
  box-shadow: none;
}

.dw-menu-btn:hover, .dw-menu-btn:focus-visible { background: var(--background-modifier-hover); }
```

Add `.dw-menubar` to the `:is(...)` focus-ring selector near styles.css:2443. Delete the `.dw-segmented` rules (and `.dw-segmented button` from the `.dw-toolbar .dw-btn` selector list).

- [ ] **Step 5: Run** `npx vitest run tests/keyboard.test.ts tests/views.test.ts tests/toolbar.test.ts`. Views will fail to typecheck until Tasks 5–7 move them off `end(theme, …)`; vitest doesn't typecheck, so the tests run. Expected: the new tests PASS.

- [ ] **Step 6: Checkpoint** — `toolbar.ts`, `keyboard.ts`, `styles.css`, tests.

---

### Task 4: Narrow panes and drawers

**Files:**
- Create: `src/presentation/common/narrow.ts`
- Modify: `styles.css`
- Create: `tests/narrow.test.ts`

**Interfaces:**
- Produces:
  - `export const NARROW = 700;`
  - `watchNarrow(root: HTMLElement, onChange: (narrow: boolean) => void, width = () => root.clientWidth): () => void` — observes `root`, toggles `is-narrow` on it, calls `onChange` only on change; returns a stop function. Exposes the check as the returned function's `check` property for tests.
  - `openDrawer(panel: HTMLElement, opts: { root: HTMLElement; toggle?: HTMLElement | null; onClose: () => void }): () => void` — shows `panel` with class `is-drawer`, closes on Escape inside it (capture phase, before `escapeToEditor`), a pointerdown outside it (except on `toggle`), or the returned close function; `onClose` runs once.

- [ ] **Step 1: Failing tests** — `tests/narrow.test.ts`:

```ts
import { afterEach, describe, expect, it } from "vitest";
import { openDrawer, watchNarrow } from "../src/presentation/common/narrow";

afterEach(() => document.body.empty());

describe("watchNarrow", () => {
  it("toggles is-narrow below 700 px and reports only changes", () => {
    const root = document.body.createDiv();
    let w = 1000;
    const seen: boolean[] = [];
    const stop = watchNarrow(root, (n) => seen.push(n), () => w);
    stop.check();
    w = 500; stop.check(); stop.check();
    w = 900; stop.check();
    expect(seen).toEqual([true, false]);
    expect(root.hasClass("is-narrow")).toBe(false);
    stop();
  });
});

describe("openDrawer", () => {
  function setup() {
    const root = document.body.createDiv();
    const panel = root.createDiv();
    panel.hide();
    const toggle = root.createEl("button");
    const outside = root.createDiv();
    let closed = 0;
    const close = openDrawer(panel, { root, toggle, onClose: () => closed++ });
    return { root, panel, toggle, outside, close, closed: () => closed };
  }

  it("shows the panel as a drawer and closes on a click outside", () => {
    const { panel, outside, closed } = setup();
    expect(panel.hasClass("is-drawer")).toBe(true);
    expect(panel.style.display).not.toBe("none");
    outside.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(panel.hasClass("is-drawer")).toBe(false);
    expect(closed()).toBe(1);
  });

  it("closes on Escape inside it, letting the Escape through to the panel's own handler", () => {
    const { panel, closed } = setup();
    const button = panel.createEl("button");
    let reached = false;
    panel.addEventListener("keydown", (e) => (reached = !e.defaultPrevented));
    button.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect([closed(), reached]).toEqual([1, true]);
  });

  it("ignores a pointerdown on its toggle (the toggle's click closes it) and closes once", () => {
    const { toggle, close, closed } = setup();
    toggle.dispatchEvent(new MouseEvent("pointerdown", { bubbles: true }));
    expect(closed()).toBe(0);
    close();
    close();
    expect(closed()).toBe(1);
  });
});
```

- [ ] **Step 2: Run** `npx vitest run tests/narrow.test.ts` — Expected: FAIL (module missing).

- [ ] **Step 3: Implement** `src/presentation/common/narrow.ts`:

```ts
/*
 * Narrow panes (Google's phone-width layout): below NARROW px the view gets `is-narrow`, docked side panels are
 * hidden by CSS and open as drawers over the document instead.
 */
export const NARROW = 700;

/** Toggle `is-narrow` on `root` as it resizes; `onChange` runs when it flips. Returns a stop function. */
export function watchNarrow(root: HTMLElement, onChange: (narrow: boolean) => void, width = () => root.clientWidth): (() => void) & { check: () => void } {
  let narrow = false;
  const check = () => {
    const w = width();
    const now = w > 0 && w < NARROW;
    if (now === narrow) return;
    narrow = now;
    root.toggleClass("is-narrow", now);
    onChange(now);
  };
  const observer = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(check);
  observer?.observe(root);
  return Object.assign(() => observer?.disconnect(), { check });
}

/** Show `panel` as a drawer over the document until Escape, a click outside it, or the returned function. */
export function openDrawer(panel: HTMLElement, opts: { root: HTMLElement; toggle?: HTMLElement | null; onClose: () => void }): () => void {
  panel.addClass("is-drawer");
  panel.show();
  let open = true;
  const close = () => {
    if (!open) return;
    open = false;
    panel.removeClass("is-drawer");
    opts.root.removeEventListener("pointerdown", outside, true);
    panel.removeEventListener("keydown", escape, true);
    opts.onClose();
  };
  const outside = (e: Event) => {
    const t = e.target as Node;
    if (!panel.contains(t) && !opts.toggle?.contains(t)) close();
  };
  // Capture: runs before escapeToEditor, which then moves focus back to the document.
  const escape = (e: KeyboardEvent) => {
    if (e.key === "Escape") close();
  };
  opts.root.addEventListener("pointerdown", outside, true);
  panel.addEventListener("keydown", escape, true);
  return close;
}
```

- [ ] **Step 4: CSS**

```css
/* Narrow panes: docked side panels hide; opened ones float over the document's edge. */
.is-narrow :is(.docx-outline, .docx-review, .pptx-props):not(.is-drawer) { display: none; }

.is-drawer {
  position: absolute;
  top: 0;
  bottom: 0;
  z-index: var(--layer-popover);
  width: min(320px, 85%);
  overflow: auto;
  background: var(--background-primary);
  box-shadow: var(--shadow-l);
}

.docx-outline.is-drawer { left: 0; }
.docx-review.is-drawer, .pptx-props.is-drawer { right: 0; }

.is-narrow .pptx-thumbs { width: 48px; min-width: 48px; }
.is-narrow .pptx-thumbs .pptx-thumb-slide { display: none; }
```

The panels' common parent (`.docx-body` / `.pptx-body`, whatever `body` is in the views at docx-view.ts:~250 and pptx-view.ts:~296) needs `position: relative`; add it to their existing rules if absent.

- [ ] **Step 5: Run** `npx vitest run tests/narrow.test.ts` — Expected: PASS.

- [ ] **Step 6: Checkpoint** — `narrow.ts`, `styles.css`, `tests/narrow.test.ts`.

---

### Task 5: Documents view

**Files:**
- Modify: `src/presentation/docx/docx-view.ts` (`buildToolbar` ~504–637, `updateFormatControls` ~662, `contextMenu` image items ~920–936, `onLoadFile` panel setup ~247–260, `setReview`, outline toggle)
- Modify: `src/presentation/docx/docx-image-view.ts` (export `WRAPS`)
- Test: `tests/views.test.ts`

**Interfaces:**
- Consumes: `Toolbar.menu/pin/context/sync/end()`, `menuItem`, `subMenu`, `showBelow`, `watchNarrow`, `openDrawer`.
- Produces (private methods reused by menu, row and context menu): `cropImage(pos: number)`, `altText(pos: number)`, `setWrap(pos: number, wrap: Wrap)`, `selectedImage(): number | null`, `inTable(): boolean`, `toggleOutline()`, `toggleReview()`, `fillTableMenu(m: Menu)` (the body of today's `tableMenu`, which becomes `showBelow` of a menu filled by it).

- [ ] **Step 1: Failing tests** — add to `tests/views.test.ts` (use the file's existing `mount(DocxView, "sample.docx")` helper and its waiting pattern for load):

```ts
describe("DocxView menus", () => {
  const menuTitles = (view: any, name: string) => {
    const b = Array.from(view.contentEl.querySelectorAll<HTMLButtonElement>(".dw-menu-btn")).find((x) => x.textContent === name)!;
    b.click();
    return Menu.last!.items.map((i) => i.title).filter((t) => t !== "---");
  };

  it("has Google's menus with every command", async () => {
    const view = mount(DocxView, "sample.docx");
    await loaded(view);
    expect(Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent)).toEqual(["File", "Edit", "View", "Insert", "Format", "Tools"]);
    expect(menuTitles(view, "Insert")).toEqual(["Image", "Table ▸", "Drawing", "Equation", "Horizontal line", "Break ▸", "Footnote", "Comment", "Link", "Table of contents ▸", "Header & footer", "Watermark"]);
    expect(menuTitles(view, "View")).toEqual(["Editing", "Suggesting", "Page view", "Ruler", "Outline", "Review panel", "Light paper", "Dark paper"]);
    expect(menuTitles(view, "File")).toEqual(["Export…"]);
  });

  it("runs the same command from a menu as from the row", async () => {
    const view = mount(DocxView, "sample.docx");
    await loaded(view);
    const editor = (view as any).editor;
    editor.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, 1, 4)));
    menuTitles(view, "Format");
    Menu.last!.items.find((i) => i.title === "Text ▸")!.run!(new MouseEvent("click"));
    Menu.last!.items.find((i) => i.title === "Bold")!.run!();
    expect(editor.state.doc.rangeHasMark(1, 4, editor.state.schema.marks.bold)).toBe(true);
  });

  it("moves the theme to View and drops the outline / review / page-view buttons from the row", async () => {
    const view = mount(DocxView, "sample.docx");
    await loaded(view);
    const label = (l: string) => view.contentEl.querySelector(`.dw-toolbar [aria-label="${l}"]`);
    expect(view.contentEl.querySelector(".dw-segmented")).toBeNull();
    expect(label("Document outline")).toBeNull();
    expect(label("Bold (Ctrl+B)")).not.toBeNull();
    menuTitles(view, "View");
    Menu.last!.items.find((i) => i.title === "Dark paper")!.run!();
    expect(view.contentEl.hasClass("docx-theme-dark") || (view as any).paperTheme === "dark").toBe(true);
  });

  it("shows picture tools while a picture is selected", async () => {
    const view = mount(DocxView, "sample.docx");  // sample.docx has an image; if not, insert one with cmd as the image tests do
    await loaded(view);
    const editor = (view as any).editor;
    let imagePos = -1;
    editor.state.doc.descendants((n: any, p: number) => { if (n.type.name === "image" && imagePos < 0) imagePos = p; });
    const crop = () => view.contentEl.querySelector<HTMLElement>('.dw-toolbar [aria-label="Crop image"]')!.closest<HTMLElement>(".dw-group")!;
    expect(crop().hidden).toBe(true);
    editor.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, imagePos)));
    expect(crop().hidden).toBe(false);
  });

  it("opens the review panel as a drawer in a narrow pane, leaving its docked state alone", async () => {
    const view = mount(DocxView, "sample.docx");
    await loaded(view);
    const wasOpen = (view as any).reviewOpen;
    (view as any).narrowWatch.check();  // jsdom width 0 = not narrow
    Object.defineProperty(view.contentEl, "clientWidth", { configurable: true, value: 500 });
    (view as any).narrowWatch.check();
    expect(view.contentEl.hasClass("is-narrow")).toBe(true);
    menuTitles(view, "View");
    Menu.last!.items.find((i) => i.title === "Review panel")!.run!();
    const review = view.contentEl.querySelector<HTMLElement>(".docx-review")!;
    expect(review.hasClass("is-drawer")).toBe(true);
    expect((view as any).reviewOpen).toBe(wasOpen);
    review.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    expect(review.hasClass("is-drawer")).toBe(false);
  });
});
```

(If the sample has no image, reuse the image-insert helper from the existing "DocxView images" tests. Use whatever `loaded()` / await pattern neighbouring DocxView tests use; import `NodeSelection`, `TextSelection` from `prosemirror-state` and `Menu` from `./obsidian-stub`.)

- [ ] **Step 2: Run** `npx vitest run tests/views.test.ts -t "DocxView menus"` — Expected: FAIL (no `.dw-menu-btn`).

- [ ] **Step 3: Implement `buildToolbar`**

Rewrite `buildToolbar(bar, ctx)` keeping every handler body as today, regrouped. The row:

```ts
    const tb = new Toolbar(bar);
    this.toolbar = tb;
    btn("↶", "Undo (Ctrl+Z)", undo);
    btn("↷", "Redo (Ctrl+Y)", redo);
    tb.pin();
    sep();
    this.styleSelect = tb.select(/* as today */);
    this.fontSelect = tb.select(/* as today */);
    this.sizeSelect = tb.select(/* as today */);
    tb.pin();
    sep();
    btn("B", "Bold (Ctrl+B)", toggleMark(m.bold), "bold");
    btn("I", "Italic (Ctrl+I)", toggleMark(m.italic), "italic");
    btn("U", "Underline (Ctrl+U)", toggleMark(m.underline), "underline");
    tb.color("A", "Text color", "#000000", (color) => run(cmd.setMark(m.color, { color })));
    tb.color("HL", "Highlight color", "#FFFF00", (color) => run(cmd.setMark(m.highlight, { color })));
    sep();
    tb.icon("link", "Add, change or remove a link (Ctrl+K; Ctrl+click follows it)", () => void this.editLink());
    tb.icon("message-square-plus", "Comment on the selected text", () => void this.addComment());
    tb.icon("image", "Insert image", () => pickImage((f) => void this.insertImage(f)));
    sep();
    tb.iconSelect("align-left", "Align", [["Left", "left"], ["Center", "center"], ["Right", "right"], ["Justify", "justify"]], (v) => run(cmd.setAlign(v as Align)));
    tb.iconSelect("unfold-vertical", "Line spacing", [["1.0", "1"], ["1.15", "1.15"], ["1.5", "1.5"], ["2.0", "2"]], (v) => run(cmd.setLineSpacing(v)));
    sep();
    tb.icon("list", "Bulleted list", () => run(cmd.toggleList(ctx, "bullet")));
    tb.icon("list-ordered", "Numbered list", () => run(cmd.toggleList(ctx, "ordered")));
    sep();
    tb.icon("outdent", "Decrease indent / list level (Shift+Tab)", () => run(cmd.shiftLevel(-1)));
    tb.icon("indent", "Increase indent / list level (Tab)", () => run(cmd.shiftLevel(1)));
    sep();
    btn("Tx", "Clear formatting", cmd.clearFormatting);
    // Picture tools (Google shows them while a picture is selected).
    sep();
    const wrap = tb.select("Text wrapping", "Wrap", WRAPS.map(([v, l]) => [l, v]), (v) => { const p = this.selectedImage(); if (p !== null) this.setWrap(p, v as Wrap); });
    tb.button("Crop", "Crop image", () => { const p = this.selectedImage(); if (p !== null) this.cropImage(p); });
    tb.button("Alt text", "Image description (alt text)", () => { const p = this.selectedImage(); if (p !== null) this.altText(p); });
    tb.context(() => this.selectedImage() !== null);
    // Table cell tools.
    sep();
    tb.color("▇", "Cell fill", "#EBF3FB", (color) => run(cmd.setCellBackground(color)));
    tb.iconSelect("square-dashed", "Table borders", [["All borders", "all"], ["Outside only", "outer"], ["No borders", "none"]], (v) => run(cmd.setTableBorders(v as "all")));
    tb.context(() => this.selectedImage() === null && this.inTable());
```

Then the end area (mode select moves there, theme switch gone):

```ts
    const end = tb.end();
    this.saveStatus = end.status;
    this.wordsEl = createSpan({ cls: "dw-status docx-words" });
    end.el.prepend(mode, this.wordsEl);
```

Create `mode` with `tb.select(...)` as today but before `tb.end()`, in its own `sep()` group at the end of the row; `prepend` moves it into `.dw-end`, leaving an empty group (hidden by `.dw-group:empty`).

Keep the ruler button's toggle logic as a method `toggleRuler()` (used by View menu); page view: `togglePageView()`. Remove `syncTheme` field usage (`setPaperTheme` no longer calls it).

Menus (all `menuItem` / `subMenu` from `../common/toolbar`; `ed = () => this.editor`):

```ts
    tb.menu("File", (mn) => menuItem(mn, "Export…", () => void this.exportAs()));
    tb.menu("Edit", (mn) => {
      menuItem(mn, "Undo", () => run(undo));
      menuItem(mn, "Redo", () => run(redo));
      mn.addSeparator();
      for (const c of ["cut", "copy", "paste"] as const) menuItem(mn, c[0].toUpperCase() + c.slice(1), () => { this.editor?.focus(); if (!runEditCommand(c)) new Notice("Paste with your keyboard."); });
      menuItem(mn, "Select all", () => run(selectAll));
      mn.addSeparator();
      menuItem(mn, "Find and replace", () => this.openFind());
    });
    tb.menu("View", (mn) => {
      menuItem(mn, "Editing", () => setMode("editing"), { checked: !this.suggesting });
      menuItem(mn, "Suggesting", () => setMode("suggesting"), { checked: this.suggesting });
      mn.addSeparator();
      menuItem(mn, "Page view", () => this.togglePageView(), { checked: this.pageView });
      menuItem(mn, "Ruler", () => this.toggleRuler(), { checked: this.rulerViewport?.style.display !== "none" });
      menuItem(mn, "Outline", () => this.toggleOutline(), { checked: this.outlineOpen });
      menuItem(mn, "Review panel", () => this.toggleReview(), { checked: this.reviewOpen });
      mn.addSeparator();
      menuItem(mn, "Light paper", () => this.setPaperTheme("light"), { checked: this.paperTheme === "light" });
      menuItem(mn, "Dark paper", () => this.setPaperTheme("dark"), { checked: this.paperTheme === "dark" });
    });
    tb.menu("Insert", (mn) => {
      menuItem(mn, "Image", () => pickImage((f) => void this.insertImage(f)));
      subMenu(mn, "Table", (s) => this.fillTableMenu(s));
      menuItem(mn, "Drawing", () => this.drawingDialog(null));
      menuItem(mn, "Equation", () => this.equationDialog(null));
      menuItem(mn, "Horizontal line", () => run(cmd.insertRule));
      subMenu(mn, "Break", (s) => {
        menuItem(s, "Page break", () => run(cmd.insertPageBreak));
        menuItem(s, "Column break", () => run(cmd.insertColumnBreak));
        menuItem(s, "Section break (next page)", () => run(cmd.insertSectionBreak("nextPage")));
        menuItem(s, "Section break (continuous)", () => run(cmd.insertSectionBreak("continuous")));
      });
      menuItem(mn, "Footnote", () => void this.addFootnote());
      menuItem(mn, "Comment", () => void this.addComment());
      menuItem(mn, "Link", () => void this.editLink());
      subMenu(mn, "Table of contents", (s) => {
        menuItem(s, "Insert table of contents", () => { this.docx?.prepareToc(); run(cmd.insertToc); });
        menuItem(s, "Update table of contents", () => { if (this.editor && !cmd.updateToc(this.editor.state, this.editor.dispatch)) new Notice("This document has no table of contents."); });
      });
      mn.addSeparator();
      menuItem(mn, "Header & footer", () => this.headerFooterDialog());
      menuItem(mn, "Watermark", () => this.watermarkDialog());
    });
    tb.menu("Format", (mn) => {
      subMenu(mn, "Text", (s) => {
        for (const [title, mark] of [["Bold", m.bold], ["Italic", m.italic], ["Underline", m.underline], ["Strikethrough", m.strike], ["Subscript", m.subscript], ["Superscript", m.superscript]] as const)
          menuItem(s, title, () => run(toggleMark(mark)));
        menuItem(s, "Code", () => run(cmd.toggleCode));
      });
      subMenu(mn, "Paragraph style", (s) => {
        for (const [label, id] of STYLES) menuItem(s, label, () => run(cmd.setStyle(id === "Normal" ? null : id)));
        s.addSeparator();
        menuItem(s, "Update style to match", () => this.updateStyle());
      });
      subMenu(mn, "Align", (s) => {
        for (const a of ["left", "center", "right", "justify"] as Align[]) menuItem(s, a[0].toUpperCase() + a.slice(1), () => run(cmd.setAlign(a)));
      });
      subMenu(mn, "Line spacing", (s) => {
        for (const v of ["1", "1.15", "1.5", "2"]) menuItem(s, v, () => run(cmd.setLineSpacing(v)));
      });
      menuItem(mn, "Paragraph…", () => this.paragraphDialog());
      menuItem(mn, "Columns…", () => this.columnsDialog());
      mn.addSeparator();
      menuItem(mn, "Bulleted list", () => run(cmd.toggleList(ctx, "bullet")));
      menuItem(mn, "Numbered list", () => run(cmd.toggleList(ctx, "ordered")));
      mn.addSeparator();
      menuItem(mn, "Clear formatting", () => run(cmd.clearFormatting));
    });
    tb.menu("Tools", (mn) => {
      menuItem(mn, "Paint format", () => this.startPainting(false));
      menuItem(mn, "Word count", () => new Notice(this.wordsEl?.textContent || "No words"));
      menuItem(mn, "Find and replace", () => this.openFind());
    });
```

`setMode(value)` is today's Mode-select handler extracted to a local function (the select calls it too, and `mode.value` is updated inside it). `selectAll` comes from `prosemirror-commands` (already a dependency). Menu calls `tb.menu(...)` must run **before** `tb.end()` so `end()` makes the menu row roving.

Paint format keeps its button? Spec's Documents row has no paint format button; it's in Tools. Keep `this.brushBtn` assignment working: create it with `tb.icon("paintbrush", …)` inside the Clear-formatting group (Google puts paint format beside undo; spec lists it in Tools only — keep the button too, it's needed for the double-click sticky mode and the `is-active` state while painting).

- [ ] **Step 4: Implement helpers**

```ts
  private toolbar: Toolbar | null = null;

  /** The selected picture's position (a NodeSelection on an image), else null. */
  private selectedImage(): number | null {
    const sel = this.editor?.state.selection;
    return sel instanceof NodeSelection && sel.node.type === schema.nodes.image ? sel.from : null;
  }

  private inTable(): boolean {
    const $from = this.editor?.state.selection.$from;
    if (!$from) return false;
    for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === schema.nodes.table_cell) return true;
    return false;
  }

  private setWrap(pos: number, wrap: Wrap): void {
    const editor = this.editor, node = editor?.state.doc.nodeAt(pos);
    if (!editor || node?.type !== schema.nodes.image) return;
    editor.dispatch(editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, wrap, ...(wrap === "inline" ? { x: 0, y: 0, align: null } : {}) }));
  }
```

`cropImage(pos)` and `altText(pos)`: move the two inline handlers from `contextMenu` (Crop… / Alt text…) into these methods unchanged, and make the context menu call them. `fillTableMenu(m)`: today's `tableMenu` body minus `new Menu()` / `showBelow`; `tableMenu(anchor)` becomes `{ const m = new Menu(); this.fillTableMenu(m); showBelow(m, anchor); }` and is deleted if nothing else calls it (the row no longer has "Table ▾").

`toggleOutline()` / `toggleReview()`: today's outline-button / review-button handlers, minus the button `is-active` toggling, plus the narrow branch:

```ts
  private narrowWatch: ReturnType<typeof watchNarrow> | null = null;
  private closeDrawer: (() => void) | null = null;

  private toggleReview(): void {
    if (this.contentEl.hasClass("is-narrow") && this.reviewEl) return this.drawer(this.reviewEl, () => this.renderReview());
    this.setReview(!this.reviewOpen);
  }

  /** In a narrow pane a panel opens over the document; its docked state is left as it was. */
  private drawer(panel: HTMLElement, render: () => void): void {
    if (this.closeDrawer) { const was = panel.hasClass("is-drawer"); this.closeDrawer(); if (was) return; }
    render();
    this.closeDrawer = openDrawer(panel, { root: this.contentEl, onClose: () => {
      this.closeDrawer = null;
      panel.toggle(panel === this.reviewEl ? this.reviewOpen : this.outlineOpen);
    } });
  }
```

`toggleOutline` is the same with `this.outlineEl` / `renderOutline` and the existing `outlineOpen` save + `fitDocument()` in the wide branch.

In `onLoadFile`, after the panels exist:

```ts
    this.narrowWatch = watchNarrow(this.contentEl, (narrow) => {
      if (!narrow) this.closeDrawer?.();
      this.toolbar?.fit();
      this.fitDocument();
    });
```

and in `onUnloadFile`: `this.narrowWatch?.(); this.closeDrawer?.();`.

At the end of `updateFormatControls()` add `this.toolbar?.sync();`, and `this.wrap` value: when a picture is selected, `showValue(wrapSelect, attrs.wrap)` — store `wrap` as `this.wrapSelect`.

- [ ] **Step 5: Fix older tests** that clicked removed row buttons (`"Table"`, `"Break ▾"`, `"Contents ▾"`, `"Header & footer"`, `"Watermark"`, `"Document outline"`, `"Review panel…"`, `"Page view…"`, `"Show ruler"`, `Light`/`Dark` segmented): switch them to the menu (`menuTitles` + `run`) or call the extracted method. Run `npx vitest run tests/views.test.ts tests/docx-*.test.ts`. Expected: PASS.

- [ ] **Step 6: Checkpoint** — `docx-view.ts`, `docx-image-view.ts`, tests.

---

### Task 6: Slides view

**Files:**
- Modify: `src/presentation/pptx/pptx-view.ts` (`buildToolbar` ~2575–2657, `setSelection` ~998, `contextMenu` picture items ~1350, `slideMenu` ~634, props toggle, `onLoadFile`)
- Modify: `src/presentation/pptx/shape-panel.ts` (export `DASHES`, `EMU_PER_PT`, `spPrOf`)
- Test: `tests/views.test.ts` (PptxView section) or `tests/slideshow.test.ts`'s mount helper, whichever existing PptxView tests use.

**Interfaces:**
- Consumes: Tasks 1–4.
- Produces: `slideMenu(i: number, menu = new Menu()): Menu` (fills a given menu), `replaceImage(el: Element)`, `selectedShape(): RenderedShape | null` (`this.selected` when not inherited), `toggleProps()`.

- [ ] **Step 1: Failing tests**

```ts
describe("PptxView menus", () => {
  it("has Google's menus", async () => {
    const view = await mountDeck();   // the helper existing PptxView tests use
    expect(Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent)).toEqual(["File", "Edit", "View", "Insert", "Format", "Slide", "Arrange"]);
    expect(menuTitles(view, "Arrange")).toEqual(["Order ▸", "Align ▸", "Distribute ▸", "Group", "Ungroup"]);
    expect(menuTitles(view, "Slide")).toContain("New slide");
    expect(menuTitles(view, "Slide")).toContain("Theme…");
  });

  it("shows shape tools for a shape and picture tools for a picture", async () => {
    const view = await mountDeck();
    const group = (label: string) => view.contentEl.querySelector<HTMLElement>(`.dw-toolbar [aria-label="${label}"]`)!.closest<HTMLElement>(".dw-group")!;
    expect(group("Shape fill").hidden).toBe(true);
    const shape = (view as any).rendered.find((r: any) => r.shape.kind === "sp" && !r.shape.inherited);
    (view as any).select(shape);
    expect(group("Shape fill").hidden).toBe(false);
    expect(group("Replace image").hidden).toBe(true);
  });

  it("sets a shape's fill from the row", async () => {
    const view = await mountDeck();
    const shape = (view as any).rendered.find((r: any) => r.shape.kind === "sp" && !r.shape.inherited);
    (view as any).select(shape);
    const input = view.contentEl.querySelector<HTMLInputElement>('.dw-toolbar [aria-label="Shape fill"] input')!;
    input.value = "#ff0000";
    input.dispatchEvent(new Event("change"));
    expect(shape.shape.el.getElementsByTagNameNS("http://schemas.openxmlformats.org/drawingml/2006/main", "srgbClr")[0].getAttribute("val")).toBe("FF0000");
  });
});
```

(`menuTitles` is the helper from Task 5; move it to the top of `views.test.ts` so both describes use it.)

- [ ] **Step 2: Run** — Expected: FAIL.

- [ ] **Step 3: Implement the row**

```ts
    btn("↶", "Undo (Ctrl+Z)", () => this.undo());
    btn("↷", "Redo (Ctrl+Y)", () => this.redo());
    this.brushBtn = tb.icon("paintbrush", "Paint format (double-click to keep painting)", () => this.startPainting(false));
    this.brushBtn.addEventListener("dblclick", () => this.startPainting(true));
    tb.pin();
    sep();
    this.zoomSelect = /* as today */;
    sep();
    btn("Text box", "Insert text box", () => this.addTextBox());
    tb.icon("image", "Insert image from file", () => pickImage((file) => void this.insertImageFile(file)));
    const shapesBtn = btn("Shapes ▾", "Insert a shape or line", () => this.toggleGallery(shapesBtn));
    tb.icon("table", "Insert table", () => void this.insertTable());
    sep();
    tb.icon("message-square", "Comments on this slide", () => this.commentsDialog());
    sep();
    this.fontSelect = /* as today */;
    this.sizeSelect = /* as today */;
    tb.pin();
    sep();
    btn("B", …); btn("I", …); btn("U", …);           // as today
    tb.color("A", "Text color", "#000000", (color) => this.setRunStyle({ color }));
    sep();
    tb.iconSelect("align-left", "Align", [["Left", "l"], ["Center", "ctr"], ["Right", "r"], ["Justify", "just"]], (v) => this.updateParas((p) => ({ ...p, align: v as Align })));
    tb.iconSelect("unfold-vertical", "Line spacing", /* as today */);
    sep();
    tb.icon("list", "Toggle bullets", /* as today */);
    tb.icon("list-ordered", "Toggle numbering", /* as today */);
    sep();
    tb.icon("outdent", …); tb.icon("indent", …);       // as today
    sep();
    btn("Tx", "Clear formatting", () => this.clearFormatting());
    // Shape tools.
    sep();
    const shapeEdit = (fn: (spPr: Element) => void) => { const rs = this.selectedShape(); if (rs) this.editShape(rs.shape.el, (el) => fn(spPrOf(el)), true); };
    tb.color("▇", "Shape fill", "#4472C4", (color) => shapeEdit((spPr) => writeFill(spPr, { type: "solid", color })));
    tb.color("▢", "Border color", "#000000", (color) => shapeEdit((spPr) => writeLine(spPr, { color })));
    tb.iconSelect("minus", "Border weight", ["0.5", "1", "2", "3", "4.5", "6"].map((v) => [`${v} pt`, v]), (v) => shapeEdit((spPr) => writeLine(spPr, { width: Number(v) * EMU_PER_PT })));
    tb.iconSelect("more-horizontal", "Border dash", DASHES, (dash) => shapeEdit((spPr) => writeLine(spPr, { dash })));
    tb.context(() => this.selectedShape()?.shape.kind === "sp" && this.selectedShape()?.shape.el.localName === "sp");
    // Picture tools.
    sep();
    btn("Crop", "Crop image", () => this.showCrop());
    btn("Replace", "Replace image", () => { const rs = this.selectedShape(); if (rs) this.replaceImage(rs.shape.el); });
    tb.iconSelect("rotate-ccw", "Reset image", [["Reset crop", "crop"], ["Reset colour adjustments", "adjust"]], (v) => {
      const rs = this.selectedShape();
      if (!rs) return;
      this.editShape(rs.shape.el, (p) => (v === "crop" ? setCrop(p, { l: 0, t: 0, r: 0, b: 0 }) : setPictureAdjust(p, NO_ADJUST)));
    });
    tb.context(() => this.selectedShape()?.shape.kind === "pic" && !this.selectedShape()?.shape.media);
    sep();
    const present = btn("▶ Present", "Present from this slide (Shift+F5). F5 presents from the beginning.", () => this.startPresentation(false), "mod-cta");
    const end = tb.end();
    this.saveStatus = end.status;
    end.el.prepend(present);
```

`showCrop()`: `this.showProps = true; this.propsEl.show(); this.propsEl.querySelector<HTMLInputElement>('[aria-label="Crop left %"]')?.focus();` (narrow pane: open the props drawer first via `toggleProps()` when it isn't open).

`selectedShape()`: `return this.selected && !this.selected.shape.inherited ? this.selected : null;`

`replaceImage(el)`: the body of the context menu's "Replace image…" item, which now calls it.

Call `this.toolbar?.sync()` at the end of `setSelection` (every return path — put it in a `try { … } finally { this.toolbar?.sync(); }` around the body) and at the end of `updateFormatControls`.

Menus:

```ts
    tb.menu("File", (mn) => menuItem(mn, "Export…", () => void this.exportAs()));
    tb.menu("Edit", (mn) => {
      const has = this.items.length > 0;
      menuItem(mn, "Undo", () => this.undo());
      menuItem(mn, "Redo", () => this.redo());
      mn.addSeparator();
      menuItem(mn, "Cut", () => void this.cutSelection(), { disabled: !has });
      menuItem(mn, "Copy", () => void this.copySelection(), { disabled: !has });
      menuItem(mn, "Paste", () => this.paste(), { disabled: !clipboard });
      menuItem(mn, "Delete", () => this.deleteSelected(), { disabled: !has });
      mn.addSeparator();
      menuItem(mn, "Find and replace", () => this.openFind());
    });
    tb.menu("View", (mn) => {
      menuItem(mn, "Present", () => this.startPresentation(false));
      menuItem(mn, "Present from start", () => this.startPresentation(true));
      mn.addSeparator();
      menuItem(mn, "Grid view", () => this.setGridView(!this.gridView), { checked: this.gridView });
      menuItem(mn, "Format panel", () => this.toggleProps(), { checked: this.showProps });
      subMenu(mn, "Zoom", (s) => { for (const z of ZOOMS) menuItem(s, z === 1 ? "Fit" : `${z * 100}%`, () => this.setZoom(z)); });
      mn.addSeparator();
      menuItem(mn, "Light", () => this.setPaperTheme("light"), { checked: this.paperTheme === "light" });
      menuItem(mn, "Dark", () => this.setPaperTheme("dark"), { checked: this.paperTheme === "dark" });
    });
    tb.menu("Insert", (mn) => {
      menuItem(mn, "Text box", () => this.addTextBox());
      menuItem(mn, "Image", () => pickImage((file) => void this.insertImageFile(file)));
      menuItem(mn, "Audio or video", () => this.pickMedia());
      menuItem(mn, "Shape", () => this.toggleGallery(shapesBtn));
      menuItem(mn, "Table", () => void this.insertTable());
      menuItem(mn, "Chart", () => this.editChart());
      menuItem(mn, "Link", () => void this.editLink());
      subMenu(mn, "Field", (s) => this.fillFieldMenu(s));
      menuItem(mn, "Comment", () => this.commentsDialog());
      mn.addSeparator();
      menuItem(mn, "New slide", () => this.mutate((d) => (this.index = d.addSlide(this.index))));
    });
    tb.menu("Format", (mn) => {
      subMenu(mn, "Text", (s) => {
        menuItem(s, "Bold", () => this.toggle("b")); menuItem(s, "Italic", () => this.toggle("i"));
        menuItem(s, "Underline", () => this.toggle("u")); menuItem(s, "Strikethrough", () => this.toggle("strike"));
      });
      subMenu(mn, "Align", (s) => { for (const [l, a] of [["Left", "l"], ["Center", "ctr"], ["Right", "r"], ["Justify", "just"]] as [string, Align][]) menuItem(s, l, () => this.updateParas((p) => ({ ...p, align: a }))); });
      subMenu(mn, "Line spacing", (s) => { for (const v of ["1.0", "1.15", "1.5", "2.0", "2.5", "3.0"]) menuItem(s, v, () => this.updateParas((p) => ({ ...p, lnSpc: { pct: Number(v) } }))); });
      subMenu(mn, "Character spacing", (s) => { for (const [l, v] of [["Very tight", -3], ["Tight", -1.5], ["Normal", 0], ["Loose", 3], ["Very loose", 6]] as const) menuItem(s, l, () => this.setRunStyle({ spc: v })); });
      menuItem(mn, "Bullets", () => this.updateParas((p) => ({ ...p, bullet: p.bullet && !p.bullet.startsWith("#") ? null : "•" })));
      menuItem(mn, "Numbering", () => this.updateParas((p) => ({ ...p, bullet: p.bullet?.startsWith("#") ? null : "#arabicPeriod" })));
      mn.addSeparator();
      menuItem(mn, "Clear formatting", () => this.clearFormatting());
      menuItem(mn, "Format options", () => { if (!this.showProps) this.toggleProps(); });
    });
    tb.menu("Slide", (mn) => {
      this.slideMenu(this.index, mn);
      mn.addSeparator();
      menuItem(mn, "Theme…", () => { this.showMotion = false; this.showTheme = true; this.showProps = true; this.setSelection([]); });
      menuItem(mn, "Motion…", () => this.openMotion());
    });
    tb.menu("Arrange", (mn) => {
      const a = this.arrangeControl(), has = this.items.length > 0;
      subMenu(mn, "Order", (s) => { for (const [l, z] of [["Bring to front", "front"], ["Bring forward", "forward"], ["Send backward", "backward"], ["Send to back", "back"]] as const) menuItem(s, l, () => a.order(z), { disabled: !has }); });
      subMenu(mn, "Align", (s) => { for (const [l, z] of [["Left", "left"], ["Center", "center"], ["Right", "right"], ["Top", "top"], ["Middle", "middle"], ["Bottom", "bottom"]] as const) menuItem(s, l, () => a.align(z), { disabled: !has }); });
      subMenu(mn, "Distribute", (s) => { menuItem(s, "Horizontally", () => a.distribute("horizontal"), { disabled: a.count < 3 }); menuItem(s, "Vertically", () => a.distribute("vertical"), { disabled: a.count < 3 }); });
      menuItem(mn, "Group", () => this.groupSelection(), { disabled: this.items.length < 2 });
      menuItem(mn, "Ungroup", () => this.ungroupSelection(), { disabled: !this.items.some((i) => i.localName === "grpSp") });
    });
```

`fillFieldMenu(m)`: today's `fieldMenu(anchor)` body without `new Menu()` / `show…`; `fieldMenu` keeps working by filling a new menu. `slideMenu(i, menu = new Menu())`: replace its first line `const menu = new Menu();` with the parameter. `toggleProps()`: today's panel-button handler (minus button state) plus the narrow branch using `openDrawer(this.propsEl, { root: this.contentEl, onClose: () => { this.closeDrawer = null; this.propsEl.toggle(this.showProps); } })` exactly like Task 5's `drawer`. Hook `watchNarrow` in `onLoadFile` (on change: close drawer, `this.toolbar?.fit()`, re-run the existing thumbnail sizing + slide fit that `onResize` does) and stop it in `onUnloadFile`.

Remove `this.gridBtn` / `panelBtn` button-state code (the grid toggle is now in View; keep `Ctrl+Alt+1`). If `setGridView` toggles `this.gridBtn`, guard it with `?.`.

- [ ] **Step 4: Run** `npx vitest run tests/views.test.ts tests/slideshow.test.ts tests/shape-panel.test.ts`; fix older tests that clicked removed buttons ("Slide ▾", "Theme", "Motion", "Grid view…", "Show or hide the format panel", "Delete shape", "Chart", "Audio/video", "Field ▾", "Link") to use menus or methods. Expected: PASS.

- [ ] **Step 5: Checkpoint** — `pptx-view.ts`, `shape-panel.ts`, tests.

---

### Task 7: Sheets view

**Files:**
- Modify: `src/presentation/xlsx/xlsx-view.ts` (`buildToolbar` 961–1032, `setTheme`, `onLoadFile`)
- Test: `tests/xlsx-view.test.ts`

**Interfaces:**
- Consumes: Tasks 1–3 (Sheets has no contextual groups and no side panels).

- [ ] **Step 1: Failing tests** (use `xlsx-view.test.ts`'s existing mount helper):

```ts
describe("sheet menus", () => {
  it("has Google's menus, and Data holds the data tools", async () => {
    const view = await mountBook();
    expect(Array.from(view.contentEl.querySelectorAll(".dw-menu-btn")).map((b) => b.textContent)).toEqual(["File", "Edit", "View", "Insert", "Format", "Data"]);
    expect(menuTitles(view, "Data")).toEqual(["Sort A → Z", "Sort Z → A", "Create a filter", "Data validation", "Remove duplicates…", "Split text to columns…", "Named ranges…", "Group rows 1 - 1", "Group columns A - A", "Import CSV as new sheet…"]);
    expect(view.contentEl.querySelector('.dw-toolbar [aria-label="Data"]')).toBeNull();
    expect(view.contentEl.querySelector('.dw-toolbar [aria-label="Freeze"]')).toBeNull();
  });

  it("freezes from View", async () => {
    const view = await mountBook();
    menuTitles(view, "View");
    Menu.last!.items.find((i) => i.title === "Freeze ▸")!.run!(new MouseEvent("click"));
    Menu.last!.items.find((i) => i.title === "1 row")!.run!();
    expect((view as any).sheet.freeze?.rows ?? (view as any).sheet.frozen?.rows).toBe(1);
  });
});
```

(Adjust the last assertion to however existing freeze tests read the frozen rows — grep `freeze` in `tests/xlsx-view.test.ts` and copy its check.)

- [ ] **Step 2: Run** — Expected: FAIL.

- [ ] **Step 3: Implement** — row (handlers as today):

```
Undo Redo [pin] · Zoom · 123▾ $ % .0← .00→ · Font Size [pin] · B I S U, text colour, fill, No fill, Borders▾, Merge · align left/center/right, V-align▾, Wrap · Comment, Chart, Filter, Functions▾ · end: Saved
```

Comment button: `tb.icon("message-square", "Insert comment", () => this.commentDialog(this.sel.r1, this.sel.c1))`. The Data and Freeze iconSelects, sort buttons, pivot, conditional formatting, validation, find buttons leave the row.

Menus:

```ts
    tb.menu("File", (mn) => {
      menuItem(mn, "Export…", () => void this.exportAs());
      menuItem(mn, "Download sheet as CSV…", () => void this.exportAs("csv"));
    });
    tb.menu("Edit", (mn) => {
      menuItem(mn, "Undo", () => this.undo());
      menuItem(mn, "Redo", () => this.redo());
      mn.addSeparator();
      menuItem(mn, "Cut", () => this.editCommand("cut"));
      menuItem(mn, "Copy", () => this.editCommand("copy"));
      menuItem(mn, "Paste", () => this.editCommand("paste"));
      menuItem(mn, "Paste values only", () => this.editCommand("paste", true));
      mn.addSeparator();
      menuItem(mn, "Find and replace", () => this.openFind());
      this.refocusAfter(mn);
    });
    tb.menu("View", (mn) => {
      subMenu(mn, "Freeze", (s) => {
        for (const [l, v] of [["No rows", "r0"], ["1 row", "r1"], ["2 rows", "r2"], ["Up to current row", "rc"], ["No columns", "c0"], ["1 column", "c1"], ["2 columns", "c2"], ["Up to current column", "cc"]]) menuItem(s, l, () => this.freeze(v));
      });
      subMenu(mn, "Zoom", (s) => { for (const z of ZOOMS) menuItem(s, `${Math.round(Number(z) * 100)}%`, () => { zoom.value = z; zoom.dispatchEvent(new Event("change")); }); });
      mn.addSeparator();
      menuItem(mn, "Light", () => this.setTheme("light"), { checked: this.theme === "light" });
      menuItem(mn, "Dark", () => this.setTheme("dark"), { checked: this.theme === "dark" });
    });
    tb.menu("Insert", (mn) => {
      menuItem(mn, "Sheet", () => { this.change(() => ops.addSheet(this.edits!, this.sheetIndex)); this.showSheet(this.sheetIndex + 1); });
      menuItem(mn, "Chart", () => this.insertChart());
      menuItem(mn, "Pivot table", () => this.insertPivotTable());
      subMenu(mn, "Function", (s) => { for (const f of ["SUM", "AVERAGE", "COUNT", "MAX", "MIN", "IF", "VLOOKUP", "CONCATENATE"]) menuItem(s, f, () => this.insertFunction(f)); });
      mn.addSeparator();
      menuItem(mn, "Comment", () => this.commentDialog(this.sel.r1, this.sel.c1));
      menuItem(mn, "Note", () => void this.editNote(this.sel.r1, this.sel.c1));
      menuItem(mn, "Dropdown", () => this.validationDialog());
    });
    tb.menu("Format", (mn) => {
      subMenu(mn, "Number", (s) => { for (const [l, code] of NUMBER_FORMATS) menuItem(s, l, () => this.format({ numFmt: code })); });
      subMenu(mn, "Text", (s) => { for (const [l, k] of [["Bold", "bold"], ["Italic", "italic"], ["Underline", "underline"], ["Strikethrough", "strike"]] as const) menuItem(s, l, () => this.toggle(k)); });
      subMenu(mn, "Alignment", (s) => {
        for (const [l, h] of [["Left", "left"], ["Center", "center"], ["Right", "right"]] as const) menuItem(s, l, () => this.format({ hAlign: h }));
        s.addSeparator();
        for (const [l, v] of [["Top", "top"], ["Middle", "center"], ["Bottom", "bottom"]] as const) menuItem(s, l, () => this.format({ vAlign: v as "top" }));
      });
      menuItem(mn, "Wrap text", () => this.toggle("wrap"));
      menuItem(mn, "Merge cells", () => this.mergeToggle());
      mn.addSeparator();
      menuItem(mn, "Conditional formatting", () => this.conditionalDialog());
    });
    tb.menu("Data", (mn) => {
      menuItem(mn, "Sort A → Z", () => this.sort(false));
      menuItem(mn, "Sort Z → A", () => this.sort(true));
      menuItem(mn, "Create a filter", () => this.toggleFilter());
      menuItem(mn, "Data validation", () => this.validationDialog());
      mn.addSeparator();
      menuItem(mn, "Remove duplicates…", () => this.removeDuplicatesDialog());
      menuItem(mn, "Split text to columns…", () => this.splitTextDialog());
      menuItem(mn, "Named ranges…", () => this.namedRangesDialog());
      mn.addSeparator();
      // "Group rows 2 - 5" / "Group columns B - C" (+ Ungroup when already grouped), as the header menus offer.
      this.groupItems(mn, "row", this.sel.r1, this.sel.r2);
      this.groupItems(mn, "col", this.sel.c1, this.sel.c2);
      mn.addSeparator();
      menuItem(mn, "Import CSV as new sheet…", () => this.pickCsv());
    });
```

Sheets has no clear-formatting command today, so Format has no such item (spec: no new commands). Make "Create a filter" read "Remove filter" when the sheet has one (mirror `toggleFilter`'s check).

`setTheme` no longer calls `syncTheme`. Each menu calls `this.refocusAfter(mn)` so the grid takes keys back after the menu closes.

- [ ] **Step 4: Run** `npx vitest run tests/xlsx-view.test.ts`; fix tests that used the Data / Freeze iconSelects, sort / pivot / conditional / validation / find buttons. Expected: PASS.

- [ ] **Step 5: Checkpoint** — `xlsx-view.ts`, tests.

---

### Task 8: Docs, full checks, live verification

**Files:**
- Modify: `CLAUDE.md` (Layout: `presentation/common/` line — add toolbar menus / overflow and `narrow.ts`)
- Modify: `AGENTS.md` status if it lists the toolbar

- [ ] **Step 1: CLAUDE.md** — in the `src/presentation/common/export-dialog.ts` bullet, append: "`toolbar.ts`: Google-style menu row (`menu`, `menuItem`, `subMenu`, built on open), compact row whose unpinned groups fold into ⋮ (`fit`, ResizeObserver; ⋮ items work the hidden controls, colours via a temporary picker), contextual groups (`context` + `sync` from the views' selection updates). `narrow.ts`: `is-narrow` below 700 px; side panels open as drawers (`openDrawer`)."

- [ ] **Step 2: Full checks**

Run: `npm test && npm run typecheck && npm run lint && npm run build`
Expected: all tests pass; typecheck clean; lint 0 errors (4 pre-existing warnings).

- [ ] **Step 3: Live check** (CLAUDE.md "Live testing"):

```bash
npm run build && cp main.js manifest.json styles.css ~/notes/.obsidian/plugins/doc-works/
obsidian plugin:reload id=doc-works 2>/dev/null && obsidian dev:errors 2>/dev/null
```

Copy the demo files to `~/notes/doc-works-test/{docx,pptx,xlsx}-scratch-menus.*`, open each in a new tab, check `view.file.path`, then for each editor:
- full width: one toolbar row, menu row above, Saved at the right; screenshot (`obsidian dev:screenshot`, twice).
- each menu opens below its button; a ▸ submenu opens; a menu item runs (e.g. Format › Text › Bold).
- narrow the pane (split the leaf vertically twice or set `leaf.containerEl.style.width = "500px"` via eval): groups fold into ⋮; ⋮ lists them; a folded select and a folded colour work; side panel toggle opens a drawer; Escape closes it and focus returns to the document.
- docx: select a picture → Wrap / Crop / Alt text appear; click in a table → cell fill / borders appear.
- pptx: select a shape → fill / border appear; select a picture → Crop / Replace / Reset.
- keyboard: Tab reaches the menu row, arrows move, ArrowDown opens; F6 cycles menu row → toolbar → document.
- dark theme via View menu.
Delete the scratch files and close the test tabs afterwards.

- [ ] **Step 4: Checkpoint** — report results; owner commits.
