# Doc Works — Obsidian .docx / .pptx editor

GPL-3.0.

## Commands

- `npm test` — vitest (jsdom). `DOC_WORKS_DECKS=<dir>` / `DOC_WORKS_DOCS=<dir>` also round-trip every real .pptx / .docx there.
- `DOC_WORKS_EXPORT=<dir> npx vitest run tests/export-samples.test.ts` writes heavily edited copies; open them with LibreOffice headless (`soffice --headless --convert-to pdf`) and compare its warnings against the originals to catch corrupt output.
- `npm run typecheck` — `tsc --noEmit` (lib ES2022).
- `npm run lint` — eslint-plugin-obsidianmd (what the community directory scanner checks); keep it at 0 errors. Use `attrsOf<T>()` for ProseMirror attrs instead of `node.attrs.x`.
- `npm run build` — bundles everything (jszip, prosemirror) into `main.js`. Ship only `main.js`, `manifest.json`, `styles.css`.

## Layout

- `src/core/` — xml namespaces/helpers, `History<T>` undo stack.
- `src/infrastructure/ooxml/ooxml-package.ts` — shared OPC zip: every XML part as a live DOM, rels, content types, media, snapshot/restore, save.
- `src/infrastructure/{pptx,docx}/` — format packages on top of it (slide ops; docx main part, images, numbering/`ensureList`).
- `src/domain/pptx/` — `color.ts`, `geometry.ts` (DrawingML formula engine: presets from generated `preset-shapes.ts` + custGeom -> SVG paths, text rects, connection sites, adjust handles), `shape-style.ts` (fill/line/shadow read incl. theme style matrix, CSS, writers in schema order), `slide-model.ts` (layout/master inheritance, text style chain), `slide-edit.ts` (minimal-diff text writer incl. numbering/spacing/links/fields, bodyPr settings, shape insert/transform/z-order/connectors), `arrange.ts` (items = direct children of spTree: bounds, group/ungroup, align/distribute, snapping, copy/paste with id + rel remapping).
- `src/domain/docx/` — `schema.ts` (PM schema; lists = paragraph numId/ilvl, headings = styles, `opaque`/`raw_block` atoms hold XML we don't edit), `docx-reader.ts` (body -> PM doc + `origins` key -> original XML), `docx-writer.ts` (verbatim copy if unchanged, else patch original pPr/rPr/tcPr in schema order), `docx-styles.ts` (display-only style CSS + list labels).
- `src/presentation/pptx/` — `text-html.ts`, `text-fit.ts` (letter-spacing to match Office font widths from generated `font-metrics.ts`), `slide-renderer.ts`, `slideshow.ts` (slideshow + presenter view, plain DOM), `shape-gallery.ts`, `shape-panel.ts` (format panel, plain DOM), `pptx-view.ts`.
- `src/presentation/docx/` — `docx-editor.ts` (PM commands, keymap, input rules, decorations; no Obsidian), `docx-image-view.ts`, `docx-view.ts` (FileView, toolbar, autosave).

## Rules

- Only `main.ts`, `*-view.ts` and `presentation/common/` import from `obsidian`. Other modules run under jsdom; they may use Obsidian's global DOM helpers `createEl`/`createDiv`/`createSpan` (shimmed in `tests/setup.ts`, which the directory scanner prefers over `document.createElement`) but nothing else Obsidian-specific.
- Writers must not touch what the user didn't change: unchanged runs/paragraphs/tables are cloned verbatim, edited ones only rewrite differing modeled props. Real-file tests assert byte-identical no-op writes for both formats.
- Manifest description must not contain "Obsidian". Avoid `!important` and `text-indent` in styles.css (directory CSS lint).
- Never dispatch transactions on load (image node views etc.) — opening a file must not save it.
- Keep OOXML child order valid (pptx: bullets before `tabLst`/`defRPr`; docx: `ORDER` table in docx-writer) — Office rejects misordered XML.
- Editor shortcuts that clash with app hotkeys (Ctrl+G...) go through the view's `Scope`, never a plain keydown listener.
- Autofit only runs after a text edit or an explicit resize/setting change; opening or clicking must not change the file.
- View type ids must be unique (`doc-works-*`); Office Suite already owns `docx-editor-view`.
- TDD: add a failing test (fixtures via `tests/fixtures/make_*_fixture.py`) before changing parser/writer behaviour.

## Live testing in Obsidian

Vault: `~/notes`, plugin dir `~/notes/.obsidian/plugins/doc-works/`, scratch files in `~/notes/doc-works-test/`.

```bash
npm run build && cp main.js manifest.json styles.css ~/notes/.obsidian/plugins/doc-works/
obsidian plugin:reload id=doc-works && obsidian dev:errors
```

`obsidian dev:screenshot` output can lag a command behind; re-run it if the image looks stale. The user works in the same vault: open test files in a new tab and check `view.file.path` inside the same `obsidian eval` before any scripted edit.
