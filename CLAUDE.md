# Doc Works — Obsidian .docx / .pptx editor

GPL-3.0.

## Commands

- `npm test` — vitest (jsdom). `DOC_WORKS_DECKS=<dir>` / `DOC_WORKS_DOCS=<dir>` also round-trip every real .pptx / .docx there.
- `npm run check:files -- <folder>` — all of the above on every .docx / .pptx under a folder (e.g. `~/notes`): byte-identical no-op writes, edited copies via `tests/export-samples.test.ts`, and, when LibreOffice is found (`SOFFICE=<path>`), the LibreOffice warnings each edit adds over its original. Run it after any writer change.
- Layout check: `soffice --headless --convert-to pdf` then `pdftotext -bbox-layout` gives LibreOffice's line positions to compare with the page view. Only trust it for fonts LibreOffice has metric clones of (Arial, Times, Courier); it renders Aptos as Liberation Serif (`pdffonts` shows what it used).
- `npm run typecheck` — `tsc --noEmit` (lib ES2022).
- `npm run lint` — eslint-plugin-obsidianmd (what the community directory scanner checks); keep it at 0 errors. Use `attrsOf<T>()` for ProseMirror attrs instead of `node.attrs.x`.
- `npm run build` — bundles everything (jszip, prosemirror) into `main.js`. Ship only `main.js`, `manifest.json`, `styles.css`.

## Layout

- `src/core/` — xml namespaces/helpers, `History<T>` undo stack, `utils/fonts.ts` (CSS font stacks with metric-compatible fallbacks, shared by both views and by text measuring).
- `src/infrastructure/ooxml/ooxml-package.ts` — shared OPC zip: every XML part as a live DOM, rels, content types, media, snapshot/restore, save.
- `src/infrastructure/blank.ts` — blank 16:9 deck / Letter document for the New presentation / New document commands.
- `src/infrastructure/{pptx,docx}/` — format packages on top of it (slide ops incl. layouts/`setLayout`, background, `setSlideSize`, p14 sections kept in step by `adjustSections`; docx main part, images, numbering/`ensureList`).
- `src/domain/pptx/` — `color.ts`, `geometry.ts` (DrawingML formula engine: presets from generated `preset-shapes.ts` + custGeom -> SVG paths, text rects, connection sites, adjust handles), `shape-style.ts` (fill/line/shadow read incl. theme style matrix, CSS, writers in schema order), `slide-model.ts` (layout/master inheritance, text style chain), `slide-edit.ts` (minimal-diff text writer incl. numbering/spacing/links/fields, bodyPr settings, shape insert/transform/z-order/connectors), `table.ts` (a:tbl read with table-style regions resolved per cell; row/col insert/delete, merge/split keeping hMerge/vMerge consistent, fills, borders, built-in style XML for tableStyles.xml; cells are `kind: "cell"` shapes listed right after their table), `find.ts` (find/replace over slide text incl. cells), `arrange.ts` (items = direct children of spTree: bounds, group/ungroup, align/distribute, snapping, copy/paste with id + rel remapping).
- `src/domain/docx/` — `pages.ts` (page setup from sectPr, `paginate` estimate), `revisions.ts` (list / accept / reject tracked changes on the XML), `schema.ts` (PM schema; lists = paragraph numId/ilvl, headings = styles, `opaque`/`raw_block` atoms hold XML we don't edit, `sdt` = editable content control around blocks (TOC is `kind: "toc"`), `field` mark = cached PAGE/NUMPAGES result the page view fills in, paragraph `tabLeader` = right tab stop drawn as a stretched flex tab; `keepNext` / `keepLines` null = from the style, row `cantSplit`), `docx-reader.ts` (body -> PM doc + `origins` key -> original XML; tracks complex fields `fldChar` begin/separate/end across runs and paragraphs), `docx-writer.ts` (verbatim copy if unchanged, else patch original pPr/rPr/tcPr in schema order), `docx-styles.ts` (display-only style CSS incl. style indents + list labels).
- `src/infrastructure/fonts/eot.ts` — Embedded OpenType -> TrueType (MTX / LZCOMP + CTF), a port of libeot kept under MPL-2.0; decks from Google Slides embed their fonts this way. Unlike libeot it writes composite glyphs' instruction length, the full glyf table and the head checksum, and drops the (unrebuilt) hdmx / VDMX, because browsers reject libeot's output. Dev check: `ots-sanitize` (nixpkgs opentype-sanitizer) on decoded fonts.
- `src/presentation/pptx/` — `embedded-fonts.ts` (registers a deck's fonts with FontFace while it's open), `text-html.ts`, `text-fit.ts` (letter-spacing to match Office font widths from generated `font-metrics.ts`), `slide-renderer.ts`, `slideshow.ts` (slideshow + presenter view, plain DOM), `shape-gallery.ts`, `shape-panel.ts` (format panel + slide panel for an empty selection, plain DOM), `pptx-view.ts`.
- `src/presentation/docx/` — `docx-editor.ts` (PM commands incl. links, spacing, table merges, comment anchors, footnote refs, TOC; decorations incl. `fitText` letter-spacing to Word font widths via `pptx/text-fit.ts`; no Obsidian), `docx-pages.ts` (page view: measures the text flow with page decorations hidden, `layoutPages` in `domain/docx/pages.ts` breaks between blocks (content controls opened up), paragraph lines and inside table cells; in-cell / in-paragraph breaks are spacer widgets carrying the page gap; keep with next / keep lines together / rows that can't split; first-page and even-page headers via `headerFooterType`), `docx-find.ts` (find/replace + word count), `docx-image-view.ts`, `docx-view.ts` (FileView, toolbar, autosave).

## Rules

- Only `main.ts`, `*-view.ts` and `presentation/common/` import from `obsidian`. Other modules run under jsdom; they may use Obsidian's global DOM helpers `createEl`/`createDiv`/`createSpan` (shimmed in `tests/setup.ts`, which the directory scanner prefers over `document.createElement`) but nothing else Obsidian-specific. View tests (`tests/views.test.ts`) run the real views against `tests/obsidian-stub.ts`, aliased as `obsidian` in `vitest.config.mjs`; extend the stub only as far as a test needs.
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

The CLI prints Electron GPU noise on stderr; add `2>/dev/null`. `obsidian dev:screenshot` output can lag a command behind; re-run it if the image looks stale. The user works in the same vault: open test files in a new tab and check `view.file.path` inside the same `obsidian eval` before any scripted edit.

## Generated files and releasing

- `tests/fixtures/sample.{pptx,docx}` come from `tests/fixtures/make_*_fixture.py` (python-pptx / python-docx). `src/presentation/pptx/font-metrics.ts` comes from `scripts/make_font_metrics.py` (fonttools), `src/domain/pptx/preset-shapes.ts` from `scripts/make_preset_shapes.py` (ECMA-376 presets).
- Release: bump `version` in `manifest.json`, `package.json` (+ both entries in `package-lock.json`) and `versions.json`, push a tag with exactly that version (no `v`). The workflow tests, builds and drafts a GitHub release; publish the draft so Obsidian sees it.
