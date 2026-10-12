# AGENTS.md — Doc Works handoff

Obsidian plugin that views and edits .docx, .pptx and .xlsx files (id `doc-works`, repo zarnuq/DocWorks, GPL-3.0).
**Read `CLAUDE.md` first**: commands, code layout, rules and live-testing steps. This file adds how the owner
works and where the project stands.

## Owner's rules

- **Never run git commands** (commit, tag, push, branch). Give the owner the exact commands to run.
- Features should match Google Docs / Google Slides behaviour. When unsure what a feature should do, do what Google does.
- Keep the plugin light and self-contained: no new runtime dependencies (LibreOffice / fonttools are dev-only checks).
- Don't mention forks or upstream projects anywhere.
- TDD: write a failing test first; write many tests when the logic is fiddly. Run `npm test`, `npm run typecheck`
  and `npm run lint` (0 errors) before calling something done.
- "Make sure it works properly": after building, check changes live in Obsidian (see CLAUDE.md), not just in tests.
- The owner uses the same vault while you test: open test files in a **new tab**, check `view.file.path` before any
  scripted edit, never touch their files, and delete scratch files / close test tabs afterwards.
- Community directory review: manifest description without "Obsidian"; no `!important`, `text-indent` or
  multi-column CSS in styles.css (set inline from code); no `window.localStorage` (use `app.loadLocalStorage`).
- Replies: short, result first.

## Status (0.9.0, not yet released when written)

Done in 0.9.0: slideshow animations + transitions, Export to PDF (docx + pptx), editable equations (LaTeX + MathML
preview), table-style cell padding, toolbar menus anchored to their buttons, vault storage for view preferences,
multi-column CSS moved inline (cleared the directory's CSS lint warning). minAppVersion is now 1.8.7.

Also implemented: slide animation authoring, theme/master/custom layouts, chart data editing, embedded audio/video,
suggestion mode, comment threads and resolution, floating picture dragging, table pagination in column sections,
document outline navigation and a margin/indent/tab ruler. Review and outline panels fit the document to available space.

Release steps are in CLAUDE.md ("Generated files and releasing"); version is already bumped to 0.9.0 in
manifest.json, package.json, package-lock.json and versions.json.

## Spreadsheets (.xlsx), from plan.md

Built 2026-10-10: plan phases 0–4 plus pictures and PDF export from phase 5 (see CLAUDE.md layout). Checked against the owner's real Google Sheets exports in `~/notes/doc-works-test/*.xlsx` (byte-identical no-op saves, cached formula results reproduced) and a LibreOffice-made fixture; driven live with ydotool.
Done 2026-10-11: conditional formatting (`domain/xlsx/conditional.ts`, drawn per cell, rules dialog), data validation (`validation.ts`: dropdown arrows + menu, refuse / flag bad input, dialog), charts (`charts.ts`: drawn from live cell values with the slide chart renderer, insert from selection, chart editor shared with slides in `presentation/common/chart-form.ts`, delete). Dialog contents are plain DOM in `presentation/xlsx/rule-forms.ts`.
Implemented in the resumed work: notes and native threaded comments (replies, editing, resolve/reopen, deletion and row/column shifts), chart dragging/resizing, print settings and PDF charts/pictures, sheet protection enforcement and cell locking, structured table references, dynamic spills (SEQUENCE/FILTER/SORT/UNIQUE/TRANSPOSE, spill references and elementwise arithmetic), zoom. Password-protected sheets unprotect with their password (`checkPassword`: legacy 16-bit hash or salted, spun SHA-n via WebCrypto). Unsupported formulas retain their caches. Icon sets are drawn as Unicode glyphs (custom x14 icon sets are not read); time-period rules are evaluated against today.

## Open ideas (not started)

- Animation presets match Google's list (appear/disappear, fade in/out, fly in/out ×4, zoom in/out, spin); compound timing and motion paths are editable.
- Done 2026-10-11: structural / header-footer suggestions, chart formatting (`domain/pptx/chart-format.ts` + renderer), per-section page sizes (`sectionSetups`, `sectionWidthsPlugin`; every page is drawn as wide as the widest, narrower ones shaded), text on both sides of pictures (`domain/docx/wrap.ts`, `presentation/docx/docx-wrap.ts`). Justified lines beside pictures are spread per side; several pictures are planned together; narrower pages get their own edge.
- Slides copy / cut / paste between decks (thumbnail menu, Ctrl+C/X/V on the thumbnails): `PptxDocument.copySlides` / `pasteSlides`; another deck's slide brings its master, layouts and theme (Google's "keep original styles"); speaker notes now travel with the slide as plain text.
- README demo files: `~/notes/Doc Works Demo/Doc Works - Field Study.{docx,pptx}` and `Doc Works - Household Budget.xlsx`.
- Word pagination (2026-10-11): line heights from font metrics (`line-heights.ts`); fonts embedded in .docx loaded (`infrastructure/fonts/odttf.ts`); right tabs measured from the drawn text; superscripts don't grow lines; when the page view makes more pages than Word saved (`wordPages`), pages may run up to half the bottom margin long (`layoutPages` `squeeze`); Word's `lastRenderedPageBreak` starts a page when the page is 80% full; drawn pages are corrected to exactly the paper height (`fixPageHeights`). Manual column pages now retain their full column height with display-only padding; forced column breaks are disabled while measuring and balance to their tallest segment. Page breaks and column breaks remain distinct. Open: `(Notes) Elements of Analysis.docx` has five page breaks (two in its 2-column section) and Word saved five pages; page-break reading gives at least six (here and in LibreOffice). Treating them as column breaks (tried 2026-10-12) overflowed the CSS columns into overlapping text and still gave six, so it was reverted; Word's rule is unconfirmed.
- Implemented: SmartArt text editing, combo and scatter chart editing, picture cropping, per-section headers/footers, inline code versus native Word code boxes, default paper theme and file-type settings. 3-D charts are edited like flat ones; changing their type makes them flat (depth axis and 3-D view dropped).
- Native Word drawings: `domain/docx/drawings.ts`, `presentation/docx/drawing-editor.ts` and `docx-drawing-view.ts`. Insert preset shapes/lines/text boxes; edit text, bounds, fill, border and order; drag/resize. Existing simple groups are editable; nested/transformed/custom groups are preserved rather than flattened. Native OOXML retains unmodeled properties and existing text styles.

## Handy dev facts

- LibreOffice for layout checks: `/nix/store/gpmxgr6l4mmgmhz8v9p2860fgw0mpi06-libreoffice-26.8.0.3-wrapped/bin/soffice`
  (or set `SOFFICE`). It substitutes Aptos with Liberation Serif, so don't trust it for Aptos documents.
- `obsidian eval` output is sometimes lost: store results in `window.__m`, then read them with a second eval.
- `obsidian dev:screenshot` can lag one command behind; take it twice.
- Electron's `remote` (dialog, printToPDF) is available; its debugger is already attached by the CLI.
- Test workbooks: `~/notes/doc-works-test/*.xlsx` (Google Sheets exports; the 29 MB discography one has 214k cells and 100+ pictures). Work on `xlsx-scratch-*` copies.
- Test decks with animations: `~/notes/SRA231/slides/*.pptx`; multi-page doc: `~/notes/doc-works-test/Red_Team_Tryout_Template_2026.docx` (16 pages, blank page 2 is real).
