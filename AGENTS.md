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
Not done yet: charts (preserved), comments/notes (preserved), conditional formatting and data validation display/editing (preserved), print settings, sheet protection enforcement, structured table references and dynamic arrays (shown from cache, marked stale after input changes), zoom.

## Open ideas (not started)

- Extend animation effect authoring beyond the basic presets; compound timing and motion paths are editable.
- Improve wrapping on both sides of floating pictures and fidelity of complex mixed section layouts.
- Expand tracked-change recording for advanced structural edits and header/footer editing.
- Chart data is editable; SmartArt data editing and advanced chart formatting remain open. For spreadsheets see "Not done yet" above.

## Handy dev facts

- LibreOffice for layout checks: `/nix/store/gpmxgr6l4mmgmhz8v9p2860fgw0mpi06-libreoffice-26.8.0.3-wrapped/bin/soffice`
  (or set `SOFFICE`). It substitutes Aptos with Liberation Serif, so don't trust it for Aptos documents.
- `obsidian eval` output is sometimes lost: store results in `window.__m`, then read them with a second eval.
- `obsidian dev:screenshot` can lag one command behind; take it twice.
- Electron's `remote` (dialog, printToPDF) is available; its debugger is already attached by the CLI.
- Test workbooks: `~/notes/doc-works-test/*.xlsx` (Google Sheets exports; the 29 MB discography one has 214k cells and 100+ pictures). Work on `xlsx-scratch-*` copies.
- Test decks with animations: `~/notes/SRA231/slides/*.pptx`; multi-page doc: `~/notes/doc-works-test/Red_Team_Tryout_Template_2026.docx` (16 pages, blank page 2 is real).
