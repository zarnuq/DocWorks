#!/usr/bin/env python3
"""GUI test suite: the built plugin in the private agent desktop (agent-desk).

Real keyboard and pointer input goes to a private Obsidian; eval only sets up,
finds coordinates and asserts. Each run edits fresh fixture copies in the
private vault and checks what was saved to disk.

    npm run build && npm run test:gui            # all tests
    npm run test:gui -- docx                     # tests whose name contains "docx"

Needs a running, unpaused `agent-desk start`. Claims the desktop for the run (other
agents are refused) and records it; the video is kept only when a test fails. Works in the real ~/notes (the
private Obsidian uses its own .obsidian-agent config there), only inside a fresh
doc-works-test/gui-<time>/ folder, deleted after a passing run. Never touches
the host desktop. Screenshots and results.json go to the printed folder.
"""
import hashlib
import json
import re
import os
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
STATE = Path(os.environ.get('XDG_STATE_HOME', Path.home() / '.local/state')) / 'agent-desk'
VAULT = Path.home() / 'notes'  # the real vault, opened with its .obsidian-agent config
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
X = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
OUT = Path(tempfile.mkdtemp(prefix='doc-works-gui-'))
# Approved screenshots of the untouched fixtures. Machine/theme specific, so
# they live outside the repo; UPDATE_BASELINES=1 rewrites them after a change
# you have looked at and accept.
BASELINES = STATE / 'baselines/doc-works'
RUN = f'doc-works-test/gui-{time.strftime("%Y%m%d-%H%M%S")}'


def desk(*args):
    # AGENT_DESK_CLAIM (set by main) lets this run act while other agents are refused.
    result = subprocess.run(['agent-desk', *args], capture_output=True, text=True, timeout=60)
    if result.returncode:
        raise RuntimeError(f'agent-desk {args[0]}: {result.stderr.strip()}')
    return result.stdout


def ev(expr):
    """Evaluate a JS expression in the private Obsidian and return it as JSON."""
    out = desk('obsidian-cli', 'eval', f'code=JSON.stringify({expr})')
    # Console lines logged during the eval come first; the per-test error
    # check reports them, so only the result line matters here.
    result = next((line for line in out.splitlines() if line.startswith('=> ')), None)
    if result is None:
        raise RuntimeError(f'eval failed: {out.strip()}')
    return json.loads(result[3:])


def wait(check, what, timeout=15):
    deadline, last = time.monotonic() + timeout, None
    while True:
        try:  # e.g. a just-launched Obsidian answers before `eval` is registered
            value = check()
        except RuntimeError as error:
            value, last = None, error
        if value:
            return value
        if time.monotonic() > deadline:
            raise AssertionError(f'timed out: {what}' + (f' (last error: {last})' if last else ''))
        time.sleep(.25)


def shot(name):
    desk('shot', str(OUT / f'{name}.png'))


def click(x, y):
    desk('move', str(round(x)), str(round(y)))
    desk('click', 'left')


def center(selector, pick='first'):
    """Screen centre of an element in the active leaf (window sits at 0,0, dpr 1)."""
    r = ev(f'''(() => {{ const all = [...document.querySelectorAll(".workspace-leaf.mod-active {selector}")];
        const el = {"all[0]" if pick == "first" else "all.sort((a,b)=>b.offsetWidth*b.offsetHeight-a.offsetWidth*a.offsetHeight)[0]"};
        if (!el) return null; const r = el.getBoundingClientRect(); return [r.x + r.width/2, r.y + r.height/2, r.right, r.y]; }})()''')
    if not r:
        raise AssertionError(f'no element {selector}')
    return r


def open_file(path, view_type):
    """Click it in the file explorer, like a user (the quick switcher may be off).
    Revealing it (expanding its folders, scrolling) is setup."""
    ev(f'''(() => {{ const f = app.vault.getAbstractFileByPath({json.dumps(path)});
        const leaf = app.workspace.getLeavesOfType("file-explorer")[0];
        app.workspace.revealLeaf(leaf); leaf.view.revealInFolder(f); return 1; }})()''')
    selector = f'.nav-file-title[data-path={json.dumps(path)}]'
    wait(lambda: ev(f'''(() => {{ const el = document.querySelector({json.dumps(selector)});
        if (!el) return null; const r = el.getBoundingClientRect();
        return r.height ? [r.x + r.width/2, r.y + r.height/2] : null; }})()'''), f'{path} in the file explorer')
    desk('wait-idle')  # reveal scrolls smoothly
    r = ev(f'''(() => {{ const r = document.querySelector({json.dumps(selector)}).getBoundingClientRect();
        return [r.x + r.width/2, r.y + r.height/2]; }})()''')
    click(*r)
    # Clicking a file that is already open leaves focus in the explorer, so
    # check the most recent editor leaf and focus it (setup, not the feature).
    wait(lambda: ev('[app.workspace.getMostRecentLeaf()?.view.file?.path, app.workspace.getMostRecentLeaf()?.view.getViewType()]')
         == [path, view_type], f'{path} open in {view_type}')
    ev('(app.workspace.setActiveLeaf(app.workspace.getMostRecentLeaf(), {focus: true}), 1)')
    wait(lambda: ev('document.querySelector(".workspace-leaf.mod-active .dw-status-text")?.textContent') == 'Saved', 'view ready')
    desk('wait-idle')


def read_zip(path, part):
    with zipfile.ZipFile(VAULT / path) as z:
        return z.read(part)


def saved(path, check, what):
    """Poll the file on disk until check(file) holds (autosave runs after edits)."""
    def attempt():
        try:
            return check(VAULT / path)
        except (zipfile.BadZipFile, KeyError, FileNotFoundError):
            return False  # mid-write
    return wait(attempt, f'saved {what}', timeout=20)


# ---------------------------------------------------------------- tests

def test_visual_baselines():
    """Each untouched fixture's page / slide / grid looks as approved (runs first:
    later tests edit the same copies)."""
    from PIL import Image
    BASELINES.mkdir(parents=True, exist_ok=True)
    problems = []
    for name, view, region in (('sample.docx', 'doc-works-docx-view', '.ProseMirror'),
                               ('sample.pptx', 'doc-works-pptx-view', '.pptx-stage'),
                               ('libreoffice-budget.xlsx', 'doc-works-xlsx-view', '.xlsx-grid')):
        open_file(f'{RUN}/{name}', view)
        if view == 'doc-works-docx-view':  # the zoom is remembered per profile; baselines are at Fit
            ev('(app.workspace.getMostRecentLeaf().view.setZoom("fit"), 1)')
        desk('move', '1', '1199')  # no hover effects
        desk('wait-idle')
        box = ev(f'''(() => {{ const r = document.querySelector(".workspace-leaf.mod-active {region}").getBoundingClientRect();
            return [Math.max(0, r.left), Math.max(0, r.top), Math.min(innerWidth, r.right), Math.min(innerHeight, r.bottom)]
                .map(Math.round); }})()''')
        desk('shot', str(OUT / 'full.png'))
        current = OUT / f'visual-{name}.png'
        Image.open(OUT / 'full.png').crop(tuple(box)).save(current)
        baseline = BASELINES / f'{name}.png'
        if os.environ.get('UPDATE_BASELINES') or not baseline.exists():
            shutil.copy(current, baseline)
            print(f'  baseline written: {baseline}')
            continue
        result = subprocess.run(['agent-desk', 'compare', str(baseline), str(current),
                                 '--diff', str(OUT / f'visual-{name}-diff.png')], capture_output=True, text=True)
        if result.returncode:
            problems.append(f'{name}: {result.stdout.strip()} (diff in {OUT})')
    assert not problems, '; '.join(problems)


def test_docx_page_break_mid_paragraph():
    """A page break inside a paragraph starts the next page with the text after it (Word / WordCraft /
    LibreOffice): "Before break" ends page 1, "after break" starts page 2."""
    open_file(f'{RUN}/sample.docx', 'doc-works-docx-view')
    pages = ev('''(() => { const dom = app.workspace.getMostRecentLeaf().view.editor.dom;
        const gaps = [...dom.querySelectorAll(".docx-page-space")].map(e => e.getBoundingClientRect().bottom);
        const page = (text) => { const w = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
            for (let n = w.nextNode(); n; n = w.nextNode()) if (n.textContent.includes(text)) {
                const r = document.createRange(); r.setStart(n, n.textContent.indexOf(text)); r.setEnd(n, n.textContent.indexOf(text) + 1);
                const top = r.getBoundingClientRect().top; return 1 + gaps.filter(g => g <= top).length; }
            return null; };
        return [page("Before break"), page("after break")]; })()''')
    assert pages == [1, 2], f'"Before break" / "after break" on pages {pages}, want [1, 2]'
    shot('docx-page-break')


def test_docx_type_bold_undo_redo():
    """Unicode typing, Ctrl+B, Ctrl+Z, Ctrl+Shift+Z (Shift chords)."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    shot('docx-before')
    # End of the "Plain bold and italic … after tab" paragraph.
    x, y, right, _ = ev('''(() => { const p = [...document.querySelectorAll(".workspace-leaf.mod-active .ProseMirror p")]
        .find(p => p.textContent.includes("after tab")); const r = p.getBoundingClientRect();
        return [r.x, r.y + r.height/2, r.right - 4, r.y]; })()''')
    click(right, y)
    desk('key', 'End')
    desk('type', '--', ' guisuite café λ')
    for _ in range(3):  # back over "λ", "café", "guisuite"
        desk('key', 'Ctrl+Shift+Left')
    assert ev('String(getSelection())').strip() == 'guisuite café λ', 'Ctrl+Shift+Left selection'
    desk('key', 'Ctrl+b')
    desk('key', 'Ctrl+z')
    desk('key', 'Ctrl+Shift+z')

    def runs(file):
        body = ET.fromstring(zipfile.ZipFile(file).read('word/document.xml'))
        out = []
        for r in body.iter(f'{W}r'):
            text = ''.join(t.text or '' for t in r.iter(f'{W}t'))
            b = r.find(f'{W}rPr/{W}b')
            bold = b is not None and b.get(f'{W}val', 'true') not in ('0', 'false')
            out.append((text, bold))
        return out
    got = saved(path, lambda f: (lambda r: r if any(b and 'guisuite café λ' in t for t, b in r) else None)(runs(f)),
                'bold "guisuite café λ"')
    text = ''.join(t for t, _ in got)
    assert 'after tab guisuite café λ' in text, text
    assert 'z' not in text.split('guisuite café λ')[1][:1], 'stray z from Ctrl+Shift+Z'
    shot('docx-after')


def test_docx_markdown_paste():
    """Markdown from the (private) clipboard: Ctrl+V converts, Ctrl+Shift+V pastes the raw text as-is."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    x, y, right, _ = ev('''(() => { const p = [...document.querySelectorAll(".workspace-leaf.mod-active .ProseMirror p")]
        .find(p => p.textContent.includes("See example page")); const r = p.getBoundingClientRect();
        return [r.x, r.y + r.height/2, r.right - 4, r.y]; })()''')
    click(right, y)
    desk('key', 'End')
    desk('key', 'Return')
    desk('clip', 'set', '--', '## Pasted heading\n\n- item one\n- item two\n')
    desk('key', 'Ctrl+v')
    desk('key', 'Return')
    desk('clip', 'set', '--', '**plain** pasted')
    desk('key', 'Ctrl+Shift+v')

    def paras(file):
        body = ET.fromstring(zipfile.ZipFile(file).read('word/document.xml'))
        out = []
        for p in body.iter(f'{W}p'):
            style = p.find(f'{W}pPr/{W}pStyle')
            bold = any(r.find(f'{W}rPr/{W}b') is not None for r in p.iter(f'{W}r'))
            out.append((''.join(t.text or '' for t in p.iter(f'{W}t')), style.get(f'{W}val') if style is not None else None,
                        p.find(f'{W}pPr/{W}numPr') is not None, bold))
        return out
    got = saved(path, lambda f: (lambda ps: ps if any(t == '**plain** pasted' for t, *_ in ps) else None)(paras(f)),
                'pasted paragraphs')
    by_text = {t: rest for t, *rest in got}
    assert (by_text.get('Pasted heading') or [''])[0].startswith('Heading'), by_text.get('Pasted heading')
    assert by_text.get('item one', [None, False])[1] and by_text.get('item two', [None, False])[1], 'list items'
    assert not by_text['**plain** pasted'][2], 'Ctrl+Shift+V made bold'
    assert not any('**' in t or t.startswith('## ') for t in by_text if t != '**plain** pasted'), 'Markdown syntax left in the text'
    shot('docx-paste-after')


def test_xlsx_name_box_formula():
    """Name box navigation, a typed formula and Unicode text."""
    path = f'{RUN}/libreoffice-budget.xlsx'
    open_file(path, 'doc-works-xlsx-view')
    shot('xlsx-before')
    for cell, value in (('H30', '=6*7'), ('H31', 'café λ')):
        x, y, *_ = center('.xlsx-name-box')
        click(x, y)
        desk('key', 'Ctrl+a')
        desk('type', '--', cell)
        desk('key', 'Return')
        desk('type', '--', value)
        desk('key', 'Return')

    def cells(file):
        z = zipfile.ZipFile(file)
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        book = ET.fromstring(z.read('xl/workbook.xml'))
        rid = book.find(f'{X}sheets/{X}sheet').get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = next(r.get('Target') for r in rels if r.get('Id') == rid)
        sheet = ET.fromstring(z.read('xl/' + target.removeprefix('/xl/')))
        strings = [''.join(t.text or '' for t in si.iter(f'{X}t'))
                   for si in ET.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else []
        out = {}
        for c in sheet.iter(f'{X}c'):
            f, v = c.find(f'{X}f'), c.find(f'{X}v')
            text = ''.join(t.text or '' for t in c.iter(f'{X}t'))
            if c.get('t') == 's' and v is not None:
                text = strings[int(v.text)]
            out[c.get('r')] = (f.text if f is not None else None, v.text if v is not None else None, text)
        return out
    got = saved(path, lambda f: (lambda c: c if 'H31' in c and 'H30' in c else None)(cells(f)), 'H30 and H31')
    assert got['H30'][:2] == ('6*7', '42'), got['H30']
    assert got['H31'][2] == 'café λ', got['H31']
    shot('xlsx-after')


def test_pptx_title_edit():
    """Click into a slide's title and type at its end."""
    path = f'{RUN}/sample.pptx'
    open_file(path, 'doc-works-pptx-view')
    shot('pptx-before')
    x, y, *_ = center('.pptx-text', pick='largest')
    click(x, y)
    wait(lambda: ev('document.activeElement?.matches(".pptx-text[contenteditable=true]")'), 'title in text edit')
    desk('key', 'Ctrl+End')
    desk('type', '--', ' GUI ✓')
    desk('key', 'Escape')

    def texts(file):
        slide = ET.fromstring(zipfile.ZipFile(file).read('ppt/slides/slide1.xml'))
        return [''.join(t.text or '' for t in p.iter(f'{A}t')) for p in slide.iter(f'{A}p')]
    got = saved(path, lambda f: (lambda t: t if 'Hello Deck GUI ✓' in t else None)(texts(f)), 'title "Hello Deck GUI ✓"')
    assert 'Subtitle here' in got, got
    shot('pptx-after')


def at(js, what=None):
    """Screen centre of the element a JS expression finds (anywhere in the window: menus, dialogs)."""
    return wait(lambda: ev(f"""(() => {{ const el = {js}; if (!el) return null; const r = el.getBoundingClientRect();
        return r.width && r.height ? [r.x + r.width/2, r.y + r.height/2] : null; }})()"""), what or f'element {js[:70]}')


def click_el(js, what=None):
    click(*at(js, what))
    desk('wait-idle')


def by_text(selector, text):
    return f'[...document.querySelectorAll({json.dumps(selector)})].find(e => e.textContent.trim() === {json.dumps(text)})'


def indexed(path):
    wait(lambda: ev(f'!!app.vault.getAbstractFileByPath({json.dumps(path)})'), f'{path} indexed')


def select_range(ref):
    x, y, *_ = center('.xlsx-name-box')
    click(x, y)
    desk('key', 'Ctrl+a')
    desk('type', '--', ref)
    desk('key', 'Return')
    desk('wait-idle')


def sheet_table(path, rows, first_row):
    """Setup: a small table typed into the open workbook from column K (the fixtures leave it empty)."""
    ev(f"""(() => {{ const v = app.workspace.getMostRecentLeaf().view; if (v.file.path !== {json.dumps(path)}) throw new Error("wrong file");
        return v.change(() => {json.dumps(rows)}.forEach((r, i) => r.forEach((t, j) => v.edits.input(v.sheet, {first_row} + i, 10 + j, t)))); }})()""")
    desk('wait-idle')


def test_template_opens_and_stays_a_template():
    """A .dotx opens in the document editor from the file explorer; an edit is saved and it is still a template."""
    path = f'{RUN}/gui-template.dotx'
    doc, tpl = 'wordprocessingml.document.main+xml', 'wordprocessingml.template.main+xml'
    with zipfile.ZipFile(REPO / 'tests/fixtures/sample.docx') as src, zipfile.ZipFile(VAULT / path, 'w', zipfile.ZIP_DEFLATED) as out:
        for info in src.infolist():
            data = src.read(info)
            if info.filename == '[Content_Types].xml':
                assert doc.encode() in data
                data = data.replace(doc.encode(), tpl.encode())
            out.writestr(info, data)
    indexed(path)
    open_file(path, 'doc-works-docx-view')
    x, y, right, _ = ev("""(() => { const p = [...document.querySelectorAll(".workspace-leaf.mod-active .ProseMirror p")]
        .find(p => p.textContent.includes("after tab")); const r = p.getBoundingClientRect();
        return [r.x, r.y + r.height/2, r.right - 4, r.y]; })()""")
    click(right, y)
    desk('key', 'End')
    desk('type', '--', ' template-edit')
    saved(path, lambda f: b'template-edit' in zipfile.ZipFile(f).read('word/document.xml'), 'the typed text')
    types = read_zip(path, '[Content_Types].xml')
    assert tpl.encode() in types and doc.encode() not in types, 'the saved file is no longer a template'
    shot('template-after')


def test_docx_markdown_paste_picture_and_math():
    """Pasted Markdown with a vault picture and math: the picture goes into the package, the math is an equation."""
    from PIL import Image
    path, pic = f'{RUN}/sample.docx', f'{RUN}/gui-ridge.png'
    Image.new('RGB', (120, 60), (40, 120, 200)).save(VAULT / pic)
    indexed(pic)
    open_file(path, 'doc-works-docx-view')
    before = sum(1 for n in zipfile.ZipFile(VAULT / path).namelist() if n.startswith('word/media/') and not n.endswith('/'))
    x, y, right, _ = ev("""(() => { const p = [...document.querySelectorAll(".workspace-leaf.mod-active .ProseMirror p")]
        .find(p => p.textContent.includes("See example page")); const r = p.getBoundingClientRect();
        return [r.x, r.y + r.height/2, r.right - 4, r.y]; })()""")
    click(right, y)
    desk('key', 'End')
    desk('key', 'Return')
    desk('clip', 'set', '--', 'Ridge ![[gui-ridge.png]] and $x^2 + 1$ end\n\n$$\n\\frac{a}{b}\n$$\n')
    desk('key', 'Ctrl+v')

    def check(file):
        z = zipfile.ZipFile(file)
        body = z.read('word/document.xml').decode()
        media = sum(1 for n in z.namelist() if n.startswith('word/media/') and not n.endswith('/'))  # the zip may list the folder too
        return (body, media) if 'Ridge ' in body and media > before else None
    body, media = saved(path, check, 'the pasted picture')
    assert media == before + 1, f'{media - before} pictures added, want 1'
    assert body.count('<m:oMath>') + body.count('<m:oMath ') >= 2, 'no equations in the saved document'
    assert '<m:oMathPara' in body, 'the display equation is not on its own line'
    text = ''.join(t.text or '' for t in ET.fromstring(body.encode()).iter(f'{W}t'))
    assert '$' not in text.split('Ridge ', 1)[1][:200], 'math left as $ text'
    assert 'gui-ridge.png' not in text, 'the picture stayed as its name'
    drawn = ev("""[...document.querySelectorAll(".workspace-leaf.mod-active .ProseMirror math")].length""")
    assert drawn >= 2, f'{drawn} equations drawn'
    shot('docx-paste-picture-math')


def test_xlsx_protected_range():
    """Data › Protected sheets and ranges: protect the selection; typing there shows "Heads up!", OK saves the edit."""
    path = f'{RUN}/libreoffice-budget.xlsx'
    open_file(path, 'doc-works-xlsx-view')
    select_range('H40:H41')
    menu_path('Data', 'Protected sheets and ranges…')
    type_into('document.querySelector(\'.modal input[aria-label="New description"]\')', 'Locked totals')
    click_el(by_text('.modal button', 'Add'))
    desk('key', 'Escape')
    select_range('H40')
    desk('type', '--', '42')
    desk('key', 'Return')
    wait(lambda: ev('[...document.querySelectorAll(".modal-title")].some(t => t.textContent === "Heads up!") ? 1 : 0'), 'warning')
    shot('xlsx-protected-warning')
    click_el(by_text('.modal button', 'OK'))

    def protected(file):
        z = zipfile.ZipFile(file)
        xml = z.read('xl/worksheets/sheet1.xml').decode()
        return 'name="Locked totals"' in xml and 'sqref="H40:H41"' in xml and re.search(r'<c r="H40"[^>]*>(<f>[^<]*</f>)?<v>42</v>', xml)
    saved(path, protected, 'H40 = 42 and the protected range')


def test_xlsx_named_function():
    """Data › Named functions: create SQUARE(x) in the form, use it in a cell; saved as an Excel LAMBDA name."""
    path = f'{RUN}/libreoffice-budget.xlsx'
    open_file(path, 'doc-works-xlsx-view')
    menu_path('Data', 'Named functions…')
    click_el(by_text('.modal button', 'Add new function'))
    type_into(setting_input('Function name'), 'SQUARE')
    type_into(setting_input('Argument placeholders'), 'x')
    type_into(setting_input('Formula definition'), '=x*x')
    click_el(by_text('.modal button', 'Create'))
    desk('key', 'Escape')
    select_range('H42')
    desk('type', '--', '=SQUARE(7)')
    desk('key', 'Return')

    def used(file):
        z = zipfile.ZipFile(file)
        book = z.read('xl/workbook.xml').decode()
        sheet = z.read('xl/worksheets/sheet1.xml').decode()
        return ('<definedName name="SQUARE">_xlfn.LAMBDA(_xlpm.x,_xlpm.x*_xlpm.x)</definedName>' in book
                and re.search(r'<c r="H42"[^>]*><f>SQUARE\(7\)</f><v>49</v>', sheet))
    saved(path, used, 'SQUARE defined and H42 = 49')


def test_xlsx_slicer():
    """Data › Add a slicer on a plain range: Make table, then a click on "West" filters the table; saved as Excel's slicer."""
    path = f'{RUN}/libreoffice-budget.xlsx'
    open_file(path, 'doc-works-xlsx-view')
    sheet_table(path, [['Region', 'Sales'], ['East', '1'], ['West', '2'], ['East', '3']], 9)
    select_range('K11')
    menu_path('Data', 'Add a slicer')
    click_el(by_text('.modal button', 'Make table'))
    click_el(by_text('.workspace-leaf.mod-active .xlsx-slicer-item', 'West'), 'slicer item West')
    wait(lambda: ev('[...document.querySelectorAll(".workspace-leaf.mod-active .xlsx-slicer-item.is-selected")].map(b => b.textContent).join()') == 'West', 'West chosen')
    shot('xlsx-slicer')

    def slicer(file):
        z = zipfile.ZipFile(file)
        names = z.namelist()
        tables = [n for n in names if n.startswith('xl/tables/') and n.endswith('.xml')]
        if not tables or not any(n.startswith('xl/slicerCaches/') for n in names) or not any(n.startswith('xl/slicers/') for n in names):
            return None
        table = z.read(tables[0]).decode()
        sheet = z.read('xl/worksheets/sheet1.xml').decode()
        return ('<filter val="West"/>' in table and 'ref="K10:L13"' in table
                and re.search(r'<row r="11"[^>]*hidden="1"', sheet) and re.search(r'<row r="13"[^>]*hidden="1"', sheet))
    saved(path, slicer, 'table filtered to West with a slicer')


def test_xlsx_pivot_filter():
    """Insert a pivot table from a range, add a filter in its editor: a report filter above the table, saved."""
    path = f'{RUN}/gui-pivot.xlsx'
    shutil.copy2(REPO / 'tests/fixtures/libreoffice-budget.xlsx', VAULT / path)
    indexed(path)
    open_file(path, 'doc-works-xlsx-view')
    sheet_table(path, [['Region', 'Product', 'Amount'], ['East', 'A', '10'], ['East', 'B', '20'], ['West', 'A', '5'], ['West', 'B', '15']], 40)
    select_range('K41:M45')
    click_el(by_text('.workspace-leaf.mod-active .dw-menu-btn', 'Insert'))
    click_el(by_text('.menu .menu-item', 'Pivot table'))
    wait(lambda: ev('document.querySelector(".modal .modal-title")?.textContent') == 'Pivot table editor', 'pivot editor')
    shot('pivot-editor')
    click_el('document.querySelector(".modal .xlsx-pivot-add-filter")')
    # The new filter starts on Region (already the rows): pick Product with the keyboard.
    field = 'document.querySelector(".modal select[aria-label=\'Filter field\']")'
    ev(f'({field}.focus(), 1)')
    desk('key', 'Down')
    wait(lambda: ev(f'{field}.selectedOptions[0].textContent') == 'Product', 'filter field Product')
    click_el('document.querySelector(".modal .xlsx-pivot-filter-status")')
    click_el(by_text('.modal .xlsx-pivot-filter-items label', 'B') + '.querySelector("input")')
    assert ev('document.querySelector(".modal .xlsx-pivot-filter-status").textContent') == 'Showing 1 of 2 items'
    shot('pivot-filter')
    click_el(by_text('.modal button.mod-cta', 'Done'))

    def check(file):
        z = zipfile.ZipFile(file)
        table = next((n for n in z.namelist() if n.startswith('xl/pivotTables/pivotTable')), None)
        return z.read(table).decode() if table and b'pageField ' in z.read(table) else None
    definition = saved(path, check, 'the pivot table with its filter')
    assert 'axis="axisPage"' in definition and 'rowPageCount="1"' in definition, definition[:400]
    shown = ev("""(() => { const v = app.workspace.getMostRecentLeaf().view, s = v.sheet;
        return [s.name, ...[0, 1, 2, 3, 4, 5].map(r => [s.at(r, 0)?.value ?? null, s.at(r, 1)?.value ?? null])]; })()""")
    assert shown == ['Pivot Table 1', ['Product', 'A'], [None, None], ['Row Labels', 'Sum of Amount'], ['East', 10], ['West', 5], ['Grand Total', 15]], shown
    shot('pivot-after')


def test_xlsx_chart_in_note_embed():
    """A sheet's chart is drawn in a note's embed of the workbook."""
    path, note = f'{RUN}/gui-chart.xlsx', f'{RUN}/gui-chart-embed.md'
    shutil.copy2(REPO / 'tests/fixtures/libreoffice-budget.xlsx', VAULT / path)
    indexed(path)
    open_file(path, 'doc-works-xlsx-view')
    sheet_table(path, [['Region', 'Sales'], ['North', '10'], ['South', '20'], ['East', '15']], 0)
    select_range('K1:L4')
    click_el('document.querySelector(".workspace-leaf.mod-active button[aria-label=\'Insert chart\']")')
    wait(lambda: ev('document.querySelector(".modal .modal-title")?.textContent') == 'Chart editor', 'chart editor')
    desk('key', 'Escape')
    saved(path, lambda f: any(n.startswith('xl/charts/chart') for n in zipfile.ZipFile(f).namelist()), 'the chart')
    (VAULT / note).write_text('Chart below.\n\n![[gui-chart.xlsx]]\n')
    indexed(note)
    ev(f"""(() => {{ const f = app.vault.getAbstractFileByPath({json.dumps(note)});
        const leaf = app.workspace.getLeavesOfType("file-explorer")[0];
        app.workspace.revealLeaf(leaf); leaf.view.revealInFolder(f); return 1; }})()""")
    desk('wait-idle')
    click_el(f'document.querySelector(\'.nav-file-title[data-path={json.dumps(note)}]\')')
    wait(lambda: ev('app.workspace.getMostRecentLeaf()?.view.file?.path') == note, 'the note open')
    found = wait(lambda: ev("""(() => { const c = document.querySelector(".workspace-leaf.mod-active .dw-embed .xlsx-chart");
        if (!c || !c.querySelector("svg")) return null; const r = c.getBoundingClientRect(), g = c.closest(".dw-embed-grid").getBoundingClientRect();
        return [r.width, r.height, r.right <= g.right + 1 && r.bottom <= g.bottom + 1]; })()"""), 'the chart drawn in the embed', timeout=20)
    assert found[0] > 100 and found[1] > 100 and found[2], f'chart box {found}: too small or outside the embedded grid'
    shot('embed-chart')


def menu_path(menu, *titles):
    """Open a menu in the active editor's menu row and click through its items, as a user."""
    click_el(by_text('.workspace-leaf.mod-active .dw-menu-btn', menu), f'menu {menu}')
    for title in titles:
        click_el(by_text('.menu .menu-item-title', title), f'menu item {title}')


def setting_input(name):
    """The input of the dialog setting called `name`."""
    return (f'[...document.querySelectorAll(".modal .setting-item")].find(s => s.querySelector(".setting-item-name")'
            f'?.textContent === {json.dumps(name)})?.querySelector("input, textarea, select")')


def type_into(js, text):
    click_el(js)
    desk('key', 'Ctrl+a')
    desk('type', '--', text)


def body_of(file):
    return ET.fromstring(zipfile.ZipFile(file).read('word/document.xml'))


def caret_after(text):
    """Click at the end of the paragraph holding `text` and move to its end."""
    x, y, right, _ = ev(f'''(() => {{ const p = [...document.querySelectorAll(".workspace-leaf.mod-active .ProseMirror p")]
        .find(p => p.textContent.includes({json.dumps(text)})); const r = p.getBoundingClientRect();
        return [r.x, r.y + r.height/2, r.right - 4, r.y]; }})()''')
    click(right, y)
    desk('key', 'End')


def test_menu_search_alt_slash():
    """Alt+/ focuses Google's "Menus" box; typing and Enter run the item (Insert › Break › Page break)."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    count = lambda f: sum(1 for br in body_of(f).iter(f'{W}br') if br.get(f'{W}type') == 'page')
    before = count(VAULT / path)
    caret_after('See example page')
    desk('key', 'Alt+slash')
    wait(lambda: ev('document.activeElement?.closest(".dw-menu-search") ? 1 : 0'), 'menu search focused')
    desk('type', '--', 'page break')
    wait(lambda: ev('document.querySelector(".dw-menu-search-results .is-selected .dw-menu-search-title")?.textContent') == 'Page break',
         'Page break first')
    shot('menu-search')
    desk('key', 'Return')
    saved(path, lambda f: count(f) == before + 1, 'one more page break')
    assert ev('document.activeElement?.closest(".ProseMirror") ? 1 : 0'), 'focus back in the document'


def test_docx_smart_chip_from_at():
    """Typing @ at a word start opens the chip menu; Today becomes a Word date picker, the @ gone."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    caret_after('after tab')
    desk('type', '--', ' due @')
    click_el(by_text('.menu .menu-item-title', 'Today'), 'chip menu')
    today = time.strftime('%Y-%m-%d')

    def chip(file):
        for sdt in body_of(file).iter(f'{W}sdt'):
            date = sdt.find(f'{W}sdtPr/{W}date')
            if date is not None and date.get(f'{W}fullDate', '').startswith(today):
                return sdt
        return None
    saved(path, chip, "today's date chip")
    text = ''.join(t.text or '' for p in body_of(VAULT / path).iter(f'{W}p') for t in p.iter(f'{W}t') if 'after tab' in ''.join(x.text or '' for x in p.iter(f'{W}t')))
    assert '@' not in text, text
    shot('docx-chip')


def test_docx_building_block():
    """@ on an empty line, Review tracker: a table below with a status dropdown; a click on it picks "Approved"."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    caret_after('after tab')
    desk('key', 'Return')
    desk('type', '--', '@')
    click_el(by_text('.menu .menu-item-title', 'Review tracker'), 'building block in the @ menu')
    desk('type', '--', 'Ada')
    click_el('document.querySelector(".workspace-leaf.mod-active .ProseMirror td .docx-chip-dropdown")', 'status chip')
    click_el(by_text('.menu .menu-item-title', 'Approved'), 'status choice')

    def tracker(file):
        tbl = next((t for t in body_of(file).iter(f'{W}tbl')
                    if [''.join(x.text or '' for x in c.iter(f'{W}t')) for c in t.find(f'{W}tr').iter(f'{W}tc')] == ['Reviewer', 'Status', 'Notes']), None)
        if tbl is None:
            return None
        row = [''.join(x.text or '' for x in c.iter(f'{W}t')) for c in tbl.findall(f'{W}tr')[1].iter(f'{W}tc')]
        return row if row[:2] == ['Ada', 'Approved'] else None
    saved(path, tracker, 'tracker row "Ada | Approved"')
    text = [''.join(t.text or '' for t in p.iter(f'{W}t')) for p in body_of(VAULT / path).find(f'{W}body').findall(f'{W}p')]
    assert '@' not in ''.join(text), text
    shot('docx-building-block')


def test_docx_citation():
    """Tools › Citations: add a source with the keyboard, cite it, insert the bibliography."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    caret_after('after tab')
    menu_path('Tools', 'Citations')
    click_el(by_text('.modal button', 'Add citation source'))
    type_into(setting_input('Contributors'), 'Ann Smith')
    type_into(setting_input('Title'), 'The Long Road')
    type_into(setting_input('Year'), '2020')
    click_el(by_text('.modal button', 'Add source'))
    wait(lambda: ev('[...document.querySelectorAll(".modal .setting-item-name")].some(e => e.textContent === "The Long Road")'), 'source listed')
    shot('citations-dialog')
    click_el(by_text('.modal button', 'Cite'))
    menu_path('Tools', 'Citations')
    click_el(by_text('.modal button', 'Insert bibliography'))

    def check(file):
        z = zipfile.ZipFile(file)
        xml = z.read('word/document.xml').decode()
        sources = [n for n in z.namelist() if n.startswith('customXml/item') and 'Props' not in n]
        return ('CITATION Smi20' in xml and '(Smith, 2020)' in xml and 'BIBLIOGRAPHY' in xml and 'References' in xml
                and any(b'The Long Road' in z.read(n) for n in sources))
    saved(path, check, 'citation, bibliography and source part')
    shot('docx-citation')


def test_docx_compare_documents():
    """Tools › Compare documents makes "Comparison of sample and other.docx" with tracked changes and opens it."""
    other = VAULT / RUN / 'other.docx'
    with zipfile.ZipFile(VAULT / RUN / 'sample.docx') as src, zipfile.ZipFile(other, 'w', zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == 'word/document.xml':
                data = data.replace(b'Before break', b'Before the big break')
            dst.writestr(item, data)
    indexed(f'{RUN}/other.docx')
    open_file(f'{RUN}/sample.docx', 'doc-works-docx-view')
    menu_path('Tools', 'Compare documents')
    click_el(by_text('.modal button', 'Choose…'))
    wait(lambda: ev('document.activeElement?.matches(".prompt-input") ? 1 : 0'), 'file picker')
    desk('type', '--', 'gui other.docx')
    desk('wait-idle')
    desk('key', 'Return')
    click_el(by_text('.modal button', 'Compare'))
    result = f'{RUN}/Comparison of sample and other.docx'
    wait(lambda: ev('app.workspace.getMostRecentLeaf()?.view.file?.path') == result, 'comparison opened', timeout=30)

    def changes(file):
        body = body_of(file)
        ins = [''.join(t.text or '' for t in e.iter(f'{W}t')) for e in body.iter(f'{W}ins')]
        return 'the big ' in ''.join(ins) or 'the big' in ''.join(ins)
    saved(result, changes, 'inserted "the big"')
    shot('docx-compare')


def test_docx_line_numbers():
    """Tools › Line numbers: numbers drawn in the left margin beside each line, restarting on each page; saved as w:lnNumType."""
    path = f'{RUN}/sample.docx'
    open_file(path, 'doc-works-docx-view')
    menu_path('Tools', 'Line numbers')
    click_el('[...document.querySelectorAll(".modal .setting-item")].find(s => s.querySelector(".setting-item-name")'
             '?.textContent === "Show line numbers")?.querySelector(".checkbox-container")', 'Show line numbers toggle')
    click_el(by_text('.modal button', 'Apply'))

    def numbering(file):
        sect = [e for e in body_of(file).find(f'{W}body') if e.tag == f'{W}sectPr'][-1]
        ln = sect.find(f'{W}lnNumType')
        return ln is not None and ln.get(f'{W}restart') == 'newPage'
    saved(path, numbering, 'w:lnNumType restart="newPage"')
    got = wait(lambda: ev('''(() => { const pm = document.querySelector(".workspace-leaf.mod-active .ProseMirror");
        const nums = [...pm.querySelectorAll(".docx-line-number")].map(n => { const r = n.getBoundingClientRect(); return [n.textContent, r.right, r.top + r.height/2]; });
        const paras = [...pm.querySelectorAll("p, h1, h2, h3, h4, h5, h6")].filter(p => !p.closest(".docx-page-gap, .docx-page-first, .docx-page-last, td")).map(p => { const r = p.getBoundingClientRect(); return [r.left, r.top, r.bottom]; });
        const pages = pm.querySelectorAll(".docx-page-space").length + 1;
        return nums.length ? { nums, paras, pages } : null; })()'''), 'line numbers drawn')
    nums, paras, pages = got['nums'], got['paras'], got['pages']
    assert nums[0][0] == '1', nums[:3]
    assert sum(1 for n in nums if n[0] == '1') == pages, (pages, [n[0] for n in nums])
    for label, right, middle in nums:
        beside = [p for p in paras if p[1] <= middle <= p[2]]
        assert beside, f'number {label} beside no paragraph'
        assert right <= beside[0][0] - 20, f'number {label} not in the left margin ({right} vs text at {beside[0][0]})'
    shot('docx-line-numbers')


def test_pptx_laser_and_pen():
    """In the slideshow: L shows the laser dot at the mouse, Ctrl+P draws ink with a drag, E erases it, Esc twice ends."""
    open_file(f'{RUN}/sample.pptx', 'doc-works-pptx-view')
    click_el(by_text('.workspace-leaf.mod-active button', '▶ Present'), 'Present button')
    wait(lambda: ev('document.querySelector(".dw-show .dw-show-slide .pptx-slide") ? 1 : 0'), 'slideshow open')
    desk('wait-idle')
    box = ev('(() => { const r = document.querySelector(".dw-show .dw-show-slide").getBoundingClientRect(); return [r.x, r.y, r.width, r.height]; })()')
    x, y, w, h = box
    desk('key', 'l')
    desk('move', str(int(x + w * 0.3)), str(int(y + h * 0.4)))
    desk('move', str(int(x + w * 0.3) + 5), str(int(y + h * 0.4) + 5))
    dot = wait(lambda: ev('''(() => { const d = document.querySelector(".dw-show.is-laser .dw-show-laser"); if (!d) return null;
        const r = d.getBoundingClientRect(); return r.width ? [r.x + r.width/2, r.y + r.height/2] : null; })()'''), 'laser dot')
    assert abs(dot[0] - (x + w * 0.3 + 5)) < 4 and abs(dot[1] - (y + h * 0.4 + 5)) < 4, (dot, box)
    shot('pptx-laser')
    desk('key', 'Ctrl+p')
    wait(lambda: ev('document.querySelector(".dw-show.is-pen") ? 1 : 0'), 'pen on (not the command palette)')
    assert not ev('document.querySelector(".prompt") ? 1 : 0'), 'Ctrl+P opened the command palette'
    desk('drag', str(int(x + w * 0.2)), str(int(y + h * 0.2)), str(int(x + w * 0.7)), str(int(y + h * 0.6)))
    d = wait(lambda: ev('document.querySelector(".dw-show-ink path")?.getAttribute("d") ?? null'), 'ink stroke')
    assert d.count('L') >= 1, d
    assert ev('document.querySelector(".dw-show .dw-show-slide") ? 1 : 0'), 'show closed'
    count = ev('document.querySelectorAll(".dw-show-ink path").length')
    assert count == 1, count
    shot('pptx-pen')
    desk('key', 'e')
    wait(lambda: ev('document.querySelector(".dw-show-ink") ? 0 : 1'), 'ink erased')
    desk('key', 'Escape')
    wait(lambda: ev('document.querySelector(".dw-show.is-pen") ? 0 : 1'), 'pen put down')
    assert ev('document.querySelector(".dw-show") ? 1 : 0'), 'first Escape ended the show'
    desk('key', 'Escape')
    wait(lambda: ev('document.querySelector(".dw-show") ? 0 : 1'), 'show ended')


def test_pptx_equation():
    """Insert › Equation types LaTeX into the dialog; a double-click edits it; both saved as PowerPoint math."""
    path = f'{RUN}/sample.pptx'
    open_file(path, 'doc-works-pptx-view')
    menu_path('Insert', 'Equation')
    wait(lambda: ev('document.activeElement?.matches(".docx-equation-input") ? 1 : 0'), 'equation input')
    desk('type', '--', 'x^2+1')
    click_el(by_text('.modal button', 'Insert'))
    A14 = '{http://schemas.microsoft.com/office/drawing/2010/main}'
    M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'

    def maths(file):
        slide = ET.fromstring(zipfile.ZipFile(file).read('ppt/slides/slide1.xml'))
        return [''.join(t.text or '' for t in m.iter(f'{M}t')) for m in slide.iter(f'{A14}m')]
    saved(path, lambda f: maths(f) == ['x2+1'], 'equation x^2+1')
    wait(lambda: ev('document.querySelector(".workspace-leaf.mod-active .pptx-stage .pptx-math math") ? 1 : 0'), 'math drawn')
    x, y = at('document.querySelector(".workspace-leaf.mod-active .pptx-stage .pptx-math")', 'drawn equation')
    desk('move', str(round(x)), str(round(y)))
    desk('click', 'left')
    desk('click', 'left')
    wait(lambda: ev('document.activeElement?.matches(".docx-equation-input") ? 1 : 0'), 'edit dialog')
    desk('key', 'Ctrl+a')
    desk('type', '--', 'y_1')
    click_el(by_text('.modal button', 'Save'))
    saved(path, lambda f: maths(f) == ['y1'], 'equation y_1')
    shot('pptx-equation')


# ---------------------------------------------------------------- runner

def setup():
    status = json.loads(desk('status', '--json'))
    if not status.get('running') or status.get('paused'):
        raise SystemExit('agent desktop is not running or is paused (user control); run `agent-desk start` / wait for handoff')
    if not (REPO / 'main.js').exists():
        raise SystemExit('build first: npm run build')
    running = any('obsidian' in w['appId'].lower() for w in status['windows'])
    desk('obsidian', '--plugin-dir', str(REPO))
    record = json.loads((STATE / 'plugin-build.json').read_text())
    for name, digest in record['sha256'].items():
        assert hashlib.sha256((REPO / name).read_bytes()).hexdigest() == digest, f'copied {name} differs'
    window = wait(lambda: next((w for w in json.loads(desk('windows', '--json'))['windows']
                                if 'obsidian' in w['appId'].lower()), None), 'Obsidian window', timeout=30)
    desk('focus', str(window['id']))  # also selects its desktop (rule: desktop 2)
    wait(lambda: ev('[app.vault.adapter.basePath, app.vault.configDir]') == [str(VAULT), '.obsidian-agent'],
         'real vault with the agent config folder')
    if running:  # an open app keeps the old bundle until the plugin reloads
        ev('(app.plugins.disablePlugin("doc-works").then(() => app.plugins.enablePlugin("doc-works")), 1)')
    wait(lambda: ev('!!app.plugins.plugins["doc-works"]?._loaded'), 'plugin loaded')
    desk('obsidian-cli', 'dev:debug', 'on')  # capture console.error, not only uncaught errors
    # One tab group, nothing open: panes left split by an earlier run halve the editor's width, which folds
    # toolbar buttons away and changes the baselines' size. (The private Obsidian's workspace is its own.)
    ev('''(() => { const leaves = []; app.workspace.iterateRootLeaves(l => leaves.push(l));
        leaves.forEach(l => l.detach()); return leaves.length; })()''')
    (VAULT / RUN).mkdir(parents=True)
    fixtures = REPO / 'tests/fixtures'
    for name in ('sample.docx', 'sample.pptx', 'libreoffice-budget.xlsx'):
        shutil.copy2(fixtures / name, VAULT / RUN / name)
    wait(lambda: ev(f'!!app.vault.getAbstractFileByPath("{RUN}/sample.pptx")'), 'fixtures indexed')


def cleanup(keep):
    ev(f'''(() => {{ const leaves = [];
        app.workspace.iterateAllLeaves(l => {{ if (l.view.file?.path.startsWith("{RUN}/")) leaves.push(l); }});
        leaves.forEach(l => l.detach()); return leaves.length; }})()''')
    if not keep:
        shutil.rmtree(VAULT / RUN)


def main():
    pattern = sys.argv[1] if len(sys.argv) > 1 else ''
    tests = [(n, f) for n, f in globals().items() if n.startswith('test_') and pattern in n]
    print(f'evidence: {OUT}')
    if not os.environ.get('AGENT_DESK_CLAIM'):
        os.environ['AGENT_DESK_CLAIM'] = desk('claim', 'doc-works gui-test', '120').strip().split('=', 1)[1]
    try:
        run(tests)
    finally:
        desk('unclaim')


def run(tests):
    setup()
    video = OUT / 'run.mp4'
    desk('record', 'start', str(video))
    results = []
    for name, test in tests:
        start = time.monotonic()
        try:
            desk('obsidian-cli', 'dev:errors', 'clear')
            desk('obsidian-cli', 'dev:console', 'clear')
            test()
            # A test also fails on errors thrown or logged while it ran, even
            # when the file saved correctly.
            errors = [out.strip() for out in (desk('obsidian-cli', 'dev:errors'),
                                              desk('obsidian-cli', 'dev:console', 'level=error'))
                      if not out.startswith(('No errors captured', 'No console messages captured'))]
            assert not errors, 'console errors:\n' + '\n'.join(errors)
            results.append(dict(test=name, ok=True))
            print(f'PASS {name} ({time.monotonic() - start:.1f}s)')
        except Exception as error:  # report every test, keep going
            results.append(dict(test=name, ok=False, error=f'{type(error).__name__}: {error}'))
            print(f'FAIL {name}: {type(error).__name__}: {error}')
            try:
                shot(f'{name}-failure')
                desk('release')
                desk('key', 'Escape')
            except RuntimeError:
                pass
    failed = [r for r in results if not r['ok']]
    desk('record', 'stop')
    if not failed:
        video.unlink()  # kept only to explain a failure
    cleanup(keep=bool(failed))
    (OUT / 'results.json').write_text(json.dumps(dict(run=RUN, results=results), indent=1, ensure_ascii=False))
    print(f'{len(results) - len(failed)}/{len(results)} passed' +
          (f'; fixtures kept in {VAULT / RUN}; video {video}' if failed else ''))
    if failed:
        sys.exit(1)


if __name__ == '__main__':
    main()
