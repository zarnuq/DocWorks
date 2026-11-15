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
    """Markdown from the (private) clipboard: Ctrl+V converts, Ctrl+Shift+V pastes plain."""
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
    got = saved(path, lambda f: (lambda ps: ps if any(t == 'plain pasted' for t, *_ in ps) else None)(paras(f)),
                'pasted paragraphs')
    by_text = {t: rest for t, *rest in got}
    assert (by_text.get('Pasted heading') or [''])[0].startswith('Heading'), by_text.get('Pasted heading')
    assert by_text.get('item one', [None, False])[1] and by_text.get('item two', [None, False])[1], 'list items'
    assert not by_text['plain pasted'][2], 'Ctrl+Shift+V kept bold'
    assert not any('**' in t or t.startswith('## ') for t in by_text), 'Markdown syntax left in the text'
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
