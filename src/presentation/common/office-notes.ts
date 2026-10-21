import { App, Notice, parseLinktext, parseYaml, TFile } from "obsidian";

/** Literal Office text in a Markdown note, keeping paragraphs without interpreting markup. */
export const noteText = (text: string): string => text.replace(/[\\`*_[\]<>!#|~$+=-]/g, "\\$&")
  .replace(/^(\s*\d+)([.)])(?=\s)/gm, "$1\\$2").replace(/\r?\n+/g, "\n\n");

function nextPath(app: App, source: TFile, suffix: string): string {
  const base = `${source.path} ${suffix}`;
  let path = `${base}.md`;
  for (let i = 2; app.vault.getAbstractFileByPath(path); i++) path = `${base} ${i}.md`;
  return path;
}

function isCompanion(app: App, note: TFile, source: TFile, properties: Record<string, unknown> | undefined): boolean {
  if (properties?.["doc-works"] !== "companion" || typeof properties.source !== "string") return false;
  const link = properties.source;
  const wiki = /^\[\[([^|]+)(?:\|[^\]]*)?\]\]$/.exec(link);
  const markdown = /\]\((?:<([^>]+)>|([^\s]+))\)$/.exec(link);
  let target = wiki?.[1] ?? markdown?.[1] ?? markdown?.[2];
  if (!target) return false;
  if (!wiki) { try { target = decodeURIComponent(target); } catch { return false; } }
  return app.metadataCache.getFirstLinkpathDest(parseLinktext(target).path, note.path)?.path === source.path;
}

async function companion(app: App, source: TFile): Promise<TFile> {
  for (const note of app.vault.getFiles()) {
    if (note.extension === "md" && isCompanion(app, note, source, app.metadataCache.getFileCache(note)?.frontmatter)) return note;
  }
  // Metadata can lag just after creation. Check our default names directly before making another note.
  const base = `${source.path} notes`;
  for (let i = 1; ; i++) {
    const candidate = app.vault.getAbstractFileByPath(`${base}${i === 1 ? "" : ` ${i}`}.md`);
    if (!candidate) break;
    if (!(candidate instanceof TFile)) continue;
    const raw = await app.vault.cachedRead(candidate);
    const yaml = /^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/.exec(raw)?.[1];
    if (!yaml) continue;
    try {
      if (isCompanion(app, candidate, source, parseYaml(yaml) as Record<string, unknown>)) return candidate;
    } catch { /* An unrelated note with invalid YAML is never replaced. */ }
  }
  const path = nextPath(app, source, "notes");
  const link = app.fileManager.generateMarkdownLink(source, path);
  return app.vault.create(path, `---\ndoc-works: companion\nsource: ${JSON.stringify(link)}\ntags: []\nproject: ""\nstatus: ""\n---\n\n# Notes\n\nSource: ${link}\n\n`);
}

const pending = new WeakMap<App, Map<string, Promise<TFile | null>>>();

/** Create or reopen the Office file's companion, always in a new tab. Never changes an existing note. */
export function openCompanionNote(app: App, source: TFile | null): Promise<TFile | null> {
  if (!source) return Promise.resolve(null);
  let jobs = pending.get(app);
  if (!jobs) pending.set(app, (jobs = new Map<string, Promise<TFile | null>>()));
  const current = jobs.get(source.path);
  if (current) return current;
  const path = source.path;
  const job = (async () => {
    try {
      const note = await companion(app, source);
      await app.workspace.getLeaf("tab").openFile(note);
      return note;
    } catch (error) {
      new Notice(`Couldn't open companion note: ${(error as Error).message}`);
      return null;
    } finally { jobs.delete(path); }
  })();
  jobs.set(path, job);
  return job;
}

/** A snapshot of a selection, with a navigable source link. Office files are never edited. */
export async function createSelectionNote(app: App, source: TFile | null, subpath: string, markdown: string): Promise<TFile | null> {
  if (!source) return null;
  if (!markdown.trim()) { new Notice("Select some text or nonempty cells first."); return null; }
  try {
    const path = nextPath(app, source, "selection");
    const link = app.fileManager.generateMarkdownLink(source, path, subpath);
    const note = await app.vault.create(path, `Source: ${link}\n\n${markdown.trim()}\n`);
    await app.workspace.getLeaf("tab").openFile(note);
    return note;
  } catch (error) {
    new Notice(`Couldn't create selection note: ${(error as Error).message}`);
    return null;
  }
}
