import { App, Component, finishRenderMath, MarkdownRenderer, parseLinktext, resolveSubpath, TFile } from "obsidian";
import { ExportPicture, officeEmbedForExport, OFFICE_EXTENSIONS } from "./office-export";

const stripProperties = (markdown: string) => markdown.replace(/^---\r?\n[\s\S]*?\r?\n---(?:\r?\n|$)/, "");
const decode = (link: string) => { try { return decodeURIComponent(link); } catch { return link; } };

/** Render note embeds from their source, independent of the asynchronous preview's readiness. */
export async function renderNoteForExport(app: App, file: TFile, component: Component): Promise<{ el: HTMLElement; warnings: string[]; pictures: Map<HTMLImageElement, ExportPicture> }> {
  const warnings = new Set<string>();
  // Pictures made for embedded Office files (slides, a document's own pictures), which aren't vault files.
  const pictures = new Map<HTMLImageElement, ExportPicture>();
  let count = 0;
  const resolve = (link: string, source: TFile) => {
    const { path, subpath } = parseLinktext(decode(link));
    return { target: path ? app.metadataCache.getFirstLinkpathDest(path, source.path) : source, subpath };
  };
  const uri = (target: TFile, subpath: string) => `obsidian://open?vault=${encodeURIComponent(app.vault.getName())}&file=${encodeURIComponent(target.path + subpath)}`;
  const replace = (embed: HTMLElement, replacement: HTMLElement) => {
    const parent = embed.parentElement;
    if (parent?.tagName !== "P") { embed.replaceWith(replacement); return; }
    // A block note embed between inline words must keep its own headings, lists and tables.
    const before = parent.cloneNode(false) as HTMLElement;
    while (parent.firstChild && parent.firstChild !== embed) before.append(parent.firstChild);
    embed.remove();
    const after = parent.cloneNode(false) as HTMLElement;
    after.append(...Array.from(parent.childNodes));
    parent.replaceWith(...(before.childNodes.length ? [before] : []), replacement, ...(after.childNodes.length ? [after] : []));
  };
  const render = async (source: TFile, subpath: string, ancestors: Set<string>): Promise<HTMLElement> => {
    let markdown = await app.vault.cachedRead(source);
    if (subpath) {
      const cache = app.metadataCache.getFileCache(source);
      const range = cache && resolveSubpath(cache, subpath);
      if (!range) throw new Error("Embedded section could not be found");
      markdown = markdown.slice(range.start.offset, range.end?.offset ?? markdown.length);
      if (subpath.startsWith("#^")) markdown = markdown.replace(new RegExp(`\\s+\\^${subpath.slice(2).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*$`), "");
    }
    const rendered = createDiv();
    await MarkdownRenderer.render(app, stripProperties(markdown), rendered, source.path, component);
    await finishRenderMath();
    // Native embeds may continue changing their preview; export owns this snapshot and replaces note embeds itself.
    const el = rendered.cloneNode(true) as HTMLElement;
    if (source === file && !subpath) {
      const properties = /^---\r?\n[\s\S]*?\r?\n---(?:\r?\n|$)/.exec(markdown)?.[0];
      if (properties) el.dataset.exportFrontmatter = properties;
    }
    for (const img of Array.from(el.querySelectorAll("img"))) img.dataset.exportSource = source.path;
    for (const link of Array.from(el.querySelectorAll<HTMLAnchorElement>("a.internal-link"))) {
      if (link.closest(".internal-embed")) continue;
      const { target, subpath: anchor } = resolve(link.getAttribute("data-href") ?? link.getAttribute("href") ?? "", source);
      if (target) link.setAttribute("href", uri(target, anchor));
    }
    // Only handle top-level embeds: a native preview may already contain its own descendants.
    const embeds = Array.from(el.querySelectorAll<HTMLElement>(".internal-embed")).filter(e => !e.parentElement?.closest(".internal-embed"));
    for (const embed of embeds) {
      const link = embed.getAttribute("src") ?? "";
      const { target, subpath: anchor } = resolve(link, source);
      if (!target && (embed.classList.contains("markdown-embed") || /\.md(?:#|$)/i.test(link))) {
        embed.replaceWith(document.createTextNode(embed.getAttribute("alt") || link));
        warnings.add("Some note embeds were unavailable; their names were kept in the document.");
        continue;
      }
      if (target && OFFICE_EXTENSIONS.has(target.extension.toLowerCase())) {
        const content = await officeEmbedForExport(app, target, anchor, component, pictures).catch(() => null);
        if (content) replace(embed, content);
        else {
          embed.replaceWith(createEl("a", { text: embed.getAttribute("alt") || link, attr: { href: uri(target, anchor) } }));
          warnings.add("Some embedded files couldn't be included and were kept as links (slides need the desktop app).");
        }
        continue;
      }
      if (target?.extension !== "md") continue;
      const key = target.path + anchor;
      const fallback = () => {
        const a = createEl("a", { text: embed.getAttribute("alt") || link, attr: { href: uri(target, anchor) } });
        embed.replaceWith(a);
        warnings.add("Some note embeds were kept as links because they were circular, too deeply nested, or unavailable.");
      };
      if (ancestors.has(key) || ancestors.size >= 16 || count >= 256) { fallback(); continue; }
      count++;
      try {
        const child = await render(target, anchor, new Set([...ancestors, key]));
        const wrapper = createDiv();
        wrapper.append(...Array.from(child.childNodes));
        replace(embed, wrapper);
      } catch { fallback(); }
    }
    return el;
  };
  return { el: await render(file, "", new Set([file.path])), warnings: [...warnings], pictures };
}
