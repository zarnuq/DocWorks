export interface TextMatch { start: number; end: number; }

/** Literal matches in original UTF-16 offsets, including when case folding changes a character's length. */
export function textMatches(text: string, query: string, matchCase = false, entire = false): TextMatch[] {
  if (!query) return [];
  const literal = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(entire ? `^(?:${literal})$` : literal, matchCase ? "gu" : "giu");
  return Array.from(text.matchAll(pattern), match => ({ start: match.index, end: match.index + match[0].length }));
}
