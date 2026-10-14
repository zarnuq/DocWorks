export interface TextMatch {
  start: number;
  end: number;
  /** A regular expression search's match and its groups, for $1 in replacements. */
  groups?: string[];
}

/** Google's find options. */
export interface SearchOptions {
  matchCase?: boolean;
  /** The query is a regular expression. */
  regex?: boolean;
  /** Only matches with no letter, digit or _ on either side. */
  wholeWord?: boolean;
  /** é matches e (Google Docs' "Ignore Latin diacritics"). */
  ignoreAccents?: boolean;
  /** The whole text must match (a sheet's "Match entire cell contents"). */
  entire?: boolean;
}

const WORD = /[\p{L}\p{N}_]/u;

/** Text without combining accents, and for each of its characters the offset it came from (plus the end). */
function fold(text: string): { text: string; at: number[] } {
  let out = "";
  const at: number[] = [];
  let i = 0;
  for (const ch of text) {
    const bare = ch.normalize("NFD").replace(/\p{M}/gu, "");
    out += bare;
    for (let k = 0; k < bare.length; k++) at.push(i);
    i += ch.length;
  }
  at.push(text.length);
  return { text: out, at };
}

/**
 * Matches in original UTF-16 offsets. `options` may be the old matchCase flag. A regular expression that doesn't
 * compile (still being typed) finds nothing.
 */
export function textMatches(text: string, query: string, options: SearchOptions | boolean = false, entire = false): TextMatch[] {
  const o: SearchOptions = typeof options === "boolean" ? { matchCase: options, entire } : options;
  if (!query) return [];
  const folded = o.ignoreAccents ? fold(text) : null;
  const q = o.ignoreAccents ? fold(query).text : query;
  const body = o.regex ? q : q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let pattern: RegExp;
  try {
    pattern = new RegExp(o.entire ? `^(?:${body})$` : body, o.matchCase ? "gu" : "giu");
  } catch {
    return [];
  }
  const searched = folded?.text ?? text;
  const out: TextMatch[] = [];
  for (const m of searched.matchAll(pattern)) {
    const start = m.index, end = start + m[0].length;
    if (end === start) continue;
    if (o.wholeWord && (WORD.test(searched[start - 1] ?? "") || WORD.test(searched[end] ?? ""))) continue;
    const match: TextMatch = folded ? { start: folded.at[start], end: folded.at[end] } : { start, end };
    if (o.regex) match.groups = Array.from(m, (g) => g ?? "");
    out.push(match);
  }
  return out;
}

/** The text a match is replaced with: $1…, $& and $$ filled in for a regular expression search. */
export function expandReplacement(replacement: string, match: TextMatch): string {
  const groups = match.groups;
  if (!groups) return replacement;
  return replacement.replace(/\$(\$|&|\d+)/g, (all, key: string) => (key === "$" ? "$" : key === "&" ? groups[0] : groups[Number(key)] ?? all));
}
