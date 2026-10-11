/** Fonts carried inside a deck, made available to the page while the deck is open. */
import { invalidateTextMetrics } from "./text-fit";

export interface EmbeddedFont {
  family: string;
  bold: boolean;
  italic: boolean;
  /** As stored in the deck (Embedded OpenType). */
  data: Uint8Array;
}

/**
 * Decode the deck's fonts to TrueType and register them under the deck's font names, so slide text
 * renders (and wraps) in the real font. Returns how many loaded and a function that removes them again.
 */
export async function loadEmbeddedFonts(fonts: EmbeddedFont[], decode: (data: Uint8Array) => Uint8Array): Promise<{ loaded: number; unload: () => void }> {
  if (typeof FontFace === "undefined" || !fonts.length) return { loaded: 0, unload: () => {} };
  // FontFaceSet is set-like; this project's DOM typings leave out its add / delete.
  const set = document.fonts as unknown as Set<FontFace>;
  const faces: FontFace[] = [];
  for (const f of fonts) {
    try {
      const ttf = decode(f.data);
      const face = new FontFace(f.family, new Uint8Array(ttf), { weight: f.bold ? "700" : "400", style: f.italic ? "italic" : "normal" });
      await face.load();
      set.add(face);
      faces.push(face);
    } catch (err) {
      console.warn(`[doc-works] could not load the embedded font ${f.family}`, err);
    }
  }
  if (faces.length) invalidateTextMetrics();
  return { loaded: faces.length, unload: () => {
    faces.forEach((face) => set.delete(face));
    if (faces.length) invalidateTextMetrics();
  } };
}
