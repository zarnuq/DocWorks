"""Generate src/presentation/pptx/font-metrics.ts: per-character advance widths
for common Office fonts, so slides wrap like PowerPoint without the fonts.

Usage: python scripts/make_font_metrics.py <fonts-dir>   (needs fonttools)

Font sources (only widths are extracted, no font data is shipped):
  Carlito, Caladea, Lato, Open Sans, Roboto, Montserrat, Fira Code  github.com/google/fonts
  Arial, Times New Roman, Courier New, Verdana, Georgia, Trebuchet,
  Comic Sans, Arial Black, Impact      sourceforge.net/projects/corefonts
  Aptos                                download.microsoft.com ("Microsoft Aptos Fonts.zip")
  Selawik (Segoe UI metrics)           github.com/microsoft/Selawik
  URW Gothic, P052, C059               github.com/ArtifexSoftware/urw-base35-fonts
  Liberation Sans Narrow               github.com/liberationfonts/liberation-sans-narrow
"""
import base64
import json
import os
import sys

from fontTools.ttLib import TTFont
from fontTools.varLib import instancer

# Printable ASCII, Latin-1, and common typographic punctuation.
CHARS = [chr(c) for c in range(0x20, 0x7F)] + [chr(c) for c in range(0xA0, 0x100)] + list("‘’‚“”„–—…•€™")
UNITS = 200  # widths stored as uint8 in 1/200 em

# family -> files for [regular, bold, italic, bold italic]; "" = missing (falls back at runtime).
# A "file@wght" entry instances a variable font.
FONTS = {
    "calibri": ["Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf"],
    "cambria": ["Caladea-Regular.ttf", "Caladea-Bold.ttf", "Caladea-Italic.ttf", "Caladea-BoldItalic.ttf"],
    "arial": ["Arial.TTF", "Arialbd.TTF", "Ariali.TTF", "Arialbi.TTF"],
    "arial narrow": [f"LiberationSansNarrow-{s}.ttf" for s in ("Regular", "Bold", "Italic", "BoldItalic")],
    "arial black": ["AriBlk.TTF", "", "", ""],
    "times new roman": ["Times.TTF", "Timesbd.TTF", "Timesi.TTF", "Timesbi.TTF"],
    "courier new": ["cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"],
    "verdana": ["Verdana.TTF", "Verdanab.TTF", "Verdanai.TTF", "Verdanaz.TTF"],
    "georgia": ["Georgia.TTF", "Georgiab.TTF", "Georgiai.TTF", "Georgiaz.TTF"],
    "trebuchet ms": ["trebuc.ttf", "Trebucbd.ttf", "trebucit.ttf", "trebucbi.ttf"],
    "comic sans ms": ["Comic.TTF", "Comicbd.TTF", "", ""],
    "impact": ["Impact.TTF", "", "", ""],
    "aptos": ["Aptos.ttf", "Aptos-Bold.ttf", "Aptos-Italic.ttf", "Aptos-Bold-Italic.ttf"],
    "aptos display": ["Aptos-Display.ttf", "Aptos-Display-Bold.ttf", "Aptos-Display-Italic.ttf", "Aptos-Display-Bold-Italic.ttf"],
    "aptos narrow": ["Aptos-Narrow.ttf", "Aptos-Narrow-Bold.ttf", "Aptos-Narrow-Italic.ttf", "Aptos-Narrow-Bold-Italic.ttf"],
    "aptos serif": ["Aptos-Serif.ttf", "Aptos-Serif-Bold.ttf", "Aptos-Serif-Italic.ttf", "Aptos-Serif-Bold-Italic.ttf"],
    "aptos mono": ["Aptos-Mono.ttf", "Aptos-Mono-Bold.ttf", "Aptos-Mono-Italic.ttf", "Aptos-Mono-Bold-Italic.ttf"],
    "segoe ui": ["selawk.ttf", "selawkb.ttf", "", ""],
    "segoe ui light": ["selawkl.ttf", "", "", ""],
    "segoe ui semibold": ["selawksb.ttf", "", "", ""],
    "segoe ui semilight": ["selawksl.ttf", "", "", ""],
    "century gothic": ["URWGothic-Book.ttf", "URWGothic-Demi.ttf", "URWGothic-BookOblique.ttf", "URWGothic-DemiOblique.ttf"],
    "palatino linotype": ["P052-Roman.ttf", "P052-Bold.ttf", "P052-Italic.ttf", "P052-BoldItalic.ttf"],
    "century schoolbook": ["C059-Roman.ttf", "C059-Bold.ttf", "C059-Italic.ttf", "C059-BdIta.ttf"],
    "lato": ["Lato-Regular.ttf", "Lato-Bold.ttf", "Lato-Italic.ttf", "Lato-BoldItalic.ttf"],
    "open sans": ["OpenSans[wdth,wght].ttf@400", "OpenSans[wdth,wght].ttf@700", "OpenSans-Italic[wdth,wght].ttf@400", "OpenSans-Italic[wdth,wght].ttf@700"],
    "roboto": ["Roboto[wdth,wght].ttf@400", "Roboto[wdth,wght].ttf@700", "Roboto-Italic[wdth,wght].ttf@400", "Roboto-Italic[wdth,wght].ttf@700"],
    "montserrat": ["Montserrat[wght].ttf@400", "Montserrat[wght].ttf@700", "Montserrat-Italic[wght].ttf@400", "Montserrat-Italic[wght].ttf@700"],
    "fira code": ["FiraCode[wght].ttf@400", "FiraCode[wght].ttf@700", "", ""],
}

# Other names that share an entry's metrics.
ALIASES = {
    "helvetica": "arial",
    "liberation sans": "arial",
    "arimo": "arial",
    "carlito": "calibri",
    "caladea": "cambria",
    "times": "times new roman",
    "liberation serif": "times new roman",
    "tinos": "times new roman",
    "courier": "courier new",
    "liberation mono": "courier new",
    "cousine": "courier new",
    "book antiqua": "palatino linotype",
    "palatino": "palatino linotype",
    "selawik": "segoe ui",
    "helvetica narrow": "arial narrow",
}


def widths(path: str) -> bytes:
    file, _, wght = path.partition("@")
    font = TTFont(file)
    if wght:
        axes = {"wght": int(wght)}
        if "wdth" in {a.axisTag for a in font["fvar"].axes}:
            axes["wdth"] = 100
        font = instancer.instantiateVariableFont(font, axes)
    cmap, hmtx, upem = font.getBestCmap(), font["hmtx"], font["head"].unitsPerEm
    out = bytearray()
    for ch in CHARS:
        glyph = cmap.get(ord(ch))
        # 0 = unknown glyph: measured with the fallback font at runtime.
        out.append(0 if glyph is None else max(1, min(255, round(hmtx[glyph][0] * UNITS / upem))))
    return bytes(out)


def main(fonts_dir: str) -> None:
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "presentation", "pptx", "font-metrics.ts")
    os.chdir(fonts_dir)
    lines = [
        "// Generated by scripts/make_font_metrics.py - do not edit.",
        f"export const UNITS_PER_EM = {UNITS};",
        f"export const CHARS = {json.dumps(''.join(CHARS), ensure_ascii=False)};",
        "/** family -> base64 uint8 widths for [regular, bold, italic, bold italic]; \"\" = missing style. */",
        "export const FONT_WIDTHS: Record<string, string[]> = {",
    ]
    for family, files in FONTS.items():
        styles = [base64.b64encode(widths(f)).decode() if f else "" for f in files]
        lines.append(f"  {json.dumps(family)}: {json.dumps(styles)},")
    lines.append("};")
    lines.append("export const FONT_ALIASES: Record<string, string> = " + json.dumps(ALIASES) + ";")
    with open(out, "w") as f:
        f.write("\n".join(lines) + "\n")


if __name__ == "__main__":
    main(sys.argv[1])
