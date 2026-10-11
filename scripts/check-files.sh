#!/usr/bin/env bash
# Check Doc Works against a folder of real .docx / .pptx / .xlsx files (searched recursively):
#   1. unchanged files write back byte-identical, and edited ones reload (tests/real-files.test.ts)
#   2. heavily edited copies are written (tests/export-samples.test.ts)
#   3. if LibreOffice is available (soffice on PATH, or SOFFICE=<path>), originals and edited copies are
#      converted to PDF and any LibreOffice warning the edit introduced is listed per file.
#
#   npm run check:files -- ~/notes
set -euo pipefail
src=$(realpath "${1:?usage: check-files.sh <folder>}")
out=$(mktemp -d)
mkdir -p "$out/files" "$out/edited"

find "$src" -name '.*' -prune -o \( -name '*.docx' -o -name '*.pptx' -o -name '*.xlsx' \) -print0 |
  while IFS= read -r -d '' f; do
    rel=${f#"$src"/}
    cp "$f" "$out/files/${rel//\//_}"
  done
# The export test also edits the fixtures; their originals are the baseline for their warnings.
cp tests/fixtures/sample.docx tests/fixtures/sample.pptx "$out/files/"
echo "$(ls "$out/files" | wc -l) files copied to $out/files"

DOC_WORKS_DOCS="$out/files" DOC_WORKS_DECKS="$out/files" DOC_WORKS_BOOKS="$out/files" DOC_WORKS_EXPORT="$out/edited" \
  npx vitest run tests/real-files.test.ts tests/export-samples.test.ts

soffice=${SOFFICE:-$(command -v soffice || command -v libreoffice || true)}
if [ -z "$soffice" ]; then
  echo "LibreOffice not found (set SOFFICE=<path>): skipped the conversion check"
  exit 0
fi
"$soffice" -env:UserInstallation="file://$out/lo-profile" --headless --convert-to pdf --outdir "$out/pdf-orig" "$out"/files/* >"$out/orig.log" 2>&1 || { cat "$out/orig.log"; exit 1; }
"$soffice" -env:UserInstallation="file://$out/lo-profile" --headless --convert-to pdf --outdir "$out/pdf-edited" "$out"/edited/* >"$out/edited.log" 2>&1 || { cat "$out/edited.log"; exit 1; }
# Some failed launches exit successfully without converting anything. Check every expected PDF.
for kind in orig edited; do
  folder=files
  if [ "$kind" = edited ]; then folder=edited; fi
  for file in "$out/$folder"/*; do
    name=${file##*/}
    if [ ! -s "$out/pdf-$kind/${name%.*}.pdf" ]; then
      cat "$out/$kind.log"
      echo "Missing PDF for $file" >&2
      exit 1
    fi
  done
done

python3 - "$out/edited.log" "$out/orig.log" <<'EOF'
import collections, re, sys

# Noise LibreOffice prints for any file, and warnings it gives for valid OOXML the export test writes
# (checked by hand: links/fields in slides, footnotes, inline page breaks and new tables all render correctly).
NOISE = ("pdfwriter", "local time", "starmath", "invalid cached glyphs",
         "URL field", "refusing to insert control character", "pop out of order",
         "createTextCursorByRange() failed", "PushFootOrEndnote() is called from another")

def parse(path, prefix):
    out, cur = collections.defaultdict(collections.Counter), []
    for line in open(path, errors="replace"):
        m = re.match(r"convert (\S.*?) as a ", line)
        if m:
            name = m.group(1).split("/")[-1].removeprefix(prefix)
            out[name].update(cur)
            cur = []
        elif line.startswith("warn:") and not any(n in line for n in NOISE):
            cur.append(re.sub(r"\d+", "N", re.sub(r":\d+:\d+:", ":", line.strip()))[:160])
        elif "Error" in line:
            print("ERROR", line.strip())
    return out

edited, orig = parse(sys.argv[1], "edited-"), parse(sys.argv[2], "")
for name in sorted(edited):
    new = [(w, c) for w, c in edited[name].items() if w not in orig.get(name, {})]
    if new:
        print("==", name)
        for w, c in sorted(new, key=lambda x: -x[1])[:8]:
            print("  ", c, w)
EOF
echo "PDFs in $out/pdf-orig and $out/pdf-edited"
