#!/bin/sh
# Fetch the open, metric-compatible substitutes for Office's core fonts into assets/fonts/ (plan section 3.4):
#   Carlito  -> Calibri          (SIL OFL 1.1, from google/fonts)
#   Caladea  -> Cambria          (SIL OFL 1.1, from google/fonts)
#   Liberation Sans / Serif / Mono -> Arial / Times New Roman / Courier New   (SIL OFL 1.1, liberation-fonts 2.1.5)
# Each family keeps its license file next to the fonts. The files are parsed only by HarfBuzz in WebAssembly.
set -eu
ROOT=$(cd "$(dirname "$0")/../.." && pwd)
DEST="$ROOT/assets/fonts"
mkdir -p "$DEST/carlito" "$DEST/caladea" "$DEST/liberation"
GF=https://raw.githubusercontent.com/google/fonts/main/ofl
for style in Regular Bold Italic BoldItalic; do
  curl -fsSL -o "$DEST/carlito/Carlito-$style.ttf" "$GF/carlito/Carlito-$style.ttf"
  curl -fsSL -o "$DEST/caladea/Caladea-$style.ttf" "$GF/caladea/Caladea-$style.ttf"
done
curl -fsSL -o "$DEST/carlito/OFL.txt" "$GF/carlito/OFL.txt"
curl -fsSL -o "$DEST/caladea/OFL.txt" "$GF/caladea/OFL.txt"

TMP=$(mktemp -d)
curl -fsSL -o "$TMP/liberation.tar.gz" https://github.com/liberationfonts/liberation-fonts/files/7261482/liberation-fonts-ttf-2.1.5.tar.gz
tar -xzf "$TMP/liberation.tar.gz" -C "$TMP"
cp "$TMP"/liberation-fonts-ttf-2.1.5/Liberation*.ttf "$DEST/liberation/"
cp "$TMP"/liberation-fonts-ttf-2.1.5/LICENSE "$DEST/liberation/LICENSE"
rm -rf "$TMP"
( cd "$DEST" && find . -name '*.ttf' | sort | xargs shasum -a 256 > SHA256SUMS )
echo "fonts: $(find "$DEST" -name '*.ttf' | wc -l | tr -d ' ') files in $DEST"
