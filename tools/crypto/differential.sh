#!/bin/sh
# tools/crypto/differential.sh
# Cross-check @pptx/crypto against the Python reference implementation, msoffcrypto-tool, in both directions:
#   (1) msoffcrypto encrypts a deck  -> our decrypt must reproduce the original byte for byte
#   (2) we encrypt a deck            -> msoffcrypto's decrypt must reproduce the original byte for byte
#   (3) our fixtures still decrypt against msoffcrypto's expectations
# Runs only when the venv exists (the committed fixtures let the Vitest suite run without Python). Create it with:
#   python3 -m venv "$VENV" && "$VENV/bin/python" -I -m pip install --require-hashes -r tools/crypto/requirements.txt
#
# Env: VENV (default below), CORPUS (default corpus/external). Exit 0 if every check passes, 1 otherwise.
set -eu

ROOT=$(cd "$(dirname "$0")/../.." && pwd)
VENV=${VENV:-/Users/soumitramishra/.claude/jobs/7baca02f/tmp/venv-crypto}
CORPUS=${CORPUS:-"$ROOT/corpus/external"}
MSOFF="$VENV/bin/msoffcrypto-tool"
PY="$VENV/bin/python"
PW='diff-check-pw-2026'
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
fails=0

if [ ! -x "$MSOFF" ]; then
  echo "msoffcrypto-tool not found at $MSOFF; set VENV or create the venv (see header). Skipping." >&2
  exit 0
fi
echo "Reference: msoffcrypto-tool $("$PY" -I -c 'import importlib.metadata as m; print(m.version("msoffcrypto-tool"))' 2>/dev/null || echo '(version unknown)')"

# A few decks: the small default plus a couple of larger ones, if present.
DECKS="
$CORPUS/python-pptx/src~pptx~templates~default.pptx
$CORPUS/python-pptx/features~steps~test_files~shp-movie-props.pptx
$CORPUS/poi/test-data~slideshow~KEY02.pptx
"

for deck in $DECKS; do
  [ -f "$deck" ] || { echo "skip (absent): $(basename "$deck")"; continue; }
  name=$(basename "$deck")

  # (1) msoffcrypto encrypt -> our decrypt == original
  "$MSOFF" -e -p "$PW" "$deck" "$TMP/enc1.bin" 2>/dev/null
  if node "$ROOT/tools/crypto/cli.ts" decrypt "$TMP/enc1.bin" "$TMP/dec1.bin" "$PW" && cmp -s "$deck" "$TMP/dec1.bin"; then
    echo "ok   msoffcrypto-encrypt -> ours-decrypt   $name"
  else
    echo "FAIL msoffcrypto-encrypt -> ours-decrypt   $name"; fails=$((fails+1))
  fi

  # (2) our encrypt -> msoffcrypto decrypt == original (lower spinCount to keep the loop quick)
  node "$ROOT/tools/crypto/cli.ts" encrypt "$deck" "$TMP/enc2.bin" "$PW" 10000
  if "$MSOFF" -p "$PW" "$TMP/enc2.bin" "$TMP/dec2.bin" 2>/dev/null && cmp -s "$deck" "$TMP/dec2.bin"; then
    echo "ok   ours-encrypt -> msoffcrypto-decrypt   $name"
  else
    echo "FAIL ours-encrypt -> msoffcrypto-decrypt   $name"; fails=$((fails+1))
  fi
done

# (3) committed msoffcrypto fixture still decrypts with the reference too (sanity on the fixture itself)
FIX="$ROOT/packages/crypto/test/fixtures/msoffcrypto/default-agile.pptx"
if [ -f "$FIX" ]; then
  if "$MSOFF" -p 'pptx-crypto-test-2026' "$FIX" "$TMP/fix.bin" 2>/dev/null && cmp -s "$ROOT/packages/crypto/test/fixtures/plain/default.pptx" "$TMP/fix.bin"; then
    echo "ok   fixture decrypts with msoffcrypto     default-agile.pptx"
  else
    echo "FAIL fixture decrypts with msoffcrypto     default-agile.pptx"; fails=$((fails+1))
  fi
fi

echo "---"
if [ "$fails" -eq 0 ]; then echo "differential: all checks passed"; else echo "differential: $fails check(s) failed"; fi
[ "$fails" -eq 0 ]
