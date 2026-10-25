#!/bin/sh
# Fetch third-party presentation test files into corpus/external/<source>/.
#
# These files are downloaded test data and are treated as untrusted input: they are only ever
# read by the parser under test, never executed, and kept out of git (see .gitignore).
# corpus/external/MANIFEST.tsv records the repository, commit and path each file came from.
#
# Sources (each a shallow, sparse, blob-filtered clone, so only the listed folders are fetched):
#   poi           Apache POI test-data/slideshow                       (Apache-2.0)
#   python-pptx   python-pptx test fixtures and its default template   (MIT)
#   openxml-sdk   Open XML SDK test assets, including Strict files     (MIT)
#   libreoffice   LibreOffice sd/qa/unit/data/pptx and oox/qa/unit/data (MPL-2.0; files from bug reports)
set -eu

ROOT=$(cd "$(dirname "$0")/../.." && pwd)
DEST="$ROOT/corpus/external"
CACHE="$ROOT/corpus/.cache"
MANIFEST="$DEST/MANIFEST.tsv"
mkdir -p "$DEST" "$CACHE"

checkout() { # <name> <url> <folder>...
  name=$1; url=$2; shift 2
  repo="$CACHE/$name"
  if [ ! -d "$repo/.git" ]; then
    echo "clone $name"
    git clone --quiet --depth 1 --filter=blob:none --sparse --no-tags "$url" "$repo"
  fi
  git -C "$repo" sparse-checkout set "$@"
}

collect() { # <name> <url>
  name=$1; url=$2
  repo="$CACHE/$name"
  commit=$(git -C "$repo" rev-parse HEAD)
  rm -rf "${DEST:?}/$name"
  mkdir -p "$DEST/$name"
  # Flatten each path into the file name so two folders' "test.pptx" cannot collide.
  git -C "$repo" ls-files | grep -i -E '\.(pptx|pptm|potx|potm|ppsx|ppsm|ppt|thmx)$' | while IFS= read -r rel; do
    [ -f "$repo/$rel" ] || continue
    flat=$(printf '%s' "$rel" | tr '/' '~')
    cp "$repo/$rel" "$DEST/$name/$flat"
    sum=$(shasum -a 256 "$DEST/$name/$flat" | cut -d' ' -f1)
    size=$(wc -c < "$DEST/$name/$flat" | tr -d ' ')
    printf '%s\t%s\t%s\t%s\t%s\t%s\n' "$name" "$name/$flat" "$url" "$commit" "$rel" "$sum $size" >> "$MANIFEST.tmp"
  done
}

: > "$MANIFEST.tmp"
checkout poi https://github.com/apache/poi.git test-data/slideshow
collect poi https://github.com/apache/poi.git
checkout python-pptx https://github.com/scanny/python-pptx.git tests/test_files features/steps/test_files src/pptx/templates
collect python-pptx https://github.com/scanny/python-pptx.git
checkout openxml-sdk https://github.com/dotnet/Open-XML-SDK.git test/DocumentFormat.OpenXml.Tests.Assets/assets
collect openxml-sdk https://github.com/dotnet/Open-XML-SDK.git
checkout libreoffice https://github.com/LibreOffice/core.git sd/qa/unit/data/pptx oox/qa/unit/data
collect libreoffice https://github.com/LibreOffice/core.git

{ printf 'source\tfile\trepository\tcommit\tpath\tsha256 bytes\n'; cat "$MANIFEST.tmp"; } > "$MANIFEST"
rm -f "$MANIFEST.tmp"
echo "corpus: $(($(wc -l < "$MANIFEST") - 1)) files in $DEST"
