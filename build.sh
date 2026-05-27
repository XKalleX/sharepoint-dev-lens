#!/usr/bin/env bash
set -e

ROOT="$(cd "$(dirname "$0")" && pwd)"
SRC="$ROOT/src"
DIST="$ROOT/dist"

echo ""
echo " SharePoint Dev Lens - Build Script"
echo " ==================================="

# Generate icons if not present
if [ ! -f "$SRC/icons/icon16.png" ]; then
  echo " [1/3] Generating icons..."
  node "$ROOT/scripts/generate-icons.js"
else
  echo " [1/3] Icons already present, skipping."
fi

mkdir -p "$DIST/chrome" "$DIST/edge"

# Chrome
echo " [2/3] Building Chrome package..."
cp -r "$SRC/." "$DIST/chrome/"
cp "$ROOT/manifests/manifest.chrome.json" "$DIST/chrome/manifest.json"

# Edge
echo " [3/3] Building Edge package..."
cp -r "$SRC/." "$DIST/edge/"
cp "$ROOT/manifests/manifest.edge.json" "$DIST/edge/manifest.json"

echo ""
echo " Done!"
echo ""
echo " Chrome  ->  dist/chrome/   (load as unpacked extension in Chrome)"
echo " Edge    ->  dist/edge/     (load as unpacked extension in Edge)"
echo ""
