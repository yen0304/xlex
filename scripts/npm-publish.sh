#!/bin/bash
# Script to publish npm packages manually
# Usage: ./scripts/npm-publish.sh <version>

set -e

VERSION=$1

if [ -z "$VERSION" ]; then
  echo "Usage: $0 <version>"
  echo "Example: $0 0.1.0"
  exit 1
fi

echo "Publishing xlex npm packages version $VERSION"
echo ""

# Update versions
echo "Updating package versions..."
for pkg in npm/*/package.json; do
  jq --arg v "$VERSION" '.version = $v' "$pkg" > tmp.json && mv tmp.json "$pkg"
done

# Update optionalDependencies
jq --arg v "$VERSION" '.optionalDependencies |= with_entries(.value = $v)' \
  npm/xlex/package.json > tmp.json && mv tmp.json npm/xlex/package.json

echo "Versions updated to $VERSION"
echo ""

# Check if binaries exist
for platform in linux-x64 linux-arm64 darwin-x64 darwin-arm64 win32-x64; do
  binary="npm/$platform/bin/xlex"
  if [ "$platform" = "win32-x64" ]; then
    binary="npm/$platform/bin/xlex.exe"
  fi
  
  if [ ! -f "$binary" ]; then
    echo "Warning: Binary not found: $binary"
    echo "Make sure to copy binaries before publishing!"
  fi
done

echo ""
echo "Ready to publish. Run the following commands:"
echo ""
echo "# Publish platform packages first"
echo "for pkg in linux-x64 linux-arm64 darwin-x64 darwin-arm64 win32-x64; do"
echo "  (cd npm/\$pkg && npm publish --access public)"
echo "done"
echo ""
echo "# Wait for registry to sync"
echo "sleep 30"
echo ""
echo "# Publish main package"
echo "(cd npm/xlex && npm publish --access public)"
