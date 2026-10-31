#!/bin/bash
# Maintainer tool: rebuilds dist/JTHotkeys.dotm from src/*.bas by driving Word.
# Word must be running. Don't touch the keyboard/mouse while it runs.
# Requires your terminal app to have Accessibility permission
# (System Settings > Privacy & Security > Accessibility).
set -euo pipefail
cd "$(dirname "$0")"
REPO="$(pwd)"
# JTHotkeys.bas first, then every other module.
FILES=(src/JTHotkeys.bas)
for f in src/*.bas; do [ "$f" = src/JTHotkeys.bas ] || FILES+=("$f"); done
osascript tools/build-dotm.applescript "$REPO" "${FILES[@]}"
