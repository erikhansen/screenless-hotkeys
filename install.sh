#!/bin/bash
# Installs (or reinstalls) JT Hotkeys for Microsoft Word on this Mac.
set -euo pipefail
cd "$(dirname "$0")"

STARTUP="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
SCRIPTS="$HOME/Library/Application Scripts/com.microsoft.Word"
TEMPLATE="dist/JTHotkeys.dotm"

say_and_echo() { echo "$1"; /usr/bin/say "$1" 2>/dev/null || true; }

if [ ! -f "$TEMPLATE" ]; then
  say_and_echo "Install failed. The template file is missing."
  exit 1
fi

WORD_WAS_RUNNING=0
if pgrep -xq "Microsoft Word"; then
  WORD_WAS_RUNNING=1
  say_and_echo "Closing Word to install the update. Word will ask to save any unsaved documents."
  osascript -e 'tell application "Microsoft Word" to quit' || true
  for _ in $(seq 1 60); do pgrep -xq "Microsoft Word" || break; sleep 1; done
  if pgrep -xq "Microsoft Word"; then
    say_and_echo "Word is still open. Please quit Word, then run the install again."
    exit 1
  fi
fi

mkdir -p "$STARTUP" "$SCRIPTS"
cp "$TEMPLATE" "$STARTUP/JTHotkeys.dotm"
osacompile -o "$SCRIPTS/JTHotkeys.scpt" applescript/JTHotkeys.applescript

if [ "$WORD_WAS_RUNNING" = 1 ]; then open -a "Microsoft Word"; fi
say_and_echo "J T hotkeys installed."
