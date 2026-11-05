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

# ChatGPT hotkey helper. It is only rebuilt when it's missing or its version
# changes: rebuilding makes macOS forget its Accessibility permission.
HELPER_DIR="$HOME/Applications"
HELPER="$HELPER_DIR/JT Hotkeys Helper.app"
HELPER_VERSION=$(sed -n 's/^property helperVersion : "\(.*\)"/\1/p' applescript/JTHotkeysHelper.applescript)
mkdir -p "$HELPER_DIR"
if [ ! -d "$HELPER" ] || [ "$(cat "$HELPER/Contents/Resources/helper-version" 2>/dev/null)" != "$HELPER_VERSION" ]; then
  tools/build-helper.sh "$HELPER_DIR" "$HELPER_VERSION"
  echo "Installed JT Hotkeys Helper. The first time you press Control+Option+Command+G,"
  echo "allow it in System Settings > Privacy & Security > Accessibility."
fi

# The ChatGPT hotkey starts the helper through launchd so that macOS gives the
# helper its own Accessibility permission (an app that Word opens borrows
# Word's permission instead).
AGENT_LABEL="com.screenless-hotkeys.helper"
AGENT="$HOME/Library/LaunchAgents/$AGENT_LABEL.plist"
mkdir -p "$HOME/Library/LaunchAgents"
cat > "$AGENT" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$AGENT_LABEL</string>
  <key>ProgramArguments</key>
  <array><string>$HELPER/Contents/MacOS/JT Hotkeys Helper</string></array>
  <key>LimitLoadToSessionType</key><string>Aqua</string>
  <key>ProcessType</key><string>Interactive</string>
  <!-- Let the spoken result finish after the helper quits. -->
  <key>AbandonProcessGroup</key><true/>
</dict>
</plist>
PLIST
launchctl bootout "gui/$(id -u)/$AGENT_LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$AGENT"

# Remove the Quick Action that older versions used for the ChatGPT hotkey.
# Only its own shortcut entry is removed; other Services settings are kept.
if [ -d "$HOME/Library/Services/Open ChatGPT in Word.workflow" ]; then
  rm -rf "$HOME/Library/Services/Open ChatGPT in Word.workflow"
  PBS_TMP=$(mktemp)
  if defaults export pbs "$PBS_TMP" 2>/dev/null &&
     /usr/libexec/PlistBuddy -c "Delete :NSServicesStatus:'com.krakencommerce.word-open-chatgpt - Open ChatGPT in Word - runWorkflowAsService'" "$PBS_TMP" 2>/dev/null; then
    defaults import pbs "$PBS_TMP"
  fi
  rm -f "$PBS_TMP"
  /System/Library/CoreServices/pbs -update >/dev/null 2>&1 || true
fi

if [ "$WORD_WAS_RUNNING" = 1 ]; then open -a "Microsoft Word"; fi
say_and_echo "J T hotkeys installed."
