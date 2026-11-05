#!/bin/bash
# Builds "JT Hotkeys Helper.app" (clicks the ChatGPT add-in button) into the
# given folder. Called by install.sh.
# Usage: tools/build-helper.sh <folder> <helper version>
set -euo pipefail
cd "$(dirname "$0")/.."
APP="$1/JT Hotkeys Helper.app"
rm -rf "$APP"
osacompile -o "$APP" applescript/JTHotkeysHelper.applescript
PLIST="$APP/Contents/Info.plist"
# Name the program "JT Hotkeys Helper" instead of "applet", so that's the name
# macOS shows in Login Items and in its "can run in the background" notice.
mv "$APP/Contents/MacOS/applet" "$APP/Contents/MacOS/JT Hotkeys Helper"
plutil -replace CFBundleExecutable -string "JT Hotkeys Helper" "$PLIST"
plutil -replace CFBundleIdentifier -string com.screenless-hotkeys.helper "$PLIST"
plutil -replace CFBundleName -string "JT Hotkeys Helper" "$PLIST"
plutil -replace LSUIElement -bool true "$PLIST"
plutil -replace NSAppleEventsUsageDescription -string \
  "JT Hotkeys Helper clicks the ChatGPT button in Microsoft Word for you." "$PLIST"
echo "$2" > "$APP/Contents/Resources/helper-version"
codesign --force --sign - "$APP" 2>/dev/null
