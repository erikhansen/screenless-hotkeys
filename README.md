# JT Hotkeys for Microsoft Word (Mac)

Keyboard shortcuts for Word that work without looking at the screen. Every action is spoken out loud.

## Hotkeys

| Keys | What it does |
|------|--------------|
| **Control + Shift + P** | Print pages. Then: **Enter** prints the current page · type **7** then Enter prints page 7 · type **3-5** then Enter prints pages 3 to 5 · **1,3,6-8** also works · **all** prints the whole document · **Escape** cancels. |
| **Option + Command + Up Arrow** | Moves the cursor to the top of the previous page and says "Page N of M". |
| **Option + Command + Down Arrow** | Moves the cursor to the top of the next page and says "Page N of M". At the first or last page, it also says "Top of document" or "End of document". |
| **Control + Option + Up Arrow** | Moves the cursor to the start of the previous paragraph and says "Paragraph N of M". |
| **Control + Option + Down Arrow** | Moves the cursor to the start of the next paragraph and says "Paragraph N of M". Empty paragraphs and tables are skipped. If VoiceOver is on, it takes these keys first. |
| **Control + Option + Command + G** | Opens the ChatGPT panel (same as clicking the ChatGPT button on the Home tab). Needs a one-time permission; see [ChatGPT hotkey does nothing](#chatgpt-hotkey-does-nothing). |
| **Control + Shift + V** | Says which version of JT Hotkeys is installed. |

Page numbers mean physical pages counted from the start of the document. If the number is out of range, it tells you and asks again.

Every printout gets **"Page N" at the bottom** of each page. It's added just for printing and removed right after, so the document itself isn't changed. Pages that already show a page number in the footer don't get a second one. To turn this off, set `JT_PRINT_NUMBERS = False` in `src/JTHotkeys.bas` and rebuild.

Paragraph numbers in the left margin are available but **off by default**. To turn them on, set `JT_PRINT_PARAGRAPH_NUMBERS = True` in `src/JTHotkeys.bas` and rebuild. They're counted from the top of the document the same way as the paragraph hotkeys (empty paragraphs and tables are skipped). To see where the numbers will go, run the `JT_PreviewPrintNumbers` macro, then run it again to remove them.

### Optional: recent documents (Keyboard Maestro)

If Keyboard Maestro is installed, **Control + Shift + O** works from any app: it opens Word (if needed), shows Word's own **Recent documents** list with the cursor already in it, and says "Recent documents". Use the arrow keys to pick a document and **Enter** to open it, or **Escape** to cancel. The list comes from Word itself, not Spotlight, so it's always up to date. (It's the same as **File > Open Recent > More…**.)

To set it up, double-click `keyboard-maestro/Word - Open Recent Documents.kmmacros`. Keyboard Maestro imports it in a group called **JT Hotkeys** that starts out **disabled**. Turn it on by running this in Terminal:

```bash
osascript -e 'tell application "Keyboard Maestro" to set enabled of macro group "JT Hotkeys" to true'
```

## First-time install

Open Terminal and run:

```bash
cd ~ && git clone https://github.com/erikhansen/screenless-hotkeys.git && ~/screenless-hotkeys/install.sh
```

(The first time you use `git`, macOS may offer to install the "command line developer tools". Accept that and run the command again.)

The first time Word opens after installing, it may ask whether to enable macros for `JTHotkeys.dotm`. Choose **Enable Macros**.

## Getting updates

```bash
~/screenless-hotkeys/update.sh
```

This downloads the latest version, closes Word (Word asks you to save any open documents), installs the update, and reopens Word. You'll hear "J T hotkeys installed" when it's done.

## Troubleshooting

### "Operation not permitted" when installing

If the installer stops with an error like this:

```
cp: /Users/<you>/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/JTHotkeys.dotm: Operation not permitted
```

macOS is blocking Terminal from writing to Office's shared folder. This is a privacy setting, not a file permission, so `sudo` does not help. **Don't run the installer with `sudo`.** It can also leave files you can't update later.

To fix it:

1. Open **System Settings > Privacy & Security > Full Disk Access**.
2. Turn on **Terminal**, or whichever terminal app you use. If it isn't listed, click **+** and add it from Applications > Utilities.
3. Quit Terminal completely (Command+Q) and open it again. The new setting only takes effect after a restart.
4. Run the installer again, **without** `sudo`:

   ```bash
   ~/screenless-hotkeys/install.sh
   ```

If macOS instead asks whether Terminal "would like to access data from other apps", click **Allow**. It does the same thing.

**If you already ran it with `sudo`**, check who owns the installed files:

```bash
ls -l "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/" "$HOME/Library/Application Scripts/com.microsoft.Word/"
```

If any line shows `root`, give them back to your account once, then run the installer again (without `sudo`):

```bash
sudo chown -R "$USER" "$HOME/Library/Application Scripts/com.microsoft.Word" "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
```

### ChatGPT hotkey does nothing

The ChatGPT hotkey clicks the ChatGPT button for you. Clicking needs the macOS **Accessibility** permission, which belongs to a small app called **JT Hotkeys Helper** (installed in your home folder's Applications folder). Word itself doesn't need it.

The first time you press **Control + Option + Command + G**, you'll hear "J T Hotkeys Helper needs Accessibility permission", and System Settings opens at the right page. Then:

1. In **System Settings > Privacy & Security > Accessibility**, turn on **JT Hotkeys Helper**. If macOS shows an "Accessibility Access" prompt naming JT Hotkeys Helper, choose **Open System Settings** and turn it on there.
2. Press **Control + Option + Command + G** again. You'll hear "ChatGPT", and the ChatGPT panel opens.
3. If macOS asks whether JT Hotkeys Helper may control **System Events**, choose **OK**.

**If JT Hotkeys Helper is already switched on but you still hear that it needs permission**, the switch belongs to an older copy of the helper. Select **JT Hotkeys Helper** in the list, click the **minus (−)** button to remove it, press the hotkey again, and turn on the new entry. (Turning the switch off and on again also works sometimes, but removing it is more reliable.)

macOS may also show a notice that "JT Hotkeys Helper can run in the background". That's expected. It only runs for a moment when you press the hotkey. Don't turn it off under **System Settings > General > Login Items & Extensions**, or the hotkey stops working.

## Uninstall

```bash
rm "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/JTHotkeys.dotm"
rm "$HOME/Library/Application Scripts/com.microsoft.Word/JTHotkeys.scpt"
launchctl bootout "gui/$(id -u)/com.screenless-hotkeys.helper"
rm "$HOME/Library/LaunchAgents/com.screenless-hotkeys.helper.plist"
rm -r "$HOME/Applications/JT Hotkeys Helper.app"
```

---

## For maintainers

* `src/*.bas` holds the macro source code, kept as text so changes can be reviewed and diffed.
* `dist/JTHotkeys.dotm` is the built Word template that actually gets installed, including its key assignments. **Rebuild it with `./build.sh` whenever `src/` changes** (see [BUILD.md](BUILD.md)).
* `applescript/JTHotkeys.applescript` lets the macros speak (Word's sandbox can only run AppleScripts kept in `~/Library/Application Scripts/com.microsoft.Word/`).
* `applescript/JTHotkeysHelper.applescript` is **JT Hotkeys Helper.app**, which clicks the ChatGPT add-in button (Word can't assign keys to add-in buttons). `install.sh` builds it into `~/Applications` with `tools/build-helper.sh`, **only when it's missing or `helperVersion` changes**: every rebuild makes macOS forget its Accessibility permission, so bump `helperVersion` only when the helper itself changes. The `JT_OpenChatGPT` macro starts it through a LaunchAgent (`launchctl kickstart`), not `open`. An app started by Word borrows Word's Accessibility permission, which is what made the old Quick Action unreliable. One started by launchd gets its own entry in the Accessibility list.

### Adding a new hotkey

1. Create `src/MyFeature.bas` with a `Public Sub JT_MyFeature()`. Use `JT_Say "..."` for spoken feedback.
2. Add one line to `JT_BakeHotkeys` in `src/JTHotkeys.bas`:
   `BindKey "JT_MyFeature", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyM)`
3. Bump `JT_VERSION`, run `./build.sh` (see [BUILD.md](BUILD.md)), commit, and push. He runs `update.sh`.

Avoid **Control + Option** combinations, because VoiceOver uses them.

### Test mode

Run the `JT_ToggleDryRun` macro (Tools > Macro > Macros…) to turn on a mode where nothing is sent to the printer. The macro only speaks what it *would* print, and the exact `Pages:=` value it would pass to Word is logged in the VBA Immediate window. Test mode resets when Word restarts.
