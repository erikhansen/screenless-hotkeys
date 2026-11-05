# JT Hotkeys for Microsoft Word (Mac)

Keyboard shortcuts for Word that work without looking at the screen. Every action is spoken out loud.

## Hotkeys

| Keys | What it does |
|------|--------------|
| **Control + Shift + P** | Print pages. Then: **Enter** prints the current page · type **7** then Enter prints page 7 · type **3-5** then Enter prints pages 3 to 5 · **1,3,6-8** also works · **all** prints the whole document · **Escape** cancels. |
| **Option + Command + Up Arrow** | Moves the cursor to the top of the previous page and says "Page N of M". |
| **Option + Command + Down Arrow** | Moves the cursor to the top of the next page and says "Page N of M". At the first or last page, it also says "Top of document" or "End of document". |
| **Control + Option + Up Arrow** | Moves the cursor to the start of the previous paragraph and says "Paragraph N of M". |
| **Control + Option + Down Arrow** | Moves the cursor to the start of the next paragraph and says "Paragraph N of M". Empty paragraphs and tables are skipped, so the number matches the paragraph numbers on printouts. If VoiceOver is on, it takes these keys first. |
| **Control + Shift + V** | Says which version of JT Hotkeys is installed. |

Page numbers mean physical pages counted from the start of the document. If the number is out of range, it tells you and asks again.

Every printout gets **"Page N" at the bottom** of each page and a **number beside each paragraph** in the left margin (counted from the top of the document, so the numbers are the same on every printout). These are added just for printing and removed right after, so the document itself isn't changed. Empty paragraphs and tables aren't numbered, and pages that already show a page number in the footer don't get a second one. To turn this off, set `JT_PRINT_NUMBERS = False` in `src/JTHotkeys.bas` and rebuild. To see where the numbers will go, run the `JT_PreviewPrintNumbers` macro, then run it again to remove them.

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

## Uninstall

```bash
rm "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/JTHotkeys.dotm"
rm "$HOME/Library/Application Scripts/com.microsoft.Word/JTHotkeys.scpt"
```

---

## For maintainers

* `src/*.bas` holds the macro source code, kept as text so changes can be reviewed and diffed.
* `dist/JTHotkeys.dotm` is the built Word template that actually gets installed, including its key assignments. **Rebuild it with `./build.sh` whenever `src/` changes** (see [BUILD.md](BUILD.md)).
* `applescript/JTHotkeys.applescript` lets the macros speak (Word's sandbox can only run AppleScripts kept in `~/Library/Application Scripts/com.microsoft.Word/`).

### Adding a new hotkey

1. Create `src/MyFeature.bas` with a `Public Sub JT_MyFeature()`. Use `JT_Say "..."` for spoken feedback.
2. Add one line to `JT_BakeHotkeys` in `src/JTHotkeys.bas`:
   `BindKey "JT_MyFeature", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyM)`
3. Bump `JT_VERSION`, run `./build.sh` (see [BUILD.md](BUILD.md)), commit, and push. He runs `update.sh`.

Avoid **Control + Option** combinations, because VoiceOver uses them.

### Test mode

Run the `JT_ToggleDryRun` macro (Tools > Macro > Macros…) to turn on a mode where nothing is sent to the printer. The macro only speaks what it *would* print, and the exact `Pages:=` value it would pass to Word is logged in the VBA Immediate window. Test mode resets when Word restarts.
