# JT Hotkeys for Microsoft Word (Mac)

Keyboard shortcuts for Word that work without looking at the screen. Every action is spoken out loud.

## Hotkeys

| Keys | What it does |
|------|--------------|
| **Control + Shift + P** | Print pages. Then: **Enter** prints the current page · type **7** then Enter prints page 7 · type **3-5** then Enter prints pages 3 to 5 · **1,3,6-8** also works · **Escape** cancels. |
| **Control + Shift + V** | Says which version of JT Hotkeys is installed. |

Page numbers mean physical pages counted from the start of the document. If the number is out of range, it tells you and asks again.

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
