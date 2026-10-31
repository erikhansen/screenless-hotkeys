# Building dist/JTHotkeys.dotm

Word for Mac has no command-line way to put VBA into a template, so the `.dotm` is built by hand in Word (about 2 minutes).

1. **Quit Word.** If an old copy is installed, remove it first so the modules don't load twice:
   `rm "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word/JTHotkeys.dotm"`
2. Open Word and create a **new blank document**.
3. **Tools > Macro > Visual Basic Editor**.
4. In the Project pane, select the new document's project (e.g. `Project (Document1)`).
5. **File > Import File…** and import each file in `src/`: `JTHotkeys.bas`, `PrintPages.bas`, and any new ones.
6. **Debug > Compile Project** should produce no errors.
7. Close the VBA editor. **File > Save As…**, set File Format to **Word Macro-Enabled Template (.dotm)**, name it `JTHotkeys.dotm`, and save it into this repo's `dist/` folder (replacing the old one).
8. Close the document, run `./install.sh`, and restart Word.
9. Test it: press **Control+Shift+V** (you should hear the version), then run the `JT_ToggleDryRun` macro and try **Control+Shift+P** with blank input, `2`, `2-3`, `99`, `abc`, and Escape.
10. Commit `src/` and `dist/JTHotkeys.dotm` together, then push.
