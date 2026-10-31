# Building dist/JTHotkeys.dotm

Word for Mac can't put VBA into a template from the command line, so `build.sh` drives Word's VBA editor for you.

## Automatic (recommended)

1. Give your terminal app Accessibility permission once: **System Settings → Privacy & Security → Accessibility**.
2. Open Word.
3. Run `./build.sh` and **don't touch the keyboard or mouse** until it finishes (about 20 seconds).

The script:
- temporarily unloads the installed add-in
- creates a new document and imports every `src/*.bas` (`JTHotkeys.bas` first)
- compiles the project
- runs `JT_BakeHotkeys`, which stores the key assignments inside the template
- saves `dist/JTHotkeys.dotm` and reloads the add-in

Then run `./install.sh` to install it on this Mac. That restarts Word, and Word asks you to save any open documents first.

## Manual fallback

1. Open Word and create a new blank document.
2. **Tools → Macro → Visual Basic Editor**. Select `Project (DocumentN)`, then use **File → Import File…** to import each file in `src/`.
3. **Debug → Compile Project** should give no errors.
4. Close the editor, open **Tools → Macro → Macros…**, and run **JT_BakeHotkeys**. This step is required, because it's what assigns the keys.
5. **File → Save As…**, choose **Word Macro-Enabled Template (.dotm)**, and save as `dist/JTHotkeys.dotm`.

## Testing

After installing, press **Control+Shift+P** in a document. You should hear the prompt. To test without printing, first run the `JT_ToggleDryRun` macro, then try blank input, `2`, `2-3`, `99`, `abc`, and Escape.
