Attribute VB_Name = "JTHotkeys"
'==============================================================================
' JT Hotkeys - core module
'
' Holds the hotkey list and shared helpers (speech feedback, test mode).
'
' HOW THE HOTKEYS GET INSTALLED: build.sh runs JT_BakeHotkeys inside the
' template while building it, which stores the key assignments in the .dotm
' itself. When Word loads the template from its Startup folder, the keys come
' with it. (Binding keys from a Startup add-in at launch does NOT work on Mac.)
'
' TO ADD A NEW HOTKEY:
'   1. Put the new macro in its own module (see PrintPages.bas as an example).
'   2. Add ONE line to JT_BakeHotkeys below.
'   3. Run build.sh.
'==============================================================================
Option Explicit

Public Const JT_VERSION As String = "1.0.2"

' Speak confirmations out loud (uses the macOS "say" command).
Public Const JT_SPEAK As Boolean = True

' When True, nothing is sent to the printer; the macro only says/logs what it
' would have printed. Toggle with the JT_ToggleDryRun macro (not saved).
Public JT_DryRun As Boolean

'------------------------------------------------------------------------------
' THE HOTKEY LIST. One line per hotkey.
'   wdKeyControl = the Control key on a Mac keyboard
'   wdKeyShift = Shift, wdKeyAlt = Option
' Avoid Control+Option combinations - VoiceOver uses those.
'------------------------------------------------------------------------------
Public Sub JT_BakeHotkeys()
    BindKey "JT_PrintPages", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyP)
    BindKey "JT_SayVersion", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyV)
    ' BindKey "JT_MyNextMacro", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyX)
End Sub

' Stores a shortcut in the document/template that contains this code.
Private Sub BindKey(ByVal macroName As String, ByVal keyCode As Long)
    CustomizationContext = ThisDocument
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:=macroName, KeyCode:=keyCode
End Sub

'------------------------------------------------------------------------------
' Helpers shared by all hotkeys
'------------------------------------------------------------------------------

' Speaks text without blocking Word. Needs JTHotkeys.scpt (installed by
' install.sh into ~/Library/Application Scripts/com.microsoft.Word/).
Public Sub JT_Say(ByVal text As String)
    Debug.Print "JT Hotkeys: " & text
    If Not JT_SPEAK Then Exit Sub
    On Error Resume Next
    AppleScriptTask "JTHotkeys.scpt", "speak", text
    If Err.Number <> 0 Then Beep
End Sub

Public Sub JT_SayVersion()
    JT_Say "J T hotkeys version " & JT_VERSION & IIf(JT_DryRun, ". Test mode is on.", ".")
End Sub

Public Sub JT_ToggleDryRun()
    JT_DryRun = Not JT_DryRun
    JT_Say IIf(JT_DryRun, "Test mode on. Nothing will print.", "Test mode off. Printing is live.")
End Sub
