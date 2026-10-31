Attribute VB_Name = "JTHotkeys"
'==============================================================================
' JT Hotkeys - core module
'
' Registers keyboard shortcuts every time Word starts (AutoExec) and provides
' shared helpers (speech feedback, beeps).
'
' TO ADD A NEW HOTKEY:
'   1. Put the new macro in its own module (see PrintPages.bas as an example).
'   2. Add ONE line to RegisterHotkeys below.
'==============================================================================
Option Explicit

Public Const JT_VERSION As String = "1.0.1"

' Speak confirmations out loud (uses the macOS "say" command).
Public Const JT_SPEAK As Boolean = True

' When True, nothing is sent to the printer; the macro only says/logs what it
' would have printed. Toggle with the JT_ToggleDryRun macro (not saved).
Public JT_DryRun As Boolean

'------------------------------------------------------------------------------
' Runs automatically when Word starts (because this template is in Word's
' Startup folder).
'------------------------------------------------------------------------------
Public Sub AutoExec()
    RegisterHotkeys
End Sub

'------------------------------------------------------------------------------
' THE HOTKEY LIST. One line per hotkey.
'   BuildKeyCode(wdKeyControl, ...) = the Control key on a Mac keyboard
'   wdKeyShift = Shift, wdKeyAlt = Option
' Avoid Control+Option combinations - VoiceOver uses those.
'------------------------------------------------------------------------------
Public Sub RegisterHotkeys()
    BindKey "JT_PrintPages", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyP)
    BindKey "JT_SayVersion", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyV)
    ' BindKey "JT_MyNextMacro", BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyX)
End Sub

' Assigns a shortcut to a macro. Bindings are added to Normal in memory and
' re-added on every launch, so Normal.dotm is never marked as changed.
Public Sub BindKey(ByVal macroName As String, ByVal keyCode As Long)
    Dim wasSaved As Boolean
    On Error GoTo Failed
    wasSaved = NormalTemplate.Saved
    CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:=macroName, KeyCode:=keyCode
    NormalTemplate.Saved = wasSaved
    Exit Sub
Failed:
    Debug.Print "JT Hotkeys: could not bind " & macroName & ": " & Err.Description
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
