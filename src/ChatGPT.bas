Attribute VB_Name = "ChatGPT"
'==============================================================================
' Control + Option + Command + G: open the ChatGPT add-in panel.
'
' Word can't assign keys to add-in buttons, so this starts "JT Hotkeys Helper"
' (~/Applications, built by install.sh), which clicks the ChatGPT button on
' the Home tab. The helper needs the macOS Accessibility permission once; it
' says so out loud and opens the right Settings page if it doesn't have it.
'==============================================================================
Option Explicit

Public Sub JT_OpenChatGPT()
    Dim result As String
    On Error GoTo Failed
    result = AppleScriptTask("JTHotkeys.scpt", "openChatGPT", "")
    If result = "missing" Then
        JT_SayNow "J T Hotkeys Helper is not installed. Please run the install script again."
    End If
    Exit Sub
Failed:
    JT_SayNow "Could not open ChatGPT. Please run the install script again."
End Sub
