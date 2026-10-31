Attribute VB_Name = "PageNav"
'==============================================================================
' Hotkeys: move by whole pages and say where you are.
'
'   Option + Command + Up Arrow    -> cursor to the top of the previous page
'   Option + Command + Down Arrow  -> cursor to the top of the next page
'
' Each press says "Page N of M". Pressing again before the speech finishes
' cuts it off, so holding the keys or tapping quickly stays responsive.
'==============================================================================
Option Explicit

Public Sub JT_PreviousPage()
    MoveByPages -1
End Sub

Public Sub JT_NextPage()
    MoveByPages 1
End Sub

Private Sub MoveByPages(ByVal delta As Long)
    Dim doc As Document, current As Long, total As Long, target As Long, note As String
    On Error GoTo Failed
    If Documents.Count = 0 Then
        JT_SayNow "No document is open."
        Exit Sub
    End If
    Set doc = ActiveDocument
    total = doc.Content.Information(wdNumberOfPagesInDocument)
    current = CursorPage(doc)

    target = current + delta
    If target < 1 Then
        target = 1
        note = " Top of document."
    ElseIf target > total Then
        target = total
        note = " End of document."
    End If

    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=target
    JT_SayNow "Page " & target & " of " & total & "." & note
    Exit Sub
Failed:
    JT_SayNow "Could not move pages. " & Err.Description
End Sub

' Page of the character at the cursor. (A bare insertion point sitting exactly
' at the top of a page can report the page before, so look at the character.)
Private Function CursorPage(ByVal doc As Document) As Long
    Dim pos As Long
    pos = Selection.Start
    If pos < doc.Content.End - 1 Then
        CursorPage = doc.Range(pos, pos + 1).Information(wdActiveEndPageNumber)
    Else
        CursorPage = Selection.Information(wdActiveEndPageNumber)
    End If
End Function
