Attribute VB_Name = "ParagraphNav"
'==============================================================================
' Hotkeys: move by paragraphs and say where you are.
'
'   Control + Option + Up Arrow    -> cursor to the start of the previous paragraph
'   Control + Option + Down Arrow  -> cursor to the start of the next paragraph
'
' Each press says "Paragraph N of M". Paragraphs are counted the same way as
' the numbers printed in the margin (see PrintPages.bas): empty paragraphs
' and table cells are skipped, so the spoken number matches the printout.
'==============================================================================
Option Explicit

Public Sub JT_PreviousParagraph()
    MoveByParagraphs -1
End Sub

Public Sub JT_NextParagraph()
    MoveByParagraphs 1
End Sub

Private Sub MoveByParagraphs(ByVal delta As Long)
    Dim doc As Document, starts() As Long, total As Long
    Dim current As Long, target As Long, i As Long, pos As Long, note As String
    On Error GoTo Failed
    If Documents.Count = 0 Then
        JT_SayNow "No document is open."
        Exit Sub
    End If
    Set doc = ActiveDocument
    total = NumberedParagraphStarts(doc, starts)
    If total = 0 Then
        JT_SayNow "This document has no paragraphs with text."
        Exit Sub
    End If

    ' Current paragraph = the last numbered one starting at or before the
    ' cursor (0 if the cursor is above the first one).
    pos = Selection.Start
    For i = 1 To total
        If starts(i) <= pos Then current = i Else Exit For
    Next i

    target = current + delta
    If delta < 0 And current = 0 Then target = 1
    If target < 1 Then
        target = 1
        note = " Top of document."
    ElseIf target > total Then
        target = total
        note = " End of document."
    End If

    Selection.SetRange starts(target), starts(target)
    ActiveWindow.ScrollIntoView Selection.Range, True
    JT_SayNow "Paragraph " & target & " of " & total & "." & note
    Exit Sub
Failed:
    JT_SayNow "Could not move paragraphs. " & Err.Description
End Sub

' Fills starts(1..n) with the start position of every paragraph that gets a
' number on the printout, and returns n.
Private Function NumberedParagraphStarts(ByVal doc As Document, starts() As Long) As Long
    Dim p As Paragraph, txt As String, n As Long
    ReDim starts(1 To doc.Paragraphs.Count)
    For Each p In doc.Paragraphs
        txt = p.Range.Text
        txt = Replace(Replace(Replace(Replace(txt, vbCr, ""), Chr(12), ""), Chr(7), ""), " ", "")
        If txt <> "" Then
            If Not p.Range.Information(wdWithInTable) Then
                n = n + 1
                starts(n) = p.Range.Start
            End If
        End If
    Next p
    NumberedParagraphStarts = n
End Function
