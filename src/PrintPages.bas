Attribute VB_Name = "PrintPages"
'==============================================================================
' Hotkey: print the current page, one page, or a range of pages.
'
'   Press the hotkey, then:
'     Enter             -> prints the page the cursor is on
'     7  Enter          -> prints page 7
'     3-5  Enter        -> prints pages 3 through 5
'     1,3,6-8  Enter    -> prints pages 1, 3 and 6 through 8
'     Escape            -> cancels
'
' Page numbers are physical pages counted from the start of the document
' (page 1 = first page), regardless of any printed page numbering.
'==============================================================================
Option Explicit

Public Sub JT_PrintPages()
    Dim doc As Document
    Dim answer As String, pageList As String, problem As String
    Dim pageCount As Long, current As Long

    If Documents.Count = 0 Then
        JT_Say "No document is open."
        Exit Sub
    End If
    Set doc = ActiveDocument
    pageCount = doc.ComputeStatistics(wdStatisticPages)
    current = Selection.Information(wdActiveEndPageNumber)

    JT_Say "Press Enter for current page or enter the page number or page number range."
    Do
        answer = InputBox("Press Enter to print the current page (" & current & ")," & vbCr & _
                          "or type a page number like 7, or a range like 3-5." & vbCr & _
                          "The document has " & pageCount & " pages.", "Print Pages")
        If StrPtr(answer) = 0 Then
            JT_Say "Print cancelled."
            Exit Sub
        End If

        pageList = NormalizePages(answer)
        If pageList = "" Then
            PrintCurrent doc, current
            Exit Sub
        End If

        problem = ValidatePages(pageList, pageCount)
        If problem = "" Then
            PrintList doc, pageList
            Exit Sub
        End If
        JT_Say problem & " Try again."
    Loop
End Sub

' "3 to 5" -> "3-5", strips spaces. Returns "" for blank input.
Private Function NormalizePages(ByVal s As String) As String
    s = LCase$(s)
    s = Replace(s, " to ", "-")
    s = Replace(s, "through", "-")
    s = Replace(s, "thru", "-")
    s = Replace(s, ChrW(8211), "-")   ' en dash
    s = Replace(s, ChrW(8212), "-")   ' em dash
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    NormalizePages = s
End Function

' Returns "" if valid, otherwise a sentence describing the problem.
Private Function ValidatePages(ByVal pageList As String, ByVal pageCount As Long) As String
    Dim parts() As String, ends() As String
    Dim i As Long, a As Long, b As Long

    parts = Split(pageList, ",")
    For i = LBound(parts) To UBound(parts)
        ends = Split(parts(i), "-")
        If UBound(ends) > 1 Or parts(i) = "" Then
            ValidatePages = "I did not understand " & SpokenList(pageList) & "."
            Exit Function
        End If
        If Not IsAllDigits(ends(0)) Or Not IsAllDigits(ends(UBound(ends))) Then
            ValidatePages = "I did not understand " & SpokenList(pageList) & "."
            Exit Function
        End If
        a = CLng(ends(0)): b = CLng(ends(UBound(ends)))
        If a < 1 Or b < 1 Or a > pageCount Or b > pageCount Then
            ValidatePages = "This document only has " & pageCount & IIf(pageCount = 1, " page.", " pages.")
            Exit Function
        End If
        If b < a Then
            ValidatePages = "The range " & a & " to " & b & " is backwards."
            Exit Function
        End If
    Next i
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 6 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Sub PrintCurrent(ByVal doc As Document, ByVal current As Long)
    JT_Say IIf(JT_DryRun, "Test mode. Would print page ", "Printing page ") & current & "."
    If JT_DryRun Then Exit Sub
    On Error GoTo Failed
    doc.PrintOut Range:=wdPrintCurrentPage, Copies:=1
    Exit Sub
Failed:
    JT_Say "Printing failed. " & Err.Description
End Sub

Private Sub PrintList(ByVal doc As Document, ByVal pageList As String)
    Dim spec As String
    spec = WordPageSpec(doc, pageList)
    Debug.Print "JT Hotkeys: PrintOut Pages:=" & spec
    JT_Say IIf(JT_DryRun, "Test mode. Would print ", "Printing ") & SpokenList(pageList) & "."
    If JT_DryRun Then Exit Sub
    On Error GoTo Failed
    doc.PrintOut Range:=wdPrintRangeOfPages, Pages:=spec, Copies:=1
    Exit Sub
Failed:
    JT_Say "Printing failed. " & Err.Description
End Sub

' Word's Pages:= argument uses the page numbers *printed* on the page and,
' when sections restart numbering, needs "p<page>s<section>". Convert our
' physical page numbers so "3" always means the 3rd physical page.
Private Function WordPageSpec(ByVal doc As Document, ByVal pageList As String) As String
    Dim parts() As String, ends() As String, out As String
    Dim i As Long, j As Long, piece As String
    parts = Split(pageList, ",")
    For i = LBound(parts) To UBound(parts)
        ends = Split(parts(i), "-")
        piece = ""
        For j = LBound(ends) To UBound(ends)
            If j > LBound(ends) Then piece = piece & "-"
            piece = piece & PhysicalToSpec(doc, CLng(ends(j)))
        Next j
        If out <> "" Then out = out & ","
        out = out & piece
    Next i
    WordPageSpec = out
End Function

Private Function PhysicalToSpec(ByVal doc As Document, ByVal n As Long) As String
    Dim rng As Range, shown As Long, sect As Long
    On Error GoTo Fallback
    Set rng = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=n)
    shown = rng.Information(wdActiveEndAdjustedPageNumber)
    sect = rng.Information(wdActiveEndSectionNumber)
    If doc.Sections.Count = 1 And shown = n Then
        PhysicalToSpec = CStr(n)
    Else
        PhysicalToSpec = "p" & shown & "s" & sect
    End If
    Exit Function
Fallback:
    PhysicalToSpec = CStr(n)
End Function

' "3-5,8" -> "pages 3 to 5 and 8";  "7" -> "page 7"
Private Function SpokenList(ByVal pageList As String) As String
    Dim s As String
    s = Replace(Replace(pageList, "-", " to "), ",", ", ")
    If InStr(pageList, "-") = 0 And InStr(pageList, ",") = 0 Then
        SpokenList = "page " & s
    Else
        SpokenList = "pages " & s
    End If
End Function
