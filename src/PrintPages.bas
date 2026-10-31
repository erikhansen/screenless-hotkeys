Attribute VB_Name = "PrintPages"
'==============================================================================
' Hotkey: print the current page, one page, or a range of pages.
'
'   Press the hotkey, then:
'     Enter             -> prints the page the cursor is on
'     7  Enter          -> prints page 7
'     3-5  Enter        -> prints pages 3 through 5
'     1,3,6-8  Enter    -> prints pages 1, 3 and 6 through 8
'     all  Enter        -> prints the whole document
'     Escape            -> cancels
'
' Page numbers are physical pages counted from the start of the document
' (page 1 = first page), regardless of any printed page numbering.
'
' The printout gets a page number at the bottom of each page and a number
' beside each paragraph (see JT_PRINT_NUMBERS in JTHotkeys.bas). These are
' added just for printing and removed right after; the document is unchanged.
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
                          "or type a page number like 7, a range like 3-5, or all." & vbCr & _
                          "The document has " & pageCount & " pages.", "Print Pages")
        If StrPtr(answer) = 0 Then
            JT_Say "Print cancelled."
            Exit Sub
        End If

        pageList = NormalizePages(answer)
        If pageList = "" Then
            PrintJob doc, "current", "", current, pageCount
            Exit Sub
        End If
        If pageList = "all" Then
            PrintJob doc, "all", "", current, pageCount
            Exit Sub
        End If

        problem = ValidatePages(pageList, pageCount)
        If problem = "" Then
            PrintJob doc, "list", pageList, current, pageCount
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

' kind: "current", "all" or "list".
Private Sub PrintJob(ByVal doc As Document, ByVal kind As String, ByVal pageList As String, _
                     ByVal current As Long, ByVal pageCount As Long)
    Dim spoken As String, spec As String, added As Collection
    Dim wasSaved As Boolean, oldBackground As Boolean, nParas As Long, nPages As Long
    Dim want() As Boolean

    Select Case kind
        Case "current": spoken = "page " & current
        Case "all": spoken = "all " & pageCount & IIf(pageCount = 1, " page", " pages")
        Case Else: spoken = SpokenList(pageList): spec = WordPageSpec(doc, pageList)
    End Select
    JT_Say IIf(JT_DryRun, "Test mode. Would print ", "Printing ") & spoken & "."
    Debug.Print "JT Hotkeys: kind=" & kind & " Pages:=" & spec

    wasSaved = doc.Saved
    Set added = New Collection
    On Error GoTo Failed

    If JT_PRINT_NUMBERS Then
        want = PagesWanted(kind, pageList, current, pageCount)
        nParas = AddParagraphNumbers(doc, want, added)
        nPages = AddPageNumbers(doc, added)
    End If

    If JT_DryRun Then
        JT_Say "Added " & nParas & " paragraph numbers and " & nPages & " page number boxes." & _
               IIf(spec <> "", " Word pages value: " & spec, "")
    Else
        ' Word for Mac keeps sending the job after PrintOut returns, so the
        ' numbers must stay until the job has reached the print queue.
        Dim oldDrawings As Boolean, jobBefore As String
        On Error Resume Next
        oldBackground = Options.PrintBackground
        Options.PrintBackground = False
        oldDrawings = Options.PrintDrawingObjects
        Options.PrintDrawingObjects = True            ' text boxes must print
        jobBefore = LastPrintJob()
        On Error GoTo Failed
        Select Case kind
            Case "current": doc.PrintOut Range:=wdPrintCurrentPage, Copies:=1
            Case "all": doc.PrintOut Range:=wdPrintAllDocument, Copies:=1
            Case Else: doc.PrintOut Range:=wdPrintRangeOfPages, Pages:=spec, Copies:=1
        End Select
        On Error Resume Next
        If JT_PRINT_NUMBERS Then WaitForPrintJob jobBefore
        Options.PrintBackground = oldBackground
        Options.PrintDrawingObjects = oldDrawings
        On Error GoTo Failed
    End If

    RemoveAdded added
    doc.Saved = wasSaved
    Exit Sub
Failed:
    JT_Say "Printing failed. " & Err.Description
    On Error Resume Next
    RemoveAdded added
    doc.Saved = wasSaved
End Sub

' True for each physical page that will be printed.
Private Function PagesWanted(ByVal kind As String, ByVal pageList As String, _
                             ByVal current As Long, ByVal pageCount As Long) As Boolean()
    Dim want() As Boolean, parts() As String, ends() As String
    Dim i As Long, j As Long
    ReDim want(1 To pageCount + 1)
    Select Case kind
        Case "current"
            If current >= 1 And current <= pageCount Then want(current) = True
        Case "all"
            For i = 1 To pageCount: want(i) = True: Next i
        Case Else
            parts = Split(pageList, ",")
            For i = LBound(parts) To UBound(parts)
                ends = Split(parts(i), "-")
                For j = CLng(ends(0)) To CLng(ends(UBound(ends)))
                    want(j) = True
                Next j
            Next i
    End Select
    PagesWanted = want
End Function

'------------------------------------------------------------------------------
' Temporary numbers for the printout. They are floating text boxes ("in front
' of text"), so they never move the document's text or change its pages.
' Everything added goes into `added` and is deleted after printing.
'------------------------------------------------------------------------------

' Numbers every paragraph that has text (1, 2, 3... through the whole document,
' skipping empty paragraphs and tables) in the left margin. Only paragraphs that
' start on a printed page get a box, but the numbering always counts from the
' top of the document so the numbers are the same on every printout.
Private Function AddParagraphNumbers(ByVal doc As Document, want() As Boolean, _
                                     ByVal added As Collection) As Long
    Dim p As Paragraph, idx As Long, pg As Long, txt As String, startRange As Range
    For Each p In doc.Paragraphs
        txt = p.Range.Text
        txt = Replace(Replace(Replace(Replace(txt, vbCr, ""), Chr(12), ""), Chr(7), ""), " ", "")
        If txt <> "" Then
            If Not p.Range.Information(wdWithInTable) Then
                idx = idx + 1
                ' Look at the paragraph's first character (a collapsed range at
                ' a page boundary can report the previous page).
                Set startRange = doc.Range(p.Range.Start, p.Range.Start + 1)
                pg = startRange.Information(wdActiveEndPageNumber)
                If pg >= LBound(want) And pg <= UBound(want) Then
                    If want(pg) Then
                        AddMarginLabel doc, p, CStr(idx), added
                        AddParagraphNumbers = AddParagraphNumbers + 1
                    End If
                End If
            End If
        End If
    Next p
End Function

Private Sub AddMarginLabel(ByVal doc As Document, ByVal p As Paragraph, ByVal label As String, _
                           ByVal added As Collection)
    Const BOX_W As Single = 30
    Dim shp As Shape, leftMargin As Single, x As Single
    leftMargin = p.Range.Sections(1).PageSetup.LeftMargin
    x = leftMargin - BOX_W - 6
    If x < 2 Then x = 2

    Set shp = doc.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, BOX_W, 14, p.Range)
    added.Add shp
    With shp
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .WrapFormat.Type = wdWrapFront
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
        .Left = x
        .Top = 0
        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .TextRange.Text = label
            .TextRange.Font.Size = 9
            .TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
            .TextRange.ParagraphFormat.SpaceBefore = 0
            .TextRange.ParagraphFormat.SpaceAfter = 0
        End With
    End With
End Sub

' Puts "Page N" at the bottom centre of every page, unless that footer already
' shows a page number.
Private Function AddPageNumbers(ByVal doc As Document, ByVal added As Collection) As Long
    Dim sec As Section
    For Each sec In doc.Sections
        AddPageNumberTo doc, sec, wdHeaderFooterPrimary, added, AddPageNumbers
        If sec.PageSetup.DifferentFirstPageHeaderFooter Then _
            AddPageNumberTo doc, sec, wdHeaderFooterFirstPage, added, AddPageNumbers
        If sec.PageSetup.OddAndEvenPagesHeaderFooter Then _
            AddPageNumberTo doc, sec, wdHeaderFooterEvenPages, added, AddPageNumbers
    Next sec
End Function

Private Sub AddPageNumberTo(ByVal doc As Document, ByVal sec As Section, ByVal which As Long, _
                            ByVal added As Collection, ByRef count As Long)
    Dim ft As HeaderFooter, shp As Shape, r As Range, f As Field
    Set ft = sec.Footers(which)
    If sec.Index > 1 And ft.LinkToPrevious Then Exit Sub   ' shares the previous section's footer
    For Each f In ft.Range.Fields
        If f.Type = wdFieldPage Then Exit Sub                ' already has a page number
    Next f

    Set shp = ft.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 120, 16, ft.Range)
    added.Add shp
    With shp
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .WrapFormat.Type = wdWrapFront
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Left = wdShapeCenter
        .Top = sec.PageSetup.PageHeight - (sec.PageSetup.BottomMargin / 2) - 8
        With .TextFrame
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .TextRange.Text = "Page "
            .TextRange.Font.Size = 10
            .TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Set r = .TextRange
            r.SetRange r.Start + Len("Page "), r.Start + Len("Page ")   ' right after "Page "
            doc.Fields.Add Range:=r, Type:=wdFieldPage
        End With
    End With
    count = count + 1
End Sub

' Highest CUPS job number so far (all printers, including finished jobs).
Private Function LastPrintJob() As String
    On Error Resume Next
    LastPrintJob = AppleScriptTask("JTHotkeys.scpt", "lastPrintJob", "")
End Function

' Waits (up to a minute) until a new job shows up in the print queue, then a
' few seconds more so Word finishes sending it.
Private Sub WaitForPrintJob(ByVal jobBefore As String)
    Dim deadline As Single
    deadline = Timer + 60
    If jobBefore <> "" Then
        Do While Timer < deadline
            Pause 0.5
            If LastPrintJob() <> jobBefore Then Exit Do
        Loop
    End If
    On Error Resume Next
    Do While Application.BackgroundPrintingStatus > 0 And Timer < deadline
        Pause 0.5
    Loop
    Pause 3
End Sub

Private Sub Pause(ByVal seconds As Single)
    Dim t As Single
    t = Timer
    Do While Timer < t + seconds And Timer >= t
        DoEvents
    Loop
End Sub

Private Sub RemoveAdded(ByVal added As Collection)
    Dim i As Long
    On Error Resume Next
    For i = added.Count To 1 Step -1
        added(i).Delete
    Next i
End Sub

' Test helper (no hotkey): adds the print numbers to the whole document so you
' can see where they go. Run it again to remove them.
Public Sub JT_PreviewPrintNumbers()
    Static previewShapes As Collection
    Dim want() As Boolean, n As Long
    If Documents.Count = 0 Then Exit Sub
    If Not previewShapes Is Nothing Then
        RemoveAdded previewShapes
        Set previewShapes = Nothing
        JT_Say "Print numbers removed."
        Exit Sub
    End If
    Set previewShapes = New Collection
    n = ActiveDocument.ComputeStatistics(wdStatisticPages)
    want = PagesWanted("all", "", 0, n)
    AddParagraphNumbers ActiveDocument, want, previewShapes
    AddPageNumbers ActiveDocument, previewShapes
    JT_Say "Print numbers shown. Run this again to remove them."
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

    ' Normal case: one section, numbering starts at 1, so the printed page
    ' number IS the physical page number. Send it as-is.
    If doc.Sections.Count = 1 And Not NumberingRestarts(doc.Sections(1)) Then
        PhysicalToSpec = CStr(n)
        Exit Function
    End If

    ' Otherwise ask Word what page number is printed on physical page n.
    ' Look at the first character OF page n: a collapsed range at the very
    ' start of a page can report the previous page (the old off-by-one bug,
    ' where "2-3" printed pages 1-2).
    Set rng = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=n)
    If rng.Start < doc.Content.End - 1 Then Set rng = doc.Range(rng.Start, rng.Start + 1)
    If rng.Information(wdActiveEndPageNumber) <> n Then GoTo Fallback
    shown = rng.Information(wdActiveEndAdjustedPageNumber)
    sect = rng.Information(wdActiveEndSectionNumber)
    PhysicalToSpec = "p" & shown & "s" & sect
    Exit Function
Fallback:
    PhysicalToSpec = CStr(n)
End Function

Private Function NumberingRestarts(ByVal sec As Section) As Boolean
    On Error Resume Next
    With sec.Headers(wdHeaderFooterPrimary).PageNumbers
        NumberingRestarts = .RestartNumberingAtSection And .StartingNumber <> 1
    End With
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
