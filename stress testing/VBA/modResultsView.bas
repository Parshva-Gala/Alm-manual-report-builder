Attribute VB_Name = "modResultsView"
Option Explicit

' ============================================================================
'  Results (Recon_Detail), design system v2
'
'  The run writes the evidence exactly as before: header on row 7, data from
'  row 8, columns A:R in the same order, so the console, the review pack, the
'  evidence export and data clearing read it unchanged. This module only dresses
'  it for reading:
'
'    row 3   title, and four quick-filter buttons on the right
'    row 4   what the run covered, in one line
'    row 5   the verdict split: checks, reconciled, differences, open, assumed
'    row 6   the same split as one stacked bar
'    row 7+  a filtered range: status pills, stage chips, difference bars,
'            hairline rows, frozen header and first two columns
'
'  Every step here is cosmetic, so each is allowed to fail on its own without
'  stopping the run that called it.
' ============================================================================

Private Const RESULTS_TABLE As String = "ResultsTable"
Private Const COLS As Long = 18

' Call before the sheet is cleared: removes the table and its slicers so the run
' can rewrite the cells freely. The values and formats stay until the clear.
Public Sub ResultsResetSheet(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    If ws.ListObjects.count > 0 Then
        DeleteResultsSlicers ws
        For i = ws.ListObjects.count To 1 Step -1
            ws.ListObjects(i).Unlist
        Next i
    End If
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).name, 8) = "Results_" Then ws.Shapes(i).Delete
    Next i
    Err.Clear
End Sub

' Call after the run has written the rows and StyleReconTable has run.
Public Sub ResultsDecorate(ByVal ws As Worksheet, ByVal rowCount As Long, ByVal runText As String)
    Dim lastRow As Long, counts As Variant
    On Error Resume Next
    lastRow = 7 + IIf(rowCount > 0, rowCount, 1)
    counts = CountVerdicts(ws, rowCount)

    ' ---- the frame -------------------------------------------------------
    ws.Range("A3:R6").Interior.Color = UI_WHITE
    ws.Cells.Font.name = UI_FONT
    With ws.Range("A3")
        .Value2 = "Results": .Font.Size = 22: .Font.Bold = True: .Font.Color = UI_INK
        .VerticalAlignment = xlBottom: .IndentLevel = 0
    End With
    With ws.Range("A4")
        .Value2 = IIf(Len(runText) > 0, "Run " & runText & "   " & ChrW(183) & "   ", "") & _
                  "Difference is the system minus the tool. Filter with the buttons above the table or any column header."
        .Font.Size = 10: .Font.Color = UI_MUTED: .VerticalAlignment = xlCenter
    End With
    ws.rows(3).RowHeight = 36: ws.rows(4).RowHeight = 22: ws.rows(5).RowHeight = 26: ws.rows(6).RowHeight = 14
    VerdictLine ws.Range("A5"), counts
    VerdictBar ws, counts

    ' ---- the table -------------------------------------------------------
    ' A plain range with AutoFilter, not an Excel table: a table with slicers
    ' that is cleared and rebuilt on every run can crash Excel outright.
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, COLS)).AutoFilter
    With ws.Range(ws.Cells(7, 1), ws.Cells(7, COLS))
        .Interior.Color = UI_NAVY: .Font.Color = UI_WHITE: .Font.Bold = True: .Font.Size = 9.5
        .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.rows(7).RowHeight = 30
    If rowCount > 0 Then
        With ws.Range(ws.Cells(8, 1), ws.Cells(lastRow, COLS))
            .Font.Size = 9.5: .Font.Color = UI_TEXT: .VerticalAlignment = xlCenter
            .Interior.Color = UI_WHITE
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous: .Borders(xlInsideHorizontal).Color = UI_LINE
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = UI_LINE
        End With
        ws.rows("8:" & lastRow).RowHeight = 21
        ws.Range("A8:A" & lastRow).HorizontalAlignment = xlCenter
        ws.Range("A8:A" & lastRow).Font.Bold = True: ws.Range("A8:A" & lastRow).Font.Size = 8.5
        ws.Range("B8:B" & lastRow).HorizontalAlignment = xlCenter: ws.Range("B8:B" & lastRow).Font.Size = 8.5
        ws.Range("E8:E" & lastRow).Font.Bold = True: ws.Range("E8:E" & lastRow).Font.Color = UI_INK
        ws.Range("H8:H" & lastRow).Font.name = "Consolas": ws.Range("H8:H" & lastRow).Font.Size = 9
        ws.Range("I8:K" & lastRow).Font.name = "Consolas": ws.Range("I8:K" & lastRow).Font.Size = 9
        ws.Range("I8:J" & lastRow).NumberFormat = "#,##0.00;-#,##0.00;0.00"
        ws.Range("K8:K" & lastRow).NumberFormat = "+#,##0.00;-#,##0.00;0.00"
        ws.Range("L8:L" & lastRow).Font.Color = UI_MUTED
        ws.Range("M8:R" & lastRow).Font.Color = UI_TEXT_2
        ws.Range("N8:N" & lastRow).Font.Color = UI_TEXT
        StatusPills ws.Range("A8:A" & lastRow)
        StageChips ws.Range("B8:B" & lastRow)
        SeverityText ws.Range("G8:G" & lastRow)
        DifferenceBars ws.Range("K8:K" & lastRow)
    End If
    ws.columns("A").ColumnWidth = 13: ws.columns("B").ColumnWidth = 12
    ws.columns("C:D").ColumnWidth = 13: ws.columns("E:F").ColumnWidth = 22: ws.columns("G").ColumnWidth = 11

    AddQuickFilters ws
    ws.Tab.Color = UI_BRAND
    ' Window settings only when a person is looking, never in the middle of a run.
    If Not JKB_Busy Then
        If UiTryActivate(ws) Then ResultsWindow ws
    End If
    Err.Clear
End Sub

Private Sub ResultsWindow(ByVal ws As Worksheet)
    On Error Resume Next
    If UiOwnWindow(ws) Is Nothing Then Exit Sub
    UiSetView ws, 7, 2, 90, False
    If UiIsActiveWindow(UiOwnWindow(ws)) Then ws.Range("A8").Select
    Err.Clear
End Sub

' ---------------------------------------------------------------- quick filters
' Four buttons above the table: everything, differences, open items, reconciled.

Public Sub ResultsFilterAll()
    QuickFilter ""
End Sub

Public Sub ResultsFilterDifferences()
    QuickFilter "FAIL"
End Sub

Public Sub ResultsFilterOpen()
    QuickFilter "OPEN"
End Sub

Public Sub ResultsFilterReconciled()
    QuickFilter "PASS"
End Sub

Private Sub QuickFilter(ByVal which As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetWorksheetSafe(ThisWorkbook, "Recon_Detail")
    If ws Is Nothing Then Exit Sub
    If ws.FilterMode Then ws.ShowAllData
    If Not ws.AutoFilterMode Then ws.Range("A7").CurrentRegion.AutoFilter
    Select Case which
        Case "FAIL", "PASS": ws.Range("A7").AutoFilter field:=1, Criteria1:=which
        Case "OPEN": ws.Range("A7").AutoFilter field:=1, Criteria1:=Array("BLOCKED", "REVIEW"), Operator:=xlFilterValues
    End Select
    If UiTryActivate(ws) Then
        On Error Resume Next
        UiOwnWindow(ws).ScrollRow = 8
        Err.Clear
    End If
End Sub

Private Sub AddQuickFilters(ByVal ws As Worksheet)
    Dim caps As Variant, procs As Variant, i As Long, x As Double, w As Double, sh As Shape
    caps = Array("All checks", "Differences", "Missing or to review", "Reconciled")
    procs = Array("ResultsFilterAll", "ResultsFilterDifferences", "ResultsFilterOpen", "ResultsFilterReconciled")
    x = ws.Range("C3").Left + 8   ' beside the title, so all four fit on screen
    For i = 0 To 3
        w = Len(caps(i)) * 5.8 + 26
        UiButton ws, "Results_Q" & i, CStr(caps(i)), CStr(procs(i)), ws.Range("C3"), IIf(i = 1, UI_BTN_PRIMARY, UI_BTN_SECONDARY), 9
        Set sh = ws.Shapes("Results_Q" & i)
        sh.Left = x: sh.Width = w: sh.Height = 22
        sh.Top = ws.Range("C3").Top + (ws.rows(3).RowHeight - 22) / 2
        x = x + w + 6
    Next i
End Sub

' Clears every filter so all rows show.
Public Sub ResultsShowAll(ByVal ws As Worksheet)
    Dim lo As ListObject
    On Error Resume Next
    For Each lo In ws.ListObjects
        If lo.ShowAutoFilter Then
            If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
        End If
    Next lo
    If ws.FilterMode Then ws.ShowAllData
    ResultsWindow ws
    Err.Clear
End Sub

' ---------------------------------------------------------------- pieces -----

' PASS, FAIL, REVIEW + BLOCKED, ASSUMPTION, total.
Private Function CountVerdicts(ByVal ws As Worksheet, ByVal rowCount As Long) As Variant
    Dim a As Variant, i As Long, p As Long, f As Long, o As Long, s As Long, v As String
    If rowCount > 0 Then
        a = ws.Range(ws.Cells(8, 1), ws.Cells(7 + rowCount, 1)).Value2
        If Not IsArray(a) Then
            v = a: ReDim a(1 To 1, 1 To 1): a(1, 1) = v
        End If
        For i = 1 To UBound(a, 1)
            If Not IsError(a(i, 1)) Then
                v = UCase$(CStr(a(i, 1)))
                Select Case v
                    Case "PASS": p = p + 1
                    Case "FAIL": f = f + 1
                    Case "ASSUMPTION": s = s + 1
                    Case "": ' blank row
                    Case Else: o = o + 1
                End Select
            End If
        Next i
    End If
    CountVerdicts = Array(p, f, o, s, p + f + o + s)
End Function

Private Sub VerdictLine(ByVal c As Range, ByVal counts As Variant)
    Dim parts As Variant, colors As Variant, t As String, i As Long, starts(0 To 4) As Long, lens(0 To 4) As Long
    parts = Array(format$(counts(4), "#,##0") & " checks", _
                  format$(counts(0), "#,##0") & " reconciled", _
                  format$(counts(1), "#,##0") & " differences", _
                  format$(counts(2), "#,##0") & " missing or to review", _
                  format$(counts(3), "#,##0") & " system-supplied assumptions")
    colors = Array(UI_INK, UI_OK, UI_BAD, UI_WARN, UI_MUTED)
    For i = 0 To 4
        If i > 0 Then t = t & "      "
        If i > 0 Then t = t & ChrW(9679) & " "
        starts(i) = Len(t) + 1: lens(i) = Len(CStr(parts(i)))
        t = t & parts(i)
    Next i
    c.Value2 = t
    c.Font.Size = 10.5: c.Font.Color = UI_TEXT_2: c.VerticalAlignment = xlCenter
    For i = 0 To 4
        With c.Characters(starts(i), lens(i)).Font
            .Bold = True: .Color = IIf(i = 0, UI_INK, UI_TEXT_2)
        End With
        If i > 0 Then c.Characters(starts(i) - 2, 1).Font.Color = colors(i)
    Next i
End Sub

' One stacked bar under the verdict line: green, red, amber, grey.
Private Sub VerdictBar(ByVal ws As Worksheet, ByVal counts As Variant)
    Dim x As Double, y As Double, w As Double, total As Double, i As Long, seg As Double, sh As Shape, colors As Variant
    total = counts(4)
    x = ws.Range("A6").Left + 2
    w = ws.Range("A6:H6").Width - 4
    y = ws.Range("A6").Top + 3
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, 6)
    sh.name = "Results_BarTrack": sh.fill.ForeColor.RGB = UI_FILL: sh.line.Visible = msoFalse: sh.Placement = xlMove
    If total <= 0 Then Exit Sub
    colors = Array(UI_OK_DOT, UI_BAD_DOT, UI_WARN_DOT, UI_LINE_2)
    For i = 0 To 3
        If counts(i) > 0 Then
            seg = w * counts(i) / total
            If seg < 2 Then seg = 2
            Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, seg, 6)
            sh.name = "Results_Bar" & i
            sh.fill.ForeColor.RGB = colors(i): sh.line.Visible = msoFalse: sh.Placement = xlMove
            x = x + seg
        End If
    Next i
End Sub

Private Sub StatusPills(ByVal rg As Range)
    Dim fc As FormatCondition
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""PASS""")
    fc.Interior.Color = UI_OK_BG: fc.Font.Color = UI_OK
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""FAIL""")
    fc.Interior.Color = UI_BAD_FILL: fc.Font.Color = UI_BAD
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""BLOCKED""")
    fc.Interior.Color = UI_WARN_FILL: fc.Font.Color = UI_WARN
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""REVIEW""")
    fc.Interior.Color = UI_WARN_BG: fc.Font.Color = UI_WARN
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""ASSUMPTION""")
    fc.Interior.Color = UI_FILL: fc.Font.Color = UI_MUTED
End Sub

Private Sub StageChips(ByVal rg As Range)
    Dim fc As FormatCondition
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlCellValue, xlNotEqual, "=""""")
    fc.Interior.Color = UI_BRAND_TINT: fc.Font.Color = UI_BRAND
End Sub

Private Sub SeverityText(ByVal rg As Range)
    Dim fc As FormatCondition
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Sev", TextOperator:=xlBeginsWith)
    fc.Font.Color = UI_BAD
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Med", TextOperator:=xlBeginsWith)
    fc.Font.Color = UI_WARN
End Sub

' Data bars on the difference: red to the left of the axis (system below the
' tool), blue to the right, both solid and without borders.
Private Sub DifferenceBars(ByVal rg As Range)
    Dim db As Databar
    On Error Resume Next
    Set db = rg.FormatConditions.AddDatabar
    If db Is Nothing Then Exit Sub
    db.BarFillType = xlDataBarFillSolid
    db.BarColor.Color = UI_UP_BAR
    db.BarBorder.Type = xlDataBarBorderNone
    db.AxisPosition = xlDataBarAxisMidpoint
    db.AxisColor.Color = UI_LINE_2
    db.NegativeBarFormat.ColorType = xlDataBarColor
    db.NegativeBarFormat.Color.Color = UI_BAD_BAR
    db.NegativeBarFormat.BorderColorType = xlDataBarSameAsPositive
    db.ShowValue = True
End Sub

' Only for workbooks whose Results was an Excel table with slicers (the first
' V15 builds): removes them once. Nothing here creates slicers any more.
Private Sub DeleteResultsSlicers(ByVal ws As Worksheet)
    Dim i As Long, sc As Object, owner As String
    On Error Resume Next
    For i = ws.Parent.SlicerCaches.count To 1 Step -1
        Set sc = ws.Parent.SlicerCaches(i)
        owner = ""
        owner = sc.ListObject.Parent.name
        If owner = ws.name Then sc.Delete
    Next i
    Err.Clear
End Sub
