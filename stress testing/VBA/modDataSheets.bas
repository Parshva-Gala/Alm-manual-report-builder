Attribute VB_Name = "modDataSheets"
Option Explicit

' ============================================================================
'  The four sheets a reviewer works in, dressed as one family (design v2):
'
'    Inputs            Pre_Shock_Sources      one card per input file
'    Test cases        Pre_Shock_Cases        one row per test case and element
'    Base & pre-shock  Derived_Values         input files beside the system
'    Value sources     Config_ValueSources    where each value comes from
'
'  Each gets the same frame: a title and a one-line purpose, a summary line that
'  says what is on the sheet and what needs attention, a navy header, white rows
'  with hairlines, status colours that follow the values, and a frozen header.
'
'  This runs after the sheet's own builder and Sheet1's layout pass, and only in
'  this workbook (Sheet1's code travels into output books, this module does
'  not). It formats only; no value is written except the summary line in the
'  frame, which sits outside every table.
' ============================================================================

Private Const HDR As Long = 7
Private Const FIRST As Long = 8
Private Const VS_NAME As String = "Config_ValueSources"
Private Const VS_HDR As Long = 9

Public Sub DressWorkSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET): If Not ws Is Nothing Then DressInputs ws
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_CASES_SHEET): If Not ws Is Nothing Then DressCases ws
    Set ws = GetWorksheetSafe(ThisWorkbook, DERIVED_SHEET): If Not ws Is Nothing Then DressBasePreShock ws
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_NAME): If Not ws Is Nothing Then DressValueSources ws
    Err.Clear
End Sub

' Dress one sheet by name, if it is one of the four.
Public Sub DressWorkSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Select Case ws.name
        Case PS_SOURCES_SHEET: DressInputs ws
        Case PS_CASES_SHEET: DressCases ws
        Case DERIVED_SHEET: DressBasePreShock ws
        Case VS_NAME: DressValueSources ws
    End Select
    Err.Clear
End Sub

' Called on every sheet activation (ThisWorkbook, late bound). Skips a sheet
' whose layout has not changed since it was last dressed, so switching sheets
' stays instant and does not reset the reader's scroll position.
Public Sub DressWorkSheetByName(ByVal sheetName As String)
    Dim ws As Worksheet, sig As String, nm As name
    Select Case sheetName
        Case PS_SOURCES_SHEET, PS_CASES_SHEET, DERIVED_SHEET, VS_NAME
        Case Else: Exit Sub
    End Select
    On Error Resume Next
    Set ws = GetWorksheetSafe(ThisWorkbook, sheetName)
    If ws Is Nothing Then Exit Sub
    sig = DressSignature(ws)
    Set nm = ws.Names("JKB_Dressed")
    If Not nm Is Nothing Then
        If nm.RefersTo = "=""" & sig & """" Then Exit Sub
    End If
    DressWorkSheet ws
    Err.Clear
End Sub

' What a builder changes when it rewrites a sheet: its extent and its header band.
Private Function DressSignature(ByVal ws As Worksheet) As String
    Dim hr As Long
    hr = IIf(ws.name = VS_NAME, VS_HDR, HDR)
    DressSignature = ws.Cells(ws.rows.count, 1).End(xlUp).row & ":" & ws.Cells(hr, ws.columns.count).End(xlToLeft).Column & ":" & _
                     ws.Cells(hr, 1).Interior.Color & ":" & Len(SafeText(ws.Range("A3").Value2)) & ":" & ws.columns(1).ColumnWidth
End Function

Private Sub MarkDressed(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Names.Add name:="JKB_Dressed", RefersTo:="=""" & DressSignature(ws) & """", Visible:=False
    Err.Clear
End Sub

' Fixed widths by header title, so no column is sized by a title or a long note.
Private Sub WidthsByHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal lastC As Long, ByVal widths As Variant, ByVal fallback As Double)
    Dim c As Long, h As String, i As Long, w As Double
    For c = 1 To lastC
        h = UCase$(SafeText(ws.Cells(r, c).Value2))
        w = fallback
        For i = 0 To UBound(widths) - 1 Step 2
            If h = UCase$(CStr(widths(i))) Then w = widths(i + 1): Exit For
        Next i
        If Len(h) > 0 Or c = 1 Then ws.columns(c).ColumnWidth = w
    Next c
End Sub

' ---------------------------------------------------------------- Inputs -----

Private Sub DressInputs(ByVal ws As Worksheet)
    Dim lastR As Long, rg As Range, fc As FormatCondition, n As Long, loaded As Long, r As Long
    Dim asOf As String, entity As String, outRows As Long
    On Error Resume Next
    EnsureSourceRows
    lastR = LastRowIn(ws, 1, FIRST)
    Frame ws, 8, "Inputs", "The files the stress test is rebuilt from. The tool reads them and never changes them."
    For r = FIRST To lastR
        If Len(SafeText(ws.Cells(r, 1).Value2)) > 0 Then
            n = n + 1
            If Left$(SafeText(ws.Cells(r, 8).Value2), 6) = "Loaded" Then loaded = loaded + 1
        End If
    Next r
    HomeSystemFacts asOf, entity, outRows
    Summary ws, Array(loaded & " of " & n & " input files loaded", HomeInputsLine(), _
                      IIf(outRows > 0, "System output " & format$(outRows, "#,##0") & " rows" & IIf(Len(asOf) > 0, " (" & asOf & ")", ""), _
                          "System output not loaded: click Load system output")), _
            Array(IIf(loaded = n And n > 0, UI_OK, UI_WARN), UI_TEXT_2, IIf(outRows > 0, UI_OK, UI_BAD))
    Table ws, 8, lastR, 30
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, 1)).Font.Bold = True
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, 1)).Font.Color = UI_INK
    ws.Range(ws.Cells(FIRST, 2), ws.Cells(lastR, 2)).Font.Color = UI_TEXT_2
    ws.Range(ws.Cells(FIRST, 3), ws.Cells(lastR, 3)).Font.Size = 8.5
    ws.Range(ws.Cells(FIRST, 3), ws.Cells(lastR, 3)).Font.Color = UI_MUTED
    ws.Range(ws.Cells(FIRST, 5), ws.Cells(lastR, 5)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(FIRST, 5), ws.Cells(lastR, 5)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, 8)).WrapText = False
    Set rg = ws.Range(ws.Cells(FIRST, 8), ws.Cells(lastR, 8))
    rg.FormatConditions.Delete
    rg.Font.Bold = True: rg.HorizontalAlignment = xlCenter
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Loaded", TextOperator:=xlBeginsWith)
    fc.Interior.Color = UI_OK_BG: fc.Font.Color = UI_OK
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Not loaded", TextOperator:=xlBeginsWith)
    fc.Interior.Color = UI_FILL: fc.Font.Color = UI_MUTED
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=LEN(" & rg.Cells(1, 1).address(False, False) & ")>0")
    fc.Interior.Color = UI_WARN_BG: fc.Font.Color = UI_WARN
    ' As-of date matching, the one setting on this sheet (J7 label, J8 value).
    With ws.Range("J7")
        .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_MUTED: .Interior.Color = UI_WHITE
        .WrapText = True: .VerticalAlignment = xlBottom
    End With
    With ws.Range("J8")
        .Interior.Color = UI_BRAND_TINT: .Font.Bold = True: .Font.Color = UI_INK
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = UI_BRAND
    End With
    WidthsByHeader ws, HDR, 8, Array("Source", 30, "What it supplies", 40, "File", 34, "Sheet", 16, "Rows", 12, _
                                     "Base scope filter", 30, "As-of date(s)", 16, "Status", 18), 16
    If ws.columns("J").ColumnWidth < 22 Then ws.columns("J").ColumnWidth = 22
    ws.columns("I").ColumnWidth = 3
    ws.rows(FIRST & ":" & Application.Max(FIRST, lastR)).RowHeight = 30
    Freeze ws, HDR, 0   ' header only: a frozen column line cut through the title and summary
    MarkDressed ws
End Sub

' ---------------------------------------------------------------- Test cases -

Private Sub DressCases(ByVal ws As Worksheet)
    Dim lastR As Long, r As Long, n As Long, nOn As Long, noFilter As Long, fams As Object, tc As String
    Dim cEn As Long, cTc As Long, cEff As Long, cMap As Long, cRows As Long, cRes As Long, lastC As Long
    Dim rg As Range, fc As FormatCondition, a1 As String
    On Error Resume Next
    lastC = 12
    cEn = HeaderCol(ws, HDR, "Enabled", lastC): cTc = HeaderCol(ws, HDR, "Test case", lastC)
    cEff = HeaderCol(ws, HDR, "Effective filter", lastC): cMap = HeaderCol(ws, HDR, "Mapping status", lastC)
    cRows = HeaderCol(ws, HDR, "Rows matched", lastC): cRes = HeaderCol(ws, HDR, "Last result", lastC)
    If cTc = 0 Then cTc = 2
    lastR = LastRowIn(ws, cTc, FIRST)
    Set fams = CreateObject("Scripting.Dictionary")
    For r = FIRST To lastR
        tc = SafeText(ws.Cells(r, cTc).Value2)
        If Len(tc) > 0 Then
            n = n + 1
            fams(HomeFamilyCode(tc)) = True
            If cEn = 0 Or UCase$(SafeText(ws.Cells(r, cEn).Value2)) <> "NO" Then
                nOn = nOn + 1
                If cEff > 0 Then
                    If Len(SafeText(ws.Cells(r, cEff).Value2)) = 0 Then noFilter = noFilter + 1
                End If
            End If
        End If
    Next r
    Frame ws, lastC, "Test cases", "One row per test case and element. Its filter picks the pre-shock population from the input files."
    Summary ws, Array(format$(n, "#,##0") & " test cases in " & fams.count & " risk families", _
                      format$(nOn, "#,##0") & " switched on", _
                      IIf(noFilter > 0, format$(noFilter, "#,##0") & " without a filter", "every one has a filter")), _
            Array(UI_INK, UI_TEXT_2, IIf(noFilter > 0, UI_WARN, UI_OK))
    Table ws, lastC, lastR, 22
    ' One line per row: long filters are cut cleanly at the column edge (the
    ' formula bar and Test filter show them whole) instead of spilling over rows.
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, lastC)).WrapText = False
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, lastC)).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(FIRST, cTc), ws.Cells(lastR, cTc)).Font.Bold = True
    ws.Range(ws.Cells(FIRST, cTc), ws.Cells(lastR, cTc)).Font.Color = UI_INK
    If cEff > 0 Then
        With ws.Range(ws.Cells(FIRST, cEff), ws.Cells(lastR, cEff))
            .Font.name = "Consolas": .Font.Size = 8.5: .Font.Color = UI_BRAND_DK: .Interior.Color = UI_BRAND_TINT
        End With
        ' A switched-on case with no effective filter is the thing to fix here.
        Set rg = ws.Range(ws.Cells(FIRST, cEff), ws.Cells(lastR, cEff))
        rg.FormatConditions.Delete
        a1 = rg.Cells(1, 1).address(False, False)
        Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=AND(LEN(" & ws.Cells(FIRST, cTc).address(False, True) & ")>0,LEN(" & a1 & ")=0)")
        fc.Interior.Color = UI_WARN_FILL
    End If
    If cEn > 0 Then
        Set rg = ws.Range(ws.Cells(FIRST, cEn), ws.Cells(lastR, cEn))
        rg.HorizontalAlignment = xlCenter: rg.Font.Bold = True: rg.Interior.Color = UI_WHITE
        rg.FormatConditions.Delete
        Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""Yes""")
        fc.Interior.Color = UI_OK_BG: fc.Font.Color = UI_OK
        Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""No""")
        fc.Interior.Color = UI_FILL: fc.Font.Color = UI_FAINT
        ' A switched-off row fades as a whole.
        Set rg = ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, lastC))
        Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=" & ws.Cells(FIRST, cEn).address(False, True) & "=""No""")
        fc.Font.Color = UI_FAINT
    End If
    If cRows > 0 Then
        ws.Range(ws.Cells(FIRST, cRows), ws.Cells(lastR, cRows)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(FIRST, cRows), ws.Cells(lastR, cRows)).HorizontalAlignment = xlRight
    End If
    If cMap > 0 Then VerdictText ws.Range(ws.Cells(FIRST, cMap), ws.Cells(lastR, cMap))
    If cRes > 0 Then VerdictText ws.Range(ws.Cells(FIRST, cRes), ws.Cells(lastR, cRes))
    WidthsByHeader ws, HDR, lastC, Array("Enabled", 9, "Test case", 16, "Test element", 14, "Element type", 14, "Source", 10, _
                                         "System filter", 34, "Override filter", 24, "Effective filter", 34, "Mapping status", 16, _
                                         "Rows matched", 11, "Last result", 16, "Notes", 24), 14
    Freeze ws, HDR, 3
    MarkDressed ws
End Sub

' ---------------------------------------------------------- Base & pre-shock -

Private Sub DressBasePreShock(ByVal ws As Worksheet)
    Dim lastC As Long, lastR As Long, cRec As Long, cB As Long, cP As Long, r As Long
    Dim nM As Long, nD As Long, nN As Long, v As String, rg As Range, fc As FormatCondition, a As Variant
    On Error Resume Next
    lastC = ws.Cells(HDR, ws.columns.count).End(xlToLeft).Column
    If lastC < 16 Then lastC = 16
    cRec = HeaderCol(ws, HDR, "Reconciles", lastC)
    cB = HeaderCol(ws, HDR, "Base difference", lastC, True)
    cP = HeaderCol(ws, HDR, "Pre-shock difference", lastC, True)
    lastR = LastRowIn(ws, 1, FIRST)
    If cRec > 0 And lastR >= FIRST Then
        a = ws.Range(ws.Cells(FIRST, cRec), ws.Cells(lastR + 1, cRec)).Value2
        For r = 1 To UBound(a, 1)
            v = SafeText(a(r, 1))
            If v = "Matches" Then
                nM = nM + 1
            ElseIf v = "Differs" Then
                nD = nD + 1
            ElseIf Len(v) > 0 Then
                nN = nN + 1
            End If
        Next r
    End If
    Frame ws, lastC, "Base & pre-shock", _
          "Base and pre-shock rebuilt from the input files beside the system values. Click Run on the home screen to refresh."
    If lastR < FIRST Or cRec = 0 Then
        Summary ws, Array("Nothing rebuilt yet"), Array(UI_MUTED)
    Else
        Summary ws, Array(format$(nM + nD + nN, "#,##0") & " values", format$(nM, "#,##0") & " match the system", _
                          format$(nD, "#,##0") & " differ", format$(nN, "#,##0") & " not comparable"), _
                Array(UI_INK, UI_OK, IIf(nD > 0, UI_BAD, UI_TEXT_2), UI_MUTED)
    End If
    Table ws, lastC, lastR, 22
    ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, lastC)).WrapText = False
    If cRec > 0 And lastR >= FIRST Then
        Set rg = ws.Range(ws.Cells(FIRST, cRec), ws.Cells(lastR, cRec))
        rg.FormatConditions.Delete
        rg.HorizontalAlignment = xlCenter: rg.Font.Bold = True: rg.Font.Size = 8.5
        Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""Matches""")
        fc.Interior.Color = UI_OK_BG: fc.Font.Color = UI_OK
        Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""Differs""")
        fc.Interior.Color = UI_BAD_FILL: fc.Font.Color = UI_BAD
        Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=LEN(" & rg.Cells(1, 1).address(False, False) & ")>0")
        fc.Interior.Color = UI_FILL: fc.Font.Color = UI_MUTED
    End If
    DiffColumn ws, cB, lastR
    DiffColumn ws, cP, lastR
    ' The builder autofits A with the title in it, which made the date column
    ' about 1,200 px wide; every column gets a fixed width by its header.
    WidthsByHeader ws, HDR, lastC, Array("As-of date", 12, "Entity", 22, "Entity ID", 12, "Test case", 16, "Element", 14, _
                                         "Metric", 30, "Source", 10, "Base (derived)", 18, "Pre-shock (derived)", 18, _
                                         "Rows matched", 12, "Filter applied", 40, "Status", 22, "Base (system)", 18, _
                                         "Pre-shock (system)", 18, "Value to use", 14, "Which one was used", 18, _
                                         "Reconciles", 14), 18
    Freeze ws, HDR, 6
    MarkDressed ws
End Sub

Private Sub DiffColumn(ByVal ws As Worksheet, ByVal c As Long, ByVal lastR As Long)
    Dim rg As Range, fc As FormatCondition
    If c = 0 Or lastR < FIRST Then Exit Sub
    Set rg = ws.Range(ws.Cells(FIRST, c), ws.Cells(lastR, c))
    rg.NumberFormat = "+#,##0.00;-#,##0.00;0.00"
    rg.Font.name = "Consolas": rg.Font.Size = 9: rg.HorizontalAlignment = xlRight
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=AND(ISNUMBER(" & rg.Cells(1, 1).address(False, False) & ")," & rg.Cells(1, 1).address(False, False) & "<>0)")
    fc.Font.Color = UI_BAD
End Sub

' ---------------------------------------------------------- Value sources ----

Private Sub DressValueSources(ByVal ws As Worksheet)
    Dim c As Long, cStat As Long, cLbl As Long, lastR As Long, r As Long, s As String
    Dim nReady As Long, nNeeds As Long, nBad As Long, nNone As Long, nUnchecked As Long
    On Error Resume Next
    For c = 1 To 12
        Select Case UCase$(SafeText(ws.Cells(VS_HDR, c).Value2))
            Case "STATUS": cStat = c
            Case "OUTPUT_ROW_LABEL": cLbl = c
        End Select
    Next c
    If cLbl = 0 Then cLbl = 1
    lastR = LastRowIn(ws, cLbl, VS_HDR + 1)
    For r = VS_HDR + 1 To lastR
        If Len(SafeText(ws.Cells(r, cLbl).Value2)) > 0 Then
            s = IIf(cStat > 0, SafeText(ws.Cells(r, cStat).Value2), "")
            If Len(s) = 0 Then
                nUnchecked = nUnchecked + 1
            ElseIf InStr(s, "Unknown") > 0 Or InStr(s, "switched off") > 0 Then
                nBad = nBad + 1
            ElseIf InStr(s, "needs") > 0 Then
                nNeeds = nNeeds + 1
            ElseIf Left$(s, 8) = "No link:" Then
                nNone = nNone + 1
            Else
                nReady = nReady + 1
            End If
        End If
    Next r
    ' The summary sits on row 8, between the two defaults and the table.
    ws.rows(8).RowHeight = 24
    If nUnchecked > 0 And nReady + nNeeds + nBad + nNone = 0 Then
        SummaryAt ws.Range("A8"), Array(format$(nUnchecked, "#,##0") & " values", "not checked yet: click Check sources"), Array(UI_INK, UI_WARN)
    Else
        SummaryAt ws.Range("A8"), Array(format$(nReady, "#,##0") & " ready", format$(nNeeds, "#,##0") & " need a file or figure", _
                                        format$(nBad, "#,##0") & " point at a missing metric", format$(nNone, "#,##0") & " keep the system value"), _
                  Array(UI_OK, IIf(nNeeds > 0, UI_WARN, UI_TEXT_2), IIf(nBad > 0, UI_BAD, UI_TEXT_2), UI_MUTED)
    End If
    ws.Range("A3").Font.Size = 20: ws.rows(3).RowHeight = 34
    ' The long note under the title wraps inside the visible columns.
    With ws.Range("A5:E5")
        If Not ws.Range("A5").MergeCells Then .Merge
        .WrapText = True: .VerticalAlignment = xlTop
    End With
    ws.rows(5).RowHeight = 30
    ' The actions sit beside the two defaults (column E), not off-screen in F:H.
    PlaceRow ws, Array("JKB_VS_Check", "JKB_VS_Load"), ws.Range("E6")
    PlaceRow ws, Array("JKB_VS_One", "JKB_VS_Figures"), ws.Range("E7")
    With ws.Range(ws.Cells(VS_HDR, 1), ws.Cells(VS_HDR, 10))
        .RowHeight = 30: .WrapText = True: .VerticalAlignment = xlCenter: .Font.Size = 9
    End With
    If lastR > VS_HDR Then
        With ws.Range(ws.Cells(VS_HDR + 1, 1), ws.Cells(lastR, 10))
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous: .Borders(xlInsideHorizontal).Color = UI_LINE
        End With
        ws.Range(ws.Cells(VS_HDR + 1, cLbl), ws.Cells(lastR, cLbl)).Font.Bold = True
        ws.Range(ws.Cells(VS_HDR + 1, cLbl), ws.Cells(lastR, cLbl)).Font.Color = UI_INK
        If cStat > 0 Then ws.Range(ws.Cells(VS_HDR + 1, cStat), ws.Cells(lastR, cStat)).Font.Size = 8.5
    End If
    Freeze ws, VS_HDR, 1
    MarkDressed ws
End Sub

Private Sub PlaceRow(ByVal ws As Worksheet, ByVal names As Variant, ByVal anchor As Range)
    Dim sh As Shape, i As Long, x As Double
    On Error Resume Next
    x = anchor.Left + 6
    For i = 0 To UBound(names)
        Set sh = Nothing: Set sh = ws.Shapes(CStr(names(i)))
        If Not sh Is Nothing Then
            sh.Left = x: sh.Top = anchor.Top + (anchor.Height - sh.Height) / 2
            x = x + sh.Width + 6
        End If
    Next i
    Err.Clear
End Sub

' ---------------------------------------------------------------- shared -----

' Title (row 3), purpose (row 4), header band (row 7). Rows 1-2 stay the toolbar's.
Private Sub Frame(ByVal ws As Worksheet, ByVal lastC As Long, ByVal title As String, ByVal purpose As String)
    Dim band As Range
    Set band = ws.Range(ws.Cells(3, 1), ws.Cells(6, lastC))
    band.Interior.Color = UI_WHITE
    band.Borders.LineStyle = xlNone
    ws.rows(3).RowHeight = 34: ws.rows(4).RowHeight = 20: ws.rows(5).RowHeight = 26: ws.rows(6).RowHeight = 8
    ' A builder's "Written ..." line on row 6 stays, small and grey.
    If ws.name <> VS_NAME And Len(SafeText(ws.Range("A6").Value2)) > 0 Then
        ws.rows(6).RowHeight = 16
        With ws.Range("A6")
            .Font.Size = 8.5: .Font.Color = UI_FAINT: .Font.Bold = False: .WrapText = False
            .VerticalAlignment = xlCenter: .IndentLevel = 1
        End With
    End If
    With ws.Range("A3")
        .Value2 = title: .Font.name = UI_FONT: .Font.Size = 20: .Font.Bold = True: .Font.Color = UI_INK
        .VerticalAlignment = xlBottom: .IndentLevel = 1
    End With
    With ws.Range("A4")
        .Value2 = purpose: .Font.name = UI_FONT: .Font.Size = 10: .Font.Bold = False: .Font.Color = UI_MUTED
        .WrapText = False: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, lastC))
        .Interior.Color = UI_NAVY: .Font.Color = UI_WHITE: .Font.Bold = True: .Font.Size = 9
        .Font.name = UI_FONT: .WrapText = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Borders.LineStyle = xlNone
    End With
    ws.rows(HDR).RowHeight = 30
    ws.Tab.Color = UI_BRAND
End Sub

' White rows with hairlines between them; nothing else.
Private Sub Table(ByVal ws As Worksheet, ByVal lastC As Long, ByVal lastR As Long, ByVal rowH As Double)
    Dim rg As Range
    If lastR < FIRST Then Exit Sub
    Set rg = ws.Range(ws.Cells(FIRST, 1), ws.Cells(lastR, lastC))
    rg.Interior.Color = UI_WHITE
    rg.Font.name = UI_FONT: rg.Font.Size = 9.5: rg.Font.Color = UI_TEXT
    rg.Borders.LineStyle = xlNone
    rg.Borders(xlInsideHorizontal).LineStyle = xlContinuous: rg.Borders(xlInsideHorizontal).Color = UI_LINE
    rg.Borders(xlEdgeBottom).LineStyle = xlContinuous: rg.Borders(xlEdgeBottom).Color = UI_LINE
    rg.IndentLevel = 1
    ws.rows(FIRST & ":" & lastR).RowHeight = rowH
End Sub

Private Sub Summary(ByVal ws As Worksheet, ByVal parts As Variant, ByVal colors As Variant)
    SummaryAt ws.Range("A5"), parts, colors
End Sub

' "part   ●  part   ●  part", each part bold in its own colour.
Private Sub SummaryAt(ByVal c As Range, ByVal parts As Variant, ByVal colors As Variant)
    Dim t As String, i As Long, starts() As Long, lens() As Long
    ReDim starts(0 To UBound(parts)): ReDim lens(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If i > 0 Then t = t & "    " & ChrW(183) & "    "
        starts(i) = Len(t) + 1: lens(i) = Len(CStr(parts(i)))
        t = t & CStr(parts(i))
    Next i
    c.WrapText = False
    c.Value2 = t
    c.Font.name = UI_FONT: c.Font.Size = 10.5: c.Font.Bold = False: c.Font.Color = UI_FAINT
    c.VerticalAlignment = xlCenter: c.IndentLevel = 1
    For i = 0 To UBound(parts)
        If lens(i) > 0 Then
            With c.Characters(starts(i), lens(i)).Font
                .Bold = True: .Color = colors(i)
            End With
        End If
    Next i
End Sub

' OK-sounding text green, problem text red, the rest amber; blanks untouched.
Private Sub VerdictText(ByVal rg As Range)
    Dim fc As FormatCondition, a1 As String
    a1 = rg.Cells(1, 1).address(False, False)
    rg.FormatConditions.Delete
    rg.Font.Size = 8.5
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=OR(LEFT(" & a1 & ",2)=""OK"",LEFT(" & a1 & ",4)=""PASS"",LEFT(" & a1 & ",7)=""Matches"",LEFT(" & a1 & ",6)=""Mapped"")")
    fc.Font.Color = UI_OK: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=OR(ISNUMBER(SEARCH(""fail""," & a1 & ")),ISNUMBER(SEARCH(""error""," & a1 & ")),ISNUMBER(SEARCH(""unknown""," & a1 & ")),LEFT(" & a1 & ",4)=""FAIL"")")
    fc.Font.Color = UI_BAD: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=LEN(" & a1 & ")>0")
    fc.Font.Color = UI_WARN
End Sub

Private Function HeaderCol(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String, ByVal lastC As Long, _
                           Optional ByVal prefix As Boolean = False) As Long
    Dim c As Long, h As String
    For c = 1 To lastC
        h = UCase$(SafeText(ws.Cells(r, c).Value2))
        If prefix Then
            If Left$(h, Len(title)) = UCase$(title) Then HeaderCol = c: Exit Function
        ElseIf h = UCase$(title) Then
            HeaderCol = c: Exit Function
        End If
    Next c
End Function

Private Function LastRowIn(ByVal ws As Worksheet, ByVal c As Long, ByVal minRow As Long) As Long
    LastRowIn = ws.Cells(ws.rows.count, c).End(xlUp).row
    If LastRowIn < minRow Then LastRowIn = minRow - 1
End Function

Private Sub Freeze(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal cols As Long)
    UiSetView ws, headerRow, cols
End Sub
