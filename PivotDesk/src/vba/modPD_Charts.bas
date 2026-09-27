Option Explicit

' ============================================================================
'  Charts: what a build draws, and where.
'
'  One row of Chart config is one chart. It is read like a pivot recipe -
'  the same field names, the same values ("Gross pre-factor sum as Exposure"),
'  the same filters and groups - because underneath it IS a pivot: each chart
'  gets its own PivotTable on the book's hidden chart sheet, over the one
'  cache every sheet shares, and is drawn as a PivotChart of it. A refresh
'  redraws every chart with every sheet.
'
'  A chart goes on Start here, in a grid of thirds, halves and full widths,
'  or on a sheet of its own. It is drawn in the Desk's colours: the dark
'  surface, hairline grid, and a palette every hue of which clears 3:1
'  against it.
' ============================================================================

' --- the table --------------------------------------------------------------
Private Const X_ON As Long = 1
Private Const X_NAME As Long = 2
Private Const X_FW As Long = 3
Private Const X_TYPE As Long = 4
Private Const X_CATS As Long = 5
Private Const X_SERIES As Long = 6
Private Const X_VALUES As Long = 7
Private Const X_FILTERS As Long = 8
Private Const X_VFILTER As Long = 9
Private Const X_GROUP As Long = 10
Private Const X_SORT As Long = 11
Private Const X_WHERE As Long = 12
Private Const X_SIZE As Long = 13
Private Const X_LABELS As Long = 14
Private Const X_LEGEND As Long = 15
Private Const X_UNITS As Long = 16
Private Const X_PALETTE As Long = 17
Private Const X_DESC As Long = 18
Private Const X_CHECK As Long = 19
Private Const X_WHY As Long = 20
Private Const X_LAST As Long = 20
Private Const CHART_ROWS As Long = 20
Private Const R_GROUPS As Long = 6

' --- Excel's chart enumerations, by value -------------------------------------
Private Const CT_COLUMN As Long = 51             ' xlColumnClustered
Private Const CT_COLUMN_STACKED As Long = 52     ' xlColumnStacked
Private Const CT_COLUMN_100 As Long = 53         ' xlColumnStacked100
Private Const CT_BAR As Long = 57                ' xlBarClustered
Private Const CT_BAR_STACKED As Long = 58        ' xlBarStacked
Private Const CT_BAR_100 As Long = 59            ' xlBarStacked100
Private Const CT_LINE As Long = 4                ' xlLine
Private Const CT_LINE_MARKERS As Long = 65       ' xlLineMarkers
Private Const CT_AREA As Long = 1                ' xlArea
Private Const CT_AREA_STACKED As Long = 76       ' xlAreaStacked
Private Const CT_PIE As Long = 5                 ' xlPie
Private Const CT_DOUGHNUT As Long = -4120        ' xlDoughnut
Private Const LEG_TOP As Long = -4160            ' xlLegendPositionTop
Private Const LEG_RIGHT As Long = -4152          ' xlLegendPositionRight
Private Const LEG_BOTTOM As Long = -4107         ' xlLegendPositionBottom
Private Const LBL_CENTER As Long = -4108         ' xlLabelPositionCenter
Private Const LBL_OUTSIDE_END As Long = 2        ' xlLabelPositionOutsideEnd
Private Const LBL_BEST_FIT As Long = 5           ' xlLabelPositionBestFit
Private Const AX_CATEGORY As Long = 1            ' xlCategory
Private Const AX_VALUE As Long = 2               ' xlValue
Private Const AX_SECONDARY As Long = 2           ' xlSecondary
Private Const TICK_LOW As Long = -4134           ' xlTickLabelPositionLow

' --- on Start here ------------------------------------------------------------
Private Const GRID_GAP As Double = 12
Private Const GUIDE_CHART_H As Double = 250
' Reserve several modest rows rather than exceeding Excel's 409.5-point row limit.
Private Const CHART_ROW_MAX As Double = 200
Private Const SHEET_CHART_W As Double = 1064
Private Const SHEET_CHART_H As Double = 520
' Each chart's pivot takes a block of columns on the hidden sheet, with room
' to grow sideways on a refresh before it would meet the next.
Private Const BLOCK_MARGIN As Long = 20

Private mGuide As Collection        ' Array(pivot, chart row) - Start here draws these
Private mSheet As Worksheet         ' the book's hidden sheet of chart pivots
Private mNext As Long               ' its next free column
Private mSeq As Long

' ===================== the sheet ============================================

Public Function ChartHeads() As Variant
    ChartHeads = Array("On", "Chart", "Frameworks", "Type", "Categories", "Series", "Values", "Show only / hide", _
                       "Top / value filter", "Group", "Sort", "Where", "Size", "Labels", "Legend", "Units", _
                       "Colours", "Description", "Check", "What to fix")
End Function

Private Function ChartWidths() As Variant
    ChartWidths = Array(7, 30, 12, 16, 16, 12, 60, 34, 22, 24, 16, 12, 12, 10, 10, 10, 11, 46, 10, 60)
End Function

Public Sub BuildChartsSheet(Optional ByVal withDefaults As Boolean = False)
    Dim ws As Worksheet, fresh As Boolean, old As Collection, ev As Boolean
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done
    Set ws = GetSheet(SH_CHARTS)
    fresh = ws Is Nothing
    Set ws = EnsureSheet(SH_CHARTS)
    If Not fresh And Not withDefaults Then
        If Not modPD_Config.HeadersMatch(ws, ChartHeads()) Then Set old = modPD_Config.Remember(ws)
    End If
    If fresh Or withDefaults Or Not old Is Nothing Then
        ws.Cells.Clear
        ws.Cells.Validation.Delete
    End If
    modPD_Theme.Dress ws, "Chart config", _
        "What a build draws. One row is one chart, read like a pivot: the same fields, values and filters. " & _
        "It goes on Start here or on a sheet of its own, and redraws whenever the book refreshes.", _
        "REPORTS  " & ChrW(183) & "  CHARTS"
    ws.Rows(R_GROUPS).RowHeight = 20
    modPD_Config.Band ws, X_ON, X_TYPE, "WHICH CHART"
    modPD_Config.Band ws, X_CATS, X_SORT, "WHAT IT PLOTS"
    modPD_Config.Band ws, X_WHERE, X_PALETTE, "HOW IT LOOKS"
    modPD_Config.Band ws, X_DESC, X_DESC, "UNDER THE TITLE"
    modPD_Config.Band ws, X_CHECK, X_WHY, "CHECK"
    modPD_Theme.Head ws, ChartHeads(), ChartWidths()
    If Not old Is Nothing Then modPD_Config.PutBack ws, old, ChartHeads()
    If fresh Or withDefaults Or Len(SafeText(ws.Cells(modPD_Theme.R_FIRST, X_NAME).Value2)) = 0 Then
        WriteDefaultCharts ws
    End If
    DressCharts ws
    ChartHints ws
    modPD_Theme.PrintReady ws, X_LAST
    CheckCharts True
Done:
    Application.EnableEvents = ev
End Sub

Private Function LastChartRow(ByVal ws As Worksheet) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.count, X_NAME).End(xlUp).Row
    b = ws.Cells(ws.Rows.count, X_CATS).End(xlUp).Row
    If b > a Then a = b
    If a < modPD_Theme.R_FIRST + CHART_ROWS - 1 Then a = modPD_Theme.R_FIRST + CHART_ROWS - 1
    LastChartRow = a
End Function

Private Sub DressCharts(ByVal ws As Worksheet)
    Dim lastR As Long
    On Error Resume Next
    lastR = LastChartRow(ws)
    modPD_Theme.DressTable ws, X_LAST, lastR
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, 1), ws.Cells(lastR, X_LAST)).NumberFormat = "@"
    With ws.Range(ws.Cells(modPD_Theme.R_FIRST, X_NAME), ws.Cells(lastR, X_NAME)).Font
        .Name = modPD_Theme.UI_SEMI
        .Color = modPD_Theme.C_TEXT
    End With
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, X_ON), ws.Cells(lastR, X_ON)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, X_DESC), ws.Cells(lastR, X_DESC)).Font.Color = modPD_Theme.C_TEXT_3
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, X_WHY), ws.Cells(lastR, X_WHY)).Font.Color = modPD_Theme.C_TEXT_2
    Err.Clear
End Sub

Private Sub ChartHints(ByVal ws As Worksheet)
    Dim r1 As Long, r2 As Long
    r1 = modPD_Theme.R_FIRST
    r2 = LastChartRow(ws) + 20
    modPD_Config.ListRule ws, r1, r2, X_ON, "Yes,No", "On", "Yes draws this chart; No keeps the row but skips it."
    modPD_Config.Hint ws, r1, r2, X_NAME, "Chart", "The chart's title. {fw} becomes the framework. On its own " & _
        "sheet, the title is the sheet's name too - 31 characters at most."
    modPD_Config.Hint ws, r1, r2, X_FW, "Frameworks", "All, or any of: LCR, NSFR, Maturity ladder."
    modPD_Config.ListRule ws, r1, r2, X_TYPE, "Column,Stacked column,100% column,Bar,Stacked bar,100% bar,Line," & _
        "Line with markers,Area,Stacked area,Pie,Doughnut,Column + line", "Type", _
        "Column + line takes two values: the first as columns, the second as a line on its own axis."
    modPD_Config.Hint ws, r1, r2, X_CATS, "Categories", "The field the chart runs along - the x axis, or the " & _
        "slices of a pie. e.g. Bucket, Counterparty, Type."
    modPD_Config.Hint ws, r1, r2, X_SERIES, "Series", "Optional. A field to split each category by - one " & _
        "colour per value, e.g. LCY / FCY. Takes a single value; not on a pie."
    modPD_Config.Hint ws, r1, r2, X_VALUES, "Values", "As on Pivot config: field, how, as caption - separated " & _
        "by ; for one series each. e.g. Gross pre-factor sum as Exposure"
    modPD_Config.Hint ws, r1, r2, X_FILTERS, "Show only / hide", "As on Pivot config: Field = a | b, " & _
        "Field <> a, Field contains text."
    modPD_Config.Hint ws, r1, r2, X_VFILTER, "Top / value filter", "As on Pivot config - Top 10 by Exposure " & _
        "keeps the ten biggest categories."
    modPD_Config.Hint ws, r1, r2, X_GROUP, "Group", "As on Pivot config - Maturity date by year."
    modPD_Config.Hint ws, r1, r2, X_SORT, "Sort", "Blank keeps the data's order (buckets by tenor). A value " & _
        "caption and desc puts the biggest first - Exposure desc - or label asc for A to Z."
    modPD_Config.ListRule ws, r1, r2, X_WHERE, "Start here,Own sheet", "Where", _
        "Start here: in the grid under the book's figures. Own sheet: a sheet of its own, full size."
    modPD_Config.ListRule ws, r1, r2, X_SIZE, "Third,Half,Two thirds,Full", "Size", _
        "On Start here: how much of the row it takes. Charts fill each row left to right."
    modPD_Config.ListRule ws, r1, r2, X_LABELS, "None,Values,Percent,Both", "Labels", _
        "Values on the bars or slices. Percent and Both are for a pie or a doughnut."
    modPD_Config.ListRule ws, r1, r2, X_LEGEND, "Auto,Top,Right,Bottom,None", "Legend", _
        "Auto: none for one series, at the top for several, on the right for a pie."
    modPD_Config.ListRule ws, r1, r2, X_UNITS, "Auto,As is,Thousands,Millions,Billions", "Units", _
        "The axis and the labels in thousands (k), millions (m) or billions (bn). Auto: whichever reads best for " & _
        "the biggest figure plotted."
    modPD_Config.ListRule ws, r1, r2, X_PALETTE, "Avati,Emerald,Two tone", "Colours", _
        "Avati: emerald, sky, amber, mint, coral, violet... Emerald: shades of the brand. Two tone: emerald and sky."
    modPD_Config.Hint ws, r1, r2, X_DESC, "Description", "On its own sheet, under the title; on Start here, " & _
        "in the index."
    modPD_Config.Hint ws, r1, r2, X_CHECK, "Check", "Written by Check. OK, or Break with the reason beside it."
End Sub

' ===================== defaults =============================================

Private Sub WriteDefaultCharts(ByVal ws As Worksheet)
    Dim r As Long
    r = modPD_Theme.R_FIRST
    ws.Range(ws.Cells(r, 1), ws.Cells(LastChartRow(ws), X_LAST)).ClearContents
    ChartRec ws, r, "On=Yes", "Chart=Where the balance sits", "Frameworks=All", "Type=Doughnut", _
        "Categories=Type", "Values=Gross pre-factor sum as Gross", "Where=Start here", "Size=Third", _
        "Labels=Percent", "Legend=Right", "Units=Auto", "Colours=Avati", _
        "Description=The gross book by balance-sheet type."
    ChartRec ws, r + 1, "On=Yes", "Chart=Largest counterparties", "Frameworks=All", "Type=Bar", _
        "Categories=Counterparty", "Values=Gross pre-factor sum as Exposure", _
        "Show only / hide=Counterparty <> (no counterparty)", "Top / value filter=Top 10 by Exposure", _
        "Sort=Exposure desc", "Where=Start here", "Size=Two thirds", "Labels=Values", "Legend=None", _
        "Units=Auto", "Colours=Avati", _
        "Description=The ten biggest counterparties, gross."
    ChartRec ws, r + 2, "On=Yes", "Chart=What the factors kept, by rule", "Frameworks=All", "Type=Bar", _
        "Categories=Rule name", "Values=Gross pre-factor sum as Read; Gross post-factor sum as Kept", _
        "Show only / hide=Rule name <> (no rule)", "Top / value filter=Top 8 by Read", "Sort=Read desc", _
        "Where=Start here", "Size=Full", "Legend=Top", "Labels=Values", "Units=Auto", "Colours=Two tone", _
        "Description=The eight rules that read the most, and how much of it their factors kept."
    ' Ready to switch on.
    ChartRec ws, r + 3, "On=No", "Chart=Across the buckets", "Frameworks=All", "Type=Stacked column", _
        "Categories=Bucket", "Series=LCY / FCY", "Values=Pre factor amount sum as Pre-factor", _
        "Show only / hide=Bucket <> (no bucket)", "Where=Start here", "Size=Full", "Legend=Top", _
        "Units=Auto", "Colours=Two tone", "Description=Balances in each maturity bucket, local and foreign."
    ChartRec ws, r + 4, "On=No", "Chart=Maturing by year", "Frameworks=All", "Type=Column + line", _
        "Categories=Maturity date", "Group=Maturity date by year", _
        "Values=Gross pre-factor sum as Maturing; Gross pre-factor %running in Maturity date as Cumulative", _
        "Where=Own sheet", "Legend=Top", "Units=Auto", "Colours=Two tone", _
        "Description=How much matures each year, and the running share of the book."
    ChartRec ws, r + 5, "On=No", "Chart=Currency mix", "Frameworks=All", "Type=Pie", "Categories=Currency", _
        "Values=Gross pre-factor sum as Gross", "Top / value filter=Top 8 by Gross", "Sort=Gross desc", _
        "Where=Start here", "Size=Third", "Labels=Percent", "Legend=Right", "Colours=Avati", _
        "Description=The eight largest currencies by gross amount."
    ChartRec ws, r + 6, "On=No", "Chart=Products across the buckets", "Frameworks=All", "Type=100% column", _
        "Categories=Bucket", "Series=Product", "Values=Gross pre-factor sum as Gross", _
        "Show only / hide=Bucket <> (no bucket)", "Where=Own sheet", "Legend=Right", "Colours=Avati", _
        "Description=Each bucket's mix of products."
    ChartRec ws, r + 7, "On=No", "Chart=Sectors", "Frameworks=All", "Type=Doughnut", "Categories=Sector", _
        "Values=Gross pre-factor sum as Gross", "Top / value filter=Top 8 by Gross", "Sort=Gross desc", _
        "Where=Start here", "Size=Third", "Labels=Percent", "Legend=Right", "Colours=Emerald", _
        "Description=The eight largest sectors by gross amount."
End Sub

' A chart row written by heading, as Pivot config's are.
Private Sub ChartRec(ByVal ws As Worksheet, ByVal r As Long, ParamArray kv() As Variant)
    Dim heads As Variant, i As Long, c As Long, p As Long, k As String
    heads = ChartHeads()
    ws.Range(ws.Cells(r, 1), ws.Cells(r, X_LAST)).NumberFormat = "@"
    For i = 0 To UBound(kv)
        p = InStr(CStr(kv(i)), "=")
        If p > 1 Then
            k = Left$(CStr(kv(i)), p - 1)
            For c = 0 To UBound(heads)
                If StrComp(CStr(heads(c)), k, vbTextCompare) = 0 Then
                    ws.Cells(r, c + 1).Value2 = Mid$(CStr(kv(i)), p + 1)
                    Exit For
                End If
            Next c
        End If
    Next i
End Sub

Public Sub PD_ChartsDefaults()
    If PD_Busy Then Exit Sub
    If MsgBox("Put Chart config back to its defaults?" & vbCrLf & vbCrLf & "Every chart row is replaced.", _
              vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    BuildChartsSheet True
    modPD_Theme.Rail GetSheet(SH_CHARTS)
    Application.ScreenUpdating = True
    CheckCharts
End Sub

' A new row, switched off, holding a working chart to change.
Public Sub PD_ChartsAdd()
    Dim ws As Worksheet, r As Long, lastR As Long, ev As Boolean
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    Set ws = GetSheet(SH_CHARTS)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    lastR = LastChartRow(ws)
    For r = modPD_Theme.R_FIRST To lastR + 1
        If Len(Cell(ws, r, X_NAME)) = 0 And Len(Cell(ws, r, X_CATS)) = 0 And Len(Cell(ws, r, X_VALUES)) = 0 Then Exit For
    Next r
    ev = Application.EnableEvents
    Application.EnableEvents = False
    ChartRec ws, r, "On=No", "Chart=New chart " & (r - modPD_Theme.R_FIRST + 1), "Frameworks=All", "Type=Column", _
        "Categories=Type", "Values=Gross pre-factor sum as Gross", "Where=Start here", "Size=Half", _
        "Units=Auto", "Colours=Avati", "Description=A new chart. Change any cell, then set On to Yes."
    If r > lastR Then
        DressCharts ws
        ChartHints ws
    End If
    CheckCharts True
    Application.EnableEvents = ev
    ws.Activate
    ws.Cells(r, X_NAME).Select
    Notify "A new chart is on row " & r & ", switched off. Change what you need, then set On to Yes.", V_OK
    Err.Clear
End Sub

' ===================== reading ==============================================

Private Function Cell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Cell = SafeText(ws.Cells(r, c).Value2)
End Function

Public Function AllCharts() As Collection
    Dim ws As Worksheet, out As Collection, r As Long, fl As Object
    Set out = New Collection
    Set AllCharts = out
    Set ws = GetSheet(SH_CHARTS)
    If ws Is Nothing Then Exit Function
    Set fl = modPD_Config.Fields()
    For r = modPD_Theme.R_FIRST To LastChartRow(ws)
        If Len(Cell(ws, r, X_NAME)) > 0 Or Len(Cell(ws, r, X_CATS)) > 0 Then out.Add ParseChart(ws, r, fl)
    Next r
End Function

' The charts a build of this framework draws: on, for it, and able to.
Public Function ChartsFor(ByVal fw As String) As Collection
    Dim out As Collection, rc As Object, t As Variant, n As String
    Set out = New Collection
    For Each rc In AllCharts()
        If CBool(rc("On")) And Len(rc("Problem")) = 0 Then
            For Each t In modPD_Recipe.SplitList(CStr(rc("Frameworks")), ",")
                n = modPD_Recipe.NormFw(CStr(t))
                If n = "ALL" Or StrComp(n, fw, vbTextCompare) = 0 Then
                    out.Add rc
                    Exit For
                End If
            Next t
        End If
    Next rc
    Set ChartsFor = out
End Function

' A chart row as the pivot builder reads a recipe - Categories are its rows,
' Series its columns - with what only a chart has beside it.
Private Function ParseChart(ByVal ws As Worksheet, ByVal r As Long, ByVal fl As Object) As Object
    Dim rc As Object, prob As String, g As Object, gOf As Object, s As String
    Set rc = NewMap()
    Set ParseChart = rc
    rc("Chart") = True
    rc("Row") = r
    rc("On") = (StrComp(Cell(ws, r, X_ON), "No", vbTextCompare) <> 0 And Len(Cell(ws, r, X_ON)) > 0)
    rc("Name") = Cell(ws, r, X_NAME)
    rc("Frameworks") = IIf(Len(Cell(ws, r, X_FW)) = 0, "All", Cell(ws, r, X_FW))
    Set rc("Split") = New Collection
    Set rc("Slicers") = New Collection
    Set rc("Rows") = modPD_Recipe.SplitList(Cell(ws, r, X_CATS), ",")
    Set rc("Cols") = modPD_Recipe.SplitList(Cell(ws, r, X_SERIES), ",")
    Set rc("Values") = modPD_Recipe.ParseValues(Cell(ws, r, X_VALUES), fl, prob)
    Set rc("Filters") = modPD_Recipe.ParseFilters(Cell(ws, r, X_FILTERS), prob)
    Set rc("VFilters") = modPD_Recipe.ParseVFilters(Cell(ws, r, X_VFILTER), prob)
    Set rc("Groups") = modPD_Recipe.ParseGroups(Cell(ws, r, X_GROUP), prob)
    Set gOf = NewMap()
    For Each g In rc("Groups")
        gOf(CStr(g("Field"))) = CStr(g("Name"))
    Next g
    Set rc("GroupOf") = gOf
    modPD_Recipe.ParseSort Cell(ws, r, X_SORT), rc
    ' What a recipe has and a chart does not, so the recipe checks read it.
    Set rc("Widths") = NewMap()
    Set rc("SubFields") = NewMap()
    rc("SubAll") = False
    rc("ExpandTo") = ""
    rc("Hilite") = ""
    rc("HiliteOn") = ""
    rc("Tab") = ""
    rc("ValuesInRows") = False
    rc("Desc") = Cell(ws, r, X_DESC)

    ' Auto, the default: the unit is chosen when the chart is drawn, from the
    ' biggest figure it plots.
    s = LCase$(Cell(ws, r, X_UNITS))
    Select Case s
        Case "", "auto": rc("Units") = "auto"
        Case "as is": rc("Units") = ""
        Case "thousands", "millions", "billions": rc("Units") = s
        Case Else
            rc("Units") = "auto"
            If Len(prob) = 0 Then prob = "Units must be Auto, As is, Thousands, Millions or Billions."
    End Select
    rc("Format") = modPD_Recipe.UnitFormat(CStr(rc("Units")))
    rc("FormatSet") = False

    rc("Type") = LCase$(Trim$(Cell(ws, r, X_TYPE)))
    If Len(rc("Type")) = 0 Then rc("Type") = "column"
    Pick rc, "Where", Cell(ws, r, X_WHERE), "start here|own sheet", "start here", "Where must be Start here or Own sheet.", prob
    Pick rc, "Size", Cell(ws, r, X_SIZE), "third|half|two thirds|full", "half", _
         "Size must be Third, Half, Two thirds or Full.", prob
    Pick rc, "Labels", Cell(ws, r, X_LABELS), "none|values|percent|both", "none", _
         "Labels must be None, Values, Percent or Both.", prob
    Pick rc, "Legend", Cell(ws, r, X_LEGEND), "auto|top|right|bottom|none", "auto", _
         "Legend must be Auto, Top, Right, Bottom or None.", prob
    Pick rc, "Palette", Cell(ws, r, X_PALETTE), "avati|emerald|two tone", "avati", _
         "Colours must be Avati, Emerald or Two tone.", prob

    If Len(prob) = 0 Then prob = ChartProblem(rc)
    If Len(prob) = 0 Then prob = InChartWords(modPD_Recipe.Validate(rc, fl))
    rc("Problem") = prob
End Function

' One of a list of words, or the blank's default; anything else is a problem.
Private Sub Pick(ByVal rc As Object, ByVal key As String, ByVal s As String, ByVal allowed As String, _
                 ByVal dflt As String, ByVal msg As String, ByRef prob As String)
    Dim w As Variant
    s = LCase$(Trim$(s))
    rc(key) = dflt
    If Len(s) = 0 Then Exit Sub
    For Each w In Split(allowed, "|")
        If s = CStr(w) Then rc(key) = s: Exit Sub
    Next w
    If Len(prob) = 0 Then prob = msg
End Sub

' What a chart needs that a pivot does not.
Private Function ChartProblem(ByVal rc As Object) As String
    Dim t As String
    t = CStr(rc("Type"))
    If TypeCode(t) = 0 Then
        ChartProblem = "Type must be Column, Stacked column, 100% column, Bar, Stacked bar, 100% bar, Line, " & _
                       "Line with markers, Area, Stacked area, Pie, Doughnut or Column + line."
    ElseIf Len(rc("Name")) = 0 Then
        ChartProblem = "Give the chart a title."
    ElseIf rc("Rows").count = 0 Then
        ChartProblem = "Name a field under Categories - what the chart runs along."
    ElseIf rc("Rows").count > 2 Then
        ChartProblem = "Categories takes one field, or two for a nested axis."
    ElseIf rc("Cols").count > 1 Then
        ChartProblem = "Series takes one field."
    ElseIf rc("Values").count = 0 Then
        ChartProblem = "Add a value - what the chart measures."
    ElseIf rc("Cols").count = 1 And rc("Values").count > 1 Then
        ChartProblem = "With a Series field, give one value - the series are its values."
    ElseIf IsRound(t) And (rc("Values").count > 1 Or rc("Cols").count > 0) Then
        ChartProblem = "A " & t & " shows one value, with no Series."
    ElseIf t = "column + line" And (rc("Values").count <> 2 Or rc("Cols").count > 0) Then
        ChartProblem = "Column + line takes two values - the columns, then the line - and no Series."
    ElseIf (rc("Labels") = "percent" Or rc("Labels") = "both") And Not IsRound(t) Then
        ChartProblem = "Percent labels are for a pie or a doughnut - use Values, or a 100% chart."
    ElseIf rc("Where") = "own sheet" And Len(Replace(rc("Name"), "{fw}", "Maturity Ladder")) > 31 Then
        ChartProblem = "On its own sheet the title is the sheet's name - 31 characters at most."
    End If
End Function

' A recipe check's words, said of a chart.
Private Function InChartWords(ByVal msg As String) As String
    msg = Replace(msg, "Rows or Columns", "Categories or Series")
    msg = Replace(msg, "one of the Rows", "one of the Categories")
    msg = Replace(msg, "in Rows or Columns", "in Categories or Series")
    InChartWords = msg
End Function

Private Function TypeCode(ByVal t As String) As Long
    Select Case t
        Case "column", "column + line": TypeCode = CT_COLUMN
        Case "stacked column": TypeCode = CT_COLUMN_STACKED
        Case "100% column": TypeCode = CT_COLUMN_100
        Case "bar": TypeCode = CT_BAR
        Case "stacked bar": TypeCode = CT_BAR_STACKED
        Case "100% bar": TypeCode = CT_BAR_100
        Case "line": TypeCode = CT_LINE
        Case "line with markers": TypeCode = CT_LINE_MARKERS
        Case "area": TypeCode = CT_AREA
        Case "stacked area": TypeCode = CT_AREA_STACKED
        Case "pie": TypeCode = CT_PIE
        Case "doughnut": TypeCode = CT_DOUGHNUT
    End Select
End Function

Private Function IsRound(ByVal t As String) As Boolean
    IsRound = (t = "pie" Or t = "doughnut")
End Function

Private Function IsLineType(ByVal t As String) As Boolean
    IsLineType = (t = "line" Or t = "line with markers")
End Function

' ===================== Check ================================================

Public Function CheckCharts(Optional ByVal quiet As Boolean = False) As Long
    Dim ws As Worksheet, rc As Object, nOn As Long, nBad As Long, r As Long, lastR As Long
    Set ws = GetSheet(SH_CHARTS)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    lastR = LastChartRow(ws)
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, X_CHECK), ws.Cells(lastR, X_WHY)).ClearContents
    For Each rc In AllCharts()
        r = CLng(rc("Row"))
        If Not CBool(rc("On")) Then
            ws.Cells(r, X_CHECK).Value2 = "Off"
            If Len(rc("Problem")) > 0 Then ws.Cells(r, X_WHY).Value2 = "When switched on: " & rc("Problem")
        ElseIf Len(rc("Problem")) > 0 Then
            nOn = nOn + 1
            nBad = nBad + 1
            ws.Cells(r, X_CHECK).Value2 = V_BREAK
            ws.Cells(r, X_WHY).Value2 = rc("Problem")
        Else
            nOn = nOn + 1
            ws.Cells(r, X_CHECK).Value2 = V_OK
        End If
    Next rc
    For r = modPD_Theme.R_FIRST To lastR
        If Len(SafeText(ws.Cells(r, X_CHECK).Value2)) > 0 Then
            modPD_Theme.PaintVerdict ws.Cells(r, X_CHECK)
        Else
            ws.Cells(r, X_CHECK).Interior.Color = IIf((r - modPD_Theme.R_FIRST) Mod 2 = 1, modPD_Theme.C_ROW_ALT, _
                                                      modPD_Theme.C_ROW)
        End If
    Next r
    If nBad > 0 Then
        modPD_Theme.SetStatus ws, nBad & " of " & nOn & " chart(s) that are on will not draw - the reason is " & _
            "beside each, under What to fix.", "Break"
    ElseIf nOn = 0 Then
        modPD_Theme.SetStatus ws, "No chart is switched on - the books will have pivots only.", "Idle"
    Else
        modPD_Theme.SetStatus ws, nOn & " chart(s) switched on, every one ready to draw.", "OK"
    End If
    If Not quiet Then
        If nBad = 0 Then
            Notify "Every chart that is on is ready to draw.", V_OK
        Else
            Notify nBad & " chart(s) will not draw - see What to fix beside each.", V_BREAK
        End If
    End If
    SettingSet "charts_on", CStr(nOn)
    CheckCharts = nBad
    Err.Clear
End Function

Public Sub PD_ChartsCheck()
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    CheckCharts
End Sub

' The status-bar line for a row just edited: read, never written.
Public Function ChartRowSays(ByVal r As Long) As String
    Dim ws As Worksheet, rc As Object
    Set ws = GetSheet(SH_CHARTS)
    If ws Is Nothing Then Exit Function
    If Len(Cell(ws, r, X_NAME)) = 0 And Len(Cell(ws, r, X_CATS)) = 0 Then
        ChartRowSays = "Row " & r & " is empty."
        Exit Function
    End If
    Set rc = ParseChart(ws, r, modPD_Config.Fields())
    If Not CBool(rc("On")) Then
        ChartRowSays = "Row " & r & " is off"
        If Len(rc("Problem")) > 0 Then
            ChartRowSays = ChartRowSays & ". When switched on it will not draw: " & rc("Problem")
        End If
    ElseIf Len(rc("Problem")) > 0 Then
        ChartRowSays = "Row " & r & " will not draw: " & rc("Problem")
    Else
        ChartRowSays = "Row " & r & " is ready to draw."
    End If
End Function

' ===================== drawing ==============================================

' In a book being built, after its pivot sheets: each chart's pivot on the
' hidden chart sheet; a chart with a sheet of its own drawn now, one for
' Start here kept for the guide to place. Returns the sheets it added.
Public Function MakeCharts(ByVal wb As Workbook, ByVal fw As String, ByVal charts As Collection) As Long
    Dim rc As Object, pt As PivotTable, n As Long, i As Long
    Set mGuide = New Collection
    Set mSheet = Nothing
    mNext = 1
    mSeq = 0
    If charts Is Nothing Then Exit Function
    For Each rc In charts
        i = i + 1
        Progress_ FwLabel(fw) & " - chart: " & Replace(CStr(rc("Name")), "{fw}", FwLabel(fw)), i / (charts.count + 1)
        Set pt = TryPivot(wb, rc, fw)
        If Not pt Is Nothing Then
            If rc("Where") = "own sheet" Then
                If DrawOnSheet(wb, fw, pt, rc) Then n = n + 1
            Else
                mGuide.Add Array(pt, rc)
            End If
        End If
    Next rc
    MakeCharts = n
End Function

' A chart's pivot, fenced: one Excel refuses is logged, and the book is built
' without that chart.
Private Function TryPivot(ByVal wb As Workbook, ByVal rc As Object, ByVal fw As String) As PivotTable
    Dim pt As PivotTable, rng As Range
    On Error GoTo Failed
    If mSheet Is Nothing Then
        Set mSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        mSheet.Name = SH_CHART
        mSheet.visible = xlSheetHidden
    End If
    Set pt = modPD_Pivot.ChartPivot(mSheet, rc, mNext)
    If pt Is Nothing Then Err.Raise vbObjectError + 520, "TryPivot", "Excel would not make its pivot."
    Set rng = pt.TableRange2
    mNext = rng.Column + rng.Columns.count + BLOCK_MARGIN
    Set TryPivot = pt
    Exit Function
Failed:
    LogIt V_BREAK, "Charts", "Chart " & Chr$(34) & Replace(CStr(rc("Name")), "{fw}", FwLabel(fw)) & Chr$(34) & _
          " (Chart config row " & rc("Row") & ") was not drawn - " & Err.Description, FwLabel(fw)
    Err.Clear
End Function

' A sheet of its own: the book's bar and title, then the chart, full size.
Private Function DrawOnSheet(ByVal wb As Workbook, ByVal fw As String, ByVal pt As PivotTable, _
                             ByVal rc As Object) As Boolean
    Dim ws As Worksheet, title As String, u As String, why As String, alerts As Boolean
    On Error GoTo Failed
    title = Replace(CStr(rc("Name")), "{fw}", FwLabel(fw))
    u = UnitsFor(pt, rc)
    Set ws = modPD_Pivot.NewBookSheet(wb, SafeSheetName(title), title, CStr(rc("Desc")), _
        UCase$(FwLabel(fw)) & "  " & ChrW(183) & "  CHART" & IIf(Len(u) > 0, "  " & ChrW(183) & "  IN " & UCase$(u), ""))
    If ws Is Nothing Then Exit Function
    ReserveChartRows ws, 4, SHEET_CHART_H + 16
    If DrawChart(ws, pt, rc, title, 14, ws.Rows(4).Top + 8, SHEET_CHART_W, SHEET_CHART_H) Is Nothing Then
        Err.Raise vbObjectError + 521, "DrawOnSheet", "Excel could not draw the chart."
    End If
    modPD_Pivot.NoteMade ws, IIf(Len(rc("Desc")) > 0, CStr(rc("Desc")), title)
    On Error Resume Next
    ws.Tab.Color = modPD_Theme.C_BRAND
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.Zoom = 100
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    modPD_Theme.PrintReady ws, 1
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .BlackAndWhite = False
    End With
    Application.PrintCommunication = True
    Err.Clear
    DrawOnSheet = True
    Exit Function
Failed:
    why = Err.Description
    On Error Resume Next
    ' A failed own-sheet chart must not leave a dressed but unindexed orphan tab.
    If Not ws Is Nothing Then
        alerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = alerts
    End If
    LogIt V_BREAK, "Charts", "Chart " & Chr$(34) & title & Chr$(34) & " was not drawn - " & why, FwLabel(fw)
    Err.Clear
End Function

' The unit a chart reads in: the one asked for, or - for Auto - the one its
' biggest figure reads best in, so 93,322,000,000 is 93.3 bn and not 93,322.2 m.
Private Function UnitsFor(ByVal pt As PivotTable, ByVal rc As Object) As String
    Dim v As Variant, x As Variant, m As Double
    UnitsFor = CStr(rc("Units"))
    If UnitsFor <> "auto" Then Exit Function
    UnitsFor = ""
    On Error Resume Next
    v = pt.DataBodyRange.Value2
    If IsArray(v) Then
        For Each x In v
            If IsNumeric(x) And Not IsEmpty(x) Then
                If Abs(CDbl(x)) > m Then m = Abs(CDbl(x))
            End If
        Next x
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        m = Abs(CDbl(v))
    End If
    If m >= 1000000000# Then
        UnitsFor = "billions"
    ElseIf m >= 1000000# Then
        UnitsFor = "millions"
    ElseIf m >= 1000# Then
        UnitsFor = "thousands"
    End If
    Err.Clear
End Function

' How many charts Start here has waiting.
Public Function GuideCharts() As Long
    If mGuide Is Nothing Then Exit Function
    GuideCharts = mGuide.count
End Function

' Start here uses a two-pass layout. All worksheet row heights are final before
' any free-floating chart or background is created. Returns the first unused
' worksheet row, so the caller's report index cannot overlap the final chart.
Public Function PlaceOnGuide(ByVal ws As Worksheet, ByVal r As Long, ByVal wide As Double, _
                             ByVal fw As String) As Long
    Dim sizes As Collection, boxes As Collection, e As Variant, box As Variant, i As Long
    Dim title As String, origin As Double, bottom As Double, pt As PivotTable
    PlaceOnGuide = r
    ClearChartPanels ws
    If GuideCharts() = 0 Then Exit Function
    Set sizes = New Collection
    For Each e In mGuide
        Set pt = e(0)
        If Not pt.Parent.Parent Is ws.Parent Then
            Err.Raise vbObjectError + 522, "PlaceOnGuide", "Chart source belongs to a different workbook."
        End If
        sizes.Add CStr(e(1)("Size"))
    Next e
    Set boxes = ChartGrid(sizes, wide, GUIDE_CHART_H)
    box = boxes(boxes.count)
    bottom = CDbl(box(1)) + CDbl(box(3)) + GRID_GAP / 2
    PlaceOnGuide = ReserveChartRows(ws, r, bottom)
    origin = ws.Rows(r).Top
    For i = 1 To mGuide.count
        e = mGuide(i): box = boxes(i)
        title = Replace(CStr(e(1)("Name")), "{fw}", FwLabel(fw))
        DrawChart ws, e(0), e(1), title, 14 + CDbl(box(0)), origin + CDbl(box(1)), CDbl(box(2)), CDbl(box(3))
    Next i
End Function

' Pure point-coordinate plan, also used by focused layout checks. The 12-point
' gutters are accounted for in widths, so thirds/halves fill a row exactly.
' Input order is preserved; an item never backfills an earlier incomplete row.
Public Function ChartGrid(ByVal sizes As Collection, ByVal wide As Double, _
                          Optional ByVal high As Double = 250) As Collection
    Dim boxes As Collection, size As Variant, x As Double, y As Double, w As Double
    If wide <= 72 Or high <= 12 Then
        Err.Raise vbObjectError + 523, "ChartGrid", "The chart area is too small."
    End If
    Set boxes = New Collection
    y = GRID_GAP / 2
    For Each size In sizes
        w = WidthFor(CStr(size), wide)
        If x > 0 And x + w > wide + 0.01 Then
            x = 0
            y = y + high + GRID_GAP
        End If
        boxes.Add Array(x, y, w, high)
        x = x + w + GRID_GAP
    Next size
    Set ChartGrid = boxes
End Function

' Preallocate physical space in rows below Excel's 409.5-point ceiling. Read
' back the actual height because Excel rounds to its display grid. At least
' one point avoids a sub-pixel final remainder becoming a zero-height row.
Private Function ReserveChartRows(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal high As Double) As Long
    Dim r As Long, remaining As Double, h As Double, actual As Double
    r = firstRow: remaining = high
    Do While remaining > 0.01
        If r > ws.Rows.count Then Err.Raise vbObjectError + 524, "ReserveChartRows", "No space remains for charts."
        h = remaining
        If h > CHART_ROW_MAX Then h = CHART_ROW_MAX
        If h < 1 Then h = 1
        ws.Rows(r).RowHeight = h
        actual = ws.Rows(r).Height
        If actual <= 0 Then Err.Raise vbObjectError + 525, "ReserveChartRows", "Excel could not reserve chart space."
        remaining = remaining - actual
        r = r + 1
    Loop
    ReserveChartRows = r
End Function

' Idempotent guide redraw: remove only our charts and their matching panels.
' The report hero, KPI cards, other shapes and source PivotTables stay intact.
Private Sub ClearChartPanels(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.ChartObjects.count To 1 Step -1
        If Left$(ws.ChartObjects(i).Name, 10) = "pdc_chart_" Then ws.ChartObjects(i).Delete
    Next i
    For i = ws.Shapes.count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 12) = "pdc_surface_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Function WidthFor(ByVal size As String, ByVal wide As Double) As Double
    Select Case size
        Case "third": WidthFor = (wide - 2 * GRID_GAP) / 3
        Case "two thirds": WidthFor = (wide - 2 * GRID_GAP) / 3 * 2 + GRID_GAP
        Case "full": WidthFor = wide
        Case Else: WidthFor = (wide - GRID_GAP) / 2
    End Select
End Function

' The chart itself: a PivotChart of pt, in the Desk's colours.
Private Function DrawChart(ByVal host As Worksheet, ByVal pt As PivotTable, ByVal rc As Object, _
                           ByVal title As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, _
                           ByVal h As Double) As Object
    Dim co As Object, ch As Object, typ As String, n As Long, u As String, card As Shape, why As String
    u = UnitsFor(pt, rc)
    On Error GoTo Failed
    mSeq = mSeq + 1
    Set card = modPD_Theme.SurfaceCard(host, "pdc_surface_" & mSeq, l, t, w, h)
    If Not card Is Nothing Then
        card.Shadow.visible = msoFalse
        card.AlternativeText = title & " chart panel"
    End If
    Set co = host.ChartObjects.Add(l + 8, t + 6, w - 16, h - 12)
    If co Is Nothing Then
        If Not card Is Nothing Then card.Delete
        Exit Function
    End If
    co.Name = "pdc_chart_" & mSeq
    co.Placement = xlFreeFloating
    co.PrintObject = True
    Set ch = co.Chart
    ch.SetSourceData pt.TableRange1
    typ = CStr(rc("Type"))
    ch.ChartType = TypeCode(typ)
    ch.ShowAllFieldButtons = False
    StyleSurface ch, title
    n = ch.SeriesCollection.count
    If typ = "column + line" And n >= 2 Then
        ch.SeriesCollection(2).ChartType = CT_LINE_MARKERS
        ch.SeriesCollection(2).AxisGroup = AX_SECONDARY
    End If
    If Not IsRound(typ) Then StyleAxes ch, rc, u
    StyleColours ch, rc
    StyleLegend ch, rc, n
    StyleLabels ch, rc, u
    If n = 0 Then
        LogIt V_CHECK, "Charts", "Chart " & Chr$(34) & title & Chr$(34) & " has nothing to plot - its pivot is empty.", ""
    End If
    ' Native chart binding/type changes can adjust its initial geometry.
    ' Reapply the planned rectangle only after the PivotChart has been styled.
    co.Placement = xlFreeFloating
    co.Left = l + 8: co.Top = t + 6
    co.Width = w - 16: co.Height = h - 12
    If Not card Is Nothing Then
        card.Placement = xlFreeFloating
        card.Left = l: card.Top = t
        card.Width = w: card.Height = h
    End If
    Set DrawChart = co
    Exit Function
Failed:
    why = Err.Description
    On Error Resume Next
    If Not co Is Nothing Then co.Delete
    If Not card Is Nothing Then card.Delete
    LogIt V_BREAK, "Charts", "Chart " & Chr$(34) & title & Chr$(34) & " was not drawn - " & why, ""
    Err.Clear
End Function

' The card: the Desk's surface, a hairline edge, the title at the top left.
Private Sub StyleSurface(ByVal ch As Object, ByVal title As String)
    On Error Resume Next
    With ch.ChartArea
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = modPD_Theme.C_SURFACE
        .Format.Line.Visible = msoFalse
        .Font.Name = modPD_Theme.UI_FONT
        .Font.Size = 8.5
        .Font.Color = modPD_Theme.C_TEXT_2
    End With
    ch.PlotArea.Format.Fill.Visible = msoFalse
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    With ch.ChartTitle.Font
        .Name = modPD_Theme.UI_SEMI
        .Size = 12
        .Bold = False
        .Color = modPD_Theme.C_TEXT
    End With
    ch.ChartTitle.Left = 10
    ch.ChartTitle.Top = 6
    Err.Clear
End Sub

' A hairline grid, no axis line, figures in the units asked for.
Private Sub StyleAxes(ByVal ch As Object, ByVal rc As Object, ByVal u As String)
    Dim typ As String, v As Object
    typ = CStr(rc("Type"))
    On Error Resume Next
    With ch.Axes(AX_VALUE)
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = modPD_Theme.C_HAIRLINE
        .MajorGridlines.Format.Line.Weight = 0.5
        .Format.Line.Visible = msoFalse
        .MajorTickMark = xlNone
        .TickLabels.Font.Name = modPD_Theme.UI_FONT
        .TickLabels.Font.Size = 8.5
        .TickLabels.Font.Color = modPD_Theme.C_TEXT_2
        If Left$(typ, 4) = "100%" Then
            .TickLabels.NumberFormat = "0%"
        Else
            .TickLabels.NumberFormat = AxisFormat(u, rc("Values")(1))
        End If
    End With
    With ch.Axes(AX_CATEGORY)
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = modPD_Theme.C_HAIRLINE_2
        .MajorTickMark = xlNone
        .TickLabelPosition = TICK_LOW
        .TickLabels.Font.Name = modPD_Theme.UI_FONT
        .TickLabels.Font.Size = 8.5
        .TickLabels.Font.Color = modPD_Theme.C_TEXT_2
        ' Bars read top down: the first category - the biggest, sorted - on top.
        If InStr(typ, "bar") > 0 Then .ReversePlotOrder = True
    End With
    If typ = "column + line" And rc("Values").count >= 2 Then
        Set v = rc("Values")(2)
        With ch.Axes(AX_VALUE, AX_SECONDARY)
            .Format.Line.Visible = msoFalse
            .MajorTickMark = xlNone
            .TickLabels.Font.Color = modPD_Theme.C_TEXT_3
            .TickLabels.NumberFormat = AxisFormat(u, v)
        End With
    End If
    If InStr(typ, "column") > 0 Or InStr(typ, "bar") > 0 Then ch.ChartGroups(1).GapWidth = 60
    Err.Clear
End Sub

' Each series - or each slice - its colour from the palette asked for.
Private Sub StyleColours(ByVal ch As Object, ByVal rc As Object)
    Dim pal As Variant, i As Long, s As Object, c As Long, typ As String, k As Long
    typ = CStr(rc("Type"))
    pal = Palette(CStr(rc("Palette")))
    k = UBound(pal) + 1
    On Error Resume Next
    If IsRound(typ) Then
        Set s = ch.SeriesCollection(1)
        For i = 1 To s.Points.count
            s.Points(i).Format.Fill.Solid
            s.Points(i).Format.Fill.ForeColor.RGB = modPD_Theme.HX(CStr(pal((i - 1) Mod k)))
            ' The surface between slices: a gap, not a line.
            s.Points(i).Format.Line.Visible = msoTrue
            s.Points(i).Format.Line.ForeColor.RGB = modPD_Theme.C_SURFACE
            s.Points(i).Format.Line.Weight = 1.5
        Next i
        If typ = "doughnut" Then ch.ChartGroups(1).DoughnutHoleSize = 60
        Exit Sub
    End If
    For i = 1 To ch.SeriesCollection.count
        Set s = ch.SeriesCollection(i)
        c = modPD_Theme.HX(CStr(pal((i - 1) Mod k)))
        If IsLineType(typ) Or (typ = "column + line" And i = 2) Then
            s.Format.Line.Visible = msoTrue
            s.Format.Line.ForeColor.RGB = c
            s.Format.Line.Weight = 2.25
            s.MarkerBackgroundColor = c
            s.MarkerForegroundColor = c
            s.MarkerSize = 5
        Else
            s.Format.Fill.Visible = msoTrue
            s.Format.Fill.Solid
            s.Format.Fill.ForeColor.RGB = c
            s.Format.Line.Visible = msoFalse
        End If
    Next i
    Err.Clear
End Sub

Private Sub StyleLegend(ByVal ch As Object, ByVal rc As Object, ByVal nSeries As Long)
    Dim pos As String
    On Error Resume Next
    pos = CStr(rc("Legend"))
    If pos = "auto" Then
        If IsRound(CStr(rc("Type"))) Then
            pos = "right"
        ElseIf nSeries > 1 Then
            pos = "top"
        Else
            pos = "none"
        End If
    End If
    If pos = "none" Then
        ch.HasLegend = False
    Else
        ch.HasLegend = True
        Select Case pos
            Case "top": ch.Legend.Position = LEG_TOP
            Case "bottom": ch.Legend.Position = LEG_BOTTOM
            Case Else: ch.Legend.Position = LEG_RIGHT
        End Select
        ch.Legend.Font.Color = modPD_Theme.C_TEXT_2
        ch.Legend.Font.Name = modPD_Theme.UI_FONT
        ch.Legend.Font.Size = 8.5
    End If
    Err.Clear
End Sub

Private Sub StyleLabels(ByVal ch As Object, ByVal rc As Object, ByVal u As String)
    Dim i As Long, s As Object, typ As String, v As Object
    typ = CStr(rc("Type"))
    If rc("Labels") = "none" Then Exit Sub
    On Error Resume Next
    For i = 1 To ch.SeriesCollection.count
        Set s = ch.SeriesCollection(i)
        If rc("Cols").count > 0 Then
            Set v = rc("Values")(1)
        ElseIf i <= rc("Values").count Then
            Set v = rc("Values")(i)
        End If
        s.HasDataLabels = True
        With s.DataLabels
            .Font.Color = modPD_Theme.C_TEXT
            .Font.Size = 8
            If IsRound(typ) Then
                .ShowPercentage = (rc("Labels") <> "values")
                .ShowValue = (rc("Labels") <> "percent")
                .ShowCategoryName = False
                If rc("Labels") = "percent" Then
                    .NumberFormat = "0%"
                Else
                    .NumberFormat = LabelFormat(u, v)
                End If
                .Position = LBL_BEST_FIT
            Else
                .ShowValue = True
                .NumberFormat = LabelFormat(u, v)
                If InStr(typ, "stacked") > 0 Or Left$(typ, 4) = "100%" Then
                    .Position = LBL_CENTER
                ElseIf Not IsLineType(typ) And InStr(typ, "area") = 0 Then
                    .Position = LBL_OUTSIDE_END
                End If
            End If
        End With
    Next i
    Err.Clear
End Sub

' The axis in the units asked for - or as a share, for a value shown as one.
Private Function AxisFormat(ByVal units As String, ByVal v As Object) As String
    If IsShare(v) Then AxisFormat = "0%": Exit Function
    Select Case units
        Case "thousands": AxisFormat = "#,##0,"" k"""
        Case "millions": AxisFormat = "#,##0,,"" m"""
        Case "billions": AxisFormat = "#,##0.0,,,"" bn"""
        Case Else: AxisFormat = "#,##0"
    End Select
End Function

Private Function LabelFormat(ByVal units As String, ByVal v As Object) As String
    If IsShare(v) Then LabelFormat = "0.0%": Exit Function
    Select Case units
        Case "thousands": LabelFormat = "#,##0,"" k"""
        Case "millions": LabelFormat = "#,##0.0,,"" m"""
        Case "billions": LabelFormat = "#,##0.00,,,"" bn"""
        Case Else: LabelFormat = "#,##0"
    End Select
End Function

Private Function IsShare(ByVal v As Object) As Boolean
    If v Is Nothing Then Exit Function
    IsShare = (Left$(CStr(v("Calc")), 1) = "%")
End Function

' Every hue clears 3:1 on the chart's surface (build/contrast.py checks).
Public Function Palette(ByVal which As String) As Variant
    Select Case which
        Case "emerald": Palette = Array("16B07F", "8FDBBE", "009060", "4FC79C", "C2EBDA", "2E9E78")
        Case "two tone": Palette = Array("16B07F", "4FA3D9")
        Case Else: Palette = Array("16B07F", "4FA3D9", "E0A43B", "8FDBBE", "E07A6B", "A58BE0", "C9B98A", "7E9388")
    End Select
End Function
