Attribute VB_Name = "modReviewPack"
Option Explicit

' ============================================================================
'  The executive front of the review pack: a Cover and an Executive summary,
'  placed before the pack's Index. Both are built from this workbook's
'  Recon_Detail (the run the pack has just made) and print on one landscape
'  page each, so the pack can go to a committee as it is.
'
'  Cover    verdict, reporting date, entity, run, contents, sign-off lines
'  Summary  four KPIs, reconciled by step (table and chart), by risk family
'           (heatmap), and the differences to explain first
'
'  Nothing here changes a value in the pack's own sheets.
' ============================================================================

Private Const DETAIL_SHEET As String = "Recon_Detail"
Private Const TOP_DIFFS As Long = 10

' Recon_Detail's eight summary columns (see DetailSlim), held while the pages are built.
Private mSlim As Variant
Private mLastR As Long

Public Sub AddExecutivePages(ByVal wb As Workbook)
    Dim d As Object, cover As Worksheet, summary As Worksheet, first As Worksheet
    On Error GoTo Quiet
    Set d = SummariseDetail()
    Set first = wb.Worksheets(1)
    Set summary = wb.Worksheets.Add(Before:=first)
    summary.name = UniqueSheetName(wb, "Summary")
    Set cover = wb.Worksheets.Add(Before:=summary)
    cover.name = UniqueSheetName(wb, "Cover")
    BuildSummary summary, d
    BuildCover cover, summary, d, wb
    cover.Activate
    mSlim = Empty
    Exit Sub
Quiet:
    mSlim = Empty
    ' The pack is still complete without its front pages; say so on the log.
    On Error Resume Next
    LogPackProblem "Executive pages were not added: " & Err.description
End Sub

' ---------------------------------------------------------------- data -------

' Keys: n pass fail open assume  run  asOf entity
'   st:<stage>:<PASS|FAIL|OPEN|ASSUME>, fam:<code>:<stage>:<...>, famcases:<code>
'   families (Collection), cases (Dictionary), diffs (array of row indexes, by size)
Private Function SummariseDetail() As Object
    Dim d As Object, ws As Worksheet, lastR As Long, a As Variant, i As Long, v As String, k As String
    Dim stg As String, tc As String, fam As String, fams As Collection, cases As Object
    Dim asOf As String, entity As String, outRows As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set fams = New Collection: Set cases = CreateObject("Scripting.Dictionary")
    d("n") = 0: d("pass") = 0: d("fail") = 0: d("open") = 0: d("assume") = 0: d("run") = ""
    Set ws = GetWorksheetSafe(ThisWorkbook, DETAIL_SHEET)
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
        If lastR >= 8 Then
            mSlim = DetailSlim(ws, lastR): mLastR = lastR
            d("rows") = True
            d("run") = HomeRunText(ws.Cells(8, 18).Value2)
            For i = 1 To UBound(mSlim, 1)
                v = VerdictKey(mSlim(i, 1))
                If Len(v) > 0 Then
                    stg = CellText(mSlim(i, 2)): tc = CellText(mSlim(i, 3)): fam = HomeFamilyCode(tc)
                    d("n") = d("n") + 1
                    Select Case v
                        Case "PASS": d("pass") = d("pass") + 1
                        Case "FAIL": d("fail") = d("fail") + 1
                        Case "ASSUME": d("assume") = d("assume") + 1
                        Case Else: d("open") = d("open") + 1
                    End Select
                    Add1 d, "st:" & stg & ":" & v
                    If Len(fam) > 0 Then
                        If Not d.Exists("famseen:" & fam) Then
                            d("famseen:" & fam) = True: fams.Add fam
                        End If
                        Add1 d, "fam:" & fam & ":" & stg & ":" & v
                        If Len(tc) > 0 And Not cases.Exists(tc) Then
                            cases(tc) = fam: Add1 d, "famcases:" & fam
                        End If
                    End If
                End If
            Next i
        End If
    End If
    HomeSystemFacts asOf, entity, outRows
    d("asOf") = asOf: d("entity") = entity
    Set d("families") = fams: Set d("cases") = cases
    d("diffs") = LargestDifferences(d)
    Set SummariseDetail = d
End Function

Private Function VerdictKey(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(CellText(v))
    Select Case s
        Case "": VerdictKey = ""
        Case "PASS", "FAIL": VerdictKey = s
        Case "ASSUMPTION": VerdictKey = "ASSUME"
        Case Else: VerdictKey = "OPEN"
    End Select
End Function

Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    CellText = Trim$(CStr(v))
End Function

Private Sub Add1(ByVal d As Object, ByVal k As String)
    If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
End Sub

Private Function N0(ByVal d As Object, ByVal k As String) As Long
    If d.Exists(k) Then N0 = d(k)
End Function

' Checked = everything but system-supplied assumptions.
Private Function Checked(ByVal d As Object, ByVal prefix As String) As Long
    Checked = N0(d, prefix & ":PASS") + N0(d, prefix & ":FAIL") + N0(d, prefix & ":OPEN")
End Function

Private Function IsRateRow(ByVal i As Long) As Boolean
    If IsNumeric(mSlim(i, 6)) And IsNumeric(mSlim(i, 7)) And Not IsEmpty(mSlim(i, 6)) And Not IsEmpty(mSlim(i, 7)) Then
        If Not IsError(mSlim(i, 6)) And Not IsError(mSlim(i, 7)) Then IsRateRow = (Abs(CDbl(mSlim(i, 6))) < 5 And Abs(CDbl(mSlim(i, 7))) < 5)
    End If
End Function

' Row indexes of the FAIL rows, largest difference first (rates ranked beside
' amounts as a share of a million, the same rule the home uses).
Private Function LargestDifferences(ByVal d As Object) As Variant
    Dim a As Variant, i As Long, j As Long, n As Long, idx() As Long, mag() As Double, p As Long, tL As Long, tD As Double
    Dim out() As Long, m As Long
    If Not d.Exists("rows") Then LargestDifferences = Empty: Exit Function
    ReDim idx(1 To UBound(mSlim, 1)): ReDim mag(1 To UBound(mSlim, 1))
    For i = 1 To UBound(mSlim, 1)
        If VerdictKey(mSlim(i, 1)) = "FAIL" Then
            If IsNumeric(mSlim(i, 8)) And Not IsEmpty(mSlim(i, 8)) And Not IsError(mSlim(i, 8)) Then
                n = n + 1: idx(n) = i
                If IsRateRow(i) Then mag(n) = Abs(CDbl(mSlim(i, 8))) * 100000000# Else mag(n) = Abs(CDbl(mSlim(i, 8)))
            End If
        End If
    Next i
    If n = 0 Then LargestDifferences = Empty: Exit Function
    m = IIf(n < TOP_DIFFS, n, TOP_DIFFS)
    For i = 1 To m
        p = i
        For j = i + 1 To n
            If mag(j) > mag(p) Then p = j
        Next j
        tD = mag(i): mag(i) = mag(p): mag(p) = tD
        tL = idx(i): idx(i) = idx(p): idx(p) = tL
    Next i
    ReDim out(1 To m)
    For i = 1 To m: out(i) = idx(i): Next i
    LargestDifferences = out
End Function

Private Function VerdictText(ByVal d As Object, ByRef detail As String, ByRef tone As Long) As String
    Dim checkedAll As Long
    checkedAll = d("pass") + d("fail") + d("open")
    If d("n") = 0 Then
        VerdictText = "No reconciliation in this pack"
        detail = "Run the reconciliation, then build the pack again."
        tone = 0
    ElseIf d("fail") = 0 And d("open") = 0 Then
        VerdictText = "Reconciled with the system output"
        detail = format$(d("pass"), "#,##0") & " of " & format$(checkedAll, "#,##0") & " checks are within tolerance. " & _
                 format$(d("assume"), "#,##0") & " values are system-supplied assumptions and are listed for information."
        tone = 1
    Else
        VerdictText = "Review required before sign-off"
        detail = format$(d("pass"), "#,##0") & " of " & format$(checkedAll, "#,##0") & " checks reconcile. " & _
                 format$(d("fail"), "#,##0") & " differences exceed tolerance" & WorstStepText(d) & ". " & _
                 format$(d("open"), "#,##0") & " checks are missing evidence or need review."
        tone = IIf(d("fail") > 0, 3, 2)
    End If
End Function

Private Function WorstStepText(ByVal d As Object) As String
    Dim s As Variant, best As Long, stepName As String
    For Each s In Array("Inputs", "Base", "Pre-shock", "Shock", "Post-shock")
        If N0(d, "st:" & s & ":FAIL") > best Then best = N0(d, "st:" & s & ":FAIL"): stepName = CStr(s)
    Next s
    If best > 0 Then WorstStepText = ", most of them at " & LCase$(stepName) & " (" & format$(best, "#,##0") & ")"
End Function

' ---------------------------------------------------------------- cover ------

Private Sub BuildCover(ByVal ws As Worksheet, ByVal summary As Worksheet, ByVal d As Object, ByVal wb As Workbook)
    Dim detail As String, tone As Long, headline As String, r As Long, i As Long, items As Variant, it As Variant
    ActiveWindowOff ws
    ws.Cells.Font.name = UI_FONT: ws.Cells.Font.Size = 10: ws.Cells.Font.Color = UI_TEXT
    ws.Cells.Interior.Color = UI_WHITE
    ws.columns("A").ColumnWidth = 2
    ws.columns("B").ColumnWidth = 34
    ws.columns("C").ColumnWidth = 5
    ws.columns("D:K").ColumnWidth = 12
    ws.columns("L").ColumnWidth = 3
    For r = 1 To 34: ws.rows(r).RowHeight = 16: Next r

    ' ---- navy side panel -------------------------------------------------
    ws.Range("A1:B34").Interior.Color = UI_NAVY
    With ws.Range("B3")
        .Value2 = "JKB": .Font.Size = 16: .Font.Bold = True: .Font.Color = UI_NAVY
        .Interior.Color = UI_WHITE: .HorizontalAlignment = xlCenter
    End With
    ws.rows(3).RowHeight = 30
    ws.Range("B3").ColumnWidth = 34
    ws.Range("B3").IndentLevel = 0
    SideFact ws, 12, "REPORTING DATE", IIf(Len(d("asOf")) > 0, d("asOf"), "Not recorded")
    SideFact ws, 15, "ENTITY", IIf(Len(d("entity")) > 0, d("entity"), "Not recorded")
    SideFact ws, 18, "RUN", IIf(Len(d("run")) > 0, d("run"), "Not run")
    SideFact ws, 21, "PREPARED", format$(Now, "dd mmm yyyy, hh:nn")
    SideFact ws, 24, "WORKBOOK", ThisWorkbook.name

    ' ---- title -----------------------------------------------------------
    With ws.Range("D4")
        .Value2 = "STRESS TESTING  " & ChrW(183) & "  REVIEW PACK"
        .Font.Size = 9: .Font.Bold = True: .Font.Color = UI_BRAND
    End With
    ws.rows(6).RowHeight = 34: ws.rows(7).RowHeight = 30
    With ws.Range("D6")
        .Value2 = "Stress test reconciliation": .Font.Size = 26: .Font.Bold = True: .Font.Color = UI_INK
    End With
    With ws.Range("D7")
        .Value2 = IIf(Len(d("asOf")) > 0, d("asOf"), "Reporting date not recorded") & IIf(Len(d("entity")) > 0, "  " & ChrW(183) & "  " & d("entity"), "")
        .Font.Size = 16: .Font.Color = UI_TEXT_2
    End With
    ws.Range("D9:K11").Merge
    With ws.Range("D9")
        .Value2 = "An independent rebuild of every scenario from the input files the stress testing system uses, " & _
                  "reconciled step by step with the system's own output: base, pre-shock, shock and post-shock."
        .WrapText = True: .VerticalAlignment = xlTop: .Font.Size = 11: .Font.Color = UI_TEXT_2
    End With

    ' ---- verdict ---------------------------------------------------------
    headline = VerdictText(d, detail, tone)
    ws.Range("D13:K13").Merge: ws.Range("D14:K16").Merge
    ws.rows(13).RowHeight = 22
    With ws.Range("D13")
        .Value2 = ChrW(9679) & "  " & headline: .Font.Bold = True: .Font.Size = 12.5: .IndentLevel = 1
        .VerticalAlignment = xlBottom
    End With
    With ws.Range("D14")
        .Value2 = detail: .WrapText = True: .VerticalAlignment = xlTop: .Font.Size = 10: .IndentLevel = 1
    End With
    ToneBox ws.Range("D13:K16"), tone

    ' ---- contents --------------------------------------------------------
    With ws.Range("D19")
        .Value2 = "CONTENTS": .Font.Size = 9: .Font.Bold = True: .Font.Color = UI_MUTED
    End With
    items = Array(Array(summary.name, "Executive summary", "KPIs, reconciled by step, by risk family, differences to explain"), _
                  Array("Index", "Index", "Every report in this pack, and each test case scope"), _
                  Array("Where_It_Breaks", "Where it breaks", "Differences and missing evidence to review"), _
                  Array("Manual_Calcs", "Every check", "Tool against system for each value, with the calculation"), _
                  Array("Base_Reconciliation", "Base", "Bank base against the system"), _
                  Array("Reconciliation", "Pre-shock", "Filtered pre-shock against the system, per test case"))
    r = 20: i = 0
    For Each it In items
        If Not GetWorksheetSafe(wb, CStr(it(0))) Is Nothing Then
            i = i + 1
            ws.Cells(r, 4).Value2 = i
            ws.Cells(r, 4).HorizontalAlignment = xlCenter: ws.Cells(r, 4).Font.Color = UI_MUTED
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 5), Address:="", SubAddress:="'" & Replace(CStr(it(0)), "'", "''") & "'!A1", TextToDisplay:=CStr(it(1))
            ws.Cells(r, 5).Font.name = UI_FONT: ws.Cells(r, 5).Font.Bold = True: ws.Cells(r, 5).Font.Color = UI_BRAND
            ws.Cells(r, 5).Font.Underline = xlUnderlineStyleNone: ws.Cells(r, 5).Font.Size = 10
            ws.Cells(r, 7).Value2 = it(2): ws.Cells(r, 7).Font.Color = UI_MUTED: ws.Cells(r, 7).Font.Size = 9.5
            ws.Range(ws.Cells(r, 4), ws.Cells(r, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range(ws.Cells(r, 4), ws.Cells(r, 11)).Borders(xlEdgeBottom).Color = UI_LINE
            ws.rows(r).RowHeight = 20
            r = r + 1
        End If
    Next it

    ' ---- sign-off --------------------------------------------------------
    r = 30
    ws.rows(r - 1).RowHeight = 34
    SignLine ws, r, 4, 5, "Prepared by"
    SignLine ws, r, 7, 8, "Reviewed by"
    SignLine ws, r, 10, 11, "Approved by"

    ws.Tab.Color = UI_NAVY
    PackPageSetup ws, "A1:L34", "Cover"
End Sub

Private Sub SideFact(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal value As String)
    With ws.Cells(r, 2)
        .Value2 = label: .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_NAVY_MUTED: .IndentLevel = 1
    End With
    With ws.Cells(r + 1, 2)
        .NumberFormat = "@": .Value2 = value: .Font.Size = 12: .Font.Bold = True: .Font.Color = UI_WHITE: .IndentLevel = 1
    End With
    ws.rows(r + 1).RowHeight = 20
End Sub

Private Sub SignLine(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal label As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Color = UI_LINE_2
    End With
    ws.Cells(r, c1).Value2 = label: ws.Cells(r, c1).Font.Size = 9: ws.Cells(r, c1).Font.Color = UI_MUTED
    ws.Cells(r + 1, c1).Value2 = "Name, date": ws.Cells(r + 1, c1).Font.Size = 8.5: ws.Cells(r + 1, c1).Font.Color = UI_FAINT
End Sub

' tone: 0 neutral, 1 ok, 2 warn, 3 bad
Private Sub ToneBox(ByVal rg As Range, ByVal tone As Long)
    Dim fillC As Long, lineC As Long, textC As Long, e As Variant
    Select Case tone
        Case 1: fillC = UI_OK_BG: lineC = UI_OK_LINE: textC = UI_OK
        Case 2: fillC = UI_WARN_BG: lineC = UI_WARN_LINE: textC = UI_WARN
        Case 3: fillC = UI_BAD_BG: lineC = UI_BAD_LINE: textC = UI_BAD
        Case Else: fillC = UI_SUBTLE: lineC = UI_LINE: textC = UI_TEXT
    End Select
    rg.Interior.Color = fillC
    For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        rg.Borders(e).LineStyle = xlContinuous: rg.Borders(e).Color = lineC
    Next e
    rg.rows(1).Font.Color = textC
    rg.Range("A2").Font.Color = UI_TEXT
End Sub

' ---------------------------------------------------------------- summary ----

Private Sub BuildSummary(ByVal ws As Worksheet, ByVal d As Object)
    Dim r As Long, i As Long, j As Long, stages As Variant, s As Variant, fams As Collection, fam As String
    Dim a As Variant, diffs As Variant, k As String, checkedAll As Long, co As ChartObject, caseCount As Long, famCount As Long
    ActiveWindowOff ws
    ws.Cells.Font.name = UI_FONT: ws.Cells.Font.Size = 9.5: ws.Cells.Font.Color = UI_TEXT
    ws.Cells.Interior.Color = UI_WHITE
    ws.columns("A").ColumnWidth = 2
    ws.columns("B").ColumnWidth = 16
    ws.columns("C:G").ColumnWidth = 10.5
    ws.columns("H").ColumnWidth = 3
    ws.columns("I").ColumnWidth = 30
    ws.columns("J").ColumnWidth = 10
    ws.columns("K:M").ColumnWidth = 14
    ws.columns("N").ColumnWidth = 38
    ws.columns("O").ColumnWidth = 2

    ' ---- header ------------------------------------------------------------
    ws.rows(1).RowHeight = 8: ws.rows(2).RowHeight = 30
    With ws.Range("B2")
        .Value2 = "Executive summary": .Font.Size = 18: .Font.Bold = True: .Font.Color = UI_INK: .VerticalAlignment = xlBottom
    End With
    With ws.Range("I2:N2")
        .Merge: .HorizontalAlignment = xlRight: .VerticalAlignment = xlBottom: .Font.Size = 9: .Font.Color = UI_MUTED
        .Value2 = "Stress test reconciliation  " & ChrW(183) & "  " & IIf(Len(d("asOf")) > 0, d("asOf"), "date not recorded") & _
                  IIf(Len(d("entity")) > 0, "  " & ChrW(183) & "  " & d("entity"), "") & "  " & ChrW(183) & "  run " & d("run")
    End With
    With ws.Range("B2:N2").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = UI_NAVY
    End With
    ws.rows(3).RowHeight = 12

    ' ---- KPIs (rows 4-6) -------------------------------------------------
    checkedAll = d("pass") + d("fail") + d("open")
    ws.rows(4).RowHeight = 16: ws.rows(5).RowHeight = 28: ws.rows(6).RowHeight = 15: ws.rows(7).RowHeight = 12
    Kpi ws, "B4:D6", "Checks reconciled", IIf(checkedAll > 0, format$(d("pass") / checkedAll, "0.0%"), "-"), _
        format$(d("pass"), "#,##0") & " of " & format$(checkedAll, "#,##0"), UI_INK
    Kpi ws, "E4:G6", "Differences", format$(d("fail"), "#,##0"), "above tolerance", IIf(d("fail") > 0, UI_BAD, UI_OK)
    Kpi ws, "I4:J6", "Missing or to review", format$(d("open"), "#,##0"), "evidence incomplete", IIf(d("open") > 0, UI_WARN, UI_OK)
    ' Same count as Home: the configured test cases and their risk families.
    caseCount = HomeCaseCounts(famCount)
    Kpi ws, "K4:N6", "Test cases", format$(caseCount, "#,##0"), famCount & " risk families", UI_INK

    ' ---- reconciled by step (table rows 9-14, chart beside it) -------------
    Eyebrow ws.Range("B9"), "Reconciled by step"
    Header ws, 10, Array("Step", "Reconciled", "Differences", "Missing", "Checked", "% reconciled"), 2
    stages = Array("Inputs", "Base", "Pre-shock", "Shock", "Post-shock")
    r = 11
    For Each s In stages
        If Checked(d, "st:" & s) > 0 Then
            ws.Cells(r, 2).Value2 = s
            ws.Cells(r, 3).Value2 = N0(d, "st:" & s & ":PASS")
            ws.Cells(r, 4).Value2 = N0(d, "st:" & s & ":FAIL")
            ws.Cells(r, 5).Value2 = N0(d, "st:" & s & ":OPEN")
            ws.Cells(r, 6).Value2 = Checked(d, "st:" & s)
            ws.Cells(r, 7).Value2 = N0(d, "st:" & s & ":PASS") / Checked(d, "st:" & s)
            BodyRow ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
            r = r + 1
        End If
    Next s
    If r = 11 Then
        ws.Cells(11, 2).Value2 = "No checks in this run.": ws.Cells(11, 2).Font.Color = UI_FAINT
    Else
        ws.Range(ws.Cells(11, 3), ws.Cells(r - 1, 6)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(11, 7), ws.Cells(r - 1, 7)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(11, 4), ws.Cells(r - 1, 4)).Font.Color = UI_BAD
        ws.Range(ws.Cells(11, 5), ws.Cells(r - 1, 5)).Font.Color = UI_WARN
        HeatScale ws.Range(ws.Cells(11, 7), ws.Cells(r - 1, 7))
        Set co = StepChart(ws, ws.Range(ws.Cells(10, 2), ws.Cells(r - 1, 5)))
    End If

    ' ---- by risk family (heatmap rows 17-26) --------------------------------
    Eyebrow ws.Range("B17"), "By risk family  " & ChrW(183) & "  % reconciled"
    Header ws, 18, Array("Risk family", "Cases", "Base", "Pre-shock", "Shock", "Post-shock"), 2
    Set fams = d("families")
    r = 19
    For i = 1 To fams.count
        If r > 26 Then Exit For
        fam = fams(i)
        If Not HomeIsRiskFamily(fam) Then GoTo NextFamily
        ws.Cells(r, 2).Value2 = HomeFamilyName(fam & "_")
        ws.Cells(r, 3).Value2 = HomeFamilyCaseCount(fam)
        If ws.Cells(r, 3).Value2 = 0 Then ws.Cells(r, 3).Value2 = N0(d, "famcases:" & fam)
        j = 4
        For Each s In Array("Base", "Pre-shock", "Shock", "Post-shock")
            k = "fam:" & fam & ":" & s
            If Checked(d, k) > 0 Then
                ws.Cells(r, j).Value2 = N0(d, k & ":PASS") / Checked(d, k)
            Else
                ws.Cells(r, j).Value2 = ChrW(8211): ws.Cells(r, j).Font.Color = UI_FAINT
            End If
            ws.Cells(r, j).HorizontalAlignment = xlCenter
            j = j + 1
        Next s
        BodyRow ws.Range(ws.Cells(r, 2), ws.Cells(r, 7))
        r = r + 1
NextFamily:
    Next i
    If r > 19 Then
        ws.Range(ws.Cells(19, 4), ws.Cells(r - 1, 7)).NumberFormat = "0%"
        ws.Range(ws.Cells(19, 4), ws.Cells(r - 1, 7)).Font.Bold = True
        HeatCells ws.Range(ws.Cells(19, 4), ws.Cells(r - 1, 7))
    Else
        ws.Cells(19, 2).Value2 = "No test cases in this run.": ws.Cells(19, 2).Font.Color = UI_FAINT
    End If

    ' ---- differences to explain (rows 17-28, right) --------------------------
    Eyebrow ws.Range("I17"), "Differences to explain first  " & ChrW(183) & "  largest " & TOP_DIFFS
    Header ws, 18, Array("Test case / field", "Step", "Tool", "System", "Difference", "Reason / next action"), 9
    diffs = d("diffs")
    r = 19
    If IsArray(diffs) Then
            For i = LBound(diffs) To UBound(diffs)
            j = diffs(i)
            ws.Cells(r, 9).Value2 = CellText(mSlim(j, 3)) & "  " & ChrW(183) & "  " & CellText(mSlim(j, 5))
            ws.Cells(r, 10).Value2 = CellText(mSlim(j, 2))
            ws.Cells(r, 11).Value2 = mSlim(j, 6): ws.Cells(r, 12).Value2 = mSlim(j, 7): ws.Cells(r, 13).Value2 = mSlim(j, 8)
            If IsRateRow(j) Then
                ws.Range(ws.Cells(r, 11), ws.Cells(r, 12)).NumberFormat = "0.00%"
                ws.Cells(r, 13).Value2 = CDbl(mSlim(j, 8)) * 100
                ws.Cells(r, 13).NumberFormat = "+0.00"" pp"";-0.00"" pp"""
            Else
                ws.Range(ws.Cells(r, 11), ws.Cells(r, 12)).NumberFormat = "#,##0"
                ws.Cells(r, 13).NumberFormat = "+#,##0;-#,##0"
            End If
            ws.Cells(r, 13).Font.Bold = True
            ws.Cells(r, 13).Font.Color = IIf(CDbl(mSlim(j, 8)) < 0, UI_BAD, UI_TEXT)
            ws.Cells(r, 14).Value2 = CellText(GetWorksheetSafe(ThisWorkbook, DETAIL_SHEET).Cells(7 + j, 14).Value2)
            ws.Cells(r, 14).Font.Color = UI_TEXT_2
            BodyRow ws.Range(ws.Cells(r, 9), ws.Cells(r, 14))
            r = r + 1
        Next i
        ws.Range(ws.Cells(19, 11), ws.Cells(r - 1, 13)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(19, 9), ws.Cells(r - 1, 9)).Font.Size = 8.5
        ws.Range(ws.Cells(19, 14), ws.Cells(r - 1, 14)).Font.Size = 8.5
    Else
        ws.Cells(19, 9).Value2 = "Nothing above tolerance.": ws.Cells(19, 9).Font.Color = UI_OK
    End If
    For r = 10 To 30: ws.rows(r).RowHeight = 18: Next r
    ' Test case, field and reason wrap onto two lines rather than being cut.
    If IsArray(diffs) Then
        With ws.Range(ws.Cells(19, 9), ws.Cells(18 + UBound(diffs) - LBound(diffs) + 1, 9))
            .WrapText = True
        End With
        With ws.Range(ws.Cells(19, 14), ws.Cells(18 + UBound(diffs) - LBound(diffs) + 1, 14))
            .WrapText = True: .IndentLevel = 1
        End With
        For r = 19 To 18 + UBound(diffs) - LBound(diffs) + 1: ws.rows(r).RowHeight = 26: Next r
    End If
    ws.Cells(18, 14).IndentLevel = 1
    ws.rows(16).RowHeight = 14: ws.rows(9).RowHeight = 20: ws.rows(17).RowHeight = 20

    ' ---- footer ----------------------------------------------------------
    With ws.Range("B31:N31")
        .Merge: .Font.Size = 8: .Font.Color = UI_FAINT: .WrapText = True: .VerticalAlignment = xlTop
        .Value2 = "A check reconciles when the tool's value, rebuilt from the input files, is within tolerance of the system output. " & _
                  "Differences are the system minus the tool. System-supplied assumptions are excluded from the percentages. " & _
                  "Every figure here comes from the Results sheet of the run named above."
    End With
    ws.rows(31).RowHeight = 40

    ' The chart sits beside the step table.
    If Not co Is Nothing Then
        co.Left = ws.Range("I9").Left: co.Top = ws.Range("I9").Top
        co.Width = ws.Range("I9:N9").Width: co.Height = ws.Range("I9:I15").Height
    End If
    ws.Tab.Color = UI_BRAND
    PackPageSetup ws, "A1:O31", "Executive summary"
End Sub

Private Sub Kpi(ByVal ws As Worksheet, ByVal address As String, ByVal label As String, ByVal value As String, ByVal note As String, ByVal valueColor As Long)
    Dim rg As Range
    Set rg = ws.Range(address)
    UiCard rg
    rg.rows(1).Merge: rg.rows(2).Merge: rg.rows(3).Merge
    With rg.rows(1)
        .Value2 = label: .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_MUTED: .IndentLevel = 1: .VerticalAlignment = xlBottom
    End With
    With rg.rows(2)
        .NumberFormat = "@": .Value2 = value: .Font.Size = 20: .Font.Bold = True: .Font.Color = valueColor: .IndentLevel = 1
        .HorizontalAlignment = xlLeft
    End With
    With rg.rows(3)
        .Value2 = note: .Font.Size = 8.5: .Font.Color = UI_FAINT: .IndentLevel = 1: .VerticalAlignment = xlTop
    End With
End Sub

Private Sub Eyebrow(ByVal c As Range, ByVal text As String)
    c.Value2 = UCase$(text)
    c.Font.Size = 8.5: c.Font.Bold = True: c.Font.Color = UI_MUTED: c.VerticalAlignment = xlBottom
End Sub

Private Sub Header(ByVal ws As Worksheet, ByVal r As Long, ByVal titles As Variant, ByVal c1 As Long)
    Dim i As Long
    For i = 0 To UBound(titles)
        With ws.Cells(r, c1 + i)
            .Value2 = titles(i): .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_MUTED
            .HorizontalAlignment = IIf(i = 0, xlLeft, IIf(c1 = 9 And i = 5, xlLeft, xlRight))
            If c1 = 2 And r = 18 And i >= 2 Then .HorizontalAlignment = xlCenter
            If c1 = 9 And i = 1 Then .HorizontalAlignment = xlLeft
        End With
    Next i
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c1 + UBound(titles))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = UI_LINE_2
    End With
End Sub

Private Sub BodyRow(ByVal rg As Range)
    With rg.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = UI_LINE
    End With
    rg.VerticalAlignment = xlCenter
End Sub

' Green at 99.5% and above, amber down to 85%, red below; text cells untouched.
Private Sub HeatCells(ByVal rg As Range)
    Dim fc As FormatCondition, tl As String
    tl = rg.Cells(1, 1).address(False, False)
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=AND(ISNUMBER(" & tl & ")," & tl & ">=0.995)")
    fc.Interior.Color = UI_OK_FILL: fc.Font.Color = UI_OK: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=AND(ISNUMBER(" & tl & ")," & tl & ">=0.85)")
    fc.Interior.Color = UI_WARN_FILL: fc.Font.Color = UI_WARN: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=ISNUMBER(" & tl & ")")
    fc.Interior.Color = UI_BAD_FILL: fc.Font.Color = UI_BAD: fc.StopIfTrue = True
    rg.Borders(xlInsideVertical).LineStyle = xlContinuous: rg.Borders(xlInsideVertical).Color = UI_WHITE
    rg.Borders(xlInsideHorizontal).LineStyle = xlContinuous: rg.Borders(xlInsideHorizontal).Color = UI_WHITE
End Sub

Private Sub HeatScale(ByVal rg As Range)
    Dim fc As FormatCondition, tl As String
    tl = rg.Cells(1, 1).address(False, False)
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=" & tl & ">=0.995")
    fc.Font.Color = UI_OK: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=" & tl & ">=0.85")
    fc.Font.Color = UI_WARN: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=ISNUMBER(" & tl & ")")
    fc.Font.Color = UI_BAD
    rg.Font.Bold = True
End Sub

' A 100% stacked bar per step: reconciled, differences, missing.
Private Function StepChart(ByVal ws As Worksheet, ByVal src As Range) As ChartObject
    Dim co As ChartObject, ch As Chart, i As Long, colors As Variant
    On Error GoTo Quiet
    Set co = ws.ChartObjects.Add(ws.Range("I9").Left, ws.Range("I9").Top, 420, 120)
    co.name = "StepChart"
    Set ch = co.Chart
    ch.ChartType = xlBarStacked100
    ch.SetSourceData Source:=src, PlotBy:=xlColumns
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    ch.Legend.Font.Size = 8: ch.Legend.Font.Color = UI_TEXT_2
    ch.ChartArea.Format.line.Visible = msoFalse
    ch.ChartArea.Format.fill.ForeColor.RGB = UI_WHITE
    ch.PlotArea.Format.fill.Visible = msoFalse
    ch.ChartArea.Font.name = UI_FONT
    colors = Array(UI_OK_DOT, UI_BAD_DOT, UI_WARN_DOT)
    For i = 1 To ch.SeriesCollection.count
        If i <= 3 Then ch.SeriesCollection(i).Format.fill.ForeColor.RGB = colors(i - 1)
        ch.SeriesCollection(i).Format.line.Visible = msoFalse
    Next i
    ch.ChartGroups(1).GapWidth = 60
    With ch.Axes(xlCategory)
        .ReversePlotOrder = True
        .TickLabels.Font.Size = 8.5: .TickLabels.Font.Color = UI_TEXT_2
        .Format.line.ForeColor.RGB = UI_LINE_2
    End With
    With ch.Axes(xlValue)
        .HasMajorGridlines = False
        .TickLabels.Font.Size = 8: .TickLabels.Font.Color = UI_FAINT
        .TickLabels.NumberFormat = "0%"
        .Format.line.Visible = msoFalse
    End With
    Set StepChart = co
    Exit Function
Quiet:
    Set StepChart = co
End Function

' ---------------------------------------------------------------- shared -----

Private Sub ActiveWindowOff(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.zoom = 100
End Sub

Private Sub PackPageSetup(ByVal ws As Worksheet, ByVal area As String, ByVal footerTitle As String)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = area
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.4): .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4): .BottomMargin = Application.InchesToPoints(0.5)
        .LeftFooter = "&8Stress test reconciliation  |  " & footerTitle
        .RightFooter = "&8Printed &D"
    End With
    Application.PrintCommunication = True
    ws.Names.Add name:="JKB_FixedPage", RefersTo:="=TRUE", Visible:=False
    ws.Range("A1").Select
End Sub

Private Sub LogPackProblem(ByVal text As String)
    Dim ws As Worksheet, r As Long
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_LOG)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1).Value2 = format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value2 = "Review pack"
    ws.Cells(r, 3).Value2 = text
End Sub
