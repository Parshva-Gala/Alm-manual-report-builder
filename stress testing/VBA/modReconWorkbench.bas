Attribute VB_Name = "modReconWorkbench"
Option Explicit

Public ReconStrictMode As Boolean
Private Const HOME As String = "Recon_Workbench"
Private Const detail As String = "Recon_Detail"
Private Const CATALOG As String = "Shock_Catalog"
Private mRun As String
Private mRows As Collection
Private mReview As Object
Private mTolerance As Double
Private mRateTolerance As Double
Private mParamIssues As Object

Public Sub OpenReconWorkbench()
    EnsureReconWorkbench
    ThisWorkbook.Worksheets(HOME).Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ' The cockpit is laid out for the full width A:AK; fit it to the window,
    ' but never so small that it stops being readable or so large it blurs.
    On Error Resume Next
    ThisWorkbook.Worksheets(HOME).Range("A1:AK1").Select
    ActiveWindow.zoom = True
    If ActiveWindow.zoom < 70 Then ActiveWindow.zoom = 70
    If ActiveWindow.zoom > 110 Then ActiveWindow.zoom = 110
    ThisWorkbook.Worksheets(HOME).Range("A2").Select
    On Error GoTo 0
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
End Sub

' The one home is the reconciliation cockpit in modCockpit. Its values are named
' (JKB_RunStatus, JKB_TolAmount, JKB_TolRate, JKB_LastActivity) and reached through
' HomeSetStatus, HomeSetActivity and HomeTolerance, never by cell address.
Public Sub EnsureReconWorkbench()
    EnsureHomeCockpit
    Sheet1.EnsureRibbon ThisWorkbook.Worksheets(HOME)
    ' The ribbon's sheet marker stays in A1 but is not something to read.
    ThisWorkbook.Worksheets(HOME).Range("A1").Font.Color = ThisWorkbook.Worksheets(HOME).Range("B1").Interior.Color
End Sub


' Native navigation only: the source facts, controls and pivots are never changed.
Public Sub EnsureJoinedOverview(ByVal wb As Workbook, Optional ByVal recordCount As Long = -1)
    Dim ws As Worksheet, dataWs As Worksheet, coverage As Worksheet, target As Worksheet
    Dim sources As Variant, labels As Variant, descriptions As Variant, bands As Variant, titles As Variant
    Dim sourceCount As Long, sourceCol As Long, testcaseCount As Long, i As Long, r As Long, c As Long, caseKeys As Object, key As String
    Set dataWs = GetWorksheetSafe(wb, "Joined_Data")
    If dataWs Is Nothing Then Err.Raise vbObjectError + 1744, , "The source evidence sheet is unavailable."
    If recordCount < 0 Then
        For c = 1 To dataWs.Cells(1, dataWs.columns.count).End(xlToLeft).Column
            If SafeUpperText(dataWs.Cells(1, c).Value2) = "ROW_SOURCE" Then sourceCol = c: Exit For
        Next c
        If sourceCol = 0 Then Err.Raise vbObjectError + 1745, , "Source record identifiers are unavailable."
        recordCount = Application.Max(0, dataWs.Cells(dataWs.rows.count, sourceCol).End(xlUp).row - 1)
    End If
    Set coverage = GetWorksheetSafe(wb, "Case_Coverage")
    Set caseKeys = CreateObject("Scripting.Dictionary"): caseKeys.CompareMode = vbTextCompare
    If Not coverage Is Nothing Then
        For r = 2 To coverage.Cells(coverage.rows.count, 1).End(xlUp).row
            key = SafeText(coverage.Cells(r, 1).Value2)
            If Len(key) > 0 Then caseKeys(key) = True
        Next r
    End If
    testcaseCount = caseKeys.count
    sources = Array("ECL", "CAPRWA", "LL", "LCR", "CAP")
    labels = Array("Credit impairment", "Capital adequacy & risk-weighted assets", "Legal liquidity", "Liquidity coverage", "Capital components")
    descriptions = Array("ECL source | balances and impairment measures", "CAPRWA source | exposures and risk measures", _
                         "LL source | assets and liabilities", "LCR source | liquidity balances", "CAP source | capital amounts and ratios")
    For i = 0 To UBound(sources)
        Set target = GetWorksheetSafe(wb, "Base_" & CStr(sources(i)))
        If Not target Is Nothing Then sourceCount = sourceCount + 1
    Next i
    Set ws = GetWorksheetSafe(wb, "Overview")
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1)): ws.name = "Overview"
    Else
        If ws.index > 1 Then ws.Move Before:=wb.Worksheets(1)
        ws.Cells.UnMerge: ws.Cells.Clear
        Do While ws.Shapes.count > 0: ws.Shapes(1).Delete: Loop
    End If
    With ws
        .columns("A").ColumnWidth = 3: .columns("B:M").ColumnWidth = 12.5: .columns("N").ColumnWidth = 3
        With .Range("A1:N37")
            .Font.name = "Aptos": .Font.Size = 11: .Font.Bold = False
            .Font.Color = RGB(52, 64, 84): .Interior.Color = RGB(245, 247, 250)
            .Borders.LineStyle = xlNone: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlLeft
        End With
        .rows("1:37").RowHeight = 18
        .rows(1).RowHeight = 12: .rows(2).RowHeight = 24: .rows(3).RowHeight = 42
        .rows(4).RowHeight = 28: .rows(5).RowHeight = 12: .rows(6).RowHeight = 14
        .rows(7).RowHeight = 28: .rows("8:9").RowHeight = 25: .rows(10).RowHeight = 9
        .rows(11).RowHeight = 9: .rows(12).RowHeight = 30: .rows(13).RowHeight = 6
        .rows(17).RowHeight = 8: .rows(21).RowHeight = 8: .rows(25).RowHeight = 8: .rows(29).RowHeight = 8
        .rows(33).RowHeight = 12: .rows(34).RowHeight = 8: .rows("35:36").RowHeight = 16: .rows(37).RowHeight = 10
        .Range("B2:M2").Merge: .Range("B3:M3").Merge: .Range("B4:M4").Merge
        .Range("B2:M5").Interior.Color = RGB(14, 34, 64): .Range("B2:M4").IndentLevel = 1
        .Range("B2").Value2 = "JKB   /   SOURCE EVIDENCE"
        .Range("B2").Font.Size = 9: .Range("B2").Font.Bold = True: .Range("B2").Font.Color = RGB(183, 196, 216)
        .Range("B3").Value2 = "Bank base & testcase populations"
        .Range("B3").Font.Size = 26: .Range("B3").Font.Bold = True: .Range("B3").Font.Color = vbWhite
        .Range("B4").Value2 = "Explore the source records behind your stress-testing review."
        .Range("B4").Font.Size = 11: .Range("B4").Font.Color = RGB(183, 196, 216)
        bands = Array("B7:E10", "F7:I10", "J7:M10")
        titles = Array("SOURCE RECORDS", "SOURCE DATASETS", "TESTCASES")
        For i = 0 To 2
            With .Range(CStr(bands(i)))
                .Interior.Color = vbWhite: .IndentLevel = 1
                .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThick
                .Borders(xlEdgeRight).Color = RGB(245, 247, 250)
            End With
            c = 2 + i * 4
            .Range(.Cells(7, c), .Cells(7, c + 3)).Merge
            .Range(.Cells(8, c), .Cells(9, c + 3)).Merge
            .Cells(7, c).Value2 = titles(i): .Cells(7, c).Font.Size = 9
            .Cells(7, c).Font.Color = RGB(102, 112, 133): .Cells(7, c).Font.Bold = True
            .Cells(8, c).Font.Size = 29: .Cells(8, c).Font.Bold = True
            .Cells(8, c).Font.Color = RGB(14, 116, 144): .Cells(8, c).NumberFormat = "#,##0"
        Next i
        .Range("B8").Value2 = recordCount: .Range("F8").Value2 = sourceCount: .Range("J8").Value2 = testcaseCount
        .Range("B12:G12").Merge: .Range("I12:M12").Merge
        .Range("B12").Value2 = "EXPLORE BANK BASE": .Range("I12").Value2 = "REVIEW THE EVIDENCE"
        .Range("B12:G12,I12:M12").Font.Size = 10: .Range("B12:G12,I12:M12").Font.Bold = True
        .Range("B12:G12,I12:M12").Font.Color = RGB(71, 84, 103)
    End With
    For i = 0 To UBound(sources)
        r = 14 + i * 4
        JoinedOverviewLink wb, ws, "B" & r & ":G" & r, "B" & (r + 1) & ":G" & (r + 2), _
                           CStr(labels(i)), CStr(descriptions(i)), "Base_" & CStr(sources(i)), False
    Next i
    JoinedOverviewLink wb, ws, "I14:M15", "I16:M18", "Testcase populations", "Open each filtered input pivot from the coverage list.", "Case_Coverage", True
    JoinedOverviewLink wb, ws, "I20:M21", "I22:M24", "Source reconciliation", "Compare original source totals with the joined evidence.", "Source_Controls", False
    JoinedOverviewLink wb, ws, "I26:M27", "I28:M30", "Source records", "Inspect the input records and testcase membership.", "Joined_Data", False
    With ws
        .Range("B35:M36").Merge
        .Range("B35").Value2 = "Start with a bank base view, then use testcase populations to review the selected records."
        .Range("B35").Font.Size = 10: .Range("B35").Font.Color = RGB(102, 112, 133)
        .Tab.Color = RGB(14, 116, 144)
        .PageSetup.PrintArea = "$A$1:$N$37": .PageSetup.orientation = xlLandscape
        .PageSetup.PaperSize = xlPaperA4: .PageSetup.zoom = False
        .PageSetup.FitToPagesWide = 1: .PageSetup.FitToPagesTall = 1
        .PageSetup.LeftMargin = 12: .PageSetup.RightMargin = 12: .PageSetup.TopMargin = 12: .PageSetup.BottomMargin = 12
        .PageSetup.PrintGridlines = False: .PageSetup.PrintHeadings = False
        .Activate: .Range("B3").Select
    End With
    With ActiveWindow
        .DisplayGridlines = False: .DisplayHeadings = False: .zoom = 90
        .ScrollRow = 1: .ScrollColumn = 1
    End With
End Sub

Private Sub JoinedOverviewLink(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal titleAddress As String, ByVal textAddress As String, _
                               ByVal title As String, ByVal description As String, ByVal sheetName As String, ByVal primary As Boolean)
    Dim titleRange As Range, textRange As Range, target As Worksheet, fill As Long, ink As Long
    Set titleRange = ws.Range(titleAddress): Set textRange = ws.Range(textAddress)
    titleRange.Merge: textRange.Merge
    fill = IIf(primary, RGB(14, 116, 144), vbWhite)
    ink = IIf(primary, vbWhite, RGB(14, 116, 144))
    titleRange.Interior.Color = fill: textRange.Interior.Color = fill
    titleRange.IndentLevel = 1: textRange.IndentLevel = 1
    titleRange.Cells(1, 1).Value2 = title & "  >"
    textRange.Cells(1, 1).Value2 = description
    Set target = GetWorksheetSafe(wb, sheetName)
    If Not target Is Nothing Then
        ws.Hyperlinks.Add anchor:=titleRange, address:="", SubAddress:="'" & sheetName & "'!A1", ScreenTip:="Open " & title
    Else
        titleRange.Cells(1, 1).Value2 = title
        textRange.Cells(1, 1).Value2 = "Source not included in this report."
        ink = RGB(102, 112, 133)
    End If
    With titleRange.Font
        .name = "Aptos": .Size = 12: .Bold = True: .Color = ink: .Underline = xlUnderlineStyleNone
    End With
    With textRange
        .Font.Size = 10: .Font.Color = IIf(primary, RGB(236, 253, 255), RGB(102, 112, 133))
        .WrapText = True: .VerticalAlignment = xlTop
    End With
End Sub

Public Sub PositionPivotSlicers(ByVal ws As Worksheet)
    Dim pt As PivotTable, sc As SlicerCache, sl As Slicer
    Dim lastCol As Long, leftPt As Double, topPt As Double
    If ws.PivotTables.count = 0 Then Exit Sub
    For Each pt In ws.PivotTables
        lastCol = Application.Max(lastCol, pt.TableRange2.Column + pt.TableRange2.columns.count - 1)
    Next pt
    leftPt = ws.Cells(1, lastCol + 2).Left
    topPt = ws.rows(3).Top
    For Each sc In ws.Parent.SlicerCaches
        For Each sl In sc.Slicers
            If sl.Shape.Parent Is ws Then
                sl.NumberOfColumns = 1
                sl.Left = leftPt: sl.Top = topPt: sl.Width = 220: sl.Height = 110
                sl.Shape.Placement = xlFreeFloating
                On Error Resume Next
                sl.Style = "SlicerStyleLight2"
                Err.Clear
                On Error GoTo 0
                topPt = topPt + 120
            End If
        Next sl
    Next sc
End Sub

Public Sub BuildManualReconPack()
    Dim folder As String, result As String, failureText As String
    If JKB_Busy Then Exit Sub
    If Not AnySourceLoaded() Then
        UploadAllSources
        If Not AnySourceLoaded() Then Exit Sub
    End If
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Save the manual reconciliation reports"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    On Error GoTo Failed
    ProgressStart "Build manual reconciliation reports", "Inputs, filtered pivots and calculation evidence in one pack", 100
    result = BuildManualReconPackTo(folder)
    ProgressDone "Your report pack is ready."
    OpenReconWorkbench
    HomeSetActivity "Reports saved: " & result
    ThisWorkbook.FollowHyperlink result
    Exit Sub
Failed:
    failureText = Err.description
    ProgressFailed failureText
    OpenReconWorkbench
    UiProblem "Build reports", "The review pack could not be finished.", failureText
End Sub

Public Function BuildManualReconPackTo(ByVal parentFolder As String) As String
    Dim folder As String, joined As String, report As String, result As String, suffix As Long, root As String, checkpoint As String
    If Not AnySourceLoaded() Then Err.Raise vbObjectError + 1740, , "Load the input extracts before building manual reports."
    root = JoinPath(parentFolder, "JKB_Manual_Recon_" & format$(Now, "yyyymmdd_hhnnss"))
    folder = root
    Do While Len(Dir$(folder, vbDirectory)) > 0
        suffix = suffix + 1: folder = root & "_" & suffix
    Loop
    MkDir folder
    ProgressStep "Reconcile base, pre-shock and stressed results", 0
    result = RunStressReconciliationQuiet()
    If Left$(result, 5) = "ERROR" Then Err.Raise vbObjectError + 1741, , result
    checkpoint = JoinPath(folder, "_Calculation_checkpoint.xlsm")
    ThisWorkbook.SaveCopyAs checkpoint
    Application.StatusBar = "Building account-level manual pivots..."
    ProgressDetail "Source and scenario calculations complete", 35
    ProgressStep "Build input-based pivots and source controls", 0
    joined = BuildJoinedInputTo(folder, False)
    If Len(joined) = 0 Then Err.Raise vbObjectError + 1742, , "The input-based pivot workbook could not be built."
    SetManualJoinedEvidence joined
    Application.StatusBar = "Building reconciliation reports..."
    ProgressStep "Compare input pivots, manual calculations and system values", 0
    ProgressDetail "Building the calculation comparison reports", 80
    report = Ps_BuildManualReportsTo(folder)
    If Len(report) = 0 Then Err.Raise vbObjectError + 1743, , "The calculation comparison workbook could not be built."
    ProgressDetail "Saving the reconciliation evidence", 95
    ExportReconEvidenceTo JoinPath(folder, "Reconciliation_evidence.xlsx")
    On Error Resume Next
    If Len(Dir$(checkpoint)) > 0 Then Kill checkpoint
    On Error GoTo 0
    BuildManualReconPackTo = folder
    Application.StatusBar = False
End Function

Private Sub ReconButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal proc As String, ByVal address As String, ByVal primary As Boolean)
    UiButton ws, nm, caption, proc, ws.Range(address), IIf(primary, UI_BTN_PRIMARY, UI_BTN_SECONDARY), IIf(primary, 12, 10.5)
End Sub

Public Sub RunStressReconciliation()
    Dim result As String
    result = RunStressReconciliationQuiet()
    OpenReconWorkbench
    If Left$(result, 5) = "ERROR" Then UiProblem "Reconcile", "The reconciliation did not run.", result
End Sub

Public Function RunStressReconciliationQuiet() As String
    Dim state As Object, a As Variant, h As Object, hr As Long, idx As Object, stats As Object
    Dim cache As Object, rulesByType As Object, grouping As Object, notes As Collection
    Dim tk As Variant, ek As Variant, tc As Object, el As Object, ws As Worksheet, wb As Workbook
    Dim sheets As Object, rules As Collection, er As String, Done As Long
    If JKB_Busy Then RunStressReconciliationQuiet = "ERROR: another operation is running.": Exit Function
    On Error GoTo Failed
    EnsureReconWorkbench
    mTolerance = ReadTolerance(False): mRateTolerance = ReadTolerance(True)
    Set state = CaptureState(): JKB_Busy = True: ReconStrictMode = True
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.DisplayAlerts = False
    mRun = format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set mRows = New Collection: Set mReview = LoadReviewMap()
    HomeSetStatus "Running", "Run started " & mRun & "."
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    ' One Run: what Run tests did first is part of Reconcile now.
    SyncPreShockTestCases idx
    Set rulesByType = LoadRulesByElementType(ThisWorkbook)
    Set grouping = LoadDetailGrouping(ThisWorkbook): Set notes = New Collection
    Set cache = PrepareBaseTestCache(idx, Nothing, a, h, notes)
    MigratePreShockConfigTokens False
    AddRuleCoverage idx
    ResetFormulaCache
    For Each tk In idx.keys
        Set tc = idx(tk): Done = Done + 1
        Application.StatusBar = "Reconciling " & Done & " of " & idx.count & ": " & tc("TestCaseCode")
        ProgressDetail "Calculating stressed results: " & Done & " of " & idx.count & " test cases", 30# + 5# * Done / idx.count
        AddBaseChecks cache, tc
        Set sheets = NewMap()
        Set wb = BuildReconCalculationBook(a, h, rulesByType, grouping, cache, tc, sheets)
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            If rulesByType.Exists(SafeUpperText(el("ElementType"))) Then
                Set rules = rulesByType(SafeUpperText(el("ElementType")))
                Set ws = sheets(CStr(ek))
                AddElementChecks ws, rules, cache, tc, el
                AddStageCoverage tc, el
            Else
                AddLine "BLOCKED", "Shock", tc, el, "ALL", "Configuration", Empty, Empty, "", "No formula configuration for this shock type.", "", "", 0
            End If
        Next ek
        wb.Close False: Set wb = Nothing
        DoEvents
    Next tk
    WriteReconDetail
    UpdateReconSummary "Completed"
    RunStressReconciliationQuiet = "Completed: " & mRows.count & " checks."
    GoTo Done
Failed:
    er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    HomeSetStatus "Failed", "Run failed: " & er & ". Any earlier detail is a prior snapshot."
    RefreshHome
    RunStressReconciliationQuiet = "ERROR: " & er
Done:
    ReconStrictMode = False: JKB_Busy = False: RestoreState state
End Function

Private Sub AddRuleCoverage(ByVal idx As Object)
    Dim ws As Worksheet, expected As Object, observed As Object, scopes As Object, tc As Object, el As Object
    Dim tk As Variant, ek As Variant, scope As Variant, pair As Variant, r As Long, tcCopy As Object, elCopy As Object, parts As Variant
    Set expected = NewMap(): Set observed = NewMap(): Set scopes = NewMap()
    For Each tk In idx.keys
        Set tc = idx(tk): scope = CStr(tc("Date")) & "|" & CStr(tc("EntityID")) & "|" & CStr(tc("EntityCode"))
        Set scopes(CStr(scope)) = tc
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            observed(CStr(scope) & "|" & SafeUpperText(tc("TestCaseCode")) & "|" & SafeUpperText(el("ScenarioCode"))) = True
        Next ek
    Next tk
    Set ws = GetWorksheetSafe(ThisWorkbook, "Rule_Register")
    If ws Is Nothing Then
        AddLine "BLOCKED", "Inputs", tc, el, "ALL", "ST rule coverage", Empty, Empty, "", "Import the authoritative ST rule export before sign-off.", "ST rules", "", 0
        Exit Sub
    End If
    For r = 6 To ws.Cells(ws.rows.count, 2).End(xlUp).row
        If Len(SafeText(ws.Cells(r, 2).Value2)) > 0 And Len(SafeText(ws.Cells(r, 5).Value2)) > 0 Then
            pair = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 5).Value2)
            expected(CStr(pair)) = True
        End If
    Next r
    For Each scope In scopes.keys
        Set tc = scopes(CStr(scope))
        For Each pair In expected.keys
            If Not observed.Exists(CStr(scope) & "|" & CStr(pair)) Then
                parts = Split(CStr(pair), "|")
                If UniqueElementCovered(idx, expected, CStr(scope), CStr(parts(1))) Then GoTo NextExpected
                Set tcCopy = NewMap(): Set elCopy = NewMap()
                tcCopy("Date") = tc("Date"): tcCopy("EntityCode") = tc("EntityCode"): tcCopy("EntityID") = tc("EntityID")
                tcCopy("TestCaseCode") = parts(0): elCopy("ScenarioCode") = parts(1)
                AddLine "BLOCKED", "Inputs", tcCopy, elCopy, "ALL", "Configured case coverage", Empty, Empty, "", "Configured ST rule is absent from this date/entity's system output.", "Rule_Register", "", 0
            End If
NextExpected:
        Next pair
    Next scope
    For Each tk In idx.keys
        Set tc = idx(tk)
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            pair = SafeUpperText(tc("TestCaseCode")) & "|" & SafeUpperText(el("ScenarioCode"))
            If Not expected.Exists(CStr(pair)) Then
                scope = CStr(tc("Date")) & "|" & CStr(tc("EntityID")) & "|" & CStr(tc("EntityCode"))
                If UniqueElementCovered(idx, expected, CStr(scope), SafeUpperText(el("ScenarioCode"))) Then
                    AddLine "REVIEW", "Inputs", tc, el, "ALL", "Filter rule mapping", Empty, Empty, "", "Matched by unique element code. Rule-export and system test-case codes differ; confirm this mapping.", "Rule_Register", "", 0
                Else
                    AddLine "BLOCKED", "Inputs", tc, el, "ALL", "Filter rule coverage", Empty, Empty, "", "System case has no matching pair in the imported ST rule register.", "Rule_Register", "", 0
                End If
            End If
        Next ek
    Next tk
End Sub

Private Function UniqueElementCovered(ByVal idx As Object, ByVal expected As Object, ByVal scope As String, ByVal element As String) As Boolean
    Dim pair As Variant, parts As Variant, sourceCount As Long, targetCount As Long, tk As Variant, tc As Object, ek As Variant
    For Each pair In expected.keys
        parts = Split(CStr(pair), "|")
        If CStr(parts(1)) = element Then sourceCount = sourceCount + 1
    Next pair
    If sourceCount <> 1 Then Exit Function
    For Each tk In idx.keys
        Set tc = idx(tk)
        If CStr(tc("Date")) & "|" & CStr(tc("EntityID")) & "|" & CStr(tc("EntityCode")) = scope Then
            For Each ek In tc("Elements").keys
                If SafeUpperText(ek) = element Then targetCount = targetCount + 1
            Next ek
        End If
    Next tk
    UniqueElementCovered = (targetCount = 1)
End Function

Private Sub AddStageCoverage(ByVal tc As Object, ByVal el As Object)
    Dim stages As Object, line As Variant, stage As Variant
    Set stages = NewMap()
    For Each line In mRows
        If CStr(line(2)) = CStr(tc("Date")) And CStr(line(3)) = tc("EntityCode") & " / " & tc("EntityID") And CStr(line(4)) = CStr(tc("TestCaseCode")) And CStr(line(5)) = CStr(el("ScenarioCode")) Then stages(CStr(line(1))) = True
    Next line
    For Each stage In Array("Base", "Pre-shock", "Shock", "Post-shock")
        If Not stages.Exists(CStr(stage)) Then AddLine "BLOCKED", CStr(stage), tc, el, "ALL", "Stage coverage", Empty, Empty, "", "This stage has no independent check in the current configuration or source population.", "", "", 0
    Next stage
End Sub

Private Function ReadTolerance(ByVal isRate As Boolean) As Double
    Dim v As Variant
    v = HomeTolerance(isRate)
    If Not ReconNumber(v) Then Err.Raise vbObjectError + 1701, , "Enter a number for the " & IIf(isRate, "rate", "amount") & " tolerance on the home screen."
    If CDbl(v) < 0 Then Err.Raise vbObjectError + 1702, , "Tolerance cannot be negative."
    ReadTolerance = CDbl(v)
End Function

Public Function ReconNumber(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    ReconNumber = IsNumeric(v)
End Function

Public Function ReconBaseRecord(ByVal cache As Object, ByVal tc As Object, ByVal el As Object, ByVal label As String) As Object
    Dim k As Variant, rec As Object, alias As String
    Select Case SafeUpperText(label)
        Case "CET1_CAPITAL_LCY": alias = "CET1_BASE"
        Case "AT1_CAPITAL_LCY": alias = "AT1_BASE"
        Case "T2_CAPITAL_LCY": alias = "T2_BASE"
        Case "TOTAL_CAPITAL_LCY": alias = "TOTAL_CAPITAL_BASE"
    End Select
    For Each k In Array(label, label & "_PRE_SHOCK", "PRE_" & label, alias)
        Set rec = DerivedLookup(cache, tc, el, CStr(k))
        If Not rec Is Nothing Then
            If CBool(rec("HasBase")) Then Set ReconBaseRecord = rec: Exit Function
        End If
    Next k
End Function

Private Sub AddBaseChecks(ByVal cache As Object, ByVal tc As Object)
    Dim k As Variant, rec As Object, el As Object, verdict As String, reason As String
    For Each k In cache.keys
        Set rec = cache(k)
        If CStr(rec("Date")) = CStr(tc("Date")) And CStr(rec("EntityID")) = CStr(tc("EntityID")) And CStr(rec("Entity")) = CStr(tc("EntityCode")) And CStr(rec("TestCase")) = CStr(tc("TestCaseCode")) Then
            Set el = tc("Elements")(SafeUpperText(rec("Element")))
            reason = "": verdict = CompareValue(rec("BaseDerived"), rec("BaseSystem"), CStr(rec("Metric")), reason)
            If Not CBool(rec("HasBase")) Then verdict = "BLOCKED": reason = "Base could not be rebuilt from the input extract."
            If verdict = "PASS" Then
                If Len(SourceDateIssue(CStr(rec("Source")))) > 0 Then verdict = "REVIEW": reason = SourceDateIssue(CStr(rec("Source")))
            End If
            AddLine verdict, "Base", tc, el, "ALL", CStr(rec("Metric")), rec("BaseDerived"), rec("BaseSystem"), "Independent source aggregation", reason, CStr(rec("Source")), "Entity and as-of date; no scenario filter", 0
        End If
    Next k
End Sub

Public Function SourceDateIssue(ByVal source As String) As String
    Dim h As Object, parts As Variant
    If IgnoreAsOfDate() Then SourceDateIssue = "REVIEW: reporting-date matching is disabled. Confirm date alignment before sign-off.": Exit Function
    parts = Split(source, "~")
    Set h = SourceHeadersFor(CStr(parts(0)))
    If h Is Nothing Then SourceDateIssue = "REVIEW: source date evidence is not loaded.": Exit Function
    If Not h.Exists("AS_OF_DATE") Then SourceDateIssue = "REVIEW: the source has no reporting date. Confirm the extract date before sign-off."
End Function

Private Sub AddElementChecks(ByVal ws As Worksheet, ByVal rules As Collection, ByVal cache As Object, ByVal tc As Object, ByVal el As Object)
    Dim ruleMap As Object, rowMap As Object, memo As Object, visiting As Object
    Dim ri As Object, r As Long, label As String, sev As Variant, j As Long, f As String
    Dim expected As Variant, actual As Variant, verdict As String, reason As String, stage As String, proof As String
    Dim rec As Object, source As String, filter As String, matched As Double, typeKey As String
    Dim grid As Variant, lastRow As Long, caseFilter As String, filterKnown As Boolean
    Set ruleMap = NewMap(): Set rowMap = NewMap()
    For Each ri In rules: Set ruleMap(SafeUpperText(ri("OutputRowLabel"))) = ri: Next ri
    ' The calculated sheet, read once. It was read a cell at a time, twice per
    ' rule per severity.
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= 7 Then grid = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 11)).Value2
    For r = 7 To lastRow
        label = SafeUpperText(grid(r, 1))
        If ruleMap.Exists(label) Then rowMap(label) = r
    Next r
    sev = SeverityList(): typeKey = SafeUpperText(el("ElementType"))
    For j = 0 To 2
        Set memo = NewMap(): Set visiting = NewMap()
        Set mParamIssues = NewMap()
        For Each ri In rules
            label = SafeUpperText(ri("OutputRowLabel"))
            If ReconAssumption(label) And rowMap.Exists(label) Then
                mParamIssues(label) = ParameterIssue(label, grid(CLng(rowMap(label)), 3 + j))
            End If
        Next ri
        If Not el("SeverityRows").Exists(sev(j)) Then
            AddLine "BLOCKED", "Inputs", tc, el, CStr(sev(j)), "Severity coverage", Empty, Empty, "", "Required severity is absent from the system output.", "System output", "", 0
        Else
            For Each ri In rules
                label = SafeUpperText(ri("OutputRowLabel"))
                If Not rowMap.Exists(label) Then
                    AddLine "BLOCKED", "Inputs", tc, el, CStr(sev(j)), label, Empty, Empty, "", "Configured field is absent from the calculation layout.", "", "", 0
                    GoTo NextRule
                End If
                r = CLng(rowMap(label)): expected = grid(r, 6 + j): actual = grid(r, 3 + j)
                f = ManualFormulaText(ri, j): stage = ReconStage(label)
                ' One element, one filter: looked up once, not once per rule and
                ' severity (each lookup scanned the Cases sheet and rewrote a cell).
                If Not filterKnown Then caseFilter = FilterForCase(CStr(tc("TestCaseCode")), CStr(el("ScenarioCode"))): filterKnown = True
                source = "": filter = caseFilter: matched = 0
                Set rec = DerivedLookup(cache, tc, el, label)
                If Not rec Is Nothing Then
                    If Not ReconUseSource(label, f) Then Set rec = Nothing
                End If
                If Not rec Is Nothing Then
                    source = CStr(rec("Source")): matched = CDbl(rec("Rows"))
                    If CBool(rec("HasPre")) Then expected = rec("PreDerived"): f = "Source aggregation WHERE " & CStr(rec("Filter"))
                End If
                reason = ""
                If ReconAssumption(label) Then
                    verdict = "ASSUMPTION": stage = "Inputs"
                    reason = "Scenario parameter supplied by the system. The ST rules file contains filter conditions, not independent shock magnitudes."
                    expected = Empty
                    proof = ParameterIssue(label, actual)
                    If Len(proof) > 0 Then verdict = "BLOCKED": reason = proof
                ElseIf ReconMetadata(label) Then
                    GoTo NextRule
                Else
                    proof = FormulaProof(label, j, ruleMap, cache, tc, el, memo, visiting)
                    If Len(proof) > 0 Then
                        verdict = "BLOCKED": reason = proof
                        If Left$(proof, 7) = "REVIEW:" Then verdict = "REVIEW"
                    Else
                        verdict = CompareValue(expected, actual, label, reason)
                        If verdict = "PASS" Then
                            If rec Is Nothing Then
                                proof = ReviewIssue(typeKey, label, f)
                                If Len(proof) > 0 Then verdict = "REVIEW": reason = proof
                            End If
                        End If
                    End If
                End If
                AddLine verdict, stage, tc, el, CStr(sev(j)), label, expected, actual, f, reason, source, filter, matched
NextRule:
            Next ri
        End If
    Next j
End Sub

Public Function ReconUseSource(ByVal label As String, ByVal formula As String) As Boolean
    Dim f As String
    f = UCase$(Replace(Replace(NormalizeConfigFormulaText(formula), "@", ""), " ", ""))
    ReconUseSource = (Len(f) = 0 Or f = "SYSTEMOUTPUT." & UCase$(label) Or f = "SYSTEM." & UCase$(label) Or f = "DERIVED." & UCase$(label))
End Function

Private Function ParameterIssue(ByVal label As String, ByVal value As Variant) As String
    If Len(SafeText(value)) = 0 Then ParameterIssue = "Scenario parameter is missing: " & label: Exit Function
    Select Case label
        Case "ECL_CALCULATION_METHOD", "ECL_METHOD"
            Select Case SafeUpperText(value)
                Case "RATIO_METHOD", "PROVISION_METHOD", "ECL_METHOD"
                Case Else: ParameterIssue = "Unknown ECL method: " & SafeText(value)
            End Select
        Case "SHOCK_TYPE", "CHANGE_TYPE"
        Case Else
            If Not ReconNumber(value) Then ParameterIssue = "Scenario parameter must be numeric: " & label
    End Select
End Function

Public Function ReconAssumption(ByVal label As String) As Boolean
    Select Case SafeUpperText(label)
        Case "AMOUNT_CHANGE", "PCT_CHANGE", "ECL_CALCULATION_METHOD", "ECL_METHOD", "SHOCK_TYPE", "CHANGE_TYPE", "TAX_RATE_PCT", "TAX_RATE", "PROVISION_PCT", "PROVISION_RATE", "LGD", "DURATION", "ADJUSTMENTS_PBT_LCY", "ADJUSTMENTS_TAX_LCY"
            ReconAssumption = True
    End Select
End Function

Private Function ReconMetadata(ByVal label As String) As Boolean
    Select Case label
        Case "AS_OF_DATE", "ENTITY_ID", "ENTITY_CODE", "RUN_ID", "SCENARIO_TEST_CASE_CODE", "SCENARIO_TEST_CASE_NAME", "SCENARIO_ELEMENT_CODE", "SCENARIO_ELEMENT_NAME", "ELEMENT_TYPE", "SEVERITY_CODE", "SEVERITY_ID", "SCENARIO_ELEMENT_FORMULA_NAME"
            ReconMetadata = True
    End Select
End Function

Private Function FormulaProof(ByVal label As String, ByVal sev As Long, ByVal rules As Object, ByVal cache As Object, ByVal tc As Object, ByVal el As Object, ByVal memo As Object, ByVal visiting As Object) As String
    Dim rec As Object, f As String, m As Object, dep As String, kind As String, issue As String, tokens As Object
    If memo.Exists(label) Then FormulaProof = CStr(memo(label)): Exit Function
    If visiting.Exists(label) Then FormulaProof = "Circular formula dependency: " & label: Exit Function
    visiting(label) = True
    Set rec = DerivedLookup(cache, tc, el, label)
    If Not rec Is Nothing Then
        If rules.Exists(label) Then
            If Not ReconUseSource(label, ManualFormulaText(rules(label), sev)) Then Set rec = Nothing
        End If
    End If
    If Not rec Is Nothing Then
        If Not CBool(rec("HasPre")) Then issue = "No independently derived pre-shock value: " & label
        If Len(issue) = 0 Then issue = SourceDateIssue(CStr(rec("Source")))
        If Len(SafeText(rec("Filter"))) = 0 Then issue = "No configured filter condition: " & label
        If Not ReconNumber(rec("Rows")) Then
            issue = "Selected population count is unavailable: " & label
        ElseIf CDbl(rec("Rows")) <= 0 Then
            issue = "Filter selected no rows: " & label
        End If
        GoTo Finish
    End If
    If ReconAssumption(label) Then
        If mParamIssues.Exists(label) Then
            issue = CStr(mParamIssues(label))
        Else
            issue = "Scenario parameter is absent: " & label
        End If
        GoTo Finish
    End If
    If Not rules.Exists(label) Then issue = "Formula references an unconfigured field: " & label: GoTo Finish
    f = ManualFormulaText(rules(label), sev)
    If Len(f) = 0 Then issue = "No independent formula or source derivation: " & label: GoTo Finish
    Set tokens = RegexExecute("@?\b(derivedbase|base_derived|derived|base|systemoutput|system|manual|man|differences|difference|diff)\.([A-Za-z0-9_]+)\b", f)
    For Each m In tokens
        kind = LCase$(m.SubMatches(0)): dep = UCase$(m.SubMatches(1))
        Select Case kind
            Case "system", "systemoutput"
                If Not ReconAssumption(dep) Then
                    issue = "Formula uses a system result as evidence: " & dep
                ElseIf Not mParamIssues.Exists(dep) Then
                    issue = "Scenario parameter is absent: " & dep
                Else
                    issue = CStr(mParamIssues(dep))
                End If
            Case "difference", "differences", "diff"
                issue = "Formula depends on a reconciliation difference: " & dep
            Case "base", "derivedbase", "base_derived"
                Set rec = ReconBaseRecord(cache, tc, el, dep)
                If rec Is Nothing Then issue = "No independently rebuilt bank base: " & dep
                If Not rec Is Nothing Then issue = SourceDateIssue(CStr(rec("Source")))
            Case "derived"
                Set rec = DerivedLookup(cache, tc, el, dep)
                If rec Is Nothing Then
                    issue = "Missing source derivation: " & dep
                ElseIf Not CBool(rec("HasPre")) Then
                    issue = "Pre-shock derivation is unavailable: " & dep
                ElseIf Len(SafeText(rec("Filter"))) = 0 Or CDbl(rec("Rows")) <= 0 Then
                    issue = "Selected population evidence is unavailable: " & dep
                End If
            Case Else
                issue = FormulaProof(dep, sev, rules, cache, tc, el, memo, visiting)
        End Select
        If Len(issue) > 0 Then Exit For
    Next m
    If Len(issue) = 0 And (InStr(f, "[") > 0 Or RegexExecute("@[A-Za-z_]+\(", f).count > 0) Then issue = "Alternative reference syntax requires a lineage review: " & label
    If Len(issue) = 0 Then
        issue = ReviewIssue(SafeUpperText(el("ElementType")), label, f)
        ' A design review issue is propagated to dependent financial results.
    End If
Finish:
    visiting.Remove label: memo(label) = issue: FormulaProof = issue
End Function

Private Function LoadReviewMap() As Object
    Dim ws As Worksheet, d As Object, r As Long, k As String
    Set d = NewMap(): Set LoadReviewMap = d
    Set ws = GetWorksheetSafe(ThisWorkbook, CATALOG)
    If ws Is Nothing Then Exit Function
    For r = 8 To ws.Cells(ws.rows.count, 1).End(xlUp).row
        k = SafeUpperText(ws.Cells(r, 1).Value2) & "|" & SafeUpperText(ws.Cells(r, 2).Value2)
        If Len(k) > 1 Then d(k) = Array(SafeUpperText(ws.Cells(r, 5).Value2), NormalizeConfigFormulaText(ws.Cells(r, 6).Value2))
    Next r
End Function

Private Function ReviewIssue(ByVal shockType As String, ByVal label As String, ByVal formula As String) As String
    Dim key As String, status As String, item As Variant
    key = shockType & "|" & label
    If mReview Is Nothing Then ReviewIssue = "REVIEW: shock design has not been checked against the configured formula.": Exit Function
    If Not mReview.Exists(key) Then ReviewIssue = "REVIEW: no reviewed shock-design mapping for " & label: Exit Function
    item = mReview(key): status = CStr(item(0))
    If status <> "ALIGNED" And status <> "CORRECTED" Then
        ReviewIssue = "REVIEW: " & status & " (" & label & "). See Shock_Catalog."
    ElseIf StrComp(Replace(NormalizeConfigFormulaText(formula), " ", ""), Replace(CStr(item(1)), " ", ""), vbTextCompare) <> 0 Then
        ReviewIssue = "REVIEW: formula differs from the reviewed definition: " & label
    End If
End Function

Private Function ReconStage(ByVal label As String) As String
    If Left$(label, 7) = "IMPACT_" Or InStr(label, "_IMPACT_") > 0 Then
        ReconStage = "Shock"
    ElseIf InStr(label, "PRE_SHOCK") > 0 Or Left$(label, 4) = "PRE_" Or InStr(label, "RATIO_STAGE") > 0 Then
        ReconStage = "Pre-shock"
    Else
        ReconStage = "Post-shock"
    End If
End Function

Private Function FieldTolerance(ByVal label As String) As Double
    If InStr(label, "RATIO") > 0 Or InStr(label, "PCT") > 0 Or Right$(label, 4) = "_CAR" Or label = "LCR" Or label = "NSFR" Or label = "IMPACT_LCR" Or label = "IMPACT_NSFR" Then
        FieldTolerance = mRateTolerance
    Else
        FieldTolerance = mTolerance
    End If
End Function

Private Function CompareValue(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String, ByRef reason As String) As String
    If Not ReconNumber(expected) Then
        CompareValue = "BLOCKED": reason = "Expected value is missing or could not be calculated."
    ElseIf Not ReconNumber(actual) Then
        CompareValue = "BLOCKED": reason = "System value is missing or non-numeric."
    ElseIf Abs(CDbl(actual) - CDbl(expected)) > FieldTolerance(label) Then
        CompareValue = "FAIL": reason = "System minus expected exceeds the absolute tolerance."
    Else
        CompareValue = "PASS": reason = "Independent value agrees within tolerance."
    End If
End Function

Private Sub AddLine(ByVal verdict As String, ByVal stage As String, ByVal tc As Object, ByVal el As Object, ByVal severity As String, ByVal label As String, ByVal expected As Variant, ByVal actual As Variant, ByVal formula As String, ByVal reason As String, ByVal source As String, ByVal filter As String, ByVal matched As Double)
    Dim Delta As Variant
    If ReconNumber(expected) And ReconNumber(actual) Then Delta = CDbl(actual) - CDbl(expected)
    If IsError(expected) Then expected = "Unavailable"
    If IsError(actual) Then actual = "Error in system output"
    mRows.Add Array(verdict, stage, tc("Date"), tc("EntityCode") & " / " & tc("EntityID"), tc("TestCaseCode"), el("ScenarioCode"), severity, label, expected, actual, Delta, FieldTolerance(label), formula, reason, source, filter, matched, mRun)
End Sub

Private Sub WriteReconDetail()
    Dim ws As Worksheet, out As Variant, line As Variant, i As Long, j As Long, hdr As Variant
    Set ws = GetWorksheetSafe(ThisWorkbook, detail)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(HOME)): ws.name = detail
    ResultsResetSheet ws
    ws.Cells.UnMerge: ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB": ws.Range("A3").Value2 = "Reconciliation evidence"
    ws.Range("A4").Value2 = "Run " & mRun & ". Filter by status, stage, case, severity or field. Difference = system minus expected."
    hdr = Array("Status", "Stage", "As-of date", "Entity", "Test case", "Element", "Severity", "Field", "Expected", "System", "Difference", "Tolerance", "Calculation", "Reason / next action", "Source", "Filter applied", "Selected rows", "Run timestamp")
    For j = 0 To UBound(hdr): ws.Cells(7, j + 1).Value2 = hdr(j): Next j
    If mRows.count > 0 Then
        ReDim out(1 To mRows.count, 1 To 18)
        For Each line In mRows
            i = i + 1
            For j = 0 To 17: out(i, j + 1) = line(j): Next j
        Next line
        WriteLiteralValues ws.Cells(8, 1).Resize(mRows.count, 18), out
    End If
    StyleReconTable ws, mRows.count + 7, 18
    ws.columns("A:G").ColumnWidth = 16: ws.columns("H").ColumnWidth = 38
    ws.columns("I:L").ColumnWidth = 18: ws.columns("M:N").ColumnWidth = 58
    ws.columns("O").ColumnWidth = 22: ws.columns("P").ColumnWidth = 55: ws.columns("Q:R").ColumnWidth = 22
    ws.Range("I8:K" & mRows.count + 8).NumberFormat = "#,##0.00;[Red](#,##0.00);0.00"
    ws.Range("L8:L" & mRows.count + 8).NumberFormat = "0.000000"
    ws.rows("8:" & mRows.count + 8).RowHeight = 25
    ResultsDecorate ws, mRows.count, mRun
End Sub

Public Sub StyleReconTable(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal columns As Long)
    Dim rg As Range, fc As FormatCondition, status As Variant
    ws.UsedRange.Font.name = "Aptos": ws.UsedRange.Font.Size = 10
    ws.UsedRange.Font.Color = RGB(14, 34, 64)
    ws.Range("A3").Font.Size = 22: ws.Range("A3").Font.Bold = True
    ws.rows(3).RowHeight = 34: ws.rows(4).RowHeight = 30: ws.Range("A4").Font.Color = RGB(102, 112, 133)
    Set rg = ws.Range(ws.Cells(7, 1), ws.Cells(7, columns))
    rg.Interior.Color = RGB(14, 34, 64): rg.Font.Color = vbWhite: rg.Font.Bold = True
    ws.rows(7).RowHeight = 30
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(7, 1), ws.Cells(IIf(lastRow > 7, lastRow, 8), columns)).AutoFilter
    Set rg = ws.Range("A8:A" & IIf(lastRow > 7, lastRow, 8))
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""PASS""")
    fc.Interior.Color = RGB(236, 253, 243): fc.Font.Color = RGB(6, 118, 71)
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""FAIL""")
    fc.Interior.Color = RGB(254, 243, 242): fc.Font.Color = RGB(180, 35, 24)
    Set fc = rg.FormatConditions.Add(xlCellValue, xlEqual, "=""BLOCKED""")
    fc.Interior.Color = RGB(253, 246, 227): fc.Font.Color = RGB(176, 138, 46)
    ws.Activate: ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False: ActiveWindow.splitRow = 7: ActiveWindow.SplitColumn = 2: ActiveWindow.FreezePanes = True
    ActiveWindow.zoom = 85
End Sub

Private Sub UpdateReconSummary(ByVal state As String)
    Dim line As Variant, pass As Long, fail As Long, blocked As Long, assumptions As Long
    For Each line In mRows
        Select Case line(0)
            Case "PASS": pass = pass + 1
            Case "FAIL": fail = fail + 1
            Case "ASSUMPTION": assumptions = assumptions + 1
            Case Else: blocked = blocked + 1
        End Select
    Next line
    HomeSetStatus IIf(fail + blocked > 0 Or pass = 0, "Review required", "Reconciled"), _
                  "Run " & mRun & ":  " & format$(mRows.count, "#,##0") & " checks, " & format$(pass, "#,##0") & " reconciled, " & _
                  format$(fail, "#,##0") & " differences, " & format$(blocked, "#,##0") & " open, " & assumptions & " system-supplied assumptions."
    RefreshHome
End Sub

Public Sub ShowReconDetail()
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, detail)
    If ws Is Nothing Then OpenReconWorkbench: Exit Sub
    ws.Activate
    ResultsShowAll ws
End Sub

Public Sub ShowReconExceptions()
    ShowReconDetail
    If ActiveSheet.name = detail Then ActiveSheet.Range("A7").AutoFilter field:=1, Criteria1:="<>PASS", Operator:=xlAnd, Criteria2:="<>ASSUMPTION"
End Sub

Public Sub ShowShockCatalog()
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, CATALOG)
    If ws Is Nothing Then
        UiNotice "Field guide", "The shock design catalog is not installed in this copy of the tool."
    Else
        ws.Activate
    End If
End Sub

Public Sub ExportReconEvidence()
    Dim target As Variant, wb As Workbook, nm As Variant, ws As Worksheet, state As Object
    If GetWorksheetSafe(ThisWorkbook, detail) Is Nothing Then UiNotice "Export evidence", "There is no reconciliation to export yet.", , "Click Reconcile on the home screen first.": Exit Sub
    target = Application.GetSaveAsFilename("JKB_Reconciliation_" & format$(Now, "yyyymmdd_hhnnss") & ".xlsx", "Excel workbook (*.xlsx), *.xlsx")
    If VarType(target) = vbBoolean Then Exit Sub
    On Error GoTo Failed
    ExportReconEvidenceTo CStr(target)
    Exit Sub
Failed:
    UiProblem "Export evidence", "The evidence file could not be written.", Err.description
End Sub

' The evidence workbook, written to a given path. Part of the review pack; raises on failure.
Public Sub ExportReconEvidenceTo(ByVal target As String)
    Dim wb As Workbook, nm As Variant, ws As Worksheet, state As Object, er As String
    On Error GoTo Failed
    Set state = CaptureState(): Application.EnableEvents = False
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    For Each nm In Array(HOME, detail, CATALOG, "Rule_Register", PS_SOURCES_SHEET, PS_METRICS_SHEET, DERIVED_SHEET)
        Set ws = GetWorksheetSafe(ThisWorkbook, CStr(nm))
        If Not ws Is Nothing Then
            ws.Copy After:=wb.Worksheets(wb.Worksheets.count)
            With wb.Worksheets(wb.Worksheets.count)
                ' Some of these are hidden in the master; a copy keeps that, the evidence must not.
                .Visible = xlSheetVisible
                Do While .Shapes.count > 0: .Shapes(1).Delete: Loop
                .UsedRange.Value2 = .UsedRange.Value2
            End With
        End If
    Next nm
    Application.DisplayAlerts = False: wb.Worksheets(1).Delete: Application.DisplayAlerts = True
    UiFinishReportBook wb, "Reconciliation evidence", "Home, detail, catalog, rules, sources, metrics and derived values at the time of export"
    wb.SaveAs fileName:=target, FileFormat:=xlOpenXMLWorkbook
    wb.Close False: RestoreState state
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    RestoreState state
    On Error GoTo 0
    Err.Raise vbObjectError + 1744, , "The evidence workbook could not be written: " & er
End Sub

Public Function ReconWorkbenchRegressionTests() As String
    Dim n As Long, emptyValue As Variant, reason As String, rules As Object, cache As Object
    Dim tc As Object, el As Object, memo As Object, visiting As Object, rule As Object, rec As Object
    Dim issue As String, oldReview As Object, oldParams As Object
    Set oldReview = mReview: Set oldParams = mParamIssues
    mTolerance = 0.005: mRateTolerance = 0.000001
    ReconCheck ReconNumber(0), "Genuine zero is numeric", n
    ReconCheck Not ReconNumber(emptyValue), "Missing is not zero", n
    ReconCheck Not ReconNumber(""), "Empty string is not zero", n
    ReconCheck Not ReconNumber(CVErr(xlErrNA)), "Error is not zero", n
    ReconCheck CompareValue(100, 100.01, "AMOUNT", reason) = "FAIL", "Amount difference", n
    ReconCheck CompareValue(0.9, 0.900002, "IMPACT_LCR", reason) = "FAIL", "LCR rate tolerance", n
    ReconCheck CompareValue(0, emptyValue, "AMOUNT", reason) = "BLOCKED", "Missing system amount", n
    ReconCheck Len(ParameterIssue("PCT_CHANGE", emptyValue)) > 0, "Missing shock magnitude", n
    ReconCheck Len(ParameterIssue("ECL_METHOD", "unknown")) > 0, "Unknown ECL method", n
    ReconCheck ReconUseSource("PRE_LCR_HQLA", "systemoutput.PRE_LCR_HQLA"), "Promote pre-shock copy", n
    ReconCheck Not ReconUseSource("CET1_CAPITAL_LCY", "base.CET1_CAPITAL_LCY+manual.IMPACT_CET1_CAPITAL_LCY"), "Retain post-shock calculation", n
    Set mReview = NewMap(): mReview("TYPE|X") = Array("ALIGNED", "manual.PRE_X")
    ReconCheck Len(ReviewIssue("TYPE", "X", "0")) > 0, "Changed formula invalidates approval", n
    Set rules = NewMap(): Set cache = NewMap(): Set tc = NewMap(): Set el = NewMap()
    tc("Date") = "2025-12-31": tc("EntityID") = "101": tc("EntityCode") = "JKB": tc("TestCaseCode") = "TC"
    el("ScenarioCode") = "EL": el("ElementType") = "TYPE"
    Set mParamIssues = NewMap(): mParamIssues("PCT_CHANGE") = "Missing shock magnitude"
    Set rule = NewMap(): rule("ManualFormulaDefault") = "system.PCT_CHANGE"
    rule("ManualFormulaModerate") = "": rule("ManualFormulaMedium") = "": rule("ManualFormulaSevere") = ""
    Set rules("X") = rule: Set memo = NewMap(): Set visiting = NewMap()
    issue = FormulaProof("X", 0, rules, cache, tc, el, memo, visiting)
    ReconCheck InStr(issue, "Missing shock") > 0, "Propagate missing parameter", n
    rule("ManualFormulaDefault") = "differences.Y": Set memo = NewMap(): Set visiting = NewMap()
    issue = FormulaProof("X", 0, rules, cache, tc, el, memo, visiting)
    ReconCheck InStr(issue, "reconciliation difference") > 0, "Reject difference dependency", n
    Set rec = NewMap(): rec("HasPre") = True: rec("Filter") = "ALL": rec("Rows") = 0
    Set cache("2025-12-31|101|JKB|TC|EL|PRE_X") = rec
    rule("ManualFormulaDefault") = "derived.PRE_X": Set memo = NewMap(): Set visiting = NewMap()
    issue = FormulaProof("X", 0, rules, cache, tc, el, memo, visiting)
    ReconCheck InStr(issue, "population evidence") > 0, "Reject empty derived population", n
    Set mReview = oldReview: Set mParamIssues = oldParams
    ReconWorkbenchRegressionTests = "PASS: " & n & " workbench regression checks"
End Function

Private Sub ReconCheck(ByVal condition As Boolean, ByVal caption As String, ByRef n As Long)
    If Not condition Then Err.Raise vbObjectError + 1749, , "Workbench regression: " & caption
    n = n + 1
End Sub

Public Function VerifyReconCoreQuiet(ByVal rulePath As String) As String
    On Error GoTo Failed
    VerifyReconCoreQuiet = ReconEngineRegressionTests() & vbCrLf & ReconCapitalDateRegressionTests() & vbCrLf & ReconDateCacheRegressionTests() & vbCrLf & ReconBreakdownCapacityRegressionTests() & vbCrLf & ReconSpeedEquivalenceTests() & vbCrLf & ReconTableHandoverTests() & vbCrLf & ReconWorkbenchRegressionTests() & vbCrLf & ManualReportRegressionTests() & vbCrLf & ReconRulesImportRegressionTests(rulePath)
    Exit Function
Failed:
    VerifyReconCoreQuiet = "ERROR: " & Err.description
End Function

Public Function ReconLoadSourceQuiet(ByVal path As String, ByVal sourceKey As String) As String
    On Error GoTo Failed
    LoadEclOutput path, False, sourceKey
    ReconLoadSourceQuiet = "Loaded " & SourceRowCount(sourceKey) & " " & sourceKey & " records."
    Exit Function
Failed:
    ReconLoadSourceQuiet = "ERROR: " & Err.description
End Function

Public Function ReconBuildPackQuiet(ByVal folder As String) As String
    On Error GoTo Failed
    ReconBuildPackQuiet = BuildManualReconPackTo(folder)
    Exit Function
Failed:
    ReconBuildPackQuiet = "ERROR: " & Err.description
End Function


Public Sub OpenDataManager()
    Dim ws As Worksheet, previous As String
    Set ws = GetWorksheetSafe(ThisWorkbook, "Data_Manager")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(HOME)): ws.name = "Data_Manager"
    previous = SafeText(ws.Range("D7").Value2)
    With ws
        .Range("B3:H19").UnMerge: .Range("B3:H19").ClearContents
        .columns("A").ColumnWidth = 3: .columns("B:H").ColumnWidth = 16
        .Range("A1:I21").Interior.Color = RGB(245, 247, 250)
        .Range("A1:I21").Font.name = "Aptos": .Range("A1:I21").Font.Size = 11
        .rows("3:19").RowHeight = 24
        .Range("B3:H4").Merge: .Range("B3").Value2 = "Data controls"
        .Range("B3:H4").Interior.Color = RGB(14, 34, 64)
        .Range("B3").Font.Color = vbWhite: .Range("B3").Font.Size = 25
        .Range("B5:H5").Merge: .Range("B5").Value2 = "Choose what to clear. A backup is saved before any change."
        .Range("B7:C7").Merge: .Range("B7").Value2 = "Clear selection"
        .Range("D7:H7").Merge: .Range("D7:H7").Interior.Color = vbWhite
        With .Range("D7").Validation
            .Delete
            .Add xlValidateList, xlValidAlertStop, xlBetween, "Calculated evidence,All input files,ECL input,CAPRWA input,LL input,LCR input,CAP input,System output,Run logs,All run data"
            .InCellDropdown = True: .ShowError = True
        End With
        .Range("D7").Value2 = IIf(Len(previous) > 0, previous, "Calculated evidence")
        .Range("B9:H12").Merge
        .Range("B9").Value2 = "Inputs: clear the selected loaded source and its saved picker values. Input or system changes also clear dependent calculation evidence. All run data includes inputs, system values, results and logs."
        .Range("B9").WrapText = True
        .Range("B17:H19").Merge
        .Range("B17").Value2 = "Rules, filters, formula cells and formatting are retained. Original input files and separately exported reports are not deleted. Cleared changes take effect immediately; save the workbook to keep them."
        .Range("B17").WrapText = True: .Range("B17").Font.Color = RGB(102, 112, 133)
        .Tab.Color = RGB(14, 116, 144)
    End With
    ReconButton ws, "DataClearConfirm", "Clear selected data", "ConfirmClearData", "B14:D15", True
    ReconButton ws, "DataClearBack", "Back to review", "OpenReconWorkbench", "F14:H15", False
    ws.Activate: ActiveWindow.DisplayGridlines = False: ActiveWindow.DisplayHeadings = False: ActiveWindow.zoom = 90
End Sub

Public Sub ConfirmClearData()
    Dim choice As String, backup As String, answer As Variant, oldEvents As Boolean
    If JKB_Busy Then Exit Sub
    On Error GoTo Failed
    choice = SafeText(ThisWorkbook.Worksheets("Data_Manager").Range("D7").Value2)
    If Not ValidClearChoice(choice) Then Err.Raise 5, , "Choose an option from the list."
    If Len(ThisWorkbook.path) = 0 Or ThisWorkbook.ReadOnly Then Err.Raise 5, , "Save a writable copy of the workbook first."
    backup = ThisWorkbook.path & Application.PathSeparator & "JKB_BeforeClear_" & format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    If Len(Dir$(backup)) > 0 Then Err.Raise 5, , "A backup with this timestamp already exists. Wait a second and retry."
    If Not UiAsk("Clear data", "Clear " & choice & "?", _
                 "Clearing inputs or system output also clears the calculation evidence that depends on them. " & _
                 "Rules, formulas and files outside the workbook are kept." & vbCrLf & vbCrLf & _
                 "A backup is saved first, to:" & vbCrLf & backup, True) Then Exit Sub
    If choice = "All run data" Then
        answer = Application.InputBox("Type CLEAR to confirm clearing all run data.", "Clear all run data", Type:=2)
        If VarType(answer) = vbBoolean Then Exit Sub
        If CStr(answer) <> "CLEAR" Then Exit Sub
    End If
    ThisWorkbook.SaveCopyAs backup
    oldEvents = Application.EnableEvents: Application.EnableEvents = False: JKB_Busy = True
    ApplyDataClear choice
    JKB_Busy = False: Application.EnableEvents = oldEvents
    OpenReconWorkbench
    HomeSetActivity "Cleared: " & choice & ". Backup: " & backup
    UiNotice "Clear data", "The data was cleared and a backup was saved.", , "Save the workbook to keep this change."
    Exit Sub
Failed:
    If JKB_Busy Then Application.EnableEvents = oldEvents
    JKB_Busy = False
    UiProblem "Clear data", "Clearing did not finish.", Err.description, IIf(Len(backup) > 0, "Your backup is at " & backup, "")
End Sub

Private Function ValidClearChoice(ByVal choice As String) As Boolean
    Select Case choice
        Case "Calculated evidence", "All input files", "ECL input", "CAPRWA input", "LL input", "LCR input", "CAP input", "System output", "Run logs", "All run data": ValidClearChoice = True
    End Select
End Function

Private Sub ClearDataRows(ByVal name As String, ByVal firstRow As Long)
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Set ws = GetWorksheetSafe(ThisWorkbook, name)
    If ws Is Nothing Then Exit Sub
    lastRow = ws.UsedRange.row + ws.UsedRange.rows.count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.columns.count - 1
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If lastRow >= firstRow Then ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).ClearContents
End Sub

Private Sub ApplyDataClear(ByVal choice As String)
    Dim allRun As Boolean, inputs As Boolean, system As Boolean, results As Boolean, ws As Worksheet, nm As Variant
    Dim firstRow As Long, firstCol As Long, lastRow As Long, constants As Range, clearCell As Range
    If Not ValidClearChoice(choice) Then Err.Raise 5, , "Unknown clear-data selection."
    allRun = (choice = "All run data")
    inputs = allRun Or choice = "All input files" Or Right$(choice, 6) = " input"
    system = allRun Or choice = "System output"
    results = inputs Or system Or choice = "Calculated evidence"
    If inputs Then
        If allRun Or choice = "All input files" Then
            ClearLoadedSourceData "ALL"
        Else
            ClearLoadedSourceData Left$(choice, Len(choice) - 6)
        End If
    End If
    If system Then
        ClearDataRows SHEET_SOURCE, 4
        Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_DASHBOARD)
        If Not ws Is Nothing Then
            For Each clearCell In ws.Range("F18:F25"): clearCell.MergeArea.ClearContents: Next clearCell
            ws.Range("F25").Value2 = "System output cleared. Import a new file."
        End If
        ClearDataRows SHEET_SELECTION, 7
    End If
    If results Then
        For Each nm In Array("Recon_Detail", "Derived_Values", "Pre_Shock_Results", "Pre_Shock_Breakdown")
            ClearDataRows CStr(nm), 8
        Next nm
        Set ws = GetWorksheetSafe(ThisWorkbook, PS_CASES_SHEET)
        If Not ws Is Nothing Then ws.Range("J8:K3007").ClearContents
        Set mRows = Nothing: Set mReview = Nothing: Set mParamIssues = Nothing: mRun = ""
        For Each ws In ThisWorkbook.Worksheets
            firstRow = 0
            If SafeText(ws.Range("N1").Value2) = "OUTPUT" Then firstRow = 7
            If SafeText(ws.Range("C3").Value2) = "System Output" And SafeText(ws.Range("F3").Value2) = "Manual Calculation" Then firstRow = 5
            If firstRow > 0 Then
                firstCol = IIf(system, 2, 6): lastRow = ws.UsedRange.row + ws.UsedRange.rows.count - 1
                If lastRow >= firstRow Then
                    Set constants = Nothing
                    On Error Resume Next
                    Set constants = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, 11)).SpecialCells(xlCellTypeConstants)
                    On Error GoTo 0
                    If Not constants Is Nothing Then
                        For Each clearCell In constants.Cells: clearCell.MergeArea.ClearContents: Next clearCell
                    End If
                End If
            End If
        Next ws
        HomeSetStatus "Not run", "Data cleared. Run again before using results."
    End If
    If allRun Or choice = "Run logs" Then ClearDataRows SHEET_LOG, 4
End Sub

' Run only on a disposable verification copy, never on the user's workbook.
Public Function DataClearRegressionTests() As String
    Dim a(1 To 2, 1 To 1) As Variant, h As Object, ws As Worksheet, Failed As Boolean, stepName As String
    On Error GoTo TestFailed
    If InStr(1, ThisWorkbook.name, "ClearData_Test", vbTextCompare) = 0 Then Err.Raise 5, , "Requires a disposable ClearData_Test workbook."
    Set h = NewMap(): h("ACCOUNT_NUMBER") = 1: a(1, 1) = "ACCOUNT_NUMBER": a(2, 1) = "proof"
    RegisterTable "ECL", a, h, "fixture", "Data": RegisterTable "LL", a, h, "fixture", "Data"
    Set ws = ThisWorkbook.Worksheets(detail): ws.Range("A8").Value2 = "evidence"
    On Error Resume Next
    ApplyDataClear "INVALID": Failed = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    If Not Failed Or ws.Range("A8").Value2 <> "evidence" Then Err.Raise 5, , "Invalid choice changed data."
    stepName = "ECL input": On Error GoTo TestFailed
    ApplyDataClear "ECL input"
    If SourceLoaded("ECL") Or Not SourceLoaded("LL") Or Len(SafeText(ws.Range("A8").Value2)) > 0 Then Err.Raise 5, , "Selective input clear failed."
    stepName = "Run logs": ws.Range("A8").Value2 = "keep evidence": ApplyDataClear "Run logs"
    If ws.Range("A8").Value2 <> "keep evidence" Then Err.Raise 5, , "Log clear changed results."
    stepName = "System output": ApplyDataClear "System output"
    If Not SourceLoaded("LL") Or Len(SafeText(ThisWorkbook.Worksheets(SHEET_SOURCE).Range("A4").Value2)) > 0 Then Err.Raise 5, , "System clear changed inputs or retained output."
    stepName = "All run data": ApplyDataClear "All run data"
    If AnySourceLoaded() Then Err.Raise 5, , "All-data clear retained inputs."
    If HomeStatus() <> "Not run" Then Err.Raise 5, , "Stale status remained."
    DataClearRegressionTests = "PASS: invalid selection, selective input retention, dependent results, logs isolation, system-only input retention, all-run clearing and status reset."
    Exit Function
TestFailed:
    DataClearRegressionTests = "FAIL: " & stepName & " | " & Err.Number & " | " & Err.description
End Function


