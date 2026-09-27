Attribute VB_Name = "modScenarioBuilder_Manual"
Option Explicit

' ============================================================================
'  Manual reports for base and pre-shock.
'
'  The ALM tool answers "rebuild the reported number and show me the rows behind
'  it" with one table of facts and a set of PivotTables over it. This is the same
'  answer for the same question on this side, because it IS the same question:
'
'      the system said X for this metric, this test case and this element;
'      rebuilding it from the extract says Y; here is every place they differ.
'
'  One Data sheet, one shared pivot cache, and every report is that table with
'  different fields on the row and column axes. A reviewer gets the field list,
'  drill-down, slicers and Show Details - none of which a range of written cells
'  offers however carefully it is formatted.
'
'  The figures come from the Derived_Values store, so this works in a session
'  that has not re-uploaded the extracts: the store is the record, and it already
'  carries both numbers and which of the two the configuration was told to read.
' ============================================================================

Private Const M_COLS As Long = 20
' The breakdown table is a different grain, so it is a different width.
Private Const W_COLS As Long = 13
Private Const PIVOT_ROW As Long = 16

Private mSlicersAdded As Long
Private mCaseSheets As Object
Private mAmountTolerance As Double
Private mRateTolerance As Double
Private mJoinedEvidencePath As String
Private mPivotMeasureMap As Object
Private mBridgeScopeProof As String

' ===================== the entry point ======================================

Public Sub Ps_BuildManualReports()
    Dim st As Object, folder As String, path As String, n As Long

    If JKB_Busy Then Exit Sub
    On Error GoTo Failed

    n = PersistedDerivedCount()
    If n = 0 Then
        UiNotice "Pivot reports", "There is nothing to report on yet.", _
                 "Base and pre-shock are stored on " & DERIVED_SHEET & " as soon as they are rebuilt, and the reports read them from there.", _
                 "Load the input files, then click Run on the home screen."
        Exit Sub
    End If

    folder = ChooseFolder("Where should the manual report workbook be written?")
    If Len(folder) = 0 Then Exit Sub

    Set st = CaptureState(): JKB_Busy = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ProgressStart "Building the manual reports", n & " derived figure(s) into " & folder, 100

    ProgressStep "Laying out the table", 20
    path = BuildWorkbook(folder)

    ProgressDone "Written."
    RestoreState st: JKB_Busy = False
    LogIssue LOG_LEVEL_INFO, "Manual reports", "Written to " & path, DERIVED_SHEET
    UiNotice "Pivot reports", "The pivot reports are ready.", _
             path & vbCrLf & vbCrLf & "There is one Data sheet with a PivotTable for each view, plus one sheet per test case. " & _
             "The buttons at the top of each sheet change the view, and the slicers filter every figure at once."
    Exit Sub

Failed:
    ProgressFailed Err.description
    RestoreState st: JKB_Busy = False
    LogIssue LOG_LEVEL_ERROR, "Manual reports", Err.description, DERIVED_SHEET
    UiProblem "Pivot reports", "The pivot reports could not be built.", Err.description
End Sub

' The headless path: same work, a folder passed in, no dialog.
Public Function Ps_BuildManualReportsTo(ByVal folder As String) As String
    Ps_BuildManualReportsTo = BuildWorkbook(folder)
End Function

Private Function ChooseFolder(ByVal title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If Not .Show = -1 Then Exit Function
        ChooseFolder = .SelectedItems(1)
    End With
End Function

' ===================== the workbook =========================================

Private Function BuildWorkbook(ByVal folder As String) As String
    Dim wb As Workbook, wsData As Worksheet, pc As PivotCache, path As String
    Dim hasMacros As Boolean, ext As String, fmt As Long, cases As Variant, i As Long, nCases As Long
    Dim failure As String
    Dim wsWhy As Worksheet, pcWhy As PivotCache, wsBase As Worksheet, pcBase As PivotCache
    Dim wsExceptions As Worksheet, pcExceptions As PivotCache, scope As Variant, caseName As String

    On Error GoTo Failed
    mSlicersAdded = 0: Set mCaseSheets = NewMap()
    ReadReportTolerances
    modShared_Pivot.ClearPivotProblem

    Set wb = NewReportWorkbook(hasMacros)
    Do While wb.Worksheets.count > 1
        wb.Worksheets(wb.Worksheets.count).Delete
    Loop

    Set wsData = WriteData(wb, cases)
    Set pc = wb.PivotCaches.Create(xlDatabase, sourceData:=DataRange(wsData))
    On Error Resume Next
    pc.RefreshOnFileOpen = False
    Err.Clear
    On Error GoTo Failed

    Set wsBase = WriteBaseData(wb, wsData)
    Set pcBase = wb.PivotCaches.Create(xlDatabase, sourceData:=DataRange(wsBase))
    pivotSheet wb, pcBase, "Base_Reconciliation", "Bank base reconciliation", _
        "One bank base per date, entity, source and metric. Repeated case copies are removed; totals across metrics are disabled.", _
        "As-of date|Entity|Source|Metric", "", "Status", _
        "BASE derived>BASE derived >#,##0.########|BASE system>BASE system >#,##0.########|BASE difference>BASE difference >#,##0.########", _
        "As-of date|Entity|Status", ""

    pivotSheet wb, pc, "Reconciliation", "Pre-shock reconciliation", _
        "Calculation snapshot versus system output at the exact case / element / date / entity grain. Source-row pivots are in Joined_Input.xlsx.", _
        "As-of date|Entity|Test case|Test element|Metric", "", "Source|Status|Value to use", _
        "PRE-SHOCK derived>PRE-SHOCK derived >#,##0.########|PRE-SHOCK system>PRE-SHOCK system >#,##0.########|PRE-SHOCK difference>PRE-SHOCK difference >#,##0.########", _
        "Test case|Source|Status", ""

    pivotSheet wb, pc, "By_Metric", "Metric comparison by scope", _
        "Metric values retain their individual scope and stored unit. No totals combine different metrics or repeated bank bases.", _
        "Source|Metric|As-of date|Entity|Test case|Test element", "", "Status", _
        "PRE-SHOCK derived>PRE-SHOCK derived >#,##0.########|PRE-SHOCK system>PRE-SHOCK system >#,##0.########", _
        "Source|Test case|Status", ""

    pivotSheet wb, pc, "Base_vs_PreShock", "Bank base and selected population", _
        "Bank base and testcase-filtered pre-shock values are shown per metric. Bank base repeats for context and is never summed across cases.", _
        "As-of date|Entity|Test case|Test element|Metric", "", "Source|Status", _
        "BASE derived>BASE derived >#,##0.########|PRE-SHOCK derived>PRE-SHOCK derived >#,##0.########|Rows matched>Rows matched >#,##0", _
        "Test case|Metric|Status", ""

    Set wsExceptions = WriteExceptionData(wb, wsData)
    If wsExceptions.Cells(wsExceptions.rows.count, 5).End(xlUp).row > 1 Then
        Set pcExceptions = wb.PivotCaches.Create(xlDatabase, sourceData:=DataRange(wsExceptions))
        pivotSheet wb, pcExceptions, "Where_It_Breaks", "Exceptions requiring review", _
            "Missing inputs, unavailable derivations, zero selected rows and differences. PASS rows are excluded from this pivot's source table.", _
            "Status|As-of date|Entity|Test case|Test element|Metric", "", "Source|Value to use", _
            "BASE difference>BASE difference >#,##0.########|PRE-SHOCK derived>PRE-SHOCK derived >#,##0.########|PRE-SHOCK system>PRE-SHOCK system >#,##0.########|PRE-SHOCK difference>PRE-SHOCK difference >#,##0.########", _
            "Status|Test case", ""
    Else
        Set wsExceptions = NewReportSheet(wb, "Where_It_Breaks", "No exceptions in this calculation snapshot", _
            "All recorded base and pre-shock comparisons passed. Check Case_Coverage in Joined_Input.xlsx for unrepresented cases.")
        wsExceptions.Range("A7").Value2 = "No exception records. No PASS rows are shown as exceptions."
    End If

    ' WHY each of those figures is what it is - a second table, at a different
    ' grain, so it gets its own cache and its own sheet rather than being forced
    ' into the shape of the first one.
    ' No "Why" sheet: Trace and the per-test-case pivots explain a value.
    Set wsWhy = Nothing
    If Not wsWhy Is Nothing Then
        Set pcWhy = wb.PivotCaches.Create(xlDatabase, sourceData:=WhyRange(wsWhy))
        On Error Resume Next
        pcWhy.RefreshOnFileOpen = False
        Err.Clear
        On Error GoTo Failed
        WhySheet wb, pcWhy
    End If

    nCases = 0
    If IsArray(cases) Then
        If UBound(cases) >= LBound(cases) Then
            For i = LBound(cases) To UBound(cases)
                scope = Split(CStr(cases(i)), vbTab)
                ProgressDetail CStr(scope(2)) & " / " & CStr(scope(3)), 82# + 10# * (i - LBound(cases) + 1) / (UBound(cases) - LBound(cases) + 1)
                caseName = "Case_" & format$(nCases + 1, "000") & "_" & Left$(CStr(scope(2)), 20)
                caseName = UniqueSheetName(wb, caseName)
                mCaseSheets(CStr(cases(i))) = caseName
                pivotSheet wb, pc, caseName, CStr(scope(2)) & " / " & CStr(scope(3)), _
                    CStr(scope(0)) & " | " & CStr(scope(1)) & " | Exact scope; the applied condition is recorded on Data.", _
                    "Metric", "", "Test case|Test element|As-of date|Entity|Source|Status", _
                    "BASE derived>BASE derived >#,##0.########|BASE system>BASE system >#,##0.########|BASE difference>BASE difference >#,##0.########|PRE-SHOCK derived>PRE-SHOCK derived >#,##0.########|PRE-SHOCK system>PRE-SHOCK system >#,##0.########|PRE-SHOCK difference>PRE-SHOCK difference >#,##0.########", _
                    "Metric|Status", CStr(scope(2)), CStr(scope(3)), CStr(scope(0)), CStr(scope(1))
                nCases = nCases + 1
            Next i
        End If
    End If

    ProgressDetail "Retaining the base, pre-shock, shock and post-shock calculation evidence", 94
    CopyCalculationEvidence wb
    ProgressDetail "Comparing the native input pivots with manual and system values", 95
    WritePivotChecks wb, folder
    ProgressDetail "Preparing report navigation and source links", 98
    WriteIndex wb, cases, nCases, hasMacros
    LinkInputPivots wb, folder
    wb.Worksheets("Index").Move Before:=wb.Worksheets(1)
    wb.Worksheets("Index").Activate
    ProgressDetail "Writing the cover and the executive summary", 99
    AddExecutivePages wb
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Sheet1").Delete
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo Failed

    If hasMacros Then
        ext = ".xlsm": fmt = xlOpenXMLWorkbookMacroEnabled
    Else
        ext = ".xlsx": fmt = xlOpenXMLWorkbook
    End If
    path = JoinPath(folder, "Pre_Shock_Manual_Reports" & ext)
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    Err.Clear
    On Error GoTo Failed
    ProgressDetail "Saving the completed comparison report", 99
    UiFinishReportBook wb, "Pre-shock manual reports", "Base and pre-shock rebuilt from the input files, compared with the system"
    wb.SaveAs fileName:=path, FileFormat:=fmt, CreateBackup:=False
    wb.Close SaveChanges:=False
    If Len(modShared_Pivot.LastPivotProblem()) > 0 Then
        LogIssue LOG_LEVEL_WARN, "Manual reports", "Pivot furniture: " & modShared_Pivot.LastPivotProblem(), ""
    End If
    LogIssue LOG_LEVEL_INFO, "Manual reports", mSlicersAdded & " slicer(s) over " & (nCases + 4) & " pivot sheet(s).", path
    BuildWorkbook = path
    Exit Function
Failed:
    failure = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 1853, , "Manual report build: " & failure
End Function

' A report workbook is born from a template carrying modReport, so its buttons
' work on a machine that has never seen this tool. Where the template is missing
' the workbook is still produced - Excel's own slicers and field list need no
' macro - and the Index says so rather than leaving dead buttons behind.
Private Function NewReportWorkbook(ByRef hasMacros As Boolean) As Workbook
    Dim p As String
    p = ReportTemplatePath()
    If Len(p) > 0 Then
        On Error Resume Next
        Set NewReportWorkbook = Workbooks.Add(Template:=p)
        If Not NewReportWorkbook Is Nothing Then hasMacros = True
        Err.Clear
        On Error GoTo 0
    End If
    If NewReportWorkbook Is Nothing Then Set NewReportWorkbook = Workbooks.Add
End Function

' A copy beside the workbook if there is one, otherwise the copy carried inside
' it. Shipping this tool is shipping one file.
Public Function ReportTemplatePath() As String
    ReportTemplatePath = modShared_Template.TemplatePath()
End Function

' ===================== the one table ========================================

Private Function WriteData(ByVal wb As Workbook, ByRef casesOut As Variant) As Worksheet
    Dim ws As Worksheet, cache As Object, k As Variant, rec As Object
    Dim a As Variant, r As Long, n As Long, seen As Object, i As Long, out As Variant
    Dim baseStatus As String, preStatus As String, scopeKey As String

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    On Error Resume Next
    ws.name = "Data"
    Err.Clear
    On Error GoTo 0

    ws.Cells(1, 1).Resize(1, M_COLS).Value2 = Array( _
        "As-of date", "Entity", "Test case", "Test element", "Metric", "Source", _
        "Value to use", "Which one was used", "Rows matched", _
        "BASE derived", "BASE system", "BASE difference", _
        "PRE-SHOCK derived", "PRE-SHOCK system", "PRE-SHOCK difference", _
        "Status", "Filter applied", "Agrees", "BASE status", "PRE-SHOCK status")

    Set cache = LoadPersistedDerived()
    n = cache.count
    Set seen = NewMap()
    If n = 0 Then
        ws.Cells(2, 1).Value2 = "(nothing has been derived yet)"
        casesOut = Array()
        Set WriteData = ws
        Exit Function
    End If

    ReDim a(1 To n, 1 To M_COLS)
    r = 0
    For Each k In cache.keys
        Set rec = cache(k)
        r = r + 1
        a(r, 1) = fld(rec, "Date")
        a(r, 2) = fld(rec, "Entity") & " / " & fld(rec, "EntityID")
        a(r, 3) = fld(rec, "TestCase")
        a(r, 4) = fld(rec, "Element")
        a(r, 5) = fld(rec, "Metric")
        a(r, 6) = fld(rec, "Source")
        a(r, 7) = BlankAs(fld(rec, "Use"), USE_DERIVED)
        a(r, 8) = BlankAs(fld(rec, "UsedPre"), USE_DERIVED)
        a(r, 9) = Num(rec, "Rows")
        If Flag(rec, "HasBase") Then a(r, 10) = NumOrEmpty(rec, "BaseDerived")
        a(r, 11) = NumOrEmpty(rec, "BaseSystem")
        If Flag(rec, "HasBase") Then a(r, 12) = Diff(rec, "BaseSystem", "BaseDerived")
        If Flag(rec, "HasPre") Then a(r, 13) = NumOrEmpty(rec, "PreDerived")
        a(r, 14) = NumOrEmpty(rec, "PreSystem")
        If Flag(rec, "HasPre") Then a(r, 15) = Diff(rec, "PreSystem", "PreDerived")
        baseStatus = StageStatus(rec, False): preStatus = StageStatus(rec, True)
        a(r, 16) = CombinedStatus(baseStatus, preStatus)
        a(r, 17) = fld(rec, "Filter")
        a(r, 18) = IIf(a(r, 16) = "PASS", "PASS", "REVIEW")
        a(r, 19) = baseStatus: a(r, 20) = preStatus
        scopeKey = CStr(a(r, 1)) & vbTab & CStr(a(r, 2)) & vbTab & CStr(a(r, 3)) & vbTab & CStr(a(r, 4))
        If Not seen.Exists(scopeKey) Then seen(scopeKey) = True
    Next k
    ws.columns("A").NumberFormat = "@"
    WriteLiteralValues ws.Cells(2, 1).Resize(n, M_COLS), a

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, M_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(14, 34, 64)
    End With
    ws.Range(ws.Cells(2, 9), ws.Cells(n + 1, 15)).NumberFormat = "#,##0.########"
    ws.columns("A:H").ColumnWidth = 24
    ws.columns("I:O").ColumnWidth = 18
    ws.columns("P:T").ColumnWidth = 26
    ws.rows(1).RowHeight = 24
    WriteViewSpecs ws

    ' Test cases in the order they were derived, which follows the source sheet.
    ReDim out(0 To WorksheetFunction.Max(seen.count - 1, 0))
    i = 0
    For Each k In seen.keys
        out(i) = CStr(k): i = i + 1
    Next k
    casesOut = out
    Set WriteData = ws
End Function

' ===================== why the figures are what they are ====================

' The breakdown store, copied in as a second fact table. Copied rather than
' referenced because a report workbook has to answer this on a machine that has
' never seen the tool, let alone the extract.
Private Function WriteWhyData(ByVal wb As Workbook) As Worksheet
    Dim src As Worksheet, ws As Worksheet, lastR As Long, n As Long, a As Variant, i As Long

    On Error GoTo Failed
    Set src = BreakdownSheet()
    If src Is Nothing Then Exit Function
    lastR = src.Cells(src.rows.count, 6).End(xlUp).row
    If lastR < PS_FIRST_ROW Then Exit Function
    n = lastR - PS_FIRST_ROW + 1

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.name = "Data_Why"
    ws.Cells(1, 1).Resize(1, W_COLS).Value2 = Array( _
        "As-of date", "Entity", "Test case", "Test element", "Source", _
        "Dimension", "Value", "Rows selected", "Rows in portfolio", _
        "Selected", "In portfolio", "Filter applied", "Measure")

    ' Columns 1-11 straight across; the filter text is column 17 on the store.
    a = src.Range(src.Cells(PS_FIRST_ROW, 1), src.Cells(lastR, 11)).Value2
    ws.Cells(2, 1).Resize(n, 11).Value2 = a
    a = src.Range(src.Cells(PS_FIRST_ROW, 17), src.Cells(lastR, 17)).Value2
    ws.Cells(2, 12).Resize(n, 1).Value2 = a
    a = src.Range(src.Cells(PS_FIRST_ROW, 15), src.Cells(lastR, 15)).Value2
    ws.Cells(2, 13).Resize(n, 1).Value2 = a

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, W_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(14, 34, 64)
    End With
    ws.Range(ws.Cells(2, 8), ws.Cells(n + 1, 11)).NumberFormat = "#,##0.########"
    ws.columns("A:G").ColumnWidth = 24
    ws.columns("H:K").ColumnWidth = 18
    ws.columns("L").ColumnWidth = 40: ws.columns("M").ColumnWidth = 30
    ws.columns("A").NumberFormat = "yyyy-mm-dd"
    ws.rows(1).RowHeight = 24
    Set WriteWhyData = ws
    Exit Function
Failed:
    On Error Resume Next
    LogIssue LOG_LEVEL_WARN, "Manual reports", "The breakdown could not be copied in: " & Err.description, "Data_Why"
End Function

Private Function WhyRange(ByVal ws As Worksheet) As Range
    Dim lastR As Long
    lastR = ws.Cells(ws.rows.count, 6).End(xlUp).row
    If lastR < 2 Then lastR = 2
    Set WhyRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, W_COLS))
End Function

' The sheet that answers "why is the pre-shock that much".
Private Sub WhySheet(ByVal wb As Workbook, ByVal pc As PivotCache)
    Dim ws As Worksheet, pt As PivotTable
    Set ws = NewReportSheet(wb, "Why", "Why the pre-shock is what it is", _
        "Selected amounts beside the bank base, separated by source, measure and dimension. Use source-pivot drilldown for the contributing input records.")
    modShared_Pivot.LayoutReportSheet ws
    On Error GoTo Done
    Set pt = modShared_Pivot.AddPivotFrom(pc, ws.Cells(modShared_Pivot.RPT_PIVOT_ROW, 1), "Why", _
        "As-of date|Entity|Test case|Test element|Source|Measure|Dimension|Value", "", "", _
        "Selected>Selected >#,##0|In portfolio>In portfolio >#,##0")
    If pt Is Nothing Then GoTo Done
    modShared_Pivot.AddRatio pt, "Share selected", "Selected", "In portfolio"
    SafeMetricPivot pt
    mSlicersAdded = mSlicersAdded + modShared_Pivot.AddSlicers(pt, "Dimension|Test case|Source", 6, _
                        modShared_Pivot.RPT_SLICER_Y, modShared_Pivot.RPT_SLICER_W, modShared_Pivot.RPT_SLICER_H)
    ws.columns("A:B").ColumnWidth = 34
    StyleManualPivotColumns pt
    modReconWorkbench.PositionPivotSlicers ws
Done:
    Err.Clear
    On Error Resume Next
    InstallButtons ws
    modShared_Pivot.LayoutReportSheet ws
    modShared_Pivot.FreezeReportSheet ws
    ws.Cells(1, 1).Select
    Err.Clear
End Sub

Private Sub InstallWhyButtons(ByVal ws As Worksheet)
    Dim x As Double, y1 As Double, y2 As Double
    y1 = modShared_Pivot.RPT_BTN1_Y
    y2 = modShared_Pivot.RPT_BTN2_Y
    x = 6
    x = Btn(ws, "By dimension", "RptView9", x, y1, 84, "Rows: dimension, then value")
    x = Btn(ws, "By test case", "RptView10", x, y1, 82, "Rows: test case, element, dimension, value")
    x = Btn(ws, "By value", "RptView11", x, y1, 66, "Rows: value alone, across every test case")
    x = x + 10
    x = Btn(ws, "Across cases", "RptView4", x, y1, 84, "Columns: one per test case")
    x = Btn(ws, "No columns", "RptView5", x, y1, 72, "Columns: none, totals only")

    x = 6
    x = Btn(ws, "Expand all", "RptExpand", x, y2, 70, "Open every level")
    x = Btn(ws, "Collapse", "RptCollapse", x, y2, 62, "Back to the top level")
    ' Mixed-unit metric snapshots have no subtotal shortcut.
    x = x + 10
    x = Btn(ws, "Row counts", "RptView12", x, y2, 74, "How many rows, rather than how much money")
    x = x + 10
    x = ZoomButtons(ws, x, y2)
    x = x + 10
    x = Btn(ws, "Index", "RptIndex", x, y2, 48, "Back to the index")
    x = Btn(ws, "Data", "RptData", x, y2, 48, "The table every figure here is built from")
End Sub

Private Function DataRange(ByVal ws As Worksheet) As Range
    Dim lastR As Long
    lastR = ws.Cells(ws.rows.count, 5).End(xlUp).row
    If lastR < 2 Then lastR = 2
    Set DataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, M_COLS))
End Function

' A record read back from an older store does not carry every key, and asking a
' Dictionary for a missing key silently ADDS it - so every optional read goes
' through here rather than growing the record it was only meant to inspect.
Private Function fld(ByVal rec As Object, ByVal key As String) As String
    If rec.Exists(key) Then fld = SafeText(rec(key))
End Function

Private Function BlankAs(ByVal s As String, ByVal ifEmpty As String) As String
    If Len(Trim$(s)) = 0 Then BlankAs = ifEmpty Else BlankAs = s
End Function

Private Function Num(ByVal rec As Object, ByVal key As String) As Double
    If Not rec.Exists(key) Then Exit Function
    If ReconNumber(rec(key)) Then Num = CDbl(rec(key))
End Function

Private Function NumOrEmpty(ByVal rec As Object, ByVal key As String) As Variant
    If Not rec.Exists(key) Then Exit Function
    If ReconNumber(rec(key)) Then NumOrEmpty = CDbl(rec(key))
End Function

Private Function Diff(ByVal rec As Object, ByVal sysKey As String, ByVal drvKey As String) As Variant
    If Not rec.Exists(sysKey) Or Not rec.Exists(drvKey) Then Exit Function
    If Not ReconNumber(rec(sysKey)) Or Not ReconNumber(rec(drvKey)) Then Exit Function
    Diff = CDbl(rec(sysKey)) - CDbl(rec(drvKey))
End Function

' ===================== the pivot sheets =====================================

Private Function pivotSheet(ByVal wb As Workbook, ByVal pc As PivotCache, ByVal nm As String, _
                            ByVal title As String, ByVal about As String, _
                            ByVal rowFields As String, ByVal colFields As String, ByVal filterFields As String, _
                            ByVal dataFields As String, ByVal slicerFields As String, _
                            ByVal caseFilter As String, Optional ByVal elementFilter As String = "", _
                            Optional ByVal dateFilter As String = "", Optional ByVal entityFilter As String = "") As Worksheet
    Dim ws As Worksheet, pt As PivotTable, failure As String
    Set ws = NewReportSheet(wb, nm, title, about)
    modShared_Pivot.LayoutReportSheet ws
    On Error GoTo Failed
    Set pt = modShared_Pivot.AddPivotFrom(pc, ws.Cells(modShared_Pivot.RPT_PIVOT_ROW, 1), nm, _
                                          rowFields, colFields, filterFields, dataFields)
    If pt Is Nothing Then Err.Raise vbObjectError + 1851, , "Pivot could not be created: " & nm
    SafeMetricPivot pt
    mSlicersAdded = mSlicersAdded + modShared_Pivot.AddSlicers(pt, slicerFields, 6, _
                        modShared_Pivot.RPT_SLICER_Y, modShared_Pivot.RPT_SLICER_W, modShared_Pivot.RPT_SLICER_H)
    If Len(caseFilter) > 0 Then SetPage pt, "Test case", caseFilter
    If Len(elementFilter) > 0 Then SetPage pt, "Test element", elementFilter
    If Len(dateFilter) > 0 Then SetPage pt, "As-of date", dateFilter
    If Len(entityFilter) > 0 Then SetPage pt, "Entity", entityFilter
    If nm = "Where_It_Breaks" Then HideMatching pt
    ws.columns("A:D").ColumnWidth = 30
    StyleManualPivotColumns pt
    modReconWorkbench.PositionPivotSlicers ws
Done:
    Err.Clear
    On Error Resume Next
    InstallButtons ws
    ' Restated AFTER the pivot lands: adding one can change a row height, and the
    ' buttons and slicers are positioned in points against this band.
    modShared_Pivot.LayoutReportSheet ws
    modShared_Pivot.FreezeReportSheet ws
    ws.Cells(1, 1).Select
    Err.Clear
    Set pivotSheet = ws
    Exit Function
Failed:
    failure = Err.description
    Err.Raise vbObjectError + 1852, , "Report " & nm & ": " & failure
End Function

Private Sub StyleManualPivotColumns(ByVal pt As PivotTable)
    Dim pf As PivotField, dataColumn As Range, targetWidth As Double
    If pt Is Nothing Then Exit Sub
    For Each pf In pt.rowFields
        targetWidth = 24
        Select Case pf.sourceName
            Case "As-of date": targetWidth = 14
            Case "Entity": targetWidth = 22
            Case "Test case": targetWidth = 16
            Case "Test element": targetWidth = 24
            Case "Source": targetWidth = 26
            Case "Metric": targetWidth = 34
            Case "Measure": targetWidth = 28
            Case "Dimension": targetWidth = 20
            Case "Value": targetWidth = 28
            Case "Status": targetWidth = 38
        End Select
        pf.DataRange.EntireColumn.ColumnWidth = targetWidth
    Next pf
    If pt.DataBodyRange Is Nothing Then Exit Sub
    For Each dataColumn In pt.DataBodyRange.columns
        dataColumn.EntireColumn.AutoFit
        dataColumn.EntireColumn.ColumnWidth = Application.Max(20, dataColumn.EntireColumn.ColumnWidth + 1)
    Next dataColumn
End Sub

Private Sub SetPage(ByVal pt As PivotTable, ByVal field As String, ByVal value As String)
    Dim pf As PivotField, item As PivotItem, nativeValue As String
    Set pf = pt.PivotFields(field)
    pf.orientation = xlPageField: pf.ClearAllFilters
    nativeValue = value
    If StrComp(field, "As-of date", vbTextCompare) = 0 Then
        nativeValue = ""
        For Each item In pf.PivotItems
            If SameReportDate(CStr(item.name), value) Then nativeValue = item.name: Exit For
        Next item
        If Len(nativeValue) = 0 Then Err.Raise vbObjectError + 1850, , "Exact reporting date is unavailable: " & value
    End If
    pf.EnableMultiplePageItems = False: pf.CurrentPage = nativeValue
    If StrComp(CStr(pf.CurrentPage), nativeValue, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1850, , "Could not apply exact scope filter: " & field & " = " & value
End Sub

Private Sub HideMatching(ByVal pt As PivotTable)
    ' This pivot uses Data_Exceptions, physically excluding all PASS rows.
    ' Keep the filter visible for audit; it cannot silently fall back to all facts.
    SetPage pt, "Agrees", "REVIEW"
End Sub

Private Sub SafeMetricPivot(ByVal pt As PivotTable)
    Dim pf As PivotField, i As Long, caption As String
    If pt Is Nothing Then Exit Sub
    pt.RowGrand = False: pt.ColumnGrand = False
    For Each pf In pt.rowFields
        For i = 1 To 12: pf.Subtotals(i) = False: Next i
    Next pf
    For Each pf In pt.dataFields
        ' Facts are already metric-level calculations, not additive source rows.
        ' IsCalculated is supported on the source field, not its data-field view.
        caption = pf.name
        If Not pt.PivotFields(pf.sourceName).IsCalculated Then pf.Function = xlMax
        pf.name = caption
        pf.NumberFormat = "#,##0.########"
    Next pf
    pt.DisplayNullString = True: pt.NullString = "Unavailable"
    pt.DisplayErrorString = True: pt.ErrorString = "Unavailable"
End Sub

' ===================== the index ============================================

Private Sub WriteIndex(ByVal wb As Workbook, ByVal cases As Variant, ByVal nCases As Long, ByVal hasMacros As Boolean)
    Dim ws As Worksheet, dat As Worksheet, r As Long, i As Long, nm As String, scope As Variant
    Dim checks As Long, passed As Long, review As Long
    Set dat = wb.Worksheets("Data")
    checks = Application.Max(0, dat.Cells(dat.rows.count, 5).End(xlUp).row - 1)
    If checks > 0 Then passed = Application.CountIf(dat.Range("P2:P" & checks + 1), "PASS")
    review = checks - passed
    Set ws = NewReportSheet(wb, "Index", "Stress testing | manual review", _
        "Prepared " & format$(Now, "dd mmm yyyy, hh:nn") & "  |  " & nCases & " case / element / date / entity scopes")
    ws.Cells(7, 1).Value2 = "METRIC PAIRS": ws.Cells(7, 2).Value2 = "VERIFIED PAIRS": ws.Cells(7, 3).Value2 = "NEEDS REVIEW"
    ws.Cells(8, 1).Value2 = checks: ws.Cells(8, 2).Value2 = passed: ws.Cells(8, 3).Value2 = review
    ws.Range("A7:C7").Font.Size = 9: ws.Range("A7:C7").Font.Bold = True
    ws.Range("A8:C8").Font.Size = 24: ws.Range("A8:C8").Font.Bold = True
    ws.Range("A7:C8").HorizontalAlignment = xlLeft
    ws.Range("A8:C8").NumberFormat = "#,##0"
    ws.Cells(9, 1).Value2 = "Includes differences and missing or unverified evidence."
    ws.Range("A9:C9").Merge: ws.Range("A9").Font.Size = 9
    ws.Range("A9").Font.Color = RGB(102, 112, 133)
    ws.Range("B8").Font.Color = RGB(6, 118, 71): ws.Range("C8").Font.Color = RGB(181, 71, 8)
    ws.rows(8).RowHeight = 34
    ws.Cells(10, 1).Value2 = "OPEN A REVIEW": ws.Cells(10, 1).Font.Bold = True
    ws.Cells(11, 1).Value2 = "Source pivots": ws.Cells(11, 2).Value2 = "Bank base, filtered testcase input rows and source coverage."
    r = 12
    IndexRow ws, r, "Base_Reconciliation", "Bank base versus system; one record per bank scope and metric.": r = r + 1
    IndexRow ws, r, "Reconciliation", "Filtered pre-shock versus system, by testcase and element.": r = r + 1
    IndexRow ws, r, "Where_It_Breaks", "Differences and missing evidence requiring review.": r = r + 1
    IndexRow ws, r, "Manual_Calcs", "Production manual formulas, expected results, system values and differences.": r = r + 1
    IndexRow ws, r, "Pivot_Checks", "Input-pivot amounts against derived and system amounts; mapping coverage is explicit.": r = r + 1
    IndexRow ws, r, "Base_vs_PreShock", "Bank base beside the selected population.": r = r + 1
    IndexRow ws, r, "Data", "Recorded source, exact filter, selected rows and comparison status.": r = r + 1
    If Not GetWorksheetSafe(wb, "Why") Is Nothing Then IndexRow ws, r, "Why", "Selected population by source dimension.": r = r + 1
    r = r + 2: ws.Cells(r, 1).Value2 = "TESTCASE SCOPES": ws.Cells(r, 1).Font.Bold = True: r = r + 1
    If IsArray(cases) Then
        For i = 0 To nCases - 1
            scope = Split(CStr(cases(i)), vbTab): nm = CaseSheetName(CStr(cases(i)))
            IndexRow ws, r, nm, scope(2) & " / " & scope(3) & "  |  " & scope(0) & "  |  " & scope(1)
            r = r + 1
        Next i
    End If
    ws.columns("A").ColumnWidth = 32: ws.columns("B").ColumnWidth = 92: ws.columns("C").ColumnWidth = 20
    ws.Range("A11:B" & r).RowHeight = 25
    ws.Range("A11:B" & r).Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
    InstallIndexButtons ws
End Sub

Private Sub IndexRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sheetName As String, ByVal what As String)
    ws.Cells(r, 1).Value2 = sheetName
    ws.Cells(r, 2).Value2 = what
    On Error Resume Next
    ws.Hyperlinks.Add anchor:=ws.Cells(r, 1), address:="", SubAddress:="'" & sheetName & "'!A1"
    Err.Clear
    On Error GoTo 0
End Sub

' ===================== the button strip =====================================

' The views this workbook offers, written onto the Data sheet as data and driven
' by the numbered RptView routines inside it. A button cannot pass an argument -
' Shape.OnAction silently refuses a macro name with one - so the field list has
' to travel on the sheet the report is already built on.
Private Sub WriteViewSpecs(ByVal ws As Worksheet)
    ws.Range("U1").Value2 = "Views the buttons apply (edit and press the button again):"
    ws.Range("U1").Font.Color = RGB(152, 162, 179)
    ws.Range("V1").Value2 = "rows=As-of date|Entity|Test case|Test element|Metric"
    ws.Range("V2").Value2 = "rows=Source|Metric|As-of date|Entity|Test case|Test element"
    ws.Range("V3").Value2 = "rows=Status|As-of date|Entity|Test case|Test element|Metric"
    ws.Range("V4").Value2 = "cols=Test case"
    ws.Range("V5").Value2 = "cols="
    ws.Range("V6").Value2 = "values=PRE-SHOCK derived|PRE-SHOCK system|PRE-SHOCK difference"
    ws.Range("V7").Value2 = "values=BASE derived|PRE-SHOCK derived|PRE-SHOCK difference"
    ws.Range("V8").Value2 = "values=Rows matched"
    ' 9 to 12 belong to the Why sheet, which is a different table with different
    ' fields. SetAxis skips a field the pivot has not got, so the two sets can
    ' share one strip of numbered routines without either knowing about the other.
    ws.Range("V9").Value2 = "rows=As-of date|Entity|Test case|Test element|Source|Dimension|Value"
    ws.Range("V10").Value2 = "rows=As-of date|Entity|Test case|Test element|Source|Dimension|Value"
    ws.Range("V11").Value2 = "rows=As-of date|Entity|Test case|Test element|Source|Dimension|Value"
    ws.Range("V12").Value2 = "values=Rows selected|Rows in portfolio"

    ' The fields the third strip offers, one per line: rows in W, columns in X.
    ' Both tables are listed - the reconciliation's and the Why sheet's - because
    ' a chip is drawn only where its field exists, so one list serves both.
    ws.Range("W1").Value2 = "Test case"
    ws.Range("W2").Value2 = "Test element"
    ws.Range("W3").Value2 = "Metric"
    ws.Range("W4").Value2 = "Source"
    ws.Range("W5").Value2 = "Status"
    ws.Range("W6").Value2 = "Agrees"
    ws.Range("W7").Value2 = "Dimension"
    ws.Range("W8").Value2 = "Value"
    ws.Range("X1").Value2 = "Test case"
    ws.Range("X2").Value2 = "Metric"
    ws.Range("X3").Value2 = "Source"
    ws.Range("X4").Value2 = "Dimension"
    ws.columns("U:V").ColumnWidth = 44
    ws.columns("W:X").ColumnWidth = 18
End Sub

Private Sub InstallButtons(ByVal ws As Worksheet)
    Dim x As Double, y As Double
    y = modShared_Pivot.RPT_BTN1_Y: x = 6
    x = Btn(ws, "Index", "RptIndex", x, y, 56, "Return to the review overview")
    x = Btn(ws, "Evidence", "RptData", x, y, 70, "Inspect filter conditions, selected rows and statuses")
    x = Btn(ws, "Expand rows", "RptExpand", x, y, 86, "Show every metric within this scope")
    x = x + 10: x = ZoomButtons(ws, x, y)
    ' Native filters and slicers remain available. Generic preset / value buttons
    ' can clear exact case filters or restore SUM over metric snapshots, so they
    ' are deliberately excluded from these controlled reconciliation reports.
End Sub

' The third strip: one chip per field, adding it to the arrangement or taking it
' away without disturbing the rest. The preset views above are whole
' arrangements; this is for the far commoner case of wanting the arrangement you
' are looking at, plus one more column.
Private Sub InstallFieldStrip(ByVal ws As Worksheet, ByVal y As Double)
    Dim dat As Worksheet, i As Long, nm As String, x As Double, pt As PivotTable
    On Error Resume Next
    Set dat = ws.Parent.Worksheets("Data")
    If dat Is Nothing Then Exit Sub
    If ws.PivotTables.count = 0 Then Exit Sub
    Set pt = ws.PivotTables(1)

    x = 6
    x = Lbl(ws, "Rows:", x, y, 34)
    For i = 1 To 8
        nm = Trim$(CStr(dat.Cells(i, 23).Value2))
        If PivotHasField(pt, nm) Then x = TogBtn(ws, "tog_R" & i, nm, "RptTogRow" & i, x, y, TogWidth(nm))
    Next i
    x = Btn(ws, "Clear", "RptClearRows", x, y, 44, "Take every field off the rows")

    x = x + 12
    x = Lbl(ws, "Columns:", x, y, 48)
    For i = 1 To 4
        nm = Trim$(CStr(dat.Cells(i, 24).Value2))
        If PivotHasField(pt, nm) Then x = TogBtn(ws, "tog_C" & i, nm, "RptTogCol" & i, x, y, TogWidth(nm))
    Next i
    x = Btn(ws, "Clear", "RptClearCols", x, y, 44, "Take every field off the columns")
    Err.Clear
End Sub

' A chip is drawn only for a field this report actually has. The reconciliation
' and the Why sheet are different tables sharing one strip, so most chips apply
' to one of them and not the other.
Private Function PivotHasField(ByVal pt As PivotTable, ByVal nm As String) As Boolean
    Dim pf As PivotField
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Set pf = Nothing
    Set pf = pt.PivotFields(nm)
    PivotHasField = Not (pf Is Nothing)
    Err.Clear
End Function

' A toggle button. NAMED, because modReport has to find it again to put the tick
' on it - Btn's positional name would change the moment a strip was re-laid out
' and the caption would silently stop tracking the pivot.
Private Function TogBtn(ByVal ws As Worksheet, ByVal shapeName As String, ByVal fieldName As String, _
                        ByVal proc As String, ByVal x As Double, ByVal y As Double, ByVal w As Double) As Double
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, modShared_Pivot.RPT_BTN_H)
    If sh Is Nothing Then TogBtn = x: Exit Function
    sh.name = shapeName
    sh.Placement = xlFreeFloating
    sh.fill.ForeColor.RGB = RGB(245, 247, 250)
    sh.line.ForeColor.RGB = RGB(208, 213, 221)
    sh.line.Weight = 0.75
    sh.Shadow.Visible = msoFalse
    sh.TextFrame2.textRange.text = "+ " & fieldName
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 2: .MarginRight = 2
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .textRange.ParagraphFormat.Alignment = msoAlignCenter
        .textRange.Font.Size = 8.5
        .textRange.Font.Bold = msoTrue
        .textRange.Font.fill.ForeColor.RGB = RGB(25, 63, 137)
    End With
    sh.AlternativeText = "Add or remove " & fieldName & ", leaving the rest of the arrangement alone"
    sh.OnAction = proc
    Err.Clear
    TogBtn = x + w + 4
End Function

' Roughly the width the caption needs, so the strip does not come out as a row of
' identical boxes with half the names clipped.
Private Function TogWidth(ByVal fieldName As String) As Double
    Dim w As Double
    w = 22 + Len(fieldName) * 5.4
    If w < 52 Then w = 52
    If w > 130 Then w = 130
    TogWidth = w
End Function

' A caption on the strip, not a button: it says what the chips beside it act on.
Private Function Lbl(ByVal ws As Worksheet, ByVal caption As String, _
                     ByVal x As Double, ByVal y As Double, ByVal w As Double) As Double
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, modShared_Pivot.RPT_BTN_H)
    If sh Is Nothing Then Lbl = x: Exit Function
    sh.name = "lbl_" & CLng(x) & "_" & CLng(y)
    sh.Placement = xlFreeFloating
    sh.fill.Visible = msoFalse
    sh.line.Visible = msoFalse
    sh.Shadow.Visible = msoFalse
    sh.TextFrame2.textRange.text = caption
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 0: .MarginRight = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .textRange.ParagraphFormat.Alignment = msoAlignRight
        .textRange.Font.Size = 8.5
        .textRange.Font.Bold = msoTrue
        .textRange.Font.fill.ForeColor.RGB = RGB(102, 112, 133)
    End With
    Err.Clear
    Lbl = x + w + 4
End Function

Private Sub InstallIndexButtons(ByVal ws As Worksheet)
    Dim x As Double, y1 As Double
    y1 = modShared_Pivot.RPT_BTN1_Y
    x = 6
    x = ZoomButtons(ws, x, y1)
    x = x + 10
    x = Btn(ws, "Fit", "RptFitBoth", x, y1, 40, "Size every column and row to what is in it")
    x = Btn(ws, "Data", "RptData", x, y1, 48, "The table every report is built from")
    x = Btn(ws, "Refresh all", "RptRefresh", x, y1, 72, "Rebuild every pivot from the Data sheet")
End Sub

Private Function ZoomButtons(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double) As Double
    Dim p As Double
    p = Btn(ws, "-", "RptZoomOut", x, y, 24, "Zoom out")
    p = Btn(ws, "+", "RptZoomIn", p, y, 24, "Zoom in")
    p = Btn(ws, "Zoom fit", "RptZoomFit", p, y, 58, "Fit the table to the window")
    p = Btn(ws, "100%", "RptZoom100", p, y, 40, "Back to actual size")
    p = Btn(ws, "Fit cols", "RptFitColumns", p, y, 54, "Size every column to what is in it")
    p = Btn(ws, "Fit rows", "RptFitRows", p, y, 54, "Size every row to what is in it")
    ZoomButtons = p
End Function

Private Function Btn(ByVal ws As Worksheet, ByVal caption As String, ByVal proc As String, _
                     ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal tip As String) As Double
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, modShared_Pivot.RPT_BTN_H)
    If sh Is Nothing Then Btn = x: Exit Function
    ' The action carries quotes and spaces - a shape name may not - so the name is
    ' built from the position, which is unique on the sheet by construction.
    sh.name = "btn_" & CLng(x) & "_" & CLng(y)
    sh.Placement = xlFreeFloating
    sh.fill.ForeColor.RGB = RGB(255, 255, 255)
    sh.line.ForeColor.RGB = RGB(183, 196, 216)
    sh.line.Weight = 0.75
    sh.Shadow.Visible = msoFalse
    sh.TextFrame2.textRange.text = caption
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 2: .MarginRight = 2
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .textRange.ParagraphFormat.Alignment = msoAlignCenter
        .textRange.Font.Size = 8.5
        .textRange.Font.Bold = msoTrue
        .textRange.Font.fill.ForeColor.RGB = RGB(18, 48, 107)
    End With
    sh.AlternativeText = tip
    sh.OnAction = proc
    Err.Clear
    Btn = x + w + 4
End Function

' ===================== shared sheet furniture ===============================

Private Function NewReportSheet(ByVal wb As Workbook, ByVal nm As String, ByVal title As String, ByVal about As String) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    On Error Resume Next
    ws.name = nm
    If Err.Number <> 0 Then
        Err.Clear
        ws.name = nm & "_" & wb.Worksheets.count
    End If
    On Error GoTo 0
    ws.Cells.Font.name = "Aptos"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(245, 247, 250)
    ws.Cells(1, 1).Value2 = title
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Color = RGB(14, 34, 64)
    ws.Cells(2, 1).Value2 = about
    ws.Cells(2, 1).Font.Size = 8.5
    ws.Cells(2, 1).Font.Color = RGB(102, 112, 133)
    ' Row heights belong to LayoutReportSheet and are NOT set here. Two places
    ' setting the same rows is how the band and the shapes drawn against it come
    ' apart - whichever ran last won, and it was not always the same one.
    ws.Tab.Color = RGB(14, 34, 64)
    Set NewReportSheet = ws
End Function

' Excel sheet names cannot exceed 31 characters or contain : \ / ? * [ ].
Private Function CaseSheetName(ByVal code As String) As String
    Dim s As String, bad As Variant, b As Variant
    If Not mCaseSheets Is Nothing Then
        If mCaseSheets.Exists(code) Then CaseSheetName = CStr(mCaseSheets(code)): Exit Function
    End If
    s = Replace(code, vbTab, "_")
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each b In bad
        s = Replace(s, CStr(b), "_")
    Next b
    If Len(s) > 31 Then s = Left$(s, 31)
    If Len(s) = 0 Then s = "case"
    CaseSheetName = s
End Function


Public Sub SetManualJoinedEvidence(ByVal path As String)
    mJoinedEvidencePath = path
End Sub

' Refresh the native-pivot comparison after a layout-only evidence upgrade.
' The dated calculation snapshot and its recorded statuses remain intact.
Public Function RefreshManualPivotEvidence(ByVal reportPath As String, ByVal joinedPath As String) As String
    Dim wb As Workbook, ws As Worksheet, folder As String, oldAlerts As Boolean, failure As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    ReadReportTolerances
    mJoinedEvidencePath = joinedPath
    folder = Left$(reportPath, InStrRev(reportPath, "\") - 1)
    Set wb = Application.Workbooks.Open(fileName:=reportPath, UpdateLinks:=0, ReadOnly:=False)
    Set ws = GetWorksheetSafe(wb, "Pivot_Checks")
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    WritePivotChecks wb, folder
    Set ws = wb.Worksheets("Pivot_Checks")
    If InStr(1, SafeText(ws.Range("N8").Value2), "could not complete", vbTextCompare) > 0 Then Err.Raise 5, , SafeText(ws.Range("N8").Value2)
    LinkInputPivots wb, folder
    wb.Worksheets("Index").Activate
    wb.Save: wb.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    RefreshManualPivotEvidence = "PASS: native input-pivot comparisons refreshed against the saved calculation snapshot."
    Exit Function
Failed:
    failure = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    RefreshManualPivotEvidence = "ERROR: " & failure
End Function

Private Sub ReadReportTolerances()
    Dim v As Variant
    mAmountTolerance = 0.005: mRateTolerance = 0.000001
    v = HomeTolerance(False)
    If ReconNumber(v) Then
        If CDbl(v) >= 0 Then mAmountTolerance = CDbl(v)
    End If
    v = HomeTolerance(True)
    If ReconNumber(v) Then
        If CDbl(v) >= 0 Then mRateTolerance = CDbl(v)
    End If
End Sub

Private Function ReportTolerance(ByVal metric As String) As Double
    metric = UCase$(metric)
    If InStr(metric, "RATIO") > 0 Or InStr(metric, "PCT") > 0 Or Right$(metric, 4) = "_CAR" Or _
       metric = "LCR" Or metric = "NSFR" Or metric = "IMPACT_LCR" Or metric = "IMPACT_NSFR" Then
        ReportTolerance = mRateTolerance
    Else
        ReportTolerance = mAmountTolerance
    End If
End Function

Private Function Flag(ByVal rec As Object, ByVal key As String) As Boolean
    If Not rec.Exists(key) Then Exit Function
    If IsNull(rec(key)) Or IsEmpty(rec(key)) Or IsError(rec(key)) Then Exit Function
    Flag = CBool(rec(key))
End Function

Private Function StageStatus(ByVal rec As Object, ByVal isPre As Boolean) As String
    Dim prefix As String, sysValue As Variant, derivedValue As Variant, statusText As String
    prefix = IIf(isPre, "Pre", "Base")
    If Not Flag(rec, "Has" & prefix) Then StageStatus = "BLOCKED - source derivation unavailable": Exit Function
    derivedValue = NumOrEmpty(rec, prefix & "Derived"): sysValue = NumOrEmpty(rec, prefix & "System")
    If Not ReconNumber(derivedValue) Then StageStatus = "BLOCKED - source value missing": Exit Function
    If Not ReconNumber(sysValue) Then StageStatus = "BLOCKED - system value missing": Exit Function
    If isPre Then
        If Len(Trim$(fld(rec, "Filter"))) = 0 Then StageStatus = "BLOCKED - filter missing": Exit Function
        If Num(rec, "Rows") <= 0 Then StageStatus = "BLOCKED - no selected rows": Exit Function
    End If
    If Abs(CDbl(sysValue) - CDbl(derivedValue)) > ReportTolerance(fld(rec, "Metric")) Then
        StageStatus = "FAIL"
    Else
        StageStatus = "PASS"
        If Len(fld(rec, "Source")) > 0 Then
            statusText = SourceDateIssue(fld(rec, "Source"))
            If Len(statusText) > 0 Then StageStatus = "REVIEW - source reporting date unverified"
        End If
    End If
End Function

Private Function CombinedStatus(ByVal baseStatus As String, ByVal preStatus As String) As String
    If Left$(baseStatus, 7) = "BLOCKED" Then
        CombinedStatus = "BASE: " & baseStatus
    ElseIf Left$(preStatus, 7) = "BLOCKED" Then
        CombinedStatus = "PRE: " & preStatus
    ElseIf baseStatus = "FAIL" Or preStatus = "FAIL" Then
        CombinedStatus = "FAIL"
    ElseIf baseStatus = "PASS" And preStatus = "PASS" Then
        CombinedStatus = "PASS"
    Else
        CombinedStatus = "REVIEW"
    End If
End Function

Private Function WriteExceptionData(ByVal wb As Workbook, ByVal src As Worksheet) As Worksheet
    Dim ws As Worksheet, a As Variant, out As Variant, i As Long, j As Long, n As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = "Data_Exceptions"
    ws.Cells(1, 1).Resize(1, M_COLS).Value2 = src.Cells(1, 1).Resize(1, M_COLS).Value2
    a = DataRange(src).Value2
    For i = 2 To UBound(a, 1)
        If Len(SafeText(a(i, 5))) > 0 And CStr(a(i, 16)) <> "PASS" Then n = n + 1
    Next i
    If n > 0 Then
        ReDim out(1 To n, 1 To M_COLS): n = 0
        For i = 2 To UBound(a, 1)
            If Len(SafeText(a(i, 5))) > 0 And CStr(a(i, 16)) <> "PASS" Then
                n = n + 1
                For j = 1 To M_COLS: out(n, j) = a(i, j): Next j
            End If
        Next i
        WriteLiteralValues ws.Cells(2, 1).Resize(n, M_COLS), out
    End If
    ws.rows(1).Font.Bold = True: ws.columns("A:T").ColumnWidth = 19
    Set WriteExceptionData = ws
End Function

Private Function WriteBaseData(ByVal wb As Workbook, ByVal src As Worksheet) As Worksheet
    Dim ws As Worksheet, d As Object, a As Variant, row As Variant, out As Variant, key As String, k As Variant
    Dim i As Long, j As Long, n As Long, prior As Variant
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = "Data_Base"
    ws.Cells(1, 1).Resize(1, M_COLS).Value2 = src.Cells(1, 1).Resize(1, M_COLS).Value2
    a = DataRange(src).Value2: Set d = NewMap()
    For i = 2 To UBound(a, 1)
        If Len(SafeText(a(i, 5))) > 0 Then
            key = CStr(a(i, 1)) & vbTab & CStr(a(i, 2)) & vbTab & CStr(a(i, 6)) & vbTab & CStr(a(i, 5))
            If Not d.Exists(key) Then
                ReDim row(1 To M_COLS)
                For j = 1 To M_COLS: row(j) = a(i, j): Next j
                row(3) = "Bank base": row(4) = "Portfolio": row(9) = Empty
                row(13) = Empty: row(14) = Empty: row(15) = Empty: row(16) = a(i, 19)
                row(17) = "Bank scope; no testcase filter": row(18) = IIf(row(16) = "PASS", "PASS", "REVIEW")
                row(20) = "Not applicable": d.Add key, row
            Else
                prior = d(key)
                If CStr(prior(10)) <> CStr(a(i, 10)) Or CStr(prior(11)) <> CStr(a(i, 11)) Or CStr(prior(19)) <> CStr(a(i, 19)) Then
                    prior(10) = Empty: prior(11) = Empty: prior(12) = Empty
                    prior(16) = "BLOCKED - inconsistent bank base across cases": prior(19) = prior(16): prior(18) = "REVIEW"
                    d(key) = prior
                End If
            End If
        End If
    Next i
    If d.count > 0 Then
        ReDim out(1 To d.count, 1 To M_COLS)
        For Each k In d.keys
            n = n + 1: row = d(k)
            For j = 1 To M_COLS: out(n, j) = row(j): Next j
        Next k
        ws.columns("A").NumberFormat = "@"
        WriteLiteralValues ws.Cells(2, 1).Resize(d.count, M_COLS), out
    End If
    ws.rows(1).Font.Bold = True: ws.columns("A:T").ColumnWidth = 19
    ws.columns("J:L").NumberFormat = "#,##0.########"
    Set WriteBaseData = ws
End Function

Private Sub CopyCalculationEvidence(ByVal wb As Workbook)
    Dim src As Worksheet, ws As Worksheet
    Set src = GetWorksheetSafe(ThisWorkbook, "Recon_Detail")
    If src Is Nothing Then
        Set ws = NewReportSheet(wb, "Manual_Calcs", "Manual calculation comparison", "Run reconciliation to populate production-formula evidence.")
        ws.Range("A7").Value2 = "BLOCKED - reconciliation has not been run."
        Exit Sub
    End If
    src.Copy After:=wb.Worksheets(wb.Worksheets.count)
    Set ws = wb.Worksheets(wb.Worksheets.count): ws.name = "Manual_Calcs"
    Do While ws.Shapes.count > 0: ws.Shapes(1).Delete: Loop
    ws.UsedRange.Value2 = ws.UsedRange.Value2
    If ws.FilterMode Then ws.ShowAllData
    ws.Range("A3").Value2 = "Manual formula comparison"
    ws.Range("A4").Value2 = "The production formulas used in scenario sheets: source-based expected result, system result, difference and lineage."
End Sub

Private Sub LinkInputPivots(ByVal wb As Workbook, ByVal folder As String)
    Dim ws As Worksheet, path As String, linkAddress As String, fso As Object
    Set ws = wb.Worksheets("Index")
    path = mJoinedEvidencePath
    If Len(path) = 0 Then path = JoinPath(folder, "Joined_Input.xlsx")
    If Len(Dir$(path)) > 0 Then
        linkAddress = path
        Set fso = CreateObject("Scripting.FileSystemObject")
        If StrComp(fso.GetFile(path).parentFolder.path, fso.GetFolder(folder).path, vbTextCompare) = 0 Then linkAddress = fso.GetFileName(path)
        ws.Hyperlinks.Add anchor:=ws.Range("A11"), address:=linkAddress, TextToDisplay:="Source pivots"
        ws.Range("B11").Value2 = "Open bank base, testcase source pivots and coverage checks."
    Else
        ws.Range("A11").Value2 = "Source pivots unavailable"
        ws.Range("B11").Value2 = "Build the complete manual report pack to include input-row pivot evidence."
    End If
End Sub


Private Sub WritePivotChecks(ByVal wb As Workbook, ByVal folder As String)
    Dim ws As Worksheet, joined As Workbook, coverage As Worksheet, metrics As Worksheet, dat As Worksheet
    Dim path As String, mapping As Object, formulas As Object, stages As Object, seen As Object, baseRecords As Object
    Dim a As Variant, cv As Variant, baseData As Variant, baseRecord As Variant, mappedCase As Variant
    Dim hdr As Variant, row As Long, i As Long, j As Long, stageNo As Long, splitAt As Long, whereAt As Long
    Dim k As String, label As String, formula As String, stage As String, measure As String, issue As String
    Dim pivotSheet As String, pv As Variant, dv As Variant, sv As Variant, result As String, fields As Object
    Dim physicalSource As String, sourcePredicate As String, membershipField As String, priorStatus As String
    On Error GoTo Failed
    Set ws = NewReportSheet(wb, "Pivot_Checks", "Input pivots | calculation comparison", _
        "Native input-pivot totals against derived and system values. Complex or unavailable mappings remain REVIEW.")
    hdr = Array("Stage", "As-of date", "Entity", "Test case", "Element", "Metric", "Source", "Input pivot", "Derived", "System", "Pivot minus derived", "System minus pivot", "Status", "Detail", "Native pivot sheet")
    For j = 0 To UBound(hdr): ws.Cells(7, j + 1).Value2 = hdr(j): Next j
    path = mJoinedEvidencePath: If Len(path) = 0 Then path = JoinPath(folder, "Joined_Input.xlsx")
    If Len(Dir$(path)) = 0 Then ws.Range("M8").Value2 = "BLOCKED": ws.Range("N8").Value2 = "Input-pivot workbook unavailable.": GoTo Finish
    ProgressDetail "Opening the saved input-pivot evidence", 95
    Set joined = Application.Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True)
    Set coverage = GetWorksheetSafe(joined, "Case_Coverage")
    Set dat = wb.Worksheets("Data"): a = DataRange(dat).Value2
    Set mapping = NewMap(): Set formulas = NewMap(): Set stages = NewMap(): Set seen = NewMap(): Set mPivotMeasureMap = NewMap()
    Set baseRecords = NewMap()
    baseData = DataRange(wb.Worksheets("Data_Base")).Value2
    For i = 2 To UBound(baseData, 1)
        If Len(SafeText(baseData(i, 5))) > 0 Then
            k = CStr(baseData(i, 1)) & vbTab & CStr(baseData(i, 2)) & vbTab & CStr(baseData(i, 6)) & vbTab & SafeUpperText(baseData(i, 5))
            If baseRecords.Exists(k) Then
                baseRecords(k) = Array(Empty, Empty, "BLOCKED - duplicate bank-base evidence")
            Else
                baseRecords(k) = Array(baseData(i, 10), baseData(i, 11), CStr(baseData(i, 19)))
            End If
        End If
    Next i
    If Not coverage Is Nothing Then
        cv = coverage.UsedRange.Value2
        For i = 2 To UBound(cv, 1)
            If UBound(cv, 2) >= 4 Then mapping(SafeUpperText(cv(i, 1)) & "|" & SafeUpperText(cv(i, 2))) = Array(SafeText(cv(i, 4)), SafeText(cv(i, 3)))
        Next i
    End If
    Set metrics = PsMetricsSheet()
    If Not metrics Is Nothing Then
        For i = PS_FIRST_ROW To metrics.Cells(metrics.rows.count, 1).End(xlUp).row
            label = SafeUpperText(metrics.Cells(i, 1).Value2)
            If Len(label) > 0 Then
                formulas(label) = SafeText(metrics.Cells(i, 4).Value2)
                stages(label) = SafeText(metrics.Cells(i, 5).Value2)
            End If
        Next i
    End If
    row = 8
    For i = 2 To UBound(a, 1)
        If i = 2 Or (i - 1) Mod 25 = 0 Or i = UBound(a, 1) Then
            ProgressDetail "Comparing input pivots " & CStr(i - 1) & " of " & CStr(UBound(a, 1) - 1), _
                95# + 3# * (i - 1) / (UBound(a, 1) - 1)
        End If
        label = SafeUpperText(a(i, 5)): If Len(label) = 0 Then GoTo NextMetric
        For stageNo = 0 To 1
            k = CStr(a(i, 1)) & vbTab & CStr(a(i, 2)) & vbTab & CStr(a(i, 6)) & vbTab & label
            If stageNo = 0 And seen.Exists(k) Then GoTo NextStage
            If stageNo = 0 Then seen(k) = True
            issue = "": mBridgeScopeProof = "": pv = Empty: dv = a(i, IIf(stageNo = 0, 10, 13)): sv = a(i, IIf(stageNo = 0, 11, 14))
            priorStatus = CStr(a(i, IIf(stageNo = 0, 19, 20)))
            If stageNo = 0 Then
                If baseRecords.Exists(k) Then
                    baseRecord = baseRecords(k): dv = baseRecord(0): sv = baseRecord(1): priorStatus = CStr(baseRecord(2))
                Else
                    dv = Empty: sv = Empty: priorStatus = "BLOCKED - bank-base evidence unavailable"
                End If
            End If
            formula = "": stage = "ALL": measure = "": pivotSheet = "": membershipField = ""
            physicalSource = Trim$(CStr(a(i, 6))): sourcePredicate = ""
            splitAt = InStr(1, physicalSource, "~", vbBinaryCompare)
            If splitAt > 0 Then
                sourcePredicate = Trim$(Mid$(physicalSource, splitAt + 1))
                physicalSource = Trim$(Left$(physicalSource, splitAt - 1))
            End If
            If formulas.Exists(label) Then formula = CStr(formulas(label))
            If stages.Exists(label) Then stage = CStr(stages(label))
            If Len(sourcePredicate) > 0 Then
                whereAt = InStr(1, formula, " WHERE ", vbTextCompare)
                If whereAt > 0 Then
                    If SafeUpperText(Trim$(Mid$(formula, whereAt + 7))) = SafeUpperText(sourcePredicate) Then
                        formula = Trim$(Left$(formula, whereAt - 1))
                    Else
                        issue = "Stored source selection does not match the configured metric policy."
                    End If
                Else
                    issue = "Stored source selection has no matching configured metric policy."
                End If
            End If
            If Len(issue) = 0 Then measure = SimpleSumMeasure(formula, stage, issue)
            If stageNo = 0 Then
                pivotSheet = "Base_" & physicalSource
            Else
                k = SafeUpperText(a(i, 3)) & "|" & SafeUpperText(a(i, 4))
                If mapping.Exists(k) Then
                    mappedCase = mapping(k): pivotSheet = CStr(mappedCase(0)): membershipField = CStr(mappedCase(1))
                End If
                If Len(membershipField) = 0 Then issue = "Exact testcase membership field is unavailable in the coverage register."
            End If
            If Len(issue) = 0 And (priorStatus = "PASS" Or priorStatus = "FAIL") Then pv = NativePivotValue(joined, pivotSheet, measure, physicalSource, stage, CStr(a(i, 1)), CStr(a(i, 2)), issue, membershipField, sourcePredicate)
            If priorStatus <> "PASS" And priorStatus <> "FAIL" Then
                result = "BLOCKED": issue = priorStatus
            ElseIf Len(issue) > 0 Then
                result = "REVIEW - unmapped"
            ElseIf Not ReconNumber(dv) Or Not ReconNumber(sv) Then
                result = "BLOCKED": issue = "Derived or system value is missing."
            ElseIf priorStatus = "FAIL" Or Abs(CDbl(pv) - CDbl(dv)) > ReportTolerance(label) Or Abs(CDbl(sv) - CDbl(pv)) > ReportTolerance(label) Or _
                   Abs(CDbl(sv) - CDbl(dv)) > ReportTolerance(label) Then
                result = "FAIL": issue = "Input pivot, derived calculation and system do not agree within tolerance."
            Else
                result = "PASS": issue = "Input pivot, derived calculation and system agree within tolerance."
            End If
            ws.Cells(row, 1).Value2 = IIf(stageNo = 0, "Base", "Pre-shock")
            ws.Cells(row, 2).Value2 = a(i, 1): ws.Cells(row, 3).Value2 = a(i, 2)
            ws.Cells(row, 4).Value2 = IIf(stageNo = 0, "Bank base", a(i, 3))
            ws.Cells(row, 5).Value2 = IIf(stageNo = 0, "Portfolio", a(i, 4))
            ws.Cells(row, 6).Value2 = label: ws.Cells(row, 7).Value2 = a(i, 6)
            ws.Cells(row, 8).Value2 = pv: ws.Cells(row, 9).Value2 = dv: ws.Cells(row, 10).Value2 = sv
            If ReconNumber(pv) And ReconNumber(dv) Then ws.Cells(row, 11).Value2 = CDbl(pv) - CDbl(dv)
            If ReconNumber(pv) And ReconNumber(sv) Then ws.Cells(row, 12).Value2 = CDbl(sv) - CDbl(pv)
            If Len(mBridgeScopeProof) > 0 And ReconNumber(pv) Then issue = issue & " Scope: " & mBridgeScopeProof
            ws.Cells(row, 13).Value2 = result: ws.Cells(row, 14).Value2 = issue
            ws.Cells(row, 15).Value2 = pivotSheet: row = row + 1
NextStage:
        Next stageNo
NextMetric:
    Next i
Finish:
    If Not joined Is Nothing Then joined.Close SaveChanges:=False
    ws.Range("A7:O7").Font.Bold = True: ws.Range("A7:O7").Interior.Color = RGB(14, 34, 64): ws.Range("A7:O7").Font.Color = vbWhite
    ws.columns("A:G").ColumnWidth = 21: ws.columns("H:L").ColumnWidth = 18
    ws.columns("M").ColumnWidth = 22: ws.columns("N").ColumnWidth = 66: ws.columns("O").ColumnWidth = 24
    ws.columns("H:L").NumberFormat = "#,##0.########"
    ws.columns("B").NumberFormat = "yyyy-mm-dd"
    ws.Range("A7:O" & Application.Max(8, row - 1)).AutoFilter
    Exit Sub
Failed:
    issue = Err.description
    On Error Resume Next
    If Not joined Is Nothing Then joined.Close SaveChanges:=False
    If Not ws Is Nothing Then
        ws.Range("M8").Value2 = "BLOCKED": ws.Range("N8").Value2 = "Input-pivot comparison could not complete: " & issue
    End If
End Sub

Private Function SimpleSumMeasure(ByVal formula As String, ByRef stage As String, ByRef issue As String) As String
    Dim matches As Object, match As Object, f As String
    f = Trim$(formula)
    Set matches = RegexExecute("^SUM\s*\(\s*([A-Za-z0-9_ |]+)\s*\)\s*(?:WHERE\s+STAGE\s*=\s*([123]))?\s*$", f)
    If matches.count = 0 Then issue = "Formula requires a ratio, weighted or conditional mapping; review the source pivots and manual calculation.": Exit Function
    Set match = matches(0)
    SimpleSumMeasure = Trim$(CStr(match.SubMatches(0)))
    If Len(CStr(match.SubMatches(1))) > 0 Then stage = CStr(match.SubMatches(1))
    If Len(stage) = 0 Then stage = "ALL"
    If stage <> "ALL" And stage <> "1" And stage <> "2" And stage <> "3" Then issue = "Metric stage requires an explicit mapping."
End Function

Private Function NativePivotValue(ByVal wb As Workbook, ByVal sheetName As String, ByVal candidates As String, _
                                  ByVal source As String, ByVal stage As String, ByVal reportingDate As String, _
                                  ByVal entity As String, ByRef issue As String, Optional ByVal membershipField As String = "", _
                                  Optional ByVal sourcePredicate As String = "") As Variant
    Dim ws As Worksheet, pt As PivotTable, pf As PivotField, df As PivotField, item As PivotItem
    Dim field As Variant, physical As String, caption As String, entityParts As Variant, found As Boolean
    Dim kept As Long, rowCaption As String, selectedRows As Variant, updating As Boolean
    On Error GoTo Failed
    Set ws = GetWorksheetSafe(wb, sheetName)
    If ws Is Nothing Then issue = "Native pivot is unavailable for this case/source.": Exit Function
    If ws.PivotTables.count = 0 Then issue = "No native pivot exists for this case/source.": Exit Function
    Set pt = ws.PivotTables(1)
    physical = OwnedPivotMeasure(wb, candidates, source)
    If Len(physical) = 0 Then issue = "SUM source field is absent from the native input pivot.": Exit Function
    ' Apply the complete exact scope before Excel materializes the input pivot.
    pt.ManualUpdate = True: updating = True
    modShared_Pivot.RemoveEmptyOptionalPageFields pt, True
    SetPage pt, "ROW_SOURCE", source
    If Len(membershipField) > 0 Then SetPage pt, membershipField, "Yes"
    If Not ApplyBridgeSourcePolicy(pt, source, sourcePredicate, issue) Then GoTo CleanExit
    If stage <> "ALL" Then
        SetPage pt, "RECON_STAGE", stage
    Else
        pt.PivotFields("RECON_STAGE").ClearAllFilters
    End If
    ' The system's logical entity 1 is not source BANK_ID 101. The supported
    ' mapping is explicit and includes an independently applied branch restriction.
    entityParts = Split(entity, " / ")
    If UBound(entityParts) < 1 Then issue = "Entity ID is unavailable for exact input-pivot scope.": GoTo CleanExit
    If IsJordanOutputMapping(entity, CurrentEntityScope(), CurrentEntityFilter()) Then
        SetPage pt, "RECON_BANK", "JKB"
        SetPage pt, "BASE_ELIGIBLE", "Yes"
        Set pf = pt.PivotFields("BRANCH_CODE")
        pf.orientation = xlPageField: pf.ClearAllFilters: pf.EnableMultiplePageItems = True
        kept = 0
        For Each item In pf.PivotItems
            If Trim$(CStr(item.name)) <> "800" Then kept = kept + 1
        Next item
        If kept = 0 Then issue = "No Jordan branches are available in the native input pivot.": GoTo CleanExit
        For Each item In pf.PivotItems
            item.Visible = (Trim$(CStr(item.name)) <> "800")
        Next item
        mBridgeScopeProof = "JKB_JORDAN / 1 = JKB bank, excluding branch 800; exact reporting date."
    Else
        issue = "No approved native-input mapping for this output entity and configured branch scope."
        GoTo CleanExit
    End If
    Set pf = pt.PivotFields("AS_OF_DATE"): pf.orientation = xlPageField: pf.ClearAllFilters
    found = False
    For Each item In pf.PivotItems
        If SameReportDate(CStr(item.name), reportingDate) Then
            pf.EnableMultiplePageItems = False: pf.CurrentPage = item.name: found = True: Exit For
        End If
    Next item
    If Not found Then issue = "Exact reporting date is unavailable in the input pivot.": GoTo CleanExit
    ' This workbook is read-only and closes unsaved. Enable a total only after
    ' its single source, date, entity and stage have been explicitly selected.
    pt.RowGrand = True: pt.ColumnGrand = True
    For Each df In pt.dataFields
        If StrComp(df.sourceName, physical, vbTextCompare) = 0 Then caption = df.name: Exit For
    Next df
    If Len(caption) = 0 Then
        Set df = pt.AddDataField(pt.PivotFields(physical), physical & " bridge", xlSum): caption = df.name
    End If
    For Each df In pt.dataFields
        If StrComp(df.sourceName, "ROW_COUNT", vbTextCompare) = 0 Then rowCaption = df.name: Exit For
    Next df
    If Len(rowCaption) = 0 Then
        Set df = pt.AddDataField(pt.PivotFields("ROW_COUNT"), "Input row count bridge", xlSum): rowCaption = df.name
    End If
    pt.ManualUpdate = False: updating = False
    selectedRows = pt.GetPivotData(rowCaption).Value2
    If Not ReconNumber(selectedRows) Then issue = "Native pivot has no measurable input-row coverage for this scope.": Exit Function
    If CDbl(selectedRows) <= 0 Then issue = "Native pivot selected no source rows for this scope.": Exit Function
    NativePivotValue = pt.GetPivotData(caption).Value2
    If Not ReconNumber(NativePivotValue) Then issue = "Native pivot did not return a numeric value for this scope."
CleanExit:
    If updating Then pt.ManualUpdate = False
    Exit Function
Failed:
    issue = "Exact input-pivot scope could not be verified: " & Err.description
    On Error Resume Next
    If updating Then pt.ManualUpdate = False
End Function

Private Function ApplyBridgeSourcePolicy(ByVal pt As PivotTable, ByVal source As String, ByVal predicate As String, ByRef issue As String) As Boolean
    Dim pf As PivotField, item As PivotItem, field As Variant, matches As Object, match As Object
    Dim wanted As Object, found As Object, value As Variant, key As String, fieldName As String
    On Error GoTo Failed
    ' A shared native pivot is reused for several metrics. Remove the prior
    ' supported metric policy before applying the current one; testcase flags remain.
    For Each field In Array("COA_BALANCESHEET_CATEGORY", "ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY")
        Set pf = Nothing
        On Error Resume Next
        Set pf = pt.PivotFields(CStr(field))
        Err.Clear
        On Error GoTo Failed
        If Not pf Is Nothing Then pf.ClearAllFilters
    Next field
    If Len(Trim$(predicate)) = 0 Then ApplyBridgeSourcePolicy = True: Exit Function
    Set matches = RegexExecute("^\s*([A-Za-z0-9_]+)\s+IN\s*\(\s*('[^']+'(?:\s*,\s*'[^']+')*)\s*\)\s*$", predicate)
    If matches.count <> 1 Then issue = "Source policy requires an explicit native-pivot mapping: " & predicate: Exit Function
    Set match = matches(0): fieldName = SafeUpperText(match.SubMatches(0))
    If Not ((SafeUpperText(source) = "LL" And fieldName = "COA_BALANCESHEET_CATEGORY") Or _
            (SafeUpperText(source) = "LCR" And fieldName = "ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY")) Then
        issue = "Source policy requires an explicit native-pivot mapping: " & predicate: Exit Function
    End If
    Set wanted = NewMap(): Set found = NewMap()
    For Each value In Split(CStr(match.SubMatches(1)), ",")
        key = Trim$(CStr(value)): key = SafeUpperText(Mid$(key, 2, Len(key) - 2))
        If SafeUpperText(source) = "LL" Then
            If key <> "ASSETS" And key <> "LIABILITIES" Then issue = "Unsupported LL source-policy value: " & key: Exit Function
        Else
            If key <> "HQLA_LEVEL1" And key <> "HQLA_LEVEL2A" And key <> "HQLA_LEVEL2B" And key <> "OUTFLOW" And key <> "INFLOW" Then
                issue = "Unsupported LCR source-policy value: " & key: Exit Function
            End If
        End If
        wanted(key) = True
    Next value
    Set pf = pt.PivotFields(fieldName)
    pf.orientation = xlPageField: pf.ClearAllFilters: pf.EnableMultiplePageItems = True
    For Each item In pf.PivotItems
        key = SafeUpperText(item.name)
        If wanted.Exists(key) Then found(key) = True
    Next item
    If found.count <> wanted.count Then issue = "Source-policy value is absent from the native input pivot: " & predicate: Exit Function
    ' Make all requested items visible first; Excel cannot hide its last item.
    For Each item In pf.PivotItems
        If wanted.Exists(SafeUpperText(item.name)) Then item.Visible = True
    Next item
    For Each item In pf.PivotItems
        If Not wanted.Exists(SafeUpperText(item.name)) Then item.Visible = False
    Next item
    ApplyBridgeSourcePolicy = True
    Exit Function
Failed:
    issue = "Source policy could not be applied to the native input pivot: " & Err.description
End Function

Private Function SameReportDate(ByVal value As String, ByVal expected As String) As Boolean
    Dim actualDay As Long, expectedDay As Long
    If Not TryReportDateDay(value, actualDay) Then Exit Function
    If Not TryReportDateDay(expected, expectedDay) Then Exit Function
    SameReportDate = (actualDay = expectedDay)
End Function

Private Function TryReportDateDay(ByVal value As String, ByRef day As Long) As Boolean
    Dim serial As Double
    On Error GoTo InvalidDate
    value = Trim$(value)
    If Len(value) = 0 Then Exit Function
    If IsNumeric(value) Then
        serial = CDbl(value)
    ElseIf IsDate(value) Then
        serial = CDbl(CDate(value))
    Else
        Exit Function
    End If
    If serial < 1 Or serial >= 2958466# Then Exit Function
    day = CLng(Fix(serial))
    TryReportDateDay = True
InvalidDate:
End Function


Private Function OwnedPivotMeasure(ByVal wb As Workbook, ByVal candidates As String, ByVal source As String) As String
    Dim ws As Worksheet, key As String, col As Long, sourceCol As Long, lastCol As Long, lastRow As Long
    Dim field As Variant, r As Long, c As Long, value As Variant, sourceValues As Variant, amounts As Variant
    key = source & "|" & candidates
    If mPivotMeasureMap.Exists(key) Then OwnedPivotMeasure = CStr(mPivotMeasureMap(key)): Exit Function
    Set ws = GetWorksheetSafe(wb, "Joined_Data")
    If ws Is Nothing Then Exit Function
    lastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If SafeUpperText(ws.Cells(1, c).Value2) = "ROW_SOURCE" Then sourceCol = c: Exit For
    Next c
    If sourceCol = 0 Then Exit Function
    lastRow = ws.Cells(ws.rows.count, sourceCol).End(xlUp).row
    If lastRow < 2 Then Exit Function
    sourceValues = ws.Range(ws.Cells(1, sourceCol), ws.Cells(lastRow, sourceCol)).Value2
    For Each field In Split(candidates, "|")
        col = 0
        For c = 1 To lastCol
            If SafeUpperText(ws.Cells(1, c).Value2) = SafeUpperText(Trim$(CStr(field))) Then col = c: Exit For
        Next c
        If col > 0 Then
            amounts = ws.Range(ws.Cells(1, col), ws.Cells(lastRow, col)).Value2
            For r = 2 To lastRow
                If SafeUpperText(sourceValues(r, 1)) = SafeUpperText(source) Then
                    value = amounts(r, 1)
                    If ReconNumber(value) Then
                        OwnedPivotMeasure = CStr(ws.Cells(1, col).Value2)
                        mPivotMeasureMap(key) = OwnedPivotMeasure: Exit Function
                    End If
                End If
            Next r
        End If
    Next field
    mPivotMeasureMap(key) = ""
End Function


Public Function ManualReportRegressionTests() As String
    Dim rec As Object, issue As String, stage As String, measure As String
    mAmountTolerance = 0.005: mRateTolerance = 0.000001
    Set rec = NewMap()
    rec("Metric") = "ECL": rec("HasBase") = True: rec("HasPre") = True
    rec("BaseDerived") = 0: rec("BaseSystem") = 0: rec("PreDerived") = 0: rec("PreSystem") = 0
    rec("Rows") = 1: rec("Filter") = "STAGE=1"
    If StageStatus(rec, False) <> "PASS" Then Err.Raise 5, , "Genuine zero bank base should compare."
    If StageStatus(rec, True) <> "PASS" Then Err.Raise 5, , "Genuine zero selected value should compare."
    rec("Rows") = 0
    If Left$(StageStatus(rec, True), 7) <> "BLOCKED" Then Err.Raise 5, , "Zero selected rows must not pass."
    rec("Rows") = 1: rec("HasPre") = False
    If Left$(StageStatus(rec, True), 7) <> "BLOCKED" Then Err.Raise 5, , "Unavailable derivation must not become zero."
    rec("HasPre") = True: rec("PreSystem") = Empty
    If Left$(StageStatus(rec, True), 7) <> "BLOCKED" Then Err.Raise 5, , "Missing system value must not pass."
    rec("PreSystem") = 0.00001: rec("Metric") = "IMPACT_LCR"
    If StageStatus(rec, True) <> "FAIL" Then Err.Raise 5, , "Rate tolerance must apply to LCR impacts."
    stage = "ALL": issue = "": measure = SimpleSumMeasure("SUM(ECL|CALCULATED_ECL) WHERE STAGE=2", stage, issue)
    If measure <> "ECL|CALCULATED_ECL" Or stage <> "2" Or Len(issue) > 0 Then Err.Raise 5, , "SUM stage mapping failed."
    stage = "ALL": issue = "": measure = SimpleSumMeasure("SUM(ECL)/SUM(OUTSTANDING_LCY)", stage, issue)
    If Len(issue) = 0 Then Err.Raise 5, , "Ratio must not pretend to be a source SUM."
    If CombinedStatus("PASS", "BLOCKED - filter missing") = "PASS" Then Err.Raise 5, , "Missing filter status lost."
    If Not IsJordanOutputMapping("JKB_JORDAN / 1", SCOPE_JORDAN, "BRANCH_CODE NOT IN ('800')") Then Err.Raise 5, , "Documented Jordan mapping was rejected."
    If IsJordanOutputMapping("JKB_JORDAN / 101", SCOPE_JORDAN, "BRANCH_CODE NOT IN ('800')") Then Err.Raise 5, , "Source bank ID must not be treated as output entity ID."
    If IsJordanOutputMapping("JKB_CYPRUS / 1", SCOPE_JORDAN, "BRANCH_CODE NOT IN ('800')") Then Err.Raise 5, , "Unmapped output entity was accepted."
    If IsJordanOutputMapping("JKB_JORDAN / 1", SCOPE_BOTH, "") Then Err.Raise 5, , "Whole-entity scope must not be mapped to Jordan."
    If IsJordanOutputMapping("JKB_JORDAN / 1", SCOPE_JORDAN, "BRANCH_CODE IN ('800')") Then Err.Raise 5, , "Overridden branch scope must not be mapped to Jordan."
    If SameReportDate("2025-09-30", "2024-12-31") Then Err.Raise 5, , "Different reporting dates were matched."
    If Not SameReportDate("31 Dec 2025", "46022") Then Err.Raise 5, , "Numeric expected reporting date was rejected."
    If Not SameReportDate("46022", "2025-12-31") Then Err.Raise 5, , "Numeric native reporting date was rejected."
    If SameReportDate("46021", "2025-12-31") Then Err.Raise 5, , "Different numeric reporting date was accepted."
    If SameReportDate("invalid", "invalid") Or SameReportDate("", "") Then Err.Raise 5, , "Invalid reporting dates matched."
    ManualReportRegressionTests = "PASS: zero/missing data, exact dates, explicit Jordan-only entity mapping, rate tolerance, source-SUM mappings and combined status."
End Function


Private Function IsJordanOutputMapping(ByVal entity As String, ByVal scope As String, ByVal branchFilter As String) As Boolean
    Dim parts As Variant, normalized As String, scopeKey As String
    parts = Split(entity, " / ")
    If UBound(parts) <> 1 Then Exit Function
    If SafeUpperText(parts(0)) <> "JKB_JORDAN" Or Trim$(CStr(parts(1))) <> "1" Then Exit Function
    scopeKey = SafeUpperText(scope)
    If scopeKey <> SafeUpperText(SCOPE_JORDAN) And scopeKey <> "JORDAN" And scopeKey <> "JO" And scopeKey <> "JOR" Then Exit Function
    normalized = Replace(Replace(SafeUpperText(branchFilter), " ", ""), vbTab, "")
    If normalized <> "BRANCH_CODENOTIN('800')" Then Exit Function
    IsJordanOutputMapping = True
End Function


