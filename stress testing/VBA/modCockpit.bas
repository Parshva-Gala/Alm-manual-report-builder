Attribute VB_Name = "modCockpit"
Option Explicit

' ============================================================================
'  Home: the reconciliation cockpit (V15, design system v2)
'
'  One screen that answers, at a glance, the four things a reviewer opens the
'  tool to know:
'
'    1  How much of the stress test reconciles with the system, and what not.
'    2  Where in the chain it breaks: scenario > test case > base > pre-shock >
'       shock > post-shock, with matched / checked at every step.
'    3  Which risk families and which values need explaining first.
'    4  What to do next: one primary action, chosen from the state of the run.
'
'  DrawHome lays the sheet out (rarely: only when the layout stamp changes).
'  RefreshHome fills it from Recon_Detail, Pre_Shock_Sources, Pre_Shock_Cases
'  and the system output, and is cheap enough to run after every load and run.
'
'  The grid is 35 narrow columns (B:AJ) so tiles, panels and cards can share one
'  rhythm: 5 KPI tiles, 3 panels, 7 input cards. Values other code reads or
'  writes are addressed by workbook names, never by cell address:
'    JKB_RunStatus  JKB_PassCount  JKB_FailCount  JKB_OpenCount
'    JKB_TolAmount  JKB_TolRate    JKB_LastActivity
' ============================================================================

Private Const HOME_SHEET As String = "Recon_Workbench"
Private Const DETAIL_SHEET As String = "Recon_Detail"
Private Const CASES_SHEET As String = "Pre_Shock_Cases"
Private Const VS_SHEET_NAME As String = "Config_ValueSources"
Private Const LAYOUT_STAMP As String = "COCKPIT_V2"
Private Const STAMP_CELL As String = "AL1"

Private Const R_BAND1 As Long = 3
Private Const R_TILE As Long = 8
Private Const R_CHAIN_HEAD As Long = 14
Private Const R_CHAIN As Long = 15
Private Const R_PANEL As Long = 19
Private Const R_INPUT_HEAD As Long = 32
Private Const R_INPUT As Long = 33
Private Const R_FOOT As Long = 36

Private Const MAX_FAMILIES As Long = 8
Private Const MAX_DIFFS As Long = 8

' Recon_Detail, only the eight columns the home reads, held for one refresh:
'   1 status  2 stage  3 test case  4 element  5 field  6 tool  7 system  8 difference
Private mSlim As Variant

' ---------------------------------------------------------------- entry ------

' Lays the home out if it is missing or older than this layout, then fills it.
Public Sub EnsureHomeCockpit(Optional ByVal forceLayout As Boolean = False)
    Dim ws As Worksheet, er As String
    Set ws = HomeSheet(True)
    On Error GoTo DrawFailed
    If forceLayout Or CStr(ws.Range(STAMP_CELL).Value2) <> LAYOUT_STAMP Then DrawHome ws
    On Error GoTo 0
    RefreshHome
    Exit Sub
DrawFailed:
    ' Drawing the home must never stop the tool from opening. Log it, clear the
    ' stamp so the next open draws again, and carry on with what is there.
    er = Err.description
    On Error Resume Next
    ws.Range(STAMP_CELL).Value2 = ""
    LogIssue LOG_LEVEL_WARN, "Home", "The home screen could not be drawn: " & er, HOME_SHEET
End Sub

Public Sub RefreshHome()
    Dim ws As Worksheet, d As Object, why As String, keepScreen As Boolean
    ' Drawn with the screen off: after a Run the screen is back on, and a few
    ' thousand visible cell and shape updates kept Excel from answering Windows
    ' long enough to be reported as hung.
    keepScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Quiet
    Set ws = HomeSheet(False)
    If ws Is Nothing Then GoTo Done
    If CStr(ws.Range(STAMP_CELL).Value2) <> LAYOUT_STAMP Then GoTo Done
    Set d = ReadDetail()
    FillBand ws, d
    FillTiles ws, d
    FillChain ws, d
    FillHeatmap ws, d
    FillDifferences ws, d
    FillInputs ws
    FillNextStep ws, d
Done:
    mSlim = Empty
    Application.ScreenUpdating = keepScreen
    Exit Sub
Quiet:
    ' The home must never be the thing that stops a run. A refresh that fails
    ' leaves the last good figures in place and says so in the activity line.
    why = Err.description   ' read before On Error Resume Next clears it
    On Error Resume Next
    If Not ws Is Nothing Then ws.Range("JKB_LastActivity").Value2 = "The home could not refresh: " & why
    mSlim = Empty
    Application.ScreenUpdating = keepScreen
End Sub

' ---------------------------------------------------------------- state ------
' Other modules read and write the home only through these, so the layout can
' change without touching the run, load or clear code.

Public Sub HomeSetStatus(ByVal status As String, Optional ByVal activity As String = vbNullString)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = HomeSheet(True)
    If CStr(ws.Range(STAMP_CELL).Value2) <> LAYOUT_STAMP Then DrawHome ws
    ws.Range("JKB_RunStatus").Value2 = status
    If Len(activity) > 0 Then ws.Range("JKB_LastActivity").Value2 = activity
End Sub

Public Sub HomeSetActivity(ByVal activity As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = HomeSheet(True)
    If CStr(ws.Range(STAMP_CELL).Value2) <> LAYOUT_STAMP Then DrawHome ws
    ws.Range("JKB_LastActivity").Value2 = activity
End Sub

Public Function HomeStatus() As String
    On Error Resume Next
    HomeStatus = CStr(HomeSheet(False).Range("JKB_RunStatus").Value2)
End Function

' The amount (isRate False) or rate tolerance as typed, Empty when the home or
' the name is missing so the caller can apply its own default.
Public Function HomeTolerance(ByVal isRate As Boolean) As Variant
    Dim ws As Worksheet
    HomeTolerance = Empty
    On Error Resume Next
    Set ws = HomeSheet(False)
    If ws Is Nothing Then Exit Function
    If isRate Then
        HomeTolerance = ws.Range("JKB_TolRate").Value2
    Else
        HomeTolerance = ws.Range("JKB_TolAmount").Value2
    End If
End Function

' Facts the review pack prints, from the same readers the home uses.
Public Sub HomeSystemFacts(ByRef asOf As String, ByRef entity As String, ByRef dataRows As Long)
    SystemOutputFacts asOf, entity, dataRows
End Sub

' "All dated 31-Dec-2025", "Reporting dates differ: check Inputs" and the like.
Public Function HomeInputsLine() As String
    Dim dates As String
    LoadedInputs dates
    HomeInputsLine = dates
End Function

Public Function HomeFamilyName(ByVal testCase As String) As String
    HomeFamilyName = FamilyName(FamilyOf(testCase))
End Function

Public Function HomeFamilyCode(ByVal testCase As String) As String
    HomeFamilyCode = FamilyOf(testCase)
End Function

' The run stamp as a date and time; Results stores it as a date serial.
Public Function HomeRunText(ByVal v As Variant) As String
    On Error GoTo AsText
    If IsDate(v) Then
        HomeRunText = format$(CDate(v), "dd-mmm-yyyy hh:nn")
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        If CDbl(v) > 0 Then HomeRunText = format$(CDate(CDbl(v)), "dd-mmm-yyyy hh:nn")
    Else
        HomeRunText = SafeText(v)
    End If
    Exit Function
AsText:
    HomeRunText = SafeText(v)
End Function

' Test cases and risk families, counted one way for Home and the review pack:
' the configured test cases on the Test cases sheet, else the ones in Results
' that belong to a known risk family.
Public Function HomeCaseCounts(ByRef famCount As Long) As Long
    Dim cfg As Object, ws As Worksheet, lastR As Long, a As Variant, i As Long, tc As String, seen As Object, fams As Object
    Set cfg = ConfiguredCases()
    If cfg.count > 0 Then
        Set fams = CreateObject("Scripting.Dictionary")
        For Each a In cfg.keys
            fams(cfg(a)) = True
        Next a
        famCount = fams.count: HomeCaseCounts = cfg.count
        Exit Function
    End If
    Set seen = CreateObject("Scripting.Dictionary"): Set fams = CreateObject("Scripting.Dictionary")
    Set ws = GetWorksheetSafe(ThisWorkbook, DETAIL_SHEET)
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
        If lastR >= 8 Then
            a = ws.Range(ws.Cells(8, 5), ws.Cells(lastR, 5)).Value2
            If Not IsArray(a) Then a = SingleCell(a)
            For i = 1 To UBound(a, 1)
                tc = UCase$(SafeText(a(i, 1)))
                If Len(tc) > 0 And KnownFamily(FamilyOf(tc)) Then
                    seen(tc) = True: fams(FamilyOf(tc)) = True
                End If
            Next i
        End If
    End If
    famCount = fams.count: HomeCaseCounts = seen.count
End Function

' Configured test cases in a risk family.
Public Function HomeFamilyCaseCount(ByVal code As String) As Long
    Dim cfg As Object, k As Variant
    Set cfg = ConfiguredCases()
    For Each k In cfg.keys
        If cfg(k) = UCase$(code) Then HomeFamilyCaseCount = HomeFamilyCaseCount + 1
    Next k
End Function

' True for a real risk family. Scenario-wide rows (SCENARIO, TRANSFER, TEST ...)
' still count in the totals but get no family row.
Public Function HomeIsRiskFamily(ByVal code As String) As Boolean
    Dim cfg As Object, k As Variant
    code = UCase$(code)
    Set cfg = ConfiguredCases()
    If cfg.count > 0 Then
        For Each k In cfg.keys
            If cfg(k) = code Then HomeIsRiskFamily = True: Exit Function
        Next k
    Else
        HomeIsRiskFamily = KnownFamily(code)
    End If
End Function

' ---------------------------------------------------------------- actions ----
' One place that maps each step to its macro, so the step buttons never need to
' change when a step's implementation does.

Public Sub HomeLoad()
    UploadAllSources
    RefreshHome
End Sub

Public Sub HomeCheckSetup()
    If GetWorksheetSafe(ThisWorkbook, VS_SHEET_NAME) Is Nothing Then
        GoValueSources
    Else
        JKB_CheckValueSources
    End If
    RefreshHome
End Sub

Public Sub HomeRun()
    ' A button's macro must never hand an error back to Excel: the run reports
    ' its own failures, and a view step after it must not end the session.
    ' RunStressReconciliation already refreshes Home (in the run's summary and
    ' again in OpenReconWorkbench); a third refresh only added time at the end.
    On Error Resume Next
    RunStressReconciliation
    Err.Clear
End Sub

Public Sub HomeReview()
    OpenStressConsole
End Sub

Public Sub HomePack()
    BuildManualReconPack
End Sub

Public Sub HomeStageInputs()
    ShowResultsStage "Inputs"
End Sub
Public Sub HomeStageBase()
    ShowResultsStage "Base"
End Sub
Public Sub HomeStagePre()
    ShowResultsStage "Pre-shock"
End Sub
Public Sub HomeStageShock()
    ShowResultsStage "Shock"
End Sub
Public Sub HomeStagePost()
    ShowResultsStage "Post-shock"
End Sub

Public Sub HomeShowResults()
    ShowReconDetail
End Sub

Public Sub HomeShowLog()
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_LOG)
    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate
End Sub

Public Sub HomeShowReference()
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, "Shock_Catalog")
    If ws Is Nothing Then Set ws = GetWorksheetSafe(ThisWorkbook, "Manual_Formula_Guide")
    If ws Is Nothing Then
        UiNotice "Reference", "The reference sheets are not in this copy of the tool."
        Exit Sub
    End If
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate
End Sub

Private Sub ShowResultsStage(ByVal stage As String)
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, DETAIL_SHEET)
    If ws Is Nothing Then
        UiNotice "Results", "There are no results yet.", , "Click Run on the home screen."
        Exit Sub
    End If
    ShowReconDetail
    On Error Resume Next
    If ActiveSheet.name = DETAIL_SHEET Then ActiveSheet.Range("A7").AutoFilter field:=2, Criteria1:=stage
End Sub

' ---------------------------------------------------------------- layout -----

Private Function HomeSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, HOME_SHEET)
    If ws Is Nothing And create Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.name = HOME_SHEET
    End If
    Set HomeSheet = ws
End Function

' Column n of the 35-column grid (1 = B).
Private Function GC(ByVal n As Long) As Long
    GC = n + 1
End Function

Private Function Blk(ByVal ws As Worksheet, ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long) As Range
    Set Blk = ws.Range(ws.Cells(r1, GC(c1)), ws.Cells(r2, GC(c2)))
End Function

Private Sub DrawHome(ByVal ws As Worksheet)
    Dim saved As Variant, sh As Shape, i As Long, r As Long, rg As Range
    Dim tiles As Variant, t As Variant, nm As Variant

    saved = ReadState(ws)

    ' A clean slate: the previous layout's shapes, merges, formats and names.
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.FormatConditions.Delete
    ws.Hyperlinks.Delete
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.ScrollArea = ""

    ws.Range("A1").Value2 = "JKB"

    ' ---- grid ------------------------------------------------------------
    ws.columns("A").ColumnWidth = 2.6
    ws.columns("B:AJ").ColumnWidth = 5.3
    ws.columns("AK").ColumnWidth = 2.6
    ws.columns("AL:AL").ColumnWidth = 2
    With ws.Range("A1:AK40")
        .Font.name = UI_FONT: .Font.Size = 10: .Font.Color = UI_TEXT
        .Interior.Color = UI_CANVAS
        .VerticalAlignment = xlCenter
    End With
    ws.Range("AL1:AL40").Interior.Color = UI_CANVAS

    ' ---- header band (rows 1-6) -----------------------------------------
    ws.Range("A1:AK6").Interior.Color = UI_NAVY
    ws.rows(3).RowHeight = 22: ws.rows(4).RowHeight = 34: ws.rows(5).RowHeight = 22: ws.rows(6).RowHeight = 16
    With ws.Range("B3")
        .Value2 = "JKB  " & ChrW(183) & "  RISK & FINANCE  " & ChrW(183) & "  STRESS TESTING"
        .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_NAVY_MUTED: .VerticalAlignment = xlBottom
    End With
    With ws.Range("B4")
        .Value2 = "Stress test reconciliation"
        .Font.Size = 22: .Font.Bold = True: .Font.Color = UI_WHITE
    End With
    With ws.Range("B5")
        .Value2 = "Rebuilds every scenario from the same input files the system uses, then reconciles each step with the system output."
        .Font.Size = 10.5: .Font.Color = UI_NAVY_TEXT: .VerticalAlignment = xlTop
    End With
    Blk(ws, R_BAND1, 21, R_BAND1, 35).Merge
    With Blk(ws, R_BAND1, 21, R_BAND1, 35)
        .HorizontalAlignment = xlRight: .Font.Size = 9.5: .Font.Color = UI_NAVY_TEXT: .VerticalAlignment = xlBottom
    End With
    ' Status pill: its fill and text colour follow the status text.
    Blk(ws, 4, 29, 4, 35).Merge
    Set rg = Blk(ws, 4, 29, 4, 35)
    With rg
        .HorizontalAlignment = xlCenter: .Font.Size = 10.5: .Font.Bold = True
        .Interior.Color = UI_FILL: .Font.Color = UI_TEXT_2
    End With
    PillFormats rg
    NameCell ws, "JKB_RunStatus", Blk(ws, 4, 29, 4, 29)
    Blk(ws, 5, 21, 5, 35).Merge
    With Blk(ws, 5, 21, 5, 35)
        .HorizontalAlignment = xlRight: .Font.Size = 9.5: .Font.Color = UI_NAVY_MUTED: .VerticalAlignment = xlTop
    End With
    ws.rows(7).RowHeight = 14

    ' ---- KPI tiles (rows 8-12) ------------------------------------------
    ws.rows(8).RowHeight = 22: ws.rows(9).RowHeight = 32: ws.rows(10).RowHeight = 10: ws.rows(11).RowHeight = 18: ws.rows(12).RowHeight = 6
    tiles = Array(Array(1, 6, "Checks reconciled"), Array(8, 13, "Differences above tolerance"), _
                  Array(15, 20, "Missing or incomplete evidence"), Array(22, 27, "Input files"), Array(29, 35, "Test cases covered"))
    For i = 0 To 4
        t = tiles(i)
        UiCard Blk(ws, R_TILE, t(0), R_TILE + 4, t(1))
        Blk(ws, R_TILE, t(0), R_TILE, t(1)).Merge
        With Blk(ws, R_TILE, t(0), R_TILE, t(1))
            .Value2 = t(2): .Font.Size = 9: .Font.Bold = True: .Font.Color = UI_MUTED: .IndentLevel = 1: .VerticalAlignment = xlBottom
        End With
        ' value (left, big) and its suffix (right, small)
        Blk(ws, R_TILE + 1, t(0), R_TILE + 1, t(0) + 2).Merge
        With Blk(ws, R_TILE + 1, t(0), R_TILE + 1, t(0) + 2)
            .Font.Size = 20: .Font.Bold = True: .Font.Color = UI_INK: .IndentLevel = 1: .NumberFormat = "#,##0"
            .HorizontalAlignment = xlLeft
        End With
        Blk(ws, R_TILE + 1, t(0) + 3, R_TILE + 1, t(1)).Merge
        With Blk(ws, R_TILE + 1, t(0) + 3, R_TILE + 1, t(1))
            .Font.Size = 10: .Font.Color = UI_MUTED: .VerticalAlignment = xlBottom
        End With
        Blk(ws, R_TILE + 3, t(0), R_TILE + 3, t(1)).Merge
        With Blk(ws, R_TILE + 3, t(0), R_TILE + 3, t(1))
            .Font.Size = 9: .Font.Color = UI_MUTED: .IndentLevel = 1
        End With
    Next i
    NameCell ws, "JKB_PassCount", Blk(ws, R_TILE + 1, 1, R_TILE + 1, 1)
    NameCell ws, "JKB_FailCount", Blk(ws, R_TILE + 1, 8, R_TILE + 1, 8)
    NameCell ws, "JKB_OpenCount", Blk(ws, R_TILE + 1, 15, R_TILE + 1, 15)
    Blk(ws, R_TILE + 1, 8, R_TILE + 1, 8).Font.Color = UI_BAD
    Blk(ws, R_TILE + 1, 15, R_TILE + 1, 15).Font.Color = UI_WARN
    ws.rows(13).RowHeight = 14

    ' ---- the chain (rows 14-17) ------------------------------------------
    ws.rows(R_CHAIN_HEAD).RowHeight = 20
    SectionHead ws, R_CHAIN_HEAD, 1, 20, "The chain, scenario to post-shock"
    Blk(ws, R_CHAIN_HEAD, 21, R_CHAIN_HEAD, 35).Merge
    With Blk(ws, R_CHAIN_HEAD, 21, R_CHAIN_HEAD, 35)
        .Value2 = "matched / checked at each step. Click a step to open its results."
        .HorizontalAlignment = xlRight: .Font.Size = 9: .Font.Color = UI_FAINT
    End With
    ws.rows(R_CHAIN).RowHeight = 20: ws.rows(R_CHAIN + 1).RowHeight = 24: ws.rows(R_CHAIN + 2).RowHeight = 18
    ws.rows(18).RowHeight = 16

    ' ---- three panels (rows 19-30) --------------------------------------
    ws.rows(R_PANEL).RowHeight = 26
    For r = R_PANEL + 1 To R_PANEL + 10: ws.rows(r).RowHeight = 21: Next r
    ws.rows(R_PANEL + 11).RowHeight = 8
    UiCard Blk(ws, R_PANEL, 1, R_PANEL + 11, 12)
    UiCard Blk(ws, R_PANEL, 14, R_PANEL + 11, 28)
    UiCard Blk(ws, R_PANEL, 30, R_PANEL + 11, 35)
    PanelHead ws, R_PANEL, 1, 12, "By risk family " & ChrW(183) & " % reconciled"
    PanelHead ws, R_PANEL, 14, 28, "Largest differences"
    With Blk(ws, R_PANEL, 24, R_PANEL, 28)
        .Merge: .Value2 = "system minus tool": .HorizontalAlignment = xlRight: .Font.Size = 8.5: .Font.Color = UI_FAINT
        .IndentLevel = 1
    End With
    PanelHead ws, R_PANEL, 30, 35, "Your next step"

    ' heatmap header: family B:D, cases E, then Base / Pre / Shock / Post in pairs
    HeadCell Blk(ws, R_PANEL + 1, 1, R_PANEL + 1, 3), "Risk family", xlLeft
    HeadCell Blk(ws, R_PANEL + 1, 4, R_PANEL + 1, 4), "Cases", xlRight
    HeadCell Blk(ws, R_PANEL + 1, 5, R_PANEL + 1, 6), "Base", xlCenter
    HeadCell Blk(ws, R_PANEL + 1, 7, R_PANEL + 1, 8), "Pre-shock", xlCenter
    HeadCell Blk(ws, R_PANEL + 1, 9, R_PANEL + 1, 10), "Shock", xlCenter
    HeadCell Blk(ws, R_PANEL + 1, 11, R_PANEL + 1, 12), "Post-shock", xlCenter
    For r = R_PANEL + 2 To R_PANEL + 1 + MAX_FAMILIES
        Blk(ws, r, 1, r, 3).Merge: Blk(ws, r, 1, r, 3).IndentLevel = 1
        Blk(ws, r, 4, r, 4).HorizontalAlignment = xlRight
        For i = 0 To 3
            Blk(ws, r, 5 + i * 2, r, 6 + i * 2).Merge
            With Blk(ws, r, 5 + i * 2, r, 6 + i * 2)
                .HorizontalAlignment = xlCenter: .NumberFormat = "0%": .Font.Bold = True: .Font.Size = 9.5
                .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Color = UI_WHITE
                .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = UI_WHITE
            End With
        Next i
        Hairline Blk(ws, r, 1, r, 4)
    Next r
    HeatFormats Blk(ws, R_PANEL + 2, 5, R_PANEL + 1 + MAX_FAMILIES, 12)

    ' differences header: case O:Q, stage R:S, field T:X, difference Y:Z, bar AA:AB
    HeadCell Blk(ws, R_PANEL + 1, 14, R_PANEL + 1, 16), "Test case", xlLeft
    HeadCell Blk(ws, R_PANEL + 1, 17, R_PANEL + 1, 18), "Stage", xlLeft
    HeadCell Blk(ws, R_PANEL + 1, 19, R_PANEL + 1, 23), "Field", xlLeft
    HeadCell Blk(ws, R_PANEL + 1, 24, R_PANEL + 1, 25), "Difference", xlRight
    HeadCell Blk(ws, R_PANEL + 1, 26, R_PANEL + 1, 28), "", xlLeft
    For r = R_PANEL + 2 To R_PANEL + 1 + MAX_DIFFS
        Blk(ws, r, 14, r, 16).Merge: Blk(ws, r, 14, r, 16).IndentLevel = 1
        Blk(ws, r, 17, r, 18).Merge
        Blk(ws, r, 19, r, 23).Merge: Blk(ws, r, 19, r, 23).Font.name = "Consolas": Blk(ws, r, 19, r, 23).Font.Size = 9
        Blk(ws, r, 24, r, 25).Merge: Blk(ws, r, 24, r, 25).HorizontalAlignment = xlRight
        Blk(ws, r, 26, r, 28).Merge: Blk(ws, r, 26, r, 28).Font.Size = 8: Blk(ws, r, 26, r, 28).IndentLevel = 1
        Hairline Blk(ws, r, 14, r, 28)
    Next r

    ' next step: four steps, two rows each (title / note) in AE:AJ, number in AD
    For i = 0 To 3
        r = R_PANEL + 1 + i * 2
        Blk(ws, r, 30, r + 1, 30).Merge
        With Blk(ws, r, 30, r + 1, 30)
            .HorizontalAlignment = xlCenter: .Font.Bold = True: .Font.Size = 10
        End With
        Blk(ws, r, 31, r, 35).Merge: Blk(ws, r, 31, r, 35).Font.Bold = True: Blk(ws, r, 31, r, 35).Font.Color = UI_INK
        Blk(ws, r, 31, r, 35).VerticalAlignment = xlBottom
        Blk(ws, r + 1, 31, r + 1, 35).Merge: Blk(ws, r + 1, 31, r + 1, 35).Font.Size = 8.5: Blk(ws, r + 1, 31, r + 1, 35).Font.Color = UI_MUTED
        Blk(ws, r + 1, 31, r + 1, 35).VerticalAlignment = xlTop
    Next i
    ' tolerances, the two review settings, beside the Run they govern
    With Blk(ws, R_PANEL + 10, 30, R_PANEL + 10, 31)
        .Merge: .Value2 = "Tolerance": .Font.Size = 8.5: .Font.Color = UI_MUTED: .IndentLevel = 1
    End With
    Blk(ws, R_PANEL + 10, 32, R_PANEL + 10, 33).Merge
    Blk(ws, R_PANEL + 10, 34, R_PANEL + 10, 35).Merge
    For Each nm In Array(32, 34)
        With Blk(ws, R_PANEL + 10, CLng(nm), R_PANEL + 10, CLng(nm) + 1)
            .Interior.Color = UI_BRAND_TINT: .Font.Bold = True: .Font.Color = UI_INK: .HorizontalAlignment = xlCenter
            .NumberFormat = "0.######": .Font.Size = 9.5
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = UI_BRAND
        End With
    Next nm
    NameCell ws, "JKB_TolAmount", Blk(ws, R_PANEL + 10, 32, R_PANEL + 10, 32)
    NameCell ws, "JKB_TolRate", Blk(ws, R_PANEL + 10, 34, R_PANEL + 10, 34)
    AddInputHint Blk(ws, R_PANEL + 10, 32, R_PANEL + 10, 32), "Amount tolerance", "A difference larger than this, in currency units, is flagged."
    AddInputHint Blk(ws, R_PANEL + 10, 34, R_PANEL + 10, 34), "Rate tolerance", "A difference larger than this, for ratios and rates, is flagged."
    ws.rows(31).RowHeight = 16

    ' ---- input files (rows 32-34) ----------------------------------------
    ws.rows(R_INPUT_HEAD).RowHeight = 20
    SectionHead ws, R_INPUT_HEAD, 1, 20, "Input files"
    Blk(ws, R_INPUT_HEAD, 21, R_INPUT_HEAD, 35).Merge
    With Blk(ws, R_INPUT_HEAD, 21, R_INPUT_HEAD, 35)
        .Value2 = "recognised by their columns. The tool never changes them."
        .HorizontalAlignment = xlRight: .Font.Size = 9: .Font.Color = UI_FAINT
    End With
    ws.rows(R_INPUT).RowHeight = 22: ws.rows(R_INPUT + 1).RowHeight = 18
    For i = 0 To 6
        Set rg = InputCard(ws, i)
        UiCard rg
        rg.rows(1).Merge: rg.rows(2).Merge
        rg.rows(1).Font.Bold = True: rg.rows(1).Font.Color = UI_INK: rg.rows(1).IndentLevel = 1: rg.rows(1).VerticalAlignment = xlBottom
        rg.rows(2).Font.Size = 8.5: rg.rows(2).Font.Color = UI_MUTED: rg.rows(2).IndentLevel = 1: rg.rows(2).VerticalAlignment = xlTop
    Next i
    ws.rows(35).RowHeight = 12

    ' ---- footer (row 36) --------------------------------------------------
    ws.rows(R_FOOT).RowHeight = 24
    Blk(ws, R_FOOT, 1, R_FOOT, 18).Merge
    With Blk(ws, R_FOOT, 1, R_FOOT, 18)
        .Font.Size = 9: .Font.Color = UI_MUTED
    End With
    NameCell ws, "JKB_LastActivity", Blk(ws, R_FOOT, 1, R_FOOT, 1)
    ws.rows(37).RowHeight = 14
    FooterLinks ws

    ' ---- restore state ----------------------------------------------------
    ws.Range("JKB_RunStatus").Value2 = IIf(Len(CStr(saved(0))) = 0, "Not run", saved(0))
    ws.Range("JKB_TolAmount").Value2 = IIf(IsNumeric(saved(1)) And Len(CStr(saved(1))) > 0, saved(1), 0.005)
    ws.Range("JKB_TolRate").Value2 = IIf(IsNumeric(saved(2)) And Len(CStr(saved(2))) > 0, saved(2), 0.000001)
    ws.Range("JKB_LastActivity").Value2 = IIf(Len(CStr(saved(3))) = 0, "Start with step 1: load the system output and its input files.", saved(3))

    ws.Tab.Color = UI_NAVY
    ws.ScrollArea = "A1:AK40"
    ' The tolerance boxes are merged; Excel sets Locked only on a whole merged area.
    On Error Resume Next
    ws.Range("JKB_TolAmount").MergeArea.Locked = False
    ws.Range("JKB_TolRate").MergeArea.Locked = False
    On Error GoTo 0
    ' Stamped last, so a draw that stops part way is retried on the next open.
    ws.Range(STAMP_CELL).Value2 = LAYOUT_STAMP
    ws.Range(STAMP_CELL).Font.Color = UI_CANVAS
End Sub

' The four values that outlive a redraw, read from the names when they exist and
' from the V14 / first V15 addresses when upgrading an older home.
Private Function ReadState(ByVal ws As Worksheet) As Variant
    Dim st As Variant, tolA As Variant, tolR As Variant, act As Variant
    On Error Resume Next
    If Left$(CStr(ws.Range(STAMP_CELL).Value2), 8) = "COCKPIT_" Then
        ' Any cockpit layout: the names, each on its own, so one missing name
        ' (an earlier half-drawn home) does not lose the others.
        st = ws.Range("JKB_RunStatus").Value2
        tolA = ws.Range("JKB_TolAmount").Value2
        tolR = ws.Range("JKB_TolRate").Value2
        act = ws.Range("JKB_LastActivity").Value2
    Else
        ' The V14 home kept these at fixed addresses.
        st = ws.Range("B9").Value2: tolA = ws.Range("C26").Value2
        tolR = ws.Range("F26").Value2: act = ws.Range("B29").Value2
    End If
    If IsError(st) Or IsEmpty(st) Then st = ""
    If IsError(tolA) Or IsEmpty(tolA) Then tolA = ""
    If IsError(tolR) Or IsEmpty(tolR) Then tolR = ""
    If IsError(act) Or IsEmpty(act) Then act = ""
    ReadState = Array(st, tolA, tolR, act)
    Err.Clear
End Function

Private Sub NameCell(ByVal ws As Worksheet, ByVal nm As String, ByVal rg As Range)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add name:=nm, RefersTo:="='" & Replace(ws.name, "'", "''") & "'!" & rg.Cells(1, 1).address
End Sub

Private Sub SectionHead(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal text As String)
    Blk(ws, r, c1, r, c2).Merge
    UiEyebrow Blk(ws, r, c1, r, c1), text
End Sub

Private Sub PanelHead(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal text As String)
    With ws.Cells(r, GC(c1))
        .Value2 = UCase$(text): .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_MUTED
        .IndentLevel = 1: .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub HeadCell(ByVal rg As Range, ByVal text As String, ByVal align As Long)
    If rg.Cells.count > 1 Then rg.Merge
    With rg
        .Value2 = text: .Font.Size = 8.5: .Font.Bold = True: .Font.Color = UI_MUTED
        .HorizontalAlignment = align
        If align = xlLeft Then .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = UI_LINE
    End With
End Sub

Private Sub Hairline(ByVal rg As Range)
    With rg.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Color = UI_FILL: .Weight = xlThin
    End With
End Sub

Private Sub AddInputHint(ByVal rg As Range, ByVal title As String, ByVal msg As String)
    On Error Resume Next
    rg.Validation.Delete
    rg.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
    rg.Validation.InputTitle = title
    rg.Validation.InputMessage = msg
    rg.Validation.ErrorTitle = title
    rg.Validation.ErrorMessage = "Enter a number of zero or more."
    rg.Validation.ShowInput = True
    Err.Clear
End Sub

Private Function InputCard(ByVal ws As Worksheet, ByVal i As Long) As Range
    Dim starts As Variant, ends As Variant
    starts = Array(1, 6, 11, 16, 21, 26, 31)
    ends = Array(4, 9, 14, 19, 24, 29, 35)
    Set InputCard = Blk(ws, R_INPUT, CLng(starts(i)), R_INPUT + 1, CLng(ends(i)))
End Function

' Status pill colours, driven by the text the run writes.
Private Sub PillFormats(ByVal rg As Range)
    Dim fc As FormatCondition
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Reconciled", TextOperator:=xlBeginsWith)
    fc.Interior.Color = UI_OK_BG: fc.Font.Color = UI_OK
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Review", TextOperator:=xlBeginsWith)
    fc.Interior.Color = UI_WARN_BG: fc.Font.Color = UI_WARN
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Failed", TextOperator:=xlBeginsWith)
    fc.Interior.Color = UI_BAD_BG: fc.Font.Color = UI_BAD
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Running", TextOperator:=xlBeginsWith)
    fc.Interior.Color = UI_BRAND_TINT: fc.Font.Color = UI_BRAND
End Sub

' Green at 99.5% and above, amber down to 85%, red below. Blank stays blank.
Private Sub HeatFormats(ByVal rg As Range)
    Dim fc As FormatCondition, tl As String
    tl = rg.Cells(1, 1).address(False, False)
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=AND(ISNUMBER(" & tl & ")," & tl & ">=0.995)")
    fc.Interior.Color = UI_OK_FILL: fc.Font.Color = UI_OK: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=AND(ISNUMBER(" & tl & ")," & tl & ">=0.85)")
    fc.Interior.Color = UI_WARN_FILL: fc.Font.Color = UI_WARN: fc.StopIfTrue = True
    Set fc = rg.FormatConditions.Add(xlExpression, Formula1:="=ISNUMBER(" & tl & ")")
    fc.Interior.Color = UI_BAD_FILL: fc.Font.Color = UI_BAD: fc.StopIfTrue = True
End Sub

Private Sub FooterLinks(ByVal ws As Worksheet)
    Dim caps As Variant, procs As Variant, i As Long, x As Double, sh As Shape, w As Double
    caps = Array("Results", "Review console", "Configuration", "Reference", "Build log")
    procs = Array("HomeShowResults", "OpenStressConsole", "Sheet1.GoConfig", "HomeShowReference", "HomeShowLog")
    x = ws.Cells(R_FOOT, GC(35)).Left + ws.Cells(R_FOOT, GC(35)).Width
    For i = UBound(caps) To 0 Step -1
        w = Len(caps(i)) * 5.6 + 24
        x = x - w
        UiButton ws, "HomeLink" & i, CStr(caps(i)), CStr(procs(i)), ws.Cells(R_FOOT, GC(35)), UI_BTN_QUIET, 9
        Set sh = ws.Shapes("HomeLink" & i)
        sh.Left = x: sh.Width = w - 4: sh.Height = 20
        sh.Top = ws.Cells(R_FOOT, 1).Top + (ws.rows(R_FOOT).RowHeight - 20) / 2
    Next i
End Sub

' ---------------------------------------------------------------- data -------

' Recon_Detail summarised once per refresh. Keys:
'   rows (array or Empty)  n  pass  fail  review  blocked  assume
'   stage:<name>:<verdict> counts, fam:<code>:<stage>:<verdict> counts
'   families (Collection of codes, first-seen order), cases (Dictionary)
'   run (text)
Private Function ReadDetail() As Object
    Dim d As Object, ws As Worksheet, lastR As Long, a As Variant, i As Long
    Dim v As String, stg As String, tc As String, fam As String, fams As Collection, cases As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set fams = New Collection
    Set cases = CreateObject("Scripting.Dictionary")
    d("n") = 0: d("pass") = 0: d("fail") = 0: d("review") = 0: d("blocked") = 0: d("assume") = 0
    d("run") = ""
    Set ws = GetWorksheetSafe(ThisWorkbook, DETAIL_SHEET)
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
        If lastR >= 8 Then
            mSlim = DetailSlim(ws, lastR)
            d("rows") = True
            d("run") = HomeRunText(ws.Cells(8, 18).Value2)
            For i = 1 To UBound(mSlim, 1)
                v = UCase$(Trim$(CStr(ValueOr(mSlim(i, 1), ""))))
                If Len(v) > 0 Then
                    stg = Trim$(CStr(ValueOr(mSlim(i, 2), "")))
                    tc = Trim$(CStr(ValueOr(mSlim(i, 3), "")))
                    fam = FamilyOf(tc)
                    d("n") = d("n") + 1
                    Select Case v
                        Case "PASS": d("pass") = d("pass") + 1
                        Case "FAIL": d("fail") = d("fail") + 1
                        Case "REVIEW": d("review") = d("review") + 1
                        Case "ASSUMPTION": d("assume") = d("assume") + 1
                        Case Else: d("blocked") = d("blocked") + 1
                    End Select
                    Bump d, "stage:" & stg & ":" & v
                    Bump d, "stage:" & stg & ":ALL"
                    If Len(fam) > 0 Then
                        If Not d.Exists("famseen:" & fam) Then d("famseen:" & fam) = True: fams.Add fam
                        Bump d, "fam:" & fam & ":" & stg & ":" & v
                        Bump d, "fam:" & fam & ":" & stg & ":ALL"
                        If Len(tc) > 0 Then
                            If Not cases.Exists(tc) Then cases(tc) = fam
                        End If
                    End If
                End If
            Next i
        End If
    End If
    Set d("families") = fams
    Set d("cases") = cases
    Set ReadDetail = d
End Function

' Recon_Detail rows 8..lastR as an n x 8 array of the columns the summaries use
' (A status, B stage, E test case, F element, H field, I tool, J system,
' K difference). Reading these instead of all 18 columns keeps a large run's
' summary to under half the memory.
Public Function DetailSlim(ByVal ws As Worksheet, ByVal lastR As Long) As Variant
    Dim out() As Variant, ab As Variant, ef As Variant, hk As Variant, n As Long, i As Long
    n = lastR - 7
    If n < 1 Then DetailSlim = Empty: Exit Function
    ReDim out(1 To n, 1 To 8)
    ab = ws.Range(ws.Cells(8, 1), ws.Cells(lastR, 2)).Value2
    If n = 1 Then ab = TwoByOne(ab, ws, 8, 1)
    For i = 1 To n: out(i, 1) = ab(i, 1): out(i, 2) = ab(i, 2): Next i
    ab = Empty
    ef = ws.Range(ws.Cells(8, 5), ws.Cells(lastR, 6)).Value2
    If n = 1 Then ef = TwoByOne(ef, ws, 8, 5)
    For i = 1 To n: out(i, 3) = ef(i, 1): out(i, 4) = ef(i, 2): Next i
    ef = Empty
    hk = ws.Range(ws.Cells(8, 8), ws.Cells(lastR, 11)).Value2
    For i = 1 To n: out(i, 5) = hk(i, 1): out(i, 6) = hk(i, 2): out(i, 7) = hk(i, 3): out(i, 8) = hk(i, 4): Next i
    hk = Empty
    DetailSlim = out
End Function

' A one-row, two-column read is already a 2-D array; kept for clarity if that changes.
Private Function TwoByOne(ByVal v As Variant, ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim out(1 To 1, 1 To 2) As Variant
    If IsArray(v) Then TwoByOne = v: Exit Function
    out(1, 1) = ws.Cells(r, c).Value2: out(1, 2) = ws.Cells(r, c + 1).Value2
    TwoByOne = out
End Function

Private Sub Bump(ByVal d As Object, ByVal k As String)
    If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
End Sub

Private Function Cnt(ByVal d As Object, ByVal k As String) As Long
    If d.Exists(k) Then Cnt = d(k)
End Function

Private Function ValueOr(ByVal v As Variant, ByVal fallback As Variant) As Variant
    If IsError(v) Or IsEmpty(v) Then ValueOr = fallback Else ValueOr = v
End Function

' CR_T003 -> CR. The family code is the part before the first underscore.
Private Function FamilyOf(ByVal tc As String) As String
    Dim p As Long
    tc = UCase$(Trim$(tc))
    p = InStr(tc, "_")
    If p > 1 Then FamilyOf = Left$(tc, p - 1) Else FamilyOf = tc
End Function

Private Function FamilyName(ByVal code As String) As String
    Select Case UCase$(code)
        Case "CR": FamilyName = "Credit"
        Case "CC": FamilyName = "Concentration"
        Case "LR": FamilyName = "Liquidity"
        Case "MR": FamilyName = "Market"
        Case "OR": FamilyName = "Operational"
        Case "GP": FamilyName = "Geopolitical"
        Case "CL": FamilyName = "Climate"
        Case "COR": FamilyName = "Cost of risk"
        Case "AS": FamilyName = "Asset"
        Case Else: FamilyName = code
    End Select
    FamilyName = FamilyName & " (" & UCase$(code) & ")"
End Function

Private Function KnownFamily(ByVal code As String) As Boolean
    Select Case UCase$(code)
        Case "CR", "CC", "LR", "MR", "OR", "GP", "CL", "COR", "AS", "IRR", "RR": KnownFamily = True
    End Select
End Function

' Test case -> family code for every active row on the Test cases sheet.
Private Function ConfiguredCases() As Object
    Dim ws As Worksheet, lastR As Long, a As Variant, i As Long, tc As String, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set ConfiguredCases = d
    Set ws = GetWorksheetSafe(ThisWorkbook, CASES_SHEET)
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.rows.count, 2).End(xlUp).row
    If lastR < 8 Then Exit Function
    a = ws.Range(ws.Cells(8, 1), ws.Cells(lastR, 2)).Value2
    For i = 1 To UBound(a, 1)
        tc = UCase$(SafeText(a(i, 2)))
        If Len(tc) > 0 And UCase$(SafeText(a(i, 1))) <> "NO" Then d(tc) = FamilyOf(tc)
    Next i
End Function

Private Function SingleCell(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    a(1, 1) = v: SingleCell = a
End Function

' Checks that count toward "reconciled": everything but system-supplied assumptions.
Private Function Checked(ByVal d As Object, ByVal prefix As String) As Long
    Checked = Cnt(d, prefix & ":ALL") - Cnt(d, prefix & ":ASSUMPTION")
End Function

' ---------------------------------------------------------------- fill -------

Private Sub FillBand(ByVal ws As Worksheet, ByVal d As Object)
    Dim asOf As String, entity As String, t As String
    SystemOutputFacts asOf, entity, 0
    t = "Reporting date  " & IIf(Len(asOf) > 0, asOf, "not loaded")
    If Len(entity) > 0 Then t = t & "      Entity  " & entity
    Blk(ws, R_BAND1, 21, R_BAND1, 21).Value2 = t
    If Len(d("run")) > 0 Then
        Blk(ws, 5, 21, 5, 21).Value2 = "Last run  " & d("run")
    Else
        Blk(ws, 5, 21, 5, 21).Value2 = "Not run yet"
    End If
    ' The pill says the verdict and, when there is one, how many to explain.
    Dim st As String
    st = CStr(ws.Range("JKB_RunStatus").Value2)
    If (st = "Review required" Or Left$(st, 16) = "Review required ") And d("fail") > 0 Then
        ws.Range("JKB_RunStatus").Value2 = "Review required " & ChrW(183) & " " & format$(d("fail"), "#,##0") & " differences"
    End If
End Sub

Private Sub FillTiles(ByVal ws As Worksheet, ByVal d As Object)
    Dim total As Long, loaded As Long, dates As String, famCount As Long, caseCount As Long
    total = d("n") - d("assume")
    ws.Range("JKB_PassCount").Value2 = d("pass")
    Blk(ws, R_TILE + 1, 4, R_TILE + 1, 4).Value2 = "of " & format$(total, "#,##0")
    ' IIf evaluates both branches, so every share goes through Share (0 when
    ' nothing has run) rather than dividing inside IIf.
    Blk(ws, R_TILE + 3, 1, R_TILE + 3, 1).Value2 = IIf(total > 0, format$(Share(d("pass"), total), "0.0%") & " within tolerance", "No run yet")
    TileBar ws, 0, 1, 6, Share(d("pass"), total), UI_OK_DOT

    ws.Range("JKB_FailCount").Value2 = d("fail")
    Blk(ws, R_TILE + 3, 8, R_TILE + 3, 8).Value2 = IIf(d("fail") > 0, "Largest: " & LargestDifferenceText(d), "None above tolerance")
    TileBar ws, 1, 8, 13, Share(d("fail"), total), UI_BAD_DOT

    ws.Range("JKB_OpenCount").Value2 = d("review") + d("blocked")
    Blk(ws, R_TILE + 3, 15, R_TILE + 3, 15).Value2 = format$(d("blocked"), "#,##0") & " missing input " & ChrW(183) & " " & format$(d("review"), "#,##0") & " to review"
    TileBar ws, 2, 15, 20, Share(d("review") + d("blocked"), total), UI_WARN_DOT

    loaded = LoadedInputs(dates)
    Blk(ws, R_TILE + 1, 22, R_TILE + 1, 22).Value2 = loaded
    Blk(ws, R_TILE + 1, 25, R_TILE + 1, 25).Value2 = "of 7 loaded"
    Blk(ws, R_TILE + 3, 22, R_TILE + 3, 22).Value2 = dates
    TileBar ws, 3, 22, 27, loaded / 7, UI_BRAND

    caseCount = CasesCovered(famCount, d)
    Blk(ws, R_TILE + 1, 29, R_TILE + 1, 29).Value2 = caseCount
    Blk(ws, R_TILE + 1, 32, R_TILE + 1, 32).Value2 = "test cases"
    Blk(ws, R_TILE + 3, 29, R_TILE + 3, 29).Value2 = famCount & " risk families"
    TileBar ws, 4, 29, 35, IIf(caseCount > 0, 1, 0), UI_BRAND
End Sub

Private Function Share(ByVal part As Double, ByVal whole As Double) As Double
    If whole <> 0 Then Share = part / whole
End Function

' A thin progress bar inside a tile: a grey track and a coloured fill, as shapes
' so the length is exact.
Private Sub TileBar(ByVal ws As Worksheet, ByVal idx As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal pct As Double, ByVal color As Long)
    Dim rg As Range, x As Double, y As Double, w As Double, track As Shape, fill As Shape
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    Set rg = Blk(ws, R_TILE + 2, c1, R_TILE + 2, c2)
    x = rg.Left + 9: w = rg.Width - 18: y = rg.Top + (rg.Height - 5) / 2
    On Error Resume Next
    ws.Shapes("HomeBarT" & idx).Delete: ws.Shapes("HomeBarF" & idx).Delete
    On Error GoTo 0
    Set track = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 5)
    track.name = "HomeBarT" & idx
    track.fill.ForeColor.RGB = UI_FILL: track.line.Visible = msoFalse: track.Shadow.Visible = msoFalse
    track.Placement = xlMove
    If pct > 0 Then
        Set fill = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, Application.Max(5, w * pct), 5)
        fill.name = "HomeBarF" & idx
        fill.fill.ForeColor.RGB = color: fill.line.Visible = msoFalse: fill.Shadow.Visible = msoFalse
        fill.Placement = xlMove
    End If
End Sub

Private Sub FillChain(ByVal ws As Worksheet, ByVal d As Object)
    Dim titles As Variant, procs As Variant, i As Long, x As Double, w As Double, y As Double, h As Double
    Dim sh As Shape, dot As Shape, big As String, note As String, state As String, stg As String
    Dim famCount As Long, caseCount As Long, emptyFilters As Long, left0 As Double, right0 As Double
    Dim fill As Long
    titles = Array("Scenarios", "Test cases", "Base", "Pre-shock", "Shock", "Post-shock")
    procs = Array("GoPreShockCases", "GoPreShockCases", "HomeStageBase", "HomeStagePre", "HomeStageShock", "HomeStagePost")
    left0 = Blk(ws, R_CHAIN, 1, R_CHAIN, 1).Left
    right0 = Blk(ws, R_CHAIN, 35, R_CHAIN, 35).Left + Blk(ws, R_CHAIN, 35, R_CHAIN, 35).Width
    y = ws.Cells(R_CHAIN, 1).Top
    h = ws.rows(R_CHAIN).RowHeight + ws.rows(R_CHAIN + 1).RowHeight + ws.rows(R_CHAIN + 2).RowHeight
    w = (right0 - left0 + 5 * 6) / 6
    caseCount = CasesCovered(famCount, d)
    emptyFilters = EmptyFilterCount()

    For i = 0 To 5
        Select Case i
            Case 0
                big = CStr(famCount)
                note = IIf(famCount > 0, "risk families in this run", "load the system output")
                state = IIf(famCount > 0, "ok", "idle")
            Case 1
                big = format$(caseCount, "#,##0")
                note = IIf(emptyFilters > 0, emptyFilters & " without a filter", "filters in place")
                state = IIf(caseCount = 0, "idle", IIf(emptyFilters > 0, "warn", "ok"))
            Case Else
                stg = CStr(titles(i))
                big = format$(Cnt(d, "stage:" & stg & ":PASS"), "#,##0") & " / " & format$(Checked(d, "stage:" & stg), "#,##0")
                If Checked(d, "stage:" & stg) = 0 Then
                    big = "-": note = "not run yet": state = "idle"
                ElseIf Cnt(d, "stage:" & stg & ":FAIL") > 0 Then
                    note = format$(Cnt(d, "stage:" & stg & ":FAIL"), "#,##0") & " differences": state = "bad"
                ElseIf Checked(d, "stage:" & stg) > Cnt(d, "stage:" & stg & ":PASS") Then
                    note = format$(Checked(d, "stage:" & stg) - Cnt(d, "stage:" & stg & ":PASS"), "#,##0") & " to review or missing": state = "warn"
                Else
                    note = "all reconciled": state = "ok"
                End If
        End Select

        On Error Resume Next
        ws.Shapes("HomeChain" & i).Delete: ws.Shapes("HomeChainDot" & i).Delete
        On Error GoTo 0
        x = left0 + i * (w - 6)
        Set sh = ws.Shapes.AddShape(IIf(i = 0, msoShapePentagon, msoShapeChevron), x, y, w, h)
        sh.name = "HomeChain" & i
        On Error Resume Next
        sh.Adjustments.item(1) = 0.18
        On Error GoTo 0
        Select Case state
            Case "ok": fill = UI_OK_SOFT
            Case "warn": fill = UI_WARN_SOFT
            Case "bad": fill = UI_BAD_SOFT
            Case Else: fill = UI_SUBTLE
        End Select
        sh.fill.ForeColor.RGB = fill
        sh.line.ForeColor.RGB = UI_LINE: sh.line.Weight = 0.75
        sh.Shadow.Visible = msoFalse
        sh.Placement = xlMove
        sh.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & CStr(procs(i))
        With sh.TextFrame2
            .MarginLeft = IIf(i = 0, 12, 24): .MarginRight = 14: .MarginTop = 4: .MarginBottom = 4
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoTrue
            .TextRange.text = CStr(titles(i)) & vbCr & big & vbCr & note
            .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            .TextRange.Font.name = UI_FONT
            .TextRange.Font.fill.ForeColor.RGB = UI_MUTED
            .TextRange.Font.Size = 8.5
            With .TextRange.Paragraphs(1).Font
                .Size = 9.5: .Bold = msoTrue: .fill.ForeColor.RGB = UI_INK
            End With
            With .TextRange.Paragraphs(2).Font
                .Size = 15: .Bold = msoTrue: .fill.ForeColor.RGB = UI_INK
            End With
        End With
        Set dot = ws.Shapes.AddShape(msoShapeOval, x + w - 30, y + 9, 8, 8)
        dot.name = "HomeChainDot" & i
        dot.line.Visible = msoFalse: dot.Shadow.Visible = msoFalse: dot.Placement = xlMove
        Select Case state
            Case "ok": dot.fill.ForeColor.RGB = UI_OK_DOT
            Case "warn": dot.fill.ForeColor.RGB = UI_WARN_DOT
            Case "bad": dot.fill.ForeColor.RGB = UI_BAD_DOT
            Case Else: dot.fill.ForeColor.RGB = UI_FAINT
        End Select
        dot.OnAction = sh.OnAction
    Next i
End Sub

Private Sub FillHeatmap(ByVal ws As Worksheet, ByVal d As Object)
    Dim fams As Collection, i As Long, r As Long, j As Long, stages As Variant, fam As String, k As String
    Dim cases As Object, key As Variant, n As Long
    stages = Array("Base", "Pre-shock", "Shock", "Post-shock")
    Set fams = d("families")
    Set cases = d("cases")
    For r = R_PANEL + 2 To R_PANEL + 1 + MAX_FAMILIES
        Blk(ws, r, 1, r, 1).Value2 = Empty: Blk(ws, r, 4, r, 4).Value2 = Empty
        For j = 0 To 3: Blk(ws, r, 5 + j * 2, r, 5 + j * 2).Value2 = Empty: Next j
    Next r
    If fams.count = 0 Then
        Blk(ws, R_PANEL + 2, 1, R_PANEL + 2, 1).Value2 = "Not run yet"   ' short: the name column is narrow
        Blk(ws, R_PANEL + 2, 1, R_PANEL + 2, 1).Font.Color = UI_FAINT
        Exit Sub
    End If
    Blk(ws, R_PANEL + 2, 1, R_PANEL + 2, 1).Font.Color = UI_TEXT
    r = R_PANEL + 1
    For i = 1 To fams.count
        If r >= R_PANEL + 1 + MAX_FAMILIES Then Exit For
        fam = fams(i)
        If Not HomeIsRiskFamily(fam) Then GoTo NextFamily
        r = r + 1
        Blk(ws, r, 1, r, 1).Value2 = FamilyName(fam)
        n = HomeFamilyCaseCount(fam)
        If n = 0 Then
            For Each key In cases.keys
                If cases(key) = fam Then n = n + 1
            Next key
        End If
        Blk(ws, r, 4, r, 4).Value2 = n
        For j = 0 To 3
            k = "fam:" & fam & ":" & stages(j)
            If Checked(d, k) > 0 Then
                Blk(ws, r, 5 + j * 2, r, 5 + j * 2).Value2 = Cnt(d, k & ":PASS") / Checked(d, k)
            Else
                Blk(ws, r, 5 + j * 2, r, 5 + j * 2).Value2 = ChrW(8211)
                Blk(ws, r, 5 + j * 2, r, 5 + j * 2).Font.Color = UI_FAINT
            End If
        Next j
NextFamily:
    Next i
End Sub

Private Sub FillDifferences(ByVal ws As Worksheet, ByVal d As Object)
    Dim i As Long, j As Long, n As Long, idx() As Long, mag() As Double, tmpL As Long, tmpD As Double
    Dim r As Long, diff As Double, rate As Boolean, maxAmt As Double, bars As Long, rg As Range, p As Long
    For r = R_PANEL + 2 To R_PANEL + 1 + MAX_DIFFS
        Blk(ws, r, 14, r, 14).Value2 = Empty: Blk(ws, r, 17, r, 17).Value2 = Empty
        Blk(ws, r, 19, r, 19).Value2 = Empty: Blk(ws, r, 24, r, 24).Value2 = Empty: Blk(ws, r, 26, r, 26).Value2 = Empty
    Next r
    If Not d.Exists("rows") Then
        Blk(ws, R_PANEL + 2, 14, R_PANEL + 2, 14).Value2 = "No differences yet."
        Exit Sub
    End If
    ReDim idx(1 To UBound(mSlim, 1)): ReDim mag(1 To UBound(mSlim, 1))
    For i = 1 To UBound(mSlim, 1)
        If UCase$(CStr(ValueOr(mSlim(i, 1), ""))) = "FAIL" And IsNumeric(ValueOr(mSlim(i, 8), "x")) Then
            n = n + 1: idx(n) = i
            diff = CDbl(mSlim(i, 8))
            ' Rates and amounts share one list; a rate difference is ranked as if
            ' it were a percentage of a million so it sits beside amounts sensibly.
            If IsRate(i) Then mag(n) = Abs(diff) * 100000000# Else mag(n) = Abs(diff)
            If Not IsRate(i) Then If Abs(diff) > maxAmt Then maxAmt = Abs(diff)
        End If
    Next i
    If n = 0 Then
        Blk(ws, R_PANEL + 2, 14, R_PANEL + 2, 14).Value2 = "Nothing above tolerance."
        Blk(ws, R_PANEL + 2, 14, R_PANEL + 2, 14).Font.Color = UI_OK
        Exit Sub
    End If
    Blk(ws, R_PANEL + 2, 14, R_PANEL + 2, 14).Font.Color = UI_TEXT
    ' partial selection sort: the top MAX_DIFFS by magnitude
    For i = 1 To Application.Min(n, MAX_DIFFS)
        p = i
        For j = i + 1 To n
            If mag(j) > mag(p) Then p = j
        Next j
        tmpD = mag(i): mag(i) = mag(p): mag(p) = tmpD
        tmpL = idx(i): idx(i) = idx(p): idx(p) = tmpL
    Next i
    For i = 1 To Application.Min(n, MAX_DIFFS)
        r = R_PANEL + 1 + i
        j = idx(i)
        diff = CDbl(mSlim(j, 8))
        rate = IsRate(j)
        Blk(ws, r, 14, r, 14).Value2 = CStr(ValueOr(mSlim(j, 4), ValueOr(mSlim(j, 3), "")))
        Blk(ws, r, 17, r, 17).Value2 = CStr(ValueOr(mSlim(j, 2), ""))
        Blk(ws, r, 19, r, 19).Value2 = CStr(ValueOr(mSlim(j, 5), ""))
        Set rg = Blk(ws, r, 24, r, 24)
        ' Compact text so a ten-digit amount fits the panel; Results has the exact figure.
        rg.NumberFormat = "@"
        If rate Then
            rg.Value2 = IIf(diff < 0, "-", "+") & format$(Abs(diff) * 100, "0.00") & " pp"
        Else
            rg.Value2 = IIf(diff < 0, "-", "+") & CompactAmount(Abs(diff))
        End If
        rg.HorizontalAlignment = xlRight
        rg.Font.Color = IIf(diff < 0, UI_BAD, UI_TEXT): rg.Font.Bold = (diff < 0)
        If rate Or maxAmt = 0 Then
            bars = 1
        Else
            bars = CLng(14 * Abs(diff) / maxAmt + 0.5)
            If bars < 1 Then bars = 1
        End If
        With Blk(ws, r, 26, r, 26)
            .Value2 = String$(bars, ChrW(&H2588))
            .Font.Color = IIf(diff < 0, UI_BAD_BAR, UI_UP_BAR)
        End With
    Next i
End Sub

' A value is treated as a rate when both sides are small (ratios, percentages).
Private Function IsRate(ByVal i As Long) As Boolean
    Dim e As Variant, s As Variant
    e = ValueOr(mSlim(i, 6), Empty): s = ValueOr(mSlim(i, 7), Empty)
    If IsNumeric(e) And IsNumeric(s) And Not IsEmpty(e) And Not IsEmpty(s) Then
        IsRate = (Abs(CDbl(e)) < 5 And Abs(CDbl(s)) < 5)
    End If
End Function

Private Function LargestDifferenceText(ByVal d As Object) As String
    Dim i As Long, best As Double, found As Boolean
    If Not d.Exists("rows") Then Exit Function
    For i = 1 To UBound(mSlim, 1)
        If UCase$(CStr(ValueOr(mSlim(i, 1), ""))) = "FAIL" And IsNumeric(ValueOr(mSlim(i, 8), "x")) Then
            If Not IsRate(i) Then
                If Abs(CDbl(mSlim(i, 8))) > best Then best = Abs(CDbl(mSlim(i, 8))): found = True
            End If
        End If
    Next i
    If found Then LargestDifferenceText = CompactAmount(best) Else LargestDifferenceText = "rate differences only"
End Function

Private Function CompactAmount(ByVal v As Double) As String
    Select Case Abs(v)
        Case Is >= 1000000000#: CompactAmount = format$(v / 1000000000#, "0.00") & "bn"
        Case Is >= 1000000#: CompactAmount = format$(v / 1000000#, "0.00") & "m"
        Case Is >= 1000#: CompactAmount = format$(v / 1000#, "0.0") & "k"
        Case Else: CompactAmount = format$(v, "#,##0")
    End Select
End Function

Private Sub FillInputs(ByVal ws As Worksheet)
    Dim keys As Variant, labels As Variant, i As Long, rg As Range, loaded As Boolean, rowsN As Double, dt As String
    Dim src As Worksheet, r As Long, asOf As String, entity As String, outRows As Long
    keys = Array("ECL", "CAPRWA", "LCR", "LL", "NSFR", "CAP")
    labels = Array("ECL output", "CAPRWA", "LCR", "Legal liquidity", "NSFR", "Capital components")
    Set src = GetWorksheetSafe(ThisWorkbook, "Pre_Shock_Sources")
    For i = 0 To 5
        Set rg = InputCard(ws, i)
        loaded = False: rowsN = 0: dt = ""
        If Not src Is Nothing Then
            r = SourceListRow(CStr(keys(i)))
            If r > 0 Then
                If IsNumeric(src.Cells(r, 5).Value2) And Not IsEmpty(src.Cells(r, 5).Value2) Then rowsN = CDbl(src.Cells(r, 5).Value2)
                dt = SafeText(src.Cells(r, 7).Value2)
                loaded = (rowsN > 0) Or (Left$(SafeText(src.Cells(r, 8).Value2), 6) = "Loaded")
            End If
        End If
        CardText rg, CStr(labels(i)), IIf(loaded, format$(rowsN, "#,##0") & " rows" & IIf(Len(dt) > 0, "  " & ChrW(183) & "  " & dt, ""), "Not loaded"), IIf(loaded, UI_OK_DOT, UI_BAD_DOT)
        rg.Interior.Color = UI_WHITE
    Next i
    Set rg = InputCard(ws, 6)
    SystemOutputFacts asOf, entity, outRows
    CardText rg, "System output", IIf(outRows > 0, format$(outRows, "#,##0") & " rows" & IIf(Len(asOf) > 0, "  " & ChrW(183) & "  " & asOf, ""), "Not loaded"), IIf(outRows > 0, UI_BRAND, UI_BAD_DOT)
    rg.Interior.Color = UI_BRAND_TINT
End Sub

Private Sub CardText(ByVal rg As Range, ByVal title As String, ByVal note As String, ByVal dotColor As Long)
    With rg.Cells(1, 1)
        .Value2 = ChrW(9679) & "  " & title
        .Characters(1, 1).Font.Color = dotColor
        .Characters(1, 1).Font.Size = 9
    End With
    rg.Cells(2, 1).Value2 = note
End Sub

Private Function SourceListRow(ByVal key As String) As Long
    Dim i As Long, k As Variant
    For Each k In SourceKeys()
        If UCase$(CStr(k)) = UCase$(key) Then SourceListRow = PS_SRC_LIST_ROW + i: Exit Function
        i = i + 1
    Next k
End Function

Private Function LoadedInputs(ByRef dates As String) As Long
    Dim src As Worksheet, k As Variant, r As Long, n As Long, dt As String, first As String, mixed As Boolean
    Dim asOf As String, entity As String, outRows As Long
    Set src = GetWorksheetSafe(ThisWorkbook, "Pre_Shock_Sources")
    If Not src Is Nothing Then
        For Each k In SourceKeys()
            r = SourceListRow(CStr(k))
            If r > 0 Then
                If (IsNumeric(src.Cells(r, 5).Value2) And Not IsEmpty(src.Cells(r, 5).Value2) And Val(CStr(src.Cells(r, 5).Value2)) > 0) _
                   Or Left$(SafeText(src.Cells(r, 8).Value2), 6) = "Loaded" Then
                    n = n + 1
                    dt = SafeText(src.Cells(r, 7).Value2)
                    If Len(dt) > 0 Then
                        If Len(first) = 0 Then
                            first = dt
                        ElseIf dt <> first Then
                            mixed = True
                        End If
                    End If
                End If
            End If
        Next k
    End If
    SystemOutputFacts asOf, entity, outRows
    If outRows > 0 Then n = n + 1
    If n = 0 Then
        dates = "Nothing loaded yet"
    ElseIf mixed Then
        dates = "Reporting dates differ: check Inputs"
    ElseIf Len(first) > 0 Then
        dates = "All dated " & first
    Else
        dates = "Dates not recorded"
    End If
    LoadedInputs = n
End Function

Private Function SystemOutputRows() As Long
    Dim asOf As String, entity As String
    SystemOutputFacts asOf, entity, SystemOutputRows
End Function

' Reporting date, entity and data row count of the system output, read from its
' own columns so the band never shows a date the data does not carry.
Private Sub SystemOutputFacts(ByRef asOf As String, ByRef entity As String, ByRef dataRows As Long)
    Dim ws As Worksheet, hr As Long, lastR As Long, lastC As Long, c As Long, h As String, v As Variant
    asOf = "": entity = "": dataRows = 0
    On Error GoTo Done
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_SOURCE)
    If ws Is Nothing Then Exit Sub
    hr = DetectHeaderRow(ws)
    If hr <= 0 Then Exit Sub
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    lastC = ws.Cells(hr, ws.columns.count).End(xlToLeft).Column
    dataRows = Application.Max(0, lastR - hr)
    If dataRows = 0 Then Exit Sub
    For c = 1 To lastC
        h = UCase$(Trim$(CStr(ws.Cells(hr, c).Value2)))
        v = ws.Cells(hr + 1, c).Value2
        If h = "AS_OF_DATE" And Len(asOf) = 0 Then
            If IsNumeric(v) And Not IsEmpty(v) Then asOf = format$(CDate(CDbl(v)), "dd-mmm-yyyy") Else asOf = SafeText(v)
        ElseIf (h = "ENTITY_NAME" Or h = "ENTITY_CODE") And Len(entity) = 0 Then
            entity = SafeText(v)
        End If
    Next c
Done:
End Sub

Private Function CasesCovered(ByRef famCount As Long, ByVal d As Object) As Long
    CasesCovered = HomeCaseCounts(famCount)
End Function

Private Function EmptyFilterCount() As Long
    Dim ws As Worksheet, lastR As Long, r As Long
    Set ws = GetWorksheetSafe(ThisWorkbook, CASES_SHEET)
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.rows.count, 2).End(xlUp).row
    For r = 8 To lastR
        If Len(SafeText(ws.Cells(r, 2).Value2)) > 0 And UCase$(SafeText(ws.Cells(r, 1).Value2)) <> "NO" Then
            If Len(SafeText(ws.Cells(r, 8).Value2)) = 0 Then EmptyFilterCount = EmptyFilterCount + 1
        End If
    Next r
End Function

' Value sources: rows still needing a file or figure, from the STATUS column.
Private Function SetupGaps(ByRef checkedYet As Boolean) As Long
    Dim ws As Worksheet, h As Long, c As Long, lastR As Long, r As Long, s As String
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_SHEET_NAME)
    If ws Is Nothing Then Exit Function
    h = 9
    For c = 1 To 12
        If UCase$(SafeText(ws.Cells(h, c).Value2)) = "STATUS" Then Exit For
    Next c
    If c > 12 Then Exit Function
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = h + 1 To lastR
        s = SafeText(ws.Cells(r, c).Value2)
        If Len(s) > 0 Then checkedYet = True
        If InStr(1, s, "needs", vbTextCompare) > 0 Or InStr(1, s, "Unknown", vbTextCompare) > 0 Then SetupGaps = SetupGaps + 1
    Next r
End Function

Private Sub FillNextStep(ByVal ws As Worksheet, ByVal d As Object)
    Dim titles As Variant, notes(0 To 3) As String, done(0 To 3) As Boolean, i As Long, r As Long, cur As Long
    Dim loaded As Long, dates As String, gaps As Long, checkedYet As Boolean, st As String
    Dim mainCap As String, mainProc As String, altCap As String, altProc As String, rg As Range
    titles = Array("Load", "Check setup", "Run", "Review and hand over")

    loaded = LoadedInputs(dates)
    done(0) = (loaded >= 2)
    notes(0) = loaded & " of 7 files"
    If InStr(1, dates, "differ", vbTextCompare) > 0 Then
        notes(0) = notes(0) & " " & ChrW(183) & " dates differ"
    ElseIf Left$(dates, 10) = "All dated " Then
        notes(0) = notes(0) & " " & ChrW(183) & " " & Mid$(dates, 11)
    End If

    gaps = SetupGaps(checkedYet)
    done(1) = checkedYet And gaps = 0
    If Not checkedYet Then
        notes(1) = "Not checked yet"
    ElseIf gaps > 0 Then
        notes(1) = gaps & " value(s) still need a file or figure"
    Else
        notes(1) = "Every value has its source"
    End If

    st = CStr(ws.Range("JKB_RunStatus").Value2)
    done(2) = (d("n") > 0) And (Left$(st, 10) = "Reconciled" Or Left$(st, 6) = "Review")
    If d("n") = 0 Then
        notes(2) = "Not run yet"
    Else
        notes(2) = format$(d("n"), "#,##0") & " checks " & ChrW(183) & " " & format$(d("fail"), "#,##0") & " differences"
    End If
    done(3) = False
    notes(3) = "Console, then one review pack"

    cur = 3
    For i = 0 To 3
        If Not done(i) Then cur = i: Exit For
    Next i
    ' Setup gaps never block a run: the run itself reports missing evidence.
    If cur = 1 And checkedYet Then cur = 2
    If cur = 1 And Not checkedYet And done(2) Then cur = 3

    For i = 0 To 3
        r = R_PANEL + 1 + i * 2
        Set rg = Blk(ws, r, 30, r + 1, 30)
        If done(i) And i <> cur Then
            rg.Value2 = ChrW(10003): rg.Font.Color = UI_OK: rg.Interior.Color = UI_OK_BG
        ElseIf i = cur Then
            rg.Value2 = i + 1: rg.Font.Color = UI_WHITE: rg.Interior.Color = UI_BRAND
        Else
            rg.Value2 = i + 1: rg.Font.Color = UI_TEXT_2: rg.Interior.Color = UI_FILL
        End If
        Blk(ws, r, 31, r, 31).Value2 = titles(i)
        Blk(ws, r, 31, r, 31).Font.Color = IIf(i = cur, UI_BRAND, UI_INK)
        Blk(ws, r + 1, 31, r + 1, 31).Value2 = notes(i)
    Next i

    Select Case cur
        Case 0
            mainCap = "Load inputs": mainProc = "HomeLoad"
            ' Without the system output there is nothing to reconcile against.
            If SystemOutputRows() = 0 Then
                altCap = "System output": altProc = "UploadScenarioOutput"
            Else
                altCap = "One file": altProc = "UiLoadOneSource"
            End If
        Case 1: mainCap = "Check setup": mainProc = "HomeCheckSetup": altCap = "Test cases": altProc = "GoPreShockCases"
        Case 2: mainCap = "Run": mainProc = "HomeRun": altCap = "Check setup": altProc = "HomeCheckSetup"
        Case Else: mainCap = "Review console": mainProc = "HomeReview": altCap = "Review pack": altProc = "HomePack"
    End Select
    UiButton ws, "HomeNextMain", ChrW(9654) & "  " & mainCap, mainProc, Blk(ws, R_PANEL + 9, 30, R_PANEL + 9, 32), UI_BTN_PRIMARY, 10
    UiButton ws, "HomeNextAlt", altCap, altProc, Blk(ws, R_PANEL + 9, 33, R_PANEL + 9, 35), UI_BTN_SECONDARY, 9.5
End Sub
