Attribute VB_Name = "modBasePreShock"
Option Explicit

' ============================================================================
'  Base and pre-shock in three places (plans/base-preshock-simplify.md).
'
'  What a user sees of base and pre-shock:
'    Inputs        Pre_Shock_Sources: files, dates, status, as-of date matching.
'    Setup         Pre_Shock_Cases and Config_ValueSources (the only switch).
'    Results       Derived_Values: base and pre-shock from the input files and
'                  from the system, the differences and Reconciles.
'
'  Behind Configuration (hidden, opened from its toolbar): Pre_Shock_Metrics and
'  Pre_Shock_Fields. Kept for the code only (very hidden): the Pre_Shock
'  overview, Pre_Shock_Results and Pre_Shock_Breakdown. EnsurePreShockWorkspace
'  still builds them, so nothing that reads them breaks.
'
'  The overview's settings: independent checking is always on, "Explain the
'  figures" is off (Trace and the joined input's pivots explain a value), and
'  as-of date matching lives on Inputs (J8). The old cell on the overview is
'  kept in step so a workspace rebuild never loses the choice.
'
'  Parshva's calls of 2026-09-26: the old Dashboard and the Generate_Output
'  selection sheet are hidden. The Dashboard stays as the template every
'  review workbook is copied from; review workbooks are one action in the
'  review console.
' ============================================================================

Private Const ASOF_LABEL_CELL As String = "J7"
Private Const ASOF_VALUE_CELL As String = "J8"
' Pre_Shock!F12: the old home of the setting (PS_SRC_FIRST_ROW + 3, column F).
Private Const ASOF_OLD_CELL As String = "F12"

' --------------------------------------------------------- visibility ------

' Hides everything that is not one of the three places. Safe to run any number
' of times: from the workbook upgrade and after a workspace rebuild.
Public Sub ApplyBasePreShockLayout()
    Dim ws As Worksheet
    On Error Resume Next
    SetVisibility PRE_SHOCK_SHEET, xlSheetVeryHidden
    SetVisibility PS_RESULTS_SHEET, xlSheetVeryHidden
    SetVisibility PS_BREAKDOWN_SHEET, xlSheetVeryHidden
    SetVisibility PS_METRICS_SHEET, xlSheetHidden
    SetVisibility PS_FIELDS_SHEET, xlSheetHidden
    SetVisibility SHEET_DASHBOARD, xlSheetHidden
    SetVisibility SHEET_SELECTION, xlSheetHidden

    ' The per-metric "Value to use" column and the "Break down by" flag no
    ' longer decide anything.
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_METRICS_SHEET)
    If Not ws Is Nothing Then ws.Columns(10).Hidden = True
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_FIELDS_SHEET)
    If Not ws Is Nothing Then ws.Columns(PS_MAP_COLS).Hidden = True

    ' The breakdown is no longer written; its old rows are dead weight.
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_BREAKDOWN_SHEET)
    If Not ws Is Nothing Then
        If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > PS_FIRST_ROW Then
            ws.Range(ws.Rows(PS_FIRST_ROW), ws.Rows(ws.Rows.Count)).ClearContents
        End If
    End If

    EnsureAsOfSetting
    Err.Clear
    On Error GoTo 0
End Sub

' Review workbooks, one per test case: an action in the review console now that
' the Dashboard that carried the button is hidden.
Public Sub GenerateReviewWorkbooks()
    Dim ws As Worksheet
    If JKB_Busy Then Exit Sub
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_SELECTION)
    If Not ws Is Nothing Then ws.Visible = xlSheetVisible
    StartScenarioCategorySelection
End Sub

' From Workbook_SheetActivate: a backstage sheet opened from Configuration
' (or by a Go* button) goes back out of sight once the user moves on.
Public Sub KeepBackstageHidden(ByVal active As Object)
    Dim nm As Variant, ws As Worksheet
    If JKB_Busy Then Exit Sub
    If Not TypeOf active Is Worksheet Then Exit Sub
    If Not active.Parent Is ThisWorkbook Then Exit Sub
    On Error Resume Next
    For Each nm In Array(PRE_SHOCK_SHEET, PS_RESULTS_SHEET, PS_BREAKDOWN_SHEET, PS_METRICS_SHEET, PS_FIELDS_SHEET, _
                         SHEET_DASHBOARD, SHEET_SELECTION)
        If CStr(nm) <> active.Name Then
            Set ws = GetWorksheetSafe(ThisWorkbook, CStr(nm))
            If Not ws Is Nothing Then
                If ws.Visible = xlSheetVisible Then
                    Select Case CStr(nm)
                        Case PRE_SHOCK_SHEET, PS_RESULTS_SHEET, PS_BREAKDOWN_SHEET: ws.Visible = xlSheetVeryHidden
                        Case Else: ws.Visible = xlSheetHidden
                    End Select
                End If
            End If
        End If
    Next nm
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub SetVisibility(ByVal nm As String, ByVal state As XlSheetVisibility)
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, nm)
    If ws Is Nothing Then Exit Sub
    If ws.Visible = state Then Exit Sub
    ' The active sheet cannot be hidden; Home is always there to move to.
    If ActiveSheet Is ws Then OpenReconWorkbench
    ws.Visible = state
End Sub

' ------------------------------------------------- as-of date matching ------

' "IGNORE", "STRICT", "MATCH" or "" when the setting has never been made.
Public Function AsOfMatchingSetting() As String
    Dim ws As Worksheet, v As String
    On Error Resume Next
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET)
    If Not ws Is Nothing Then v = SafeUpperText(ws.Range(ASOF_VALUE_CELL).Value2)
    If Len(v) = 0 Then
        Set ws = GetWorksheetSafe(ThisWorkbook, PRE_SHOCK_SHEET)
        If Not ws Is Nothing Then v = SafeUpperText(ws.Range(ASOF_OLD_CELL).Value2)
    End If
    Err.Clear
    On Error GoTo 0
    AsOfMatchingSetting = v
End Function

' Writes both cells, so either one read back gives the same answer.
Public Sub WriteAsOfMatching(ByVal mode As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET)
    If Not ws Is Nothing Then ws.Range(ASOF_VALUE_CELL).Value2 = mode
    Set ws = GetWorksheetSafe(ThisWorkbook, PRE_SHOCK_SHEET)
    If Not ws Is Nothing Then ws.Range(ASOF_OLD_CELL).Value2 = mode
    Err.Clear
    On Error GoTo 0
End Sub

' Before a workspace rebuild: the rebuild carries the overview's cell across,
' so the choice made on Inputs is copied there first.
Public Sub KeepAsOfSetting()
    Dim ws As Worksheet, v As String
    On Error Resume Next
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET)
    If Not ws Is Nothing Then v = SafeText(ws.Range(ASOF_VALUE_CELL).Value2)
    If Len(v) > 0 Then
        Set ws = GetWorksheetSafe(ThisWorkbook, PRE_SHOCK_SHEET)
        If Not ws Is Nothing Then ws.Range(ASOF_OLD_CELL).Value2 = v
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' The label, dropdown and note on Inputs, filled from the old cell the first time.
Private Sub EnsureAsOfSetting()
    Dim ws As Worksheet, v As String
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET)
    If ws Is Nothing Then Exit Sub
    v = AsOfMatchingSetting()
    If v = "IGNORE" Then v = "Ignore" Else v = "Strict"
    If Len(SafeText(ws.Range(ASOF_VALUE_CELL).Value2)) = 0 Then ws.Range(ASOF_VALUE_CELL).Value2 = v
    With ws.Range(ASOF_LABEL_CELL)
        .Value2 = "As-of date matching"
        .Font.Bold = True: .Font.Color = UI_WHITE: .Interior.Color = UI_NAVY
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ASOF_VALUE_CELL)
        .Font.Bold = True: .HorizontalAlignment = xlCenter
        .Borders.Color = UI_LINE
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Strict,Ignore"
    End With
    ws.Columns("J").ColumnWidth = 22
    If ws.Range(ASOF_VALUE_CELL).Comment Is Nothing Then
        ws.Range(ASOF_VALUE_CELL).AddComment "Strict: each input file must carry the test case's as-of date. " & _
            "Ignore: dates are not checked, and the results are flagged for review."
    End If
End Sub
