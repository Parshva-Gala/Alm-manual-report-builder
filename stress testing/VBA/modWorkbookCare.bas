Attribute VB_Name = "modWorkbookCare"
Option Explicit

' ============================================================================
'  Workbook care: the one-time upgrade that used to run on every open, the
'  shared operation guard and problem report, and the build-time cleanup of
'  duplicated sheet code.
'
'  Opening the workbook now only rebinds buttons and shows Home. The slow,
'  data-changing work (metric defaults, the test-case sync, restyling every
'  sheet) runs once per layout stamp: when a workbook's stamp is older than
'  WORKBOOK_LAYOUT_STAMP, on the next open, or whenever JKB_UpgradeWorkbook
'  is run. Bump the stamp when a change must be re-applied to workbooks
'  already in use.
' ============================================================================

Private Const WORKBOOK_LAYOUT_STAMP As String = "2026-09-26c"
Private Const STAMP_NAME As String = "JKB_LayoutStamp"
Private Const CARE_PROC As String = "WorkbookCare"

Private mUpgradePending As Boolean

' ---------------------------------------------------------- upgrade ---------

Public Function WorkbookNeedsUpgrade() As Boolean
    WorkbookNeedsUpgrade = (StoredStamp() <> WORKBOOK_LAYOUT_STAMP)
End Function

' The on-demand repair: forgets the stamp and reopens the tool, so the full
' upgrade runs exactly as it would on a workbook that has never had one.
Public Sub JKB_UpgradeWorkbook()
    If JKB_Busy Then Exit Sub
    ClearStamp
    InitializeJKBTool
End Sub

' Runs from InitializeJKBTool, inside its captured application state, before the
' buttons are rebound. Every step is logged rather than stopping the open: a
' workbook with no system output loaded yet cannot sync test cases, and that
' must not stop it opening.
Public Sub RunWorkbookUpgrade(ByVal reason As String)
    Application.StatusBar = "JKB: updating this workbook to the current layout..."
    CareLog LOG_LEVEL_INFO, "Upgrading workbook layout (" & reason & ") to " & WORKBOOK_LAYOUT_STAMP & "."

    On Error Resume Next
    modScenarioBuilder_Multi.UpgradeCapitalMetricDefaults PsMetricsSheet()
    If Err.Number <> 0 Then CareLog LOG_LEVEL_WARN, "Capital metric defaults not updated: " & Err.Description
    Err.Clear

    SyncPreShockTestCasesFromCurrentSource
    If Err.Number <> 0 Then CareLog LOG_LEVEL_INFO, "Test cases not refreshed (normal before the first system output upload): " & Err.Description
    Err.Clear

    Sheet1.StyleWorkbookUI
    If Err.Number <> 0 Then CareLog LOG_LEVEL_WARN, "Sheet styling not refreshed: " & Err.Description
    Err.Clear

    StylePreShockSheet PreShockSheet()
    If Err.Number <> 0 Then CareLog LOG_LEVEL_WARN, "Pre-shock styling not refreshed: " & Err.Description
    Err.Clear

    ApplyBasePreShockLayout
    If Err.Number <> 0 Then CareLog LOG_LEVEL_WARN, "Base and pre-shock layout not applied: " & Err.Description
    Err.Clear
    On Error GoTo 0

    mUpgradePending = True
End Sub

' The part of the upgrade that must follow the rebind: gridlines are a per-window
' setting, so each visible sheet has to be activated once. Then the stamp.
Public Sub FinishWorkbookUpgrade(Optional ByVal quiet As Boolean = True)
    Dim ws As Worksheet
    If Not mUpgradePending Then Exit Sub
    mUpgradePending = False
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayOutline = True
        End If
    Next ws
    If Err.Number <> 0 Then CareLog LOG_LEVEL_WARN, "Gridlines not updated on every sheet: " & Err.Description
    Err.Clear
    On Error GoTo 0
    WriteStamp
End Sub

Private Function StoredStamp() As String
    Dim nm As Object
    On Error Resume Next
    Set nm = ThisWorkbook.Names(STAMP_NAME)
    On Error GoTo 0
    If nm Is Nothing Then Exit Function
    StoredStamp = Replace(Mid$(CStr(nm.RefersTo), 2), """", "")
End Function

Private Sub WriteStamp()
    ClearStamp
    On Error Resume Next
    ThisWorkbook.Names.Add STAMP_NAME, "=""" & WORKBOOK_LAYOUT_STAMP & """", False
    If Err.Number <> 0 Then CareLog LOG_LEVEL_WARN, "Layout stamp not saved; the upgrade will run again next open: " & Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ClearStamp()
    On Error Resume Next
    ThisWorkbook.Names(STAMP_NAME).Delete
    Err.Clear
    On Error GoTo 0
End Sub

' ---------------------------------------------------------- operations ------

' The guard every long-running action needs, in one place:
'
'     Dim op As Object
'     Set op = BeginOperation("JKB: building reports...")
'     If op Is Nothing Then Exit Sub          ' another action is running
'     On Error GoTo Failed
'     ...
'     EndOperation op
'     Exit Sub
' Failed:
'     EndOperation op
'     ReportProblem "Build reports", "The reports were not built.", Err.Description, "Check Build_Log.", "BuildReports"
'
' EndOperation always releases the busy flag and restores Excel's settings, so a
' failure can no longer leave the tool locked or screen updating switched off.
Public Function BeginOperation(Optional ByVal statusText As String = "") As Object
    If JKB_Busy Then Exit Function
    Set BeginOperation = CaptureState()
    JKB_Busy = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If Len(statusText) > 0 Then Application.DisplayStatusBar = True: Application.StatusBar = statusText
End Function

Public Sub EndOperation(ByVal state As Object)
    If state Is Nothing Then Exit Sub
    JKB_Busy = False
    RestoreState state
End Sub

' Writes the problem to Build_Log first, then shows it, so every failure a user
' saw can be found again afterwards. The wording goes through UiProblem.
Public Sub ReportProblem(ByVal title As String, ByVal headline As String, Optional ByVal reason As String = "", _
                         Optional ByVal nextStep As String = "", Optional ByVal procedure As String = "")
    Dim where As String
    where = procedure: If Len(where) = 0 Then where = title
    On Error Resume Next
    LogIssue LOG_LEVEL_ERROR, where, headline, reason
    Err.Clear
    On Error GoTo 0
    UiProblem title, headline, reason, nextStep
End Sub

Private Sub CareLog(ByVal level As String, ByVal message As String)
    On Error Resume Next
    LogIssue level, CARE_PROC, message
    Err.Clear
    On Error GoTo 0
End Sub

' ---------------------------------------------------------- build-time ------

' Run once by JKB_ApplyReleaseSetup. It deliberately does not write the layout
' stamp: the build never opens the workbook, so the first real open must still
' run the upgrade that applies the current styling.
Public Sub JKB_ApplySimplifySetup(Optional ByVal quiet As Boolean = True)
    JKB_RemoveDuplicateSheetHandlers
End Sub

' Every element sheet module carried the same Worksheet_Activate handler, which
' redraws the sheet toolbar through the Dashboard. ThisWorkbook's
' Workbook_SheetActivate already does that for every sheet, so each tab switch
' drew the toolbar twice. This removes the copies, and only exact copies: a
' handler with anything else in it is left alone. Dashboard (Sheet1) keeps its
' own, because it travels into the generated review workbooks, which have no
' ThisWorkbook code.
'
' Needs "Trust access to the VBA project object model", which the build turns
' on. Without it this logs and does nothing, so it is safe to run anywhere.
Public Sub JKB_RemoveDuplicateSheetHandlers(Optional ByVal quiet As Boolean = True)
    Dim comp As Object, cm As Object, first As Long, n As Long, removed As Long, kept As Long
    On Error GoTo NoAccess
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 100 And comp.Name <> "ThisWorkbook" And comp.Name <> "Sheet1" Then
            Set cm = comp.CodeModule
            first = 0: n = 0
            On Error Resume Next
            first = cm.ProcStartLine("Worksheet_Activate", 0)
            n = cm.ProcCountLines("Worksheet_Activate", 0)
            Err.Clear
            On Error GoTo NoAccess
            If first > 0 And n > 0 Then
                If IsToolbarOnlyHandler(cm.Lines(first, n)) Then
                    cm.DeleteLines first, n
                    removed = removed + 1
                Else
                    kept = kept + 1
                End If
            End If
        End If
    Next comp
    CareLog LOG_LEVEL_INFO, "Removed " & removed & " duplicate sheet toolbar handlers; kept " & kept & " that do more."
    Exit Sub
NoAccess:
    CareLog LOG_LEVEL_WARN, "Duplicate sheet handlers not removed (needs trust access to the VBA project): " & Err.Description
End Sub

' True only for the exact boilerplate: open, Dashboard lookup, one
' Application.Run of EnsureRibbon, the error exit and its message.
Private Function IsToolbarOnlyHandler(ByVal code As String) As Boolean
    Dim ln As Variant, t As String, sawRun As Boolean
    For Each ln In Split(Replace(code, vbCr, ""), vbLf)
        t = LCase$(Trim$(CStr(ln)))
        If Len(t) = 0 Or Left$(t, 1) = "'" Then
            ' blank or comment
        ElseIf t = "private sub worksheet_activate()" Or t = "dim dash as worksheet" _
            Or t = "on error goto failed" Or t = "set dash = me.parent.worksheets(""dashboard"")" _
            Or t = "exit sub" Or t = "failed:" Or t = "end sub" Then
            ' part of the boilerplate
        ElseIf Left$(t, 16) = "application.run " And InStr(t, ".ensureribbon"", me") > 0 Then
            sawRun = True
        ElseIf Left$(t, 7) = "msgbox " And InStr(t, "err.description") > 0 Then
            ' the handler's own error message
        Else
            Exit Function
        End If
    Next ln
    IsToolbarOnlyHandler = sawRun
End Function
