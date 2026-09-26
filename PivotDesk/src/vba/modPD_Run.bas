Option Explicit

' ============================================================================
'  Setting the desk up, and running what the Desk asks for.
' ============================================================================

Public Sub PD_Setup()
    Dim ws As Worksheet
    On Error Resume Next
    modPD_Theme.DarkNormal ThisWorkbook
    BuildDesk
    modPD_Files.BuildFilesSheet
    modPD_Config.BuildConfigSheet
    modPD_Config.CheckAll True
    modPD_Recon.BuildReconSheet
    BuildLogSheet
    OrderSheets
    modPD_Theme.RailEverywhere
    PrintSetup
    SettingSet "styled_version", TOOL_VERSION
    Set ws = GetSheet(SH_HOME)
    If Not ws Is Nothing Then ws.Activate
    modPD_Desk.RefreshDesk
    Err.Clear
End Sub

' A workbook saved by an older version still carries the older look on its
' table sheets. Re-dressing them is safe - the data is not touched - but it
' is not free, so it happens once per version rather than on every open.
Public Sub Restyle()
    Dim ws As Worksheet, lastR As Long
    On Error Resume Next
    If SettingGet("styled_version") = TOOL_VERSION Then Exit Sub

    modPD_Theme.DarkNormal ThisWorkbook
    modPD_Files.BuildFilesSheet
    modPD_Config.BuildConfigSheet
    modPD_Config.CheckAll True
    OrderSheets

    Set ws = GetSheet(SH_RECON)
    If Not ws Is Nothing Then
        ReconChrome ws
        lastR = LastRow(ws, 3)
        If lastR >= modPD_Theme.R_FIRST Then
            ws.Range(ws.Cells(modPD_Theme.R_FIRST, 4), ws.Cells(lastR, 6)).NumberFormat = NUM_FMT
            modPD_Theme.DressTable ws, 8, lastR, 7
        End If
    End If

    Set ws = GetSheet(SH_LOG)
    If Not ws Is Nothing Then
        LogChrome ws
        lastR = LastRow(ws, 4)
        If lastR >= modPD_Theme.R_FIRST Then DressLogBlock ws, lastR
    End If

    modPD_Theme.RailEverywhere
    PrintSetup
    SettingSet "styled_version", TOOL_VERSION
    SettingSet "welcome", "1"
    Err.Clear
End Sub

' ===================== the desk sheet =======================================

' The Desk is designed and ships in the workbook; nothing here draws it. If
' the sheet is somehow gone, a plain one says so rather than the tool
' quietly having no front page.
Private Sub BuildDesk()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_HOME)
    If ws Is Nothing Then
        Set ws = EnsureSheet(SH_HOME)
        modPD_Theme.Dress ws, TOOL_NAME & "   " & BANK_NAME, _
            "The Desk's design is missing from this copy of the workbook. The Files, Reconciliation " & _
            "and Activity sheets still work; download a fresh copy of Avati to get the Desk back."
        modPD_Theme.SetStatus ws, "The Desk could not be found.", "Check"
        Exit Sub
    End If
    modPD_Desk.RefreshDesk
End Sub

Private Sub BuildLogSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_LOG)
    LogChrome ws
End Sub

Private Sub LogChrome(ByVal ws As Worksheet)
    modPD_Theme.Dress ws, "Activity", _
        "What the desk did and why, newest first. A file that was not recognised says here what was missing."
    modPD_Theme.Head ws, Array("When", "Level", "Stage", "What happened", "Which file"), _
                        Array(20, 12, 14, 110, 40)
    modPD_Theme.SetStatus ws, "The newest entry is at the top. The Desk shows the latest three.", "Idle"
End Sub

Private Sub ReconChrome(ByVal ws As Worksheet)
    Dim msg As String
    msg = SafeText(ws.Cells(modPD_Theme.R_STATUS, 1).Value2)
    modPD_Theme.Dress ws, "Reconciliation", _
        "Each loaded output against control report 3 (the ledger, by COA) and control report 6 " & _
        "(the reporting balance, by account)."
    modPD_Theme.Head ws, Array("Control", "Framework", "Key", "Output", "Control", "Difference", _
                               "Verdict", "What it means"), _
                        Array(12, 18, 26, 18, 18, 18, 13, 74)
    If Len(msg) = 0 Then msg = "Nothing reconciled yet."
    modPD_Theme.SetStatus ws, msg, SettingGet("recon_level", "Idle")
End Sub

' Every existing log line dressed in one pass - the first open after an
' upgrade, when there can be a couple of thousand of them.
Private Sub DressLogBlock(ByVal ws As Worksheet, ByVal lastR As Long)
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(modPD_Theme.R_FIRST, 1), ws.Cells(lastR, 5))
    With rng
        .Interior.Color = modPD_Theme.C_ROW
        .Font.Name = modPD_Theme.UI_FONT
        .Font.Size = 9.5
        .Font.Color = modPD_Theme.C_TEXT
        .Font.Bold = False
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .WrapText = False
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = modPD_Theme.C_LINE
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = modPD_Theme.C_LINE
    End With
    With ws.Range(ws.Cells(modPD_Theme.R_FIRST, 1), ws.Cells(lastR, 1)).Font
        .Name = modPD_Theme.UI_MONO
        .Size = 9
        .Color = modPD_Theme.C_TEXT_3
    End With
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, 3), ws.Cells(lastR, 3)).Font.Name = modPD_Theme.UI_SEMI
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, 5), ws.Cells(lastR, 5)).Font.Color = modPD_Theme.C_TEXT_3
    ws.Range(ws.Rows(modPD_Theme.R_FIRST), ws.Rows(lastR)).RowHeight = 21
    modPD_Theme.PaintVerdictColumn ws, 2, lastR
    Err.Clear
End Sub

Private Sub PrintSetup()
    Dim ws As Worksheet
    Set ws = GetSheet(SH_SOURCES)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, modPD_Files.S_COLS
    Set ws = GetSheet(SH_RECON)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, 8
    Set ws = GetSheet(SH_LOG)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, 5
    Set ws = GetSheet(SH_CONFIG)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, 21
    Set ws = GetSheet(SH_FIELDS)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, 8
    Set ws = GetSheet(SH_BOOKS)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, 8
    Set ws = GetSheet(SH_CHARTS)
    If Not ws Is Nothing Then modPD_Theme.PrintReady ws, 20
End Sub

Private Sub OrderSheets()
    Dim order As Variant, i As Long, ws As Worksheet
    On Error Resume Next
    order = Array(SH_HOME, SH_SOURCES, SH_CONFIG, SH_CHARTS, SH_BOOKS, SH_FIELDS, SH_GALLERY, SH_RECON, SH_LOG)
    For i = 0 To UBound(order)
        Set ws = GetSheet(CStr(order(i)))
        If Not ws Is Nothing Then ws.Move Before:=ThisWorkbook.Worksheets(i + 1)
    Next i
    ' The seed's starter sheet, once the real ones exist.
    Set ws = GetSheet("Home")
    If Not ws Is Nothing And ThisWorkbook.Worksheets.count > 1 Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Err.Clear
End Sub

' ===================== the three actions ====================================
'
' PD_LoadFolder, PD_LoadFiles and PD_Reconcile are NOT re-declared here.
'
' They used to be, as one-line wrappers onto modPD_Files and modPD_Recon, and a
' Shape's OnAction is an unqualified name - so VBA saw two public procedures
' called PD_LoadFolder and refused to pick one: "Ambiguous name detected", on
' every button press. The buttons name the real procedures directly.

' Kept for anything still wired to the 1.0 name: builds what the Desk has
' switched on.
Public Sub PD_Pivots()
    modPD_Desk.PD_BuildSelected
End Sub

' ============================================================================
'  Build one workbook per framework named in args ("LCR,NSFR").
' ============================================================================
Public Sub BuildPivotsFor(ByVal args As String)
    Dim st As Object, parts As Variant, i As Long, fw As String, folder As String
    Dim made As String, notBuilt As String, errOut As String, Path As String, n As Long, t0 As Single

    If PD_Busy Then Exit Sub
    parts = Split(args, ",")
    If UBound(parts) < 0 Then Exit Sub

    ' Build what the Pivot config says, or nothing: a recipe that will not
    ' build is reported before twenty minutes are spent finding out.
    If modPD_Config.Engine() = "recipes" Then
        ' A deleted Pivot config comes back as the defaults - which build
        ' exactly what 1.0 built - rather than a build that makes nothing.
        If GetSheet(SH_CONFIG) Is Nothing Then
            modPD_Config.BuildConfigSheet
            modPD_Theme.Rail GetSheet(SH_CONFIG)
            modPD_Theme.Rail GetSheet(SH_FIELDS)
            modPD_Theme.Rail GetSheet(SH_BOOKS)
            modPD_Theme.Rail GetSheet(SH_CHARTS)
            modPD_Theme.Rail GetSheet(SH_GALLERY)
        End If
        If GetSheet(SH_BOOKS) Is Nothing Then modPD_Books.BuildBooksSheet
        If GetSheet(SH_CHARTS) Is Nothing Then modPD_Charts.BuildChartsSheet
        If modPD_Charts.CheckCharts(True) > 0 Then
            Notify "A chart that is on will not draw - the reason is beside it on Chart config, under What to fix.", _
                   V_BREAK
            modPD_Theme.GoTo_ SH_CHARTS
            Exit Sub
        End If
        If modPD_Books.CheckBooks(True) > 0 Then
            Notify "A row on Workbooks will not build - the reason is beside it, under What to fix.", V_BREAK
            modPD_Theme.GoTo_ SH_BOOKS
            Exit Sub
        End If
        If modPD_Config.CheckAll(True) > 0 Then
            modPD_Desk.RefreshDesk
            Notify "Some pivots on Pivot config will not build - the reason is beside each, under What to fix. " & _
                   "Fix them or switch them off, then build again.", V_BREAK
            modPD_Theme.GoTo_ SH_CONFIG
            Exit Sub
        End If
    End If

    folder = AskFolder("Where should the framework workbooks be written?", SettingGet("last_out_folder"))
    If Len(folder) = 0 Then Exit Sub

    On Error GoTo Failed
    modPD_Desk.BusyOn
    Set st = CaptureState(): PD_Busy = True
    t0 = Timer

    For i = 0 To UBound(parts)
        fw = Trim$(CStr(parts(i)))
        If Len(fw) > 0 Then
            errOut = ""
            Progress_ FwLabel(fw) & " - starting", i / (UBound(parts) + 1)
            Path = modPD_Build.BuildFramework(fw, folder, errOut)
            If Len(Path) > 0 Then
                n = n + modPD_Build.BooksMade()
                If modPD_Build.BooksMade() > 1 Then
                    made = made & IIf(Len(made) > 0, ", ", "") & FwLabel(fw) & " (" & modPD_Build.BooksMade() & ")"
                Else
                    made = made & IIf(Len(made) > 0, ", ", "") & FileLeaf(Path)
                End If
            Else
                LogIt V_BREAK, "Pivots", "Not built - " & errOut, FwLabel(fw)
                notBuilt = notBuilt & IIf(Len(notBuilt) > 0, ", ", "") & FwLabel(fw)
            End If
        End If
    Next i

    RestoreState st: PD_Busy = False
    modPD_Desk.NoteBuild folder, n
    modPD_Files.RefreshStatuses
    modPD_Desk.BusyOff

    If Len(notBuilt) = 0 Then
        Notify Words(n) & IIf(n = 1, " workbook", " workbooks") & " written to " & MidTrim(folder, 60) & _
               " in " & Format$(Timer - t0, "0") & "s. Each opens on its " & Chr$(34) & SH_GUIDE & Chr$(34) & " sheet.", V_OK
    Else
        Notify n & " built, " & notBuilt & " not built - Activity says why.", V_CHECK
    End If
    Exit Sub

Failed:
    Step_ "BUILD FAILED: " & Err.Number & " " & Err.Description
    RestoreState st: PD_Busy = False
    modPD_Desk.BusyOff
    LogIt V_BREAK, "Pivots", Err.Number & " " & Err.Description, ""
    Tell "The build stopped:" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
         "Activity has what had been done up to that point.", vbExclamation
End Sub

Public Function AskFolder(ByVal title As String, Optional ByVal start As String = "") As String
    On Error Resume Next
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If Len(start) > 0 Then .InitialFileName = PathJoin(start, "")
        If .Show = -1 Then AskFolder = .SelectedItems(1)
    End With
    Err.Clear
End Function

' ===================== starting over ========================================

Public Sub PD_Reset()
    Dim st As Object, ws As Worksheet, k As Variant, r As Long
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    If MsgBox("Clear every loaded file and every result?" & vbCrLf & vbCrLf & _
              "The files themselves are not touched, and workbooks already written stay where " & _
              "they are. Your preferences - which frameworks are switched on, the folders " & _
              "last used - are kept.", vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
    Set st = CaptureState(): PD_Busy = True
    On Error Resume Next
    Set ws = GetSheet(SH_SOURCES)
    If Not ws Is Nothing Then
        r = modPD_Theme.R_FIRST
        For Each k In modPD_Files.Slots.keys
            ws.Range(ws.Cells(r, modPD_Files.S_STATUS), ws.Cells(r, modPD_Files.S_COLS)).ClearContents
            r = r + 1
        Next k
    End If
    modPD_Files.ClearRefused
    SettingClear "m_"
    SettingClear "recon_"
    SettingClear "last_build_"
    SettingClear "build_sig"
    SettingClear "ctl3_"
    SettingClear "ctl6_"
    SettingClear "gap_"
    SettingClear "rb_"
    SettingSet "styled_version", ""
    PD_Setup
    RestoreState st: PD_Busy = False
    LogIt V_OK, "Reset", "Cleared.", ""
    modPD_Desk.RefreshDesk
    Notify "The desk is clear. Add files to start again.", V_OK
    Err.Clear
End Sub

Public Sub PD_ClearLog()
    Dim ws As Worksheet, lastR As Long
    If PD_Busy Then Exit Sub
    Set ws = GetSheet(SH_LOG)
    If ws Is Nothing Then Exit Sub
    If MsgBox("Clear the activity log?" & vbCrLf & vbCrLf & _
              "Loaded files and results are not affected.", vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
    On Error Resume Next
    lastR = LastRow(ws, 4)
    If lastR >= modPD_Theme.R_FIRST Then ws.Range(ws.Rows(modPD_Theme.R_FIRST), ws.Rows(lastR)).Delete
    LogIt V_OK, "Activity", "The log was cleared.", ""
    modPD_Desk.RefreshDesk
    Err.Clear
End Sub
