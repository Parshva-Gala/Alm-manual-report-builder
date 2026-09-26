Option Explicit

' ============================================================================
'  Setting the desk up, and running what the console asked for.
' ============================================================================

Public Sub PD_Setup()
    Dim ws As Worksheet
    On Error Resume Next
    BuildDesk
    modPD_Files.BuildFilesSheet
    modPD_Recon.BuildReconSheet
    BuildLogSheet
    OrderSheets
    modPD_Theme.RailEverywhere
    Set ws = GetSheet(SH_HOME)
    If Not ws Is Nothing Then ws.Activate
    Err.Clear
End Sub

' ===================== the desk sheet =======================================

Private Sub BuildDesk()
    Dim ws As Worksheet, r As Long
    Set ws = EnsureSheet(SH_HOME)
    ws.Cells.Clear
    modPD_Theme.ClearButtons ws, "pd_card_"

    modPD_Theme.Dress ws, TOOL_NAME & "   " & BANK_NAME, _
        "A desk for daily analysis. Everything it builds is a live PivotTable you can drag, " & _
        "slice and drill - not a picture of an analysis."

    modPD_Theme.SetStatus ws, "Press " & Chr$(34) & "Open the console" & Chr$(34) & _
        " to choose what to do.", "Idle"

    r = modPD_Theme.R_HDR
    Card ws, r, "1", "Add files", _
        "Point at a folder or pick files. Each is identified by the columns it carries, never by " & _
        "its name - the exports in this bank are routinely misnamed.", "PD_LoadFolder", "Choose a folder", _
        "PD_LoadFiles", "Pick files"
    r = r + 5
    Card ws, r, "2", "Build pivots", _
        "One workbook per framework: the rule-level Output, the Balance sheet, and one sheet per " & _
        "rule. All on a single pivot cache, so a slicer drives the lot.", "PD_Pivots", "Choose frameworks", _
        "", ""
    r = r + 5
    Card ws, r, "3", "Reconcile", _
        "The loaded outputs against control reports 3 and 6. Scope is established before any " & _
        "difference is called a break.", "PD_Reconcile", "Reconcile now", "", ""

    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 96
    ws.Columns(4).ColumnWidth = 2
End Sub

' A numbered block with its own actions. Three of these IS the desk - the whole
' tool is three things and the sheet should look like three things.
Private Sub Card(ByVal ws As Worksheet, ByVal r As Long, ByVal num As String, ByVal title As String, _
                 ByVal body As String, ByVal proc1 As String, ByVal cap1 As String, _
                 ByVal proc2 As String, ByVal cap2 As String)
    Dim x As Double, topPt As Double
    With ws.Cells(r, 1)
        .Value2 = num
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = modPD_Theme.C_HAIR
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(r, 2)
        .Value2 = title
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = modPD_Theme.C_INK
    End With
    With ws.Cells(r, 3)
        .Value2 = body
        .Font.Size = 10
        .Font.Color = modPD_Theme.C_MUTED
        .WrapText = False
    End With
    ws.Rows(r).RowHeight = 24
    ws.Rows(r + 1).RowHeight = 6
    ws.Rows(r + 2).RowHeight = 26
    ws.Rows(r + 3).RowHeight = 12

    With ws.Range(ws.Cells(r, 2), ws.Cells(r + 2, 3)).Borders(xlEdgeBottom)
        .Color = modPD_Theme.C_HAIR
    End With

    topPt = ws.Cells(r + 2, 2).Top + 2
    x = ws.Cells(r + 2, 2).Left
    If Len(proc1) > 0 Then x = PlaceButton(ws, cap1, proc1, x, topPt, 118, 1)
    If Len(proc2) > 0 Then x = PlaceButton(ws, cap2, proc2, x, topPt, 92, 0)
End Sub

Private Function PlaceButton(ByVal ws As Worksheet, ByVal caption As String, ByVal proc As String, _
                             ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                             ByVal kind As Long) As Double
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 22)
    If sh Is Nothing Then PlaceButton = x: Exit Function
    sh.Name = "pd_card_" & CLng(x) & "_" & CLng(y)
    sh.Placement = xlFreeFloating
    sh.Adjustments(1) = 0.24
    sh.Line.visible = msoFalse
    With sh.Shadow
        .visible = msoTrue
        .style = msoShadowStyleOuterShadow
        .Blur = 5
        .Transparency = 0.8
        .Size = 100
        .OffsetX = 0
        .OffsetY = 1.5
        .ForeColor.RGB = modPD_Theme.C_INK
    End With
    If kind = 1 Then
        sh.Fill.ForeColor.RGB = modPD_Theme.C_BRAND
    Else
        sh.Fill.ForeColor.RGB = RGB(238, 243, 240)
    End If
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Text = caption
        .TextRange.Font.Size = 9.5
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = modPD_Theme.UI_FONT
        If kind = 1 Then
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            .TextRange.Font.Fill.ForeColor.RGB = modPD_Theme.C_INK
        End If
    End With
    sh.OnAction = proc
    Err.Clear
    PlaceButton = x + w + 7
End Function

Private Sub BuildLogSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_LOG)
    modPD_Theme.Dress ws, "Activity", _
        "What the desk did and why, newest first. A file that was not recognised says here what was missing."
    modPD_Theme.Head ws, Array("When", "Level", "Stage", "What happened", "Which file"), _
                        Array(20, 12, 14, 96, 40)
    modPD_Theme.SetStatus ws, "Ready.", "Idle"
End Sub

Private Sub OrderSheets()
    Dim order As Variant, i As Long, ws As Worksheet
    On Error Resume Next
    order = Array(SH_HOME, SH_SOURCES, SH_RECON, SH_LOG)
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

' Straight to the console with the framework chooser already open.
Public Sub PD_Pivots()
    Dim cmd As String
    If Not modPD_Files.AnyFrameworkLoaded() Then
        Tell "No framework output has been loaded yet." & vbCrLf & vbCrLf & _
             "Add files first - the desk needs at least one LCR, NSFR or maturity ladder output " & _
             "before it can build pivots.", vbInformation
        Exit Sub
    End If
    cmd = modPD_Console.ShowConsole()
    If Len(cmd) > 0 Then modPD_Console.Dispatch cmd
End Sub

' ============================================================================
'  Build one workbook per framework named in args ("LCR,NSFR").
' ============================================================================
Public Sub BuildPivotsFor(ByVal args As String)
    Dim st As Object, parts As Variant, i As Long, fw As String, folder As String
    Dim made As String, errOut As String, Path As String, n As Long, t0 As Single

    If PD_Busy Then Exit Sub
    parts = Split(args, ",")
    If UBound(parts) < 0 Then Exit Sub

    folder = AskFolder("Where should the framework workbooks be written?")
    If Len(folder) = 0 Then Exit Sub

    On Error GoTo Failed
    Set st = CaptureState(): PD_Busy = True
    t0 = Timer

    For i = 0 To UBound(parts)
        fw = Trim$(CStr(parts(i)))
        If Len(fw) > 0 Then
            errOut = ""
            Step_ FwLabel(fw) & " - starting"
            Path = modPD_Build.BuildFramework(fw, folder, errOut)
            If Len(Path) > 0 Then
                n = n + 1
                made = made & vbCrLf & "   " & FileLeaf(Path)
            Else
                LogIt V_BREAK, "Pivots", "Not built - " & errOut, FwLabel(fw)
                made = made & vbCrLf & "   " & FwLabel(fw) & " - NOT BUILT: " & errOut
            End If
        End If
    Next i

    RestoreState st: PD_Busy = False
    modPD_Files.RefreshStatuses
    modPD_Theme.RailEverywhere

    Tell n & " workbook(s) written to:" & vbCrLf & vbCrLf & folder & vbCrLf & made & vbCrLf & vbCrLf & _
         "Each one opens on a " & Chr$(34) & SH_GUIDE & Chr$(34) & " sheet listing what is inside. " & _
         "Every sheet is a live pivot - drag a field, drop a slicer, double-click a total to see " & _
         "the rows behind it.", vbInformation
    Exit Sub

Failed:
    Step_ "BUILD FAILED: " & Err.Number & " " & Err.Description
    RestoreState st: PD_Busy = False
    LogIt V_BREAK, "Pivots", Err.Number & " " & Err.Description, ""
    Tell "The build stopped:" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
         "Activity has what had been done up to that point.", vbExclamation
End Sub

Public Function AskFolder(ByVal title As String) As String
    On Error Resume Next
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If .Show = -1 Then AskFolder = .SelectedItems(1)
    End With
    Err.Clear
End Function

' ===================== starting over ========================================

Public Sub PD_Reset()
    Dim st As Object, ws As Worksheet, k As Variant, r As Long
    If PD_Busy Then Exit Sub
    If MsgBox("Clear every loaded file and every result?" & vbCrLf & vbCrLf & _
              "The files themselves are not touched, and workbooks already written stay where " & _
              "they are.", vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
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
    PD_Setup
    RestoreState st: PD_Busy = False
    LogIt V_OK, "Reset", "Cleared.", ""
    Err.Clear
End Sub
