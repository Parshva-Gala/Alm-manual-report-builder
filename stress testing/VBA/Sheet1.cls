Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private mRules As Variant
Private mRuleColumns As Object
Private mNumberFormats As Object

' Name of the appearance sheet carried by a generated workbook.
Private Const OUT_FORMAT_SHEET As String = "Format"

' Last element sheet the reviewer visited, for the Back button.
Private mLastElement As String

Private Sub Worksheet_Activate()
    On Error GoTo Failed
    EnsureRibbon Me
    Exit Sub
Failed:
    MsgBox "Could not refresh the sheet controls." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub

Public Sub RebindAll()
    Dim ws As Worksheet
    For Each ws In Me.Parent.Worksheets
        If ws.Visible = xlSheetVisible Then EnsureRibbon ws
    Next ws
End Sub

Public Sub EnsureRibbon(ByVal ws As Worksheet)
    Dim n As Long, hidden As Boolean, gn As Long, groupsHidden As Boolean, isOutput As Boolean, isMaster As Boolean
    Dim host As Workbook, marker As String
    If CStr(ws.Range("A1").Value2) <> "JKB" Then
        ws.rows("1:2").Insert Shift:=xlDown
        ws.Range("A1").Value2 = "JKB"
    End If
    ' The master formats sheets that live in a generated workbook, so the ribbon must
    ' describe the sheet's OWN workbook, not the workbook this code happens to run in.
    Set host = ws.Parent
    marker = ""
    On Error Resume Next
    marker = UCase$(Trim$(CStr(host.Worksheets("Dashboard").Range("N1").Value2)))
    On Error GoTo 0
    isOutput = (marker = "OUTPUT")
    isMaster = Not isOutput
    ' Two strips float over row 1: what this sheet DOES, and how to READ it. 42
    ' points is exactly enough for both without either touching the content, which
    ' is the sort of thing that looks broken the moment it is a point out.
    ws.rows(1).RowHeight = 42: ws.rows(2).RowHeight = 4
    ' The band runs the full width the strips can reach. Stopping it at column K
    ' left every button past that floating over bare gridlines, which is what made
    ' the toolbar look like something that had come loose rather than a toolbar.
    With ws.Range("A1:BZ1")
        .Interior.Color = RGB(249, 250, 251)
        .Font.name = "Aptos": .Font.Size = 8.5: .Font.Bold = True
        .Font.Color = RGB(52, 64, 84)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(214, 226, 245)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    On Error Resume Next
    ws.Shapes("shpToggleFormulaRows").Delete
    ws.Buttons("btnToggleFormulaRows").Delete
    ws.Shapes("JKB_GroupToggle").Delete
    ws.Shapes("JKB_Config").Delete
    ws.Shapes("JKB_SideConfig").Delete
    ws.Shapes("JKB_SideFormat").Delete
    ws.Shapes("JKB_SideBase").Delete
    ws.Shapes("JKB_SidePreShock").Delete
    ws.Shapes("JKB_PreShockStudio").Delete
    ws.Shapes("JKB_PreShockUpload").Delete
    ws.Shapes("JKB_PreShockUploadCap").Delete
    ws.Shapes("JKB_PreShockRefresh").Delete
    ws.Shapes("JKB_PsTest").Delete
    ws.Shapes("JKB_PsGoFormulas").Delete
    ws.Shapes("JKB_PsGoDims").Delete
    ws.Shapes("JKB_PsGoCases").Delete
    ws.Shapes("JKB_PsGoResults").Delete
    ' Buttons from the retired Base_Testing / Pre_Shock_Config pair.
    ws.Shapes("JKB_BasePreShock").Delete
    ws.Shapes("JKB_BaseUpload").Delete
    ws.Shapes("JKB_BaseRefresh").Delete
    ws.Shapes("JKB_FilterLoad").Delete
    ws.Shapes("JKB_FilterBuilder").Delete
    ws.Shapes("JKB_FilterSave").Delete
    ws.Shapes("JKB_SevMod").Delete
    ws.Shapes("JKB_SevMed").Delete
    ws.Shapes("JKB_SevSev").Delete
    ws.Shapes("JKB_OutFormat").Delete
    ws.Shapes("JKB_OutApply").Delete
    ws.Shapes("JKB_BackElement").Delete
    ' Retired in design v2: the view strip (Excel's own zoom does this) and the
    ' per-type loaders, which are now one "Load inputs" plus "Load one input".
    Dim vw As Variant
    ' Retired with the three-place base and pre-shock layout.
    For Each vw In Array("JKB_PsNavHome", "JKB_PsNavMetrics", "JKB_PsNavFields", "JKB_PsNavResults", "JKB_PsUseDerived", _
                         "JKB_PsJoin", "JKB_PsManual")
        ws.Shapes(CStr(vw)).Delete
    Next vw
    For Each vw In Array("VW_ZoomOut", "VW_ZoomIn", "VW_ZoomFit", "VW_Zoom100", "VW_FitCols", "VW_FitRows", "VW_Copy", "VW_Manual", _
                         "JKB_ReconHome", "JKB_PsUpFolder", "JKB_PsUpECL", "JKB_PsUpCap", "JKB_PsUpLL", "JKB_PsUpLCR", "JKB_PsUpCapC")
        ws.Shapes(CStr(vw)).Delete
    Next vw
    If Not isMaster Then
        ws.Shapes("JKB_ReconHome").Delete
        ws.Shapes("JKB_ReconStart").Delete
        ws.Shapes("JKB_ConsoleHome").Delete
    End If
    On Error GoTo 0
    Button ws, "JKB_Dashboard", "Home", "GoHome", 42, 2, 80, 17, True
    If isMaster Then Button ws, "JKB_Config", "Config", "GoConfig", 127, 2, 62, 17, True
    Button ws, "JKB_Formula", FormulaCaption(ws, n, hidden), "ToggleFormulaRows", IIf(isMaster, 194, 127), 2, 111, 17, True
    If n = 0 Then
        ws.Shapes("JKB_Formula").OnAction = ""
        ws.Shapes("JKB_Formula").fill.ForeColor.RGB = RGB(152, 162, 179)
    End If
    If CountDetailGroups(ws, groupsHidden) > 0 Then
        Button ws, "JKB_GroupToggle", IIf(groupsHidden, "Expand groups", "Collapse groups"), "ToggleDetailGroups", IIf(isMaster, 310, 243), 2, 112, 17, True
    End If
    ' Generated workbooks get severity visibility controls and their own appearance sheet.
    If isOutput Then
        If IsElementSheet(ws) Then
            Button ws, "JKB_SevMod", IIf(SeverityHidden(ws, 0), "+ Moderate", "- Moderate"), "ToggleModerate", 360, 2, 82, 17, True
            Button ws, "JKB_SevMed", IIf(SeverityHidden(ws, 1), "+ Medium", "- Medium"), "ToggleMedium", 445, 2, 78, 17, True
            Button ws, "JKB_SevSev", IIf(SeverityHidden(ws, 2), "+ Severe", "- Severe"), "ToggleSevere", 526, 2, 74, 17, True
        Else
            ' Anywhere that is not an element sheet, offer the way back to one.
            Button ws, "JKB_BackElement", "Back to sheet", "GoElementSheet", 360, 2, 96, 17, True
        End If
        Button ws, "JKB_OutFormat", "Appearance", "GoOutputFormat", 605, 2, 84, 17, True
        If ws.name = OUT_FORMAT_SHEET Then Button ws, "JKB_OutApply", "Apply formatting", "ApplyOutputFormatting", 694, 2, 112, 17, True
    End If
    If ws.name = "Dashboard" Then
        If isMaster Then
            Button ws, "JKB_Upload", "Import system output", "UploadScenarioOutput", ws.Range("F8").Left, ws.Range("F8:J10").Top, ws.Range("F8:J10").Width, ws.Range("F8:J10").Height, False
            Button ws, "JKB_Generate", "Generate review workbooks", "BuildAllScenarioSheets", ws.Range("F12").Left, ws.Range("F12:J14").Top, ws.Range("F12:J14").Width, ws.Range("F12:J14").Height, False
            Button ws, "JKB_SideConfig", "Configuration", "GoConfig", ws.Range("B8").Left, ws.Range("B8:C9").Top, ws.Range("B8:C9").Width, ws.Range("B8:C9").Height, True
            Button ws, "JKB_SideFormat", "Formatting", "GoFormat", ws.Range("B11").Left, ws.Range("B11:C12").Top, ws.Range("B11:C12").Width, ws.Range("B11:C12").Height, True
            Button ws, "JKB_SidePreShock", "Base / pre-shock", "GoPreShock", ws.Range("B14").Left, ws.Range("B14:C15").Top, ws.Range("B14:C15").Width, ws.Range("B14:C15").Height, False
        End If
    End If
    If isMaster And ws.name = "Config_MasterFormatting" Then Button ws, "JKB_ApplyFormat", "Apply formatting", "ApplyFormatting", ws.Range("F3").Left, ws.Range("F3").Top, ws.Range("F3:G3").Width, ws.Range("F3:G3").Height, True
    ' The pre-shock workspace is six sheets, and every one of them carries the same
    ' navigation strip followed by its own actions - so moving between them costs a
    ' click and never a scroll. The single crowded sheet these replaced put five
    ' differently-shaped blocks at rows 26, 279, 342 and 395 of one document.
    If isMaster And IsPreShockSheet(ws) Then
        Dim px As Double
        px = 427
        ' The console first, because it is the interface: the sheets are the record.
        Button ws, "JKB_PsConsole", "Open console", "OpenStressConsole", px, 2, 92, 17, False: px = px + 100
        ' Three places: Inputs, Setup (Test cases, Value sources) and Results.
        Button ws, "JKB_PsNavSrc", "Inputs", "GoPreShockSources", px, 2, 64, 17, False: px = px + 68
        Button ws, "JKB_PsNavCases", "Test cases", "GoPreShockCases", px, 2, 76, 17, False: px = px + 80
        Button ws, "JKB_PsNavVS", "Value sources", "GoValueSources", px, 2, 94, 17, False: px = px + 98
        Button ws, "JKB_PsNavDerived", "Base & pre-shock results", "GoDerivedValues", px, 2, 150, 17, False: px = px + 154

        Select Case ws.name
            Case "Pre_Shock"
                Button ws, "JKB_PsUpAll", "Load inputs", "UploadAllSources", px, 2, 88, 17, False: px = px + 92
                Button ws, "JKB_PsRun", "Run", "RunStressReconciliation", px, 2, 74, 17, False
            Case "Pre_Shock_Sources"
                Button ws, "JKB_PsUpAll", "Load inputs", "UploadAllSources", px, 2, 88, 17, False: px = px + 92
                Button ws, "JKB_PsUpOne", "Load one input", "UiLoadOneSource", px, 2, 100, 17, False: px = px + 104
                Button ws, "JKB_PsEntity", "Entity scope", "CycleEntityScope", px, 2, 88, 17, False: px = px + 92
            Case "Pre_Shock_Metrics"
                Button ws, "JKB_PsBuilder", "Formula builder", "OpenPreShockStudio", px, 2, 100, 17, False
            Case "Pre_Shock_Cases"
                Button ws, "JKB_PsBuilder", "Filter builder", "OpenPreShockStudio", px, 2, 92, 17, False: px = px + 96
                Button ws, "JKB_PsTest", "Test filter", "TestSelectedFilter", px, 2, 76, 17, False: px = px + 80
                Button ws, "JKB_PsTrace", "Trace", "TraceSelectedValue", px, 2, 60, 17, False: px = px + 64
                Button ws, "JKB_PsRules", "Import filter conditions", "UploadScenarioRules", px, 2, 142, 17, False: px = px + 146
                Button ws, "JKB_PsRun", "Run", "RunStressReconciliation", px, 2, 74, 17, False
            Case "Pre_Shock_Results", "Derived_Values"
                Button ws, "JKB_PsTrace", "Trace this value", "TraceSelectedValue", px, 2, 108, 17, False: px = px + 112
                Button ws, "JKB_PsClearTr", "Clear traces", "ClearAllTraces", px, 2, 88, 17, False: px = px + 92
                Button ws, "JKB_PsRun", "Run", "RunStressReconciliation", px, 2, 74, 17, False
        End Select
    End If
    If isMaster And ws.name = "Config_ElementTypeRules" Then
        Button ws, "JKB_Editor", "Formula Studio", "OpenFormulaEditor", 427, 2, 129, 17, True
        Button ws, "JKB_GroupSetup", "Group setup", "GoGroups", 561, 2, 94, 17, True
        Button ws, "JKB_CfgMetrics", "Metric formulas", "GoPreShockFormulas", 659, 2, 104, 17, True
        Button ws, "JKB_CfgFields", "Field mapping", "GoPreShockDimensions", 767, 2, 94, 17, True
    End If

    EnsureViewStrip ws, isMaster
    ' Last: the strips above place their buttons at fixed x positions computed
    ' from captions that vary by sheet, which leaves ragged gaps. ProRibbon packs
    ' whatever actually ended up on this sheet, left to right, and sizes the band
    ' to fit it.
    ProRibbon ws, isMaster
End Sub

' The second strip: what a person does to a table once it is in front of them.
'
' It is on EVERY sheet of both the tool and every workbook it generates, and it
' is the same six controls in the same order everywhere, so it is learned once.
' Everything here is local to this module - a generated workbook carries a copy
' of it and nothing else, so a button that called a standard module would be a
' dead button the moment the file left this machine.
Private Sub EnsureViewStrip(ByVal ws As Worksheet, ByVal isMaster As Boolean)
    ' Retired in design v2. The zoom, fit and copy macros below stay public so
    ' buttons in review workbooks generated before v2 keep working.
End Sub

' ===================== reading a sheet ======================================
'
' Excel's own zoom is a slider in a corner and a modifier-plus-wheel nobody
' remembers, and "fit this to my screen" is a selection-then-menu dance. These
' are the same things where the eye already is.

Public Sub UiZoomIn()
    UiZoomBy 10
End Sub

Public Sub UiZoomOut()
    UiZoomBy -10
End Sub

Public Sub UiZoom100()
    On Error Resume Next
    ActiveWindow.zoom = 100
    Err.Clear
End Sub

Public Sub UiZoomFit()
    Dim ws As Worksheet, keep As Range, lastR As Long, lastC As Long
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Set keep = Selection
    lastC = ws.UsedRange.Column + ws.UsedRange.columns.count - 1
    If lastC > 60 Then lastC = 60
    lastR = ws.UsedRange.row + ws.UsedRange.rows.count - 1
    ' Width is what a person is trying to fit; forty rows is enough to find it
    ' without zooming out to read a twelve-thousand-row table at four per cent.
    If lastR > 48 Then lastR = 48
    ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC)).Select
    ActiveWindow.zoom = True
    If ActiveWindow.zoom > 130 Then ActiveWindow.zoom = 130
    If ActiveWindow.zoom < 40 Then ActiveWindow.zoom = 40
    If Not keep Is Nothing Then keep.Select
    Err.Clear
End Sub

Private Sub UiZoomBy(ByVal Delta As Long)
    Dim z As Long
    On Error Resume Next
    z = ActiveWindow.zoom + Delta
    If z < 30 Then z = 30
    If z > 250 Then z = 250
    ActiveWindow.zoom = z
    Err.Clear
End Sub

' Columns sized to what is in them, with a ceiling: one 900-character formula
' otherwise makes a column wider than the screen and every other one useless.
Public Sub UiFitColumns()
    Dim ws As Worksheet, c As Long, lastC As Long
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    lastC = ws.UsedRange.Column + ws.UsedRange.columns.count - 1
    If lastC > 60 Then lastC = 60
    For c = 1 To lastC
        If Not ws.columns(c).hidden Then
            ws.columns(c).AutoFit
            If ws.columns(c).ColumnWidth > 70 Then ws.columns(c).ColumnWidth = 70
            If ws.columns(c).ColumnWidth < 7 Then ws.columns(c).ColumnWidth = 7
        End If
    Next c
    Application.ScreenUpdating = True
    Err.Clear
End Sub

' Rows sized to what is in them - the other half of fitting a table, and the half
' that matters once a column has been narrowed and its text has wrapped.
'
' Rows 1 and 2 are left alone: row 1 is the toolbar band the two button strips
' are positioned against, and autofitting it would pull them out of alignment
' with the sheet underneath.
Public Sub UiFitRows()
    Dim ws As Worksheet, lastR As Long, r As Long, firstR As Long
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    firstR = 3
    lastR = ws.UsedRange.row + ws.UsedRange.rows.count - 1
    If lastR > 20000 Then lastR = 20000
    If lastR > firstR Then
        ws.Range(ws.rows(firstR), ws.rows(lastR)).AutoFit
        ' A wrapped formula can autofit to several inches; cap it so one long rule
        ' cannot push the next twenty rows off the screen.
        For r = firstR To lastR
            If ws.rows(r).RowHeight > 60 Then ws.rows(r).RowHeight = 60
        Next r
    End If
    ws.rows(1).RowHeight = 42
    ws.rows(2).RowHeight = 4
    Application.ScreenUpdating = True
    Err.Clear
End Sub

' What is on screen, on the clipboard - which is how a finding reaches an email.
Public Sub UiCopyView()
    Dim ws As Worksheet, rng As Range
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Set rng = ws.UsedRange
    If ws.FilterMode Then Set rng = rng.SpecialCells(xlCellTypeVisible)
    rng.Copy
    Err.Clear
End Sub

' Reads the pre-shock home sheet's own status line rather than guessing at cell
' addresses. Literals only: this module is copied into generated workbooks.
Private Function PreShockHomeStatus() As String
    Dim ws As Worksheet, s As String
    On Error Resume Next
    Set ws = Me.Parent.Worksheets("Pre_Shock")
    If Not ws Is Nothing Then s = CStr(ws.Range("A6").Value2)
    Err.Clear
    On Error GoTo 0
    If Len(Trim$(s)) = 0 Then s = "Ready - upload an extract when independent base and pre-shock derivation is required"
    PreShockHomeStatus = s
End Function

' The six sheets of the pre-shock workspace, named as literals because this module
' is copied into every generated workbook where the constants do not exist.
Private Function IsPreShockSheet(ByVal ws As Worksheet) As Boolean
    Select Case ws.name
        Case "Pre_Shock", "Pre_Shock_Sources", "Pre_Shock_Metrics", _
             "Pre_Shock_Fields", "Pre_Shock_Cases", "Pre_Shock_Results", _
             "Derived_Values", "Pre_Shock_Breakdown", "Config_ValueSources"
            IsPreShockSheet = True
    End Select
End Function

' `subtle` draws the second strip - the read-this-sheet controls - as quiet white
' chips rather than solid blue. Two strips of identical blue buttons would read as
' one wall of twenty things to do; the eye has to be able to tell the actions from
' the view controls without reading either.
Private Sub Button(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal proc As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal localProc As Boolean, Optional ByVal subtle As Boolean = False)
    Dim sh As Shape, target As String, fillColor As Long
    On Error Resume Next
    Set sh = ws.Shapes(nm)
    On Error GoTo 0
    If sh Is Nothing Then Set sh = ws.Shapes.AddShape(5, x, y, w, h): sh.name = nm
    sh.Left = x: sh.Top = y: sh.Width = w: sh.Height = h
    sh.Placement = xlFreeFloating: sh.LockAspectRatio = 0: sh.Visible = msoTrue
    ' Design system v2 (same tokens as modUI_Theme; literals because a generated
    ' workbook carries this module alone):
    '   primary     filled navy, white text - the one thing that starts the work
    '   secondary   white, grey hairline, navy text - every other action
    '   tab         the pre-shock navigation; the tab for this sheet is tinted
    Dim kind As String, isTab As Boolean
    isTab = (Left$(nm, 9) = "JKB_PsNav")
    kind = "secondary"
    If nm = "JKB_PsRun" Or nm = "JKB_ReconStart" Or nm = "JKB_ApplyFormat" Or nm = "JKB_OutApply" Then kind = "primary"
    If isTab Then kind = IIf(NavTabSheet(nm) = ws.name, "tabOn", "tab")
    sh.Shadow.Visible = msoFalse: sh.fill.Transparency = 0: sh.line.Weight = 0.75
    Select Case kind
        Case "primary"
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = RGB(25, 63, 137)
            sh.line.Visible = msoTrue: sh.line.ForeColor.RGB = RGB(25, 63, 137)
        Case "tab"
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = RGB(255, 255, 255)
            sh.line.Visible = msoFalse
        Case "tabOn"
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = RGB(238, 243, 251)
            sh.line.Visible = msoFalse
        Case Else
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = RGB(255, 255, 255)
            sh.line.Visible = msoTrue: sh.line.ForeColor.RGB = RGB(208, 213, 221)
    End Select
    On Error Resume Next: sh.Adjustments.item(1) = 0.16: On Error GoTo 0
    sh.TextFrame.Characters.text = caption
    With sh.TextFrame
        .HorizontalAlignment = xlHAlignCenter: .VerticalAlignment = xlVAlignCenter
        .MarginTop = 1: .MarginBottom = 1: .MarginLeft = 6: .MarginRight = 6
        .Characters.Font.name = "Aptos": .Characters.Font.Size = IIf(h < 24, 9.5, 11)
        .Characters.Font.Bold = True
        Select Case kind
            Case "primary": .Characters.Font.Color = RGB(255, 255, 255)
            Case "tab": .Characters.Font.Color = RGB(102, 112, 133)
            Case Else: .Characters.Font.Color = RGB(25, 63, 137)
        End Select
    End With
    ' Bind to the workbook that owns the sheet. When the master formats a sheet inside a
    ' generated workbook, the button must still call that workbook's own copy of the code.
    target = "'" & Replace(ws.Parent.name, "'", "''") & "'!"
    If localProc Then
        Dim hostDash As Worksheet
        On Error Resume Next
        Set hostDash = ws.Parent.Worksheets("Dashboard")
        On Error GoTo 0
        If hostDash Is Nothing Then target = target & Me.CodeName & "." Else target = target & hostDash.CodeName & "."
    End If
    sh.OnAction = target & proc
End Sub

Private Function FormulaCaption(ByVal ws As Worksheet, ByRef count As Long, ByRef allHidden As Boolean) As String
    Dim a As Variant, r As Long, lastRow As Long
    allHidden = True: count = 0
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= 7 Then
        a = ws.Range("L7:L" & lastRow).Value2
        If lastRow = 7 Then
            If CStr(a) = "TOGGLE_ROW" Then count = 1: allHidden = ws.rows(7).hidden
        Else
            For r = 1 To UBound(a, 1)
                If Not IsError(a(r, 1)) Then
                    If CStr(a(r, 1)) = "TOGGLE_ROW" Then
                        count = count + 1
                        If Not ws.rows(r + 6).hidden Then allHidden = False
                    End If
                End If
            Next r
        End If
    End If
    If count = 0 Then
        FormulaCaption = "No formula rows"
    ElseIf allHidden Then
        FormulaCaption = "Show formulas"
    Else
        FormulaCaption = "Hide formulas"
    End If
End Function

Public Sub ToggleFormulaRows()
    Dim ws As Worksheet, n As Long, hidden As Boolean, caption As String, r As Long, lastRow As Long, a As Variant, area As String
    Dim oldScreen As Boolean, er As String
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set ws = ActiveSheet
    If Not ws.Parent Is Me.Parent Then Exit Sub
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    caption = FormulaCaption(ws, n, hidden)
    If n = 0 Then Exit Sub
    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 7 To lastRow
        If CStr(ws.Cells(r, 12).Value2) = "TOGGLE_ROW" Then
            If Len(area) > 0 Then area = area & ","
            area = area & CStr(r) & ":" & CStr(r)
            If Len(area) > 180 Then ws.Range(area).EntireRow.hidden = Not hidden: area = ""
        End If
    Next r
    If Len(area) > 0 Then ws.Range(area).EntireRow.hidden = Not hidden
    EnsureRibbon ws
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    er = Err.description: Application.ScreenUpdating = oldScreen
    MsgBox "Could not change formula-row visibility." & vbCrLf & vbCrLf & "Reason: " & er, vbExclamation, "JKB Stress Testing"
End Sub

' ===================== generated-workbook controls =========================
'
' Everything below runs inside a GENERATED workbook, which only receives this sheet's
' code module - never the standard modules. It therefore uses literals and Excel
' built-ins only. Keep it that way.

' Freezes the header rows and the label columns of an element sheet, reliably.
'
' SplitRow / SplitColumn are measured from the window's CURRENT top-left visible cell, not
' from A1. If the window happens to be scrolled when the split is set - which it is during
' generation, because sheets are activated in sequence while columns and rows are being
' hidden - the freeze lands on the wrong rows and the sheet then scrolls oddly. Scrolling
' home first, and setting zoom before the split rather than after, makes it deterministic.
Public Sub EnsureFrozenLayout(ByVal ws As Worksheet, Optional ByVal zoom As Long = 0, _
                              Optional ByVal splitRow As Long = 6, Optional ByVal splitCol As Long = 2)
    Dim wnd As Window, oldScreen As Boolean
    On Error Resume Next
    If ws.Parent.Windows.count = 0 Then Exit Sub
    ' FreezePanes is a window-manager operation: with ScreenUpdating off it silently leaves
    ' the pane SPLIT but not FROZEN, which is what made these sheets scroll around. Generation
    ' runs with screen updating off, so it has to be turned back on for this one step.
    oldScreen = Application.ScreenUpdating
    Application.ScreenUpdating = True
    ' Pane operations apply to the ACTIVE window. During generation the master workbook is
    ' still active, so activating only the sheet leaves the split set but never frozen -
    ' which is precisely a pane that scrolls instead of staying put. Activate the workbook.
    ws.Parent.Activate
    ws.Activate
    Set wnd = ws.Parent.Windows(1)
    If Not wnd Is Nothing Then
        With wnd
            .FreezePanes = False
            .Split = False
            If zoom >= 50 And zoom <= 200 Then .zoom = zoom
            ' The split is measured from the top-left VISIBLE cell, so scroll home first.
            .ScrollRow = 1
            .ScrollColumn = 1
            ws.Range("A1").Select
            .splitRow = splitRow
            .SplitColumn = splitCol
            Err.Clear
            .FreezePanes = True
            ' A split that never became a freeze is worse than no split at all: the panes
            ' scroll independently and the sheet feels broken. If freezing is refused -
            ' Excel rejects it outright when the application is not visible - drop the
            ' split rather than leave the sheet in that state.
            If Err.Number <> 0 Or Not .FreezePanes Then
                Err.Clear
                .Split = False
                .FreezePanes = False
            End If
            .ScrollRow = 1
            .ScrollColumn = 1
        End With
    End If
    Application.ScreenUpdating = oldScreen
    Err.Clear
    On Error GoTo 0
End Sub

Private Function IsElementSheet(ByVal ws As Worksheet) As Boolean
    IsElementSheet = (Left$(CStr(ws.Range("A4").Value2), 14) = "Element type: ")
End Function

' Remembers the element sheet the reviewer last looked at, so the Back button returns
' there rather than guessing.
Public Sub RememberElementSheet(ByVal ws As Worksheet)
    If IsElementSheet(ws) Then mLastElement = ws.name
End Sub

Public Sub GoElementSheet()
    Dim ws As Worksheet, target As Worksheet
    On Error GoTo Failed
    If Len(mLastElement) > 0 Then
        On Error Resume Next
        Set target = Me.Parent.Worksheets(mLastElement)
        On Error GoTo Failed
    End If
    If target Is Nothing Then
        For Each ws In Me.Parent.Worksheets
            If ws.Visible = xlSheetVisible Then
                If IsElementSheet(ws) Then Set target = ws: Exit For
            End If
        Next ws
    End If
    If target Is Nothing Then
        MsgBox "This workbook has no element sheet to return to.", vbInformation, "JKB Stress Testing"
        Exit Sub
    End If
    Me.Parent.Activate
    target.Activate
    Application.Goto target.Range("A1"), True
    EnsureFrozenLayout target
    Exit Sub
Failed:
    MsgBox "Could not open the element sheet." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub

' Each severity owns one column in each of the three sections:
'   Moderate  C (system)  F (manual)  I (difference)
'   Medium    D           G           J
'   Severe    E           H           K
Public Sub ToggleModerate()
    ToggleSeverityColumns 0, "Moderate"
End Sub
Public Sub ToggleMedium()
    ToggleSeverityColumns 1, "Medium"
End Sub
Public Sub ToggleSevere()
    ToggleSeverityColumns 2, "Severe"
End Sub

Private Sub ToggleSeverityColumns(ByVal offset As Long, ByVal caption As String)
    Dim ws As Worksheet, hideNow As Boolean, er As String
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set ws = ActiveSheet
    If Left$(CStr(ws.Range("A4").Value2), 14) <> "Element type: " Then
        MsgBox "Severity columns can only be shown or hidden on a generated element sheet.", vbInformation, "JKB Stress Testing"
        Exit Sub
    End If
    On Error GoTo Failed
    RememberElementSheet ws
    hideNow = Not ws.columns(3 + offset).hidden
    ws.columns(3 + offset).hidden = hideNow
    ws.columns(6 + offset).hidden = hideNow
    ws.columns(9 + offset).hidden = hideNow
    ' Hiding columns can disturb a frozen split, so re-assert it.
    EnsureFrozenLayout ws
    EnsureRibbon ws
    Exit Sub
Failed:
    er = Err.description
    MsgBox "Could not change the " & caption & " columns." & vbCrLf & vbCrLf & "Reason: " & er, vbExclamation, "JKB Stress Testing"
End Sub

Private Function SeverityHidden(ByVal ws As Worksheet, ByVal offset As Long) As Boolean
    On Error Resume Next
    SeverityHidden = ws.columns(3 + offset).hidden
    On Error GoTo 0
End Function

' A compact appearance sheet that travels with the generated workbook so the reviewer can
' restyle it after the fact. It deliberately carries only view settings - the master's rule
' table is not a deliverable and is not copied.
Public Sub EnsureOutputFormatSheet(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(OUT_FORMAT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.name = OUT_FORMAT_SHEET
    End If
    ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB"
    ws.Range("A3").Value2 = "Appearance"
    ws.Range("A4").Value2 = "Change any value below, then click Apply formatting. Only the look of the element sheets changes - values and formulas are never touched."
    ws.Range("A6:B6").value = Array("Setting", "Value")
    ws.Range("A7:B18").value = Application.Transpose(Array( _
        Array("Label column width", "Value column width", "Minimum row height", "Maximum row height", _
              "Zoom (%)", "Show gridlines", "Wrap row labels", "Auto-fit row heights", _
              "Difference highlight tolerance", "Freeze panes", "Number format", "Difference number format"), _
        Array(40, 18, 18, 120, 90, "No", "Yes", "Yes", 0, "Yes", "#,##0.00", "#,##0.00")))
    ws.Range("A3").Font.Size = 15: ws.Range("A3").Font.Bold = True
    ws.Range("A3").Font.Color = RGB(18, 48, 107)
    ws.Range("A4").Font.Size = 9: ws.Range("A4").Font.Color = RGB(102, 112, 133): ws.Range("A4").WrapText = True
    ws.Range("A3:F3").Interior.Color = RGB(238, 243, 251)
    With ws.Range("A6:B6")
        .Interior.Color = RGB(59, 99, 176): .Font.Color = vbWhite: .Font.Bold = True
    End With
    ws.Range("A7:A18").Font.Bold = True
    ws.Range("B7:B18").Interior.Color = RGB(253, 246, 227)
    ws.Range("A6:B18").Borders.Color = RGB(208, 213, 221)
    ws.columns("A").ColumnWidth = 34: ws.columns("B").ColumnWidth = 22
    ws.rows(3).RowHeight = 28: ws.rows(4).RowHeight = 30
    EnsureRibbon ws
End Sub

Public Function OutputSetting(ByVal key As String, ByVal fallback As Variant) As Variant
    Dim ws As Worksheet, a As Variant, r As Long
    OutputSetting = fallback
    On Error Resume Next
    Set ws = Me.Parent.Worksheets(OUT_FORMAT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    a = ws.Range("A7:B30").Value2
    For r = 1 To UBound(a, 1)
        If UCase$(Trim$(CStr(a(r, 1)))) = UCase$(key) Then
            If Len(CStr(a(r, 2))) > 0 Then OutputSetting = a(r, 2): Exit Function
        End If
    Next r
End Function

' Re-applies the appearance settings to every element sheet in this workbook.
Public Sub ApplyOutputFormatting()
    Dim ws As Worksheet, lastRow As Long, r As Long, h As Double, n As Long
    Dim minH As Double, maxH As Double, gridlines As Boolean, zoom As Long, er As String
    Dim oldScreen As Boolean, oldSheet As Object
    oldScreen = Application.ScreenUpdating
    Set oldSheet = ActiveSheet
    On Error GoTo Failed
    Application.ScreenUpdating = False
    minH = CDbl(OutputSetting("Minimum row height", 18))
    maxH = CDbl(OutputSetting("Maximum row height", 120))
    If minH > maxH Then Err.Raise vbObjectError + 581, , "Minimum row height must not exceed maximum row height."
    gridlines = (UCase$(CStr(OutputSetting("Show gridlines", "No"))) = "YES")
    zoom = CLng(OutputSetting("Zoom (%)", 90))
    If zoom < 50 Or zoom > 200 Then zoom = 90

    For Each ws In Me.Parent.Worksheets
        If Left$(CStr(ws.Range("A4").Value2), 14) = "Element type: " Then
            n = n + 1
            lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
            If lastRow >= 7 Then
                ws.columns("A").ColumnWidth = CDbl(OutputSetting("Label column width", 40))
                ws.columns("B:K").ColumnWidth = CDbl(OutputSetting("Value column width", 18))
                ws.Range("B7:H" & lastRow).NumberFormat = CStr(OutputSetting("Number format", "#,##0.00"))
                ws.Range("I7:K" & lastRow).NumberFormat = CStr(OutputSetting("Difference number format", "#,##0.00"))
                ws.Range("A7:A" & lastRow).WrapText = (UCase$(CStr(OutputSetting("Wrap row labels", "Yes"))) = "YES")
                If UCase$(CStr(OutputSetting("Auto-fit row heights", "Yes"))) = "YES" Then ws.rows("7:" & lastRow).AutoFit
                For r = 7 To lastRow
                    If Not ws.rows(r).hidden Then
                        h = ws.rows(r).RowHeight
                        If h < minH Then h = minH
                        If h > maxH Then h = maxH
                        ws.rows(r).RowHeight = h
                    End If
                Next r
                With ws.Range("I7:K" & lastRow)
                    .FormatConditions.Delete
                    .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(ISNUMBER(I7),ABS(I7)>" & Trim$(Str$(CDbl(OutputSetting("Difference highlight tolerance", 0)))) & ")"
                    .FormatConditions(1).Interior.Color = RGB(254, 243, 242)
                    .FormatConditions(1).Font.Color = RGB(180, 35, 24)
                    .FormatConditions(1).Font.Bold = True
                End With
            End If
            On Error Resume Next
            ws.Activate
            ActiveWindow.DisplayGridlines = gridlines
            ActiveWindow.zoom = zoom
            If UCase$(CStr(OutputSetting("Freeze panes", "Yes"))) = "YES" Then
                ActiveWindow.FreezePanes = False
                ActiveWindow.SplitColumn = 2: ActiveWindow.splitRow = 6
                ActiveWindow.FreezePanes = True
            Else
                ActiveWindow.FreezePanes = False
            End If
            Err.Clear
            On Error GoTo Failed
        End If
    Next ws
    On Error Resume Next
    oldSheet.Activate
    On Error GoTo 0
    Application.ScreenUpdating = oldScreen
    MsgBox "Formatting applied to " & n & " element sheet(s)." & vbCrLf & vbCrLf & "Values and formulas are unchanged.", vbInformation, "JKB Stress Testing"
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    oldSheet.Activate
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    MsgBox "Formatting stopped." & vbCrLf & vbCrLf & "Reason: " & er, vbExclamation, "JKB Stress Testing"
End Sub

Public Sub GoOutputFormat()
    On Error GoTo Failed
    ' Remember where we came from so Back returns to the same element sheet.
    If TypeOf ActiveSheet Is Worksheet Then RememberElementSheet ActiveSheet
    Me.Parent.Activate
    Me.Parent.Worksheets(OUT_FORMAT_SHEET).Activate
    Application.Goto Me.Parent.Worksheets(OUT_FORMAT_SHEET).Range("A1"), True
    Exit Sub
Failed:
    MsgBox "Could not open the appearance sheet." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub

' Home is the reconciliation workbench in the tool, and the Dashboard in a
' generated review workbook (which has no workbench).
Public Sub GoHome()
    Dim marker As String
    If TypeOf ActiveSheet Is Worksheet Then RememberElementSheet ActiveSheet
    On Error Resume Next
    marker = UCase$(Trim$(CStr(Me.Parent.Worksheets("Dashboard").Range("N1").Value2)))
    On Error GoTo Failed
    If marker = "OUTPUT" Then
        Navigate "Dashboard"
    Else
        Application.Run "'" & Replace(Me.Parent.name, "'", "''") & "'!OpenReconWorkbench"
    End If
    Exit Sub
Failed:
    Navigate "Dashboard"
End Sub

' Which sheet a pre-shock navigation tab opens, so the tab for the sheet in
' front of you can be shown as the current one.
Private Function NavTabSheet(ByVal nm As String) As String
    Select Case nm
        Case "JKB_PsNavVS": NavTabSheet = "Config_ValueSources"
        Case "JKB_PsNavSrc": NavTabSheet = "Pre_Shock_Sources"
        Case "JKB_PsNavCases": NavTabSheet = "Pre_Shock_Cases"
        Case "JKB_PsNavDerived": NavTabSheet = "Derived_Values"
    End Select
End Function

Public Sub GoDashboard()
    If TypeOf ActiveSheet Is Worksheet Then RememberElementSheet ActiveSheet
    Navigate "Dashboard"
End Sub
Public Sub GoConfig()
    Navigate "Config_ElementTypeRules"
End Sub
Public Sub GoFormat()
    Navigate "Config_MasterFormatting"
End Sub
Public Sub GoGroups()
    On Error GoTo Failed
    Me.Parent.Activate
    Me.Parent.Worksheets("Config_ElementTypeRules").Activate
    Application.Goto Me.Parent.Worksheets("Config_ElementTypeRules").Range("Y3"), True
    Exit Sub
Failed:
    MsgBox "Could not open detail grouping setup." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub
Private Sub Navigate(ByVal nm As String)
    On Error GoTo Failed
    Me.Parent.Activate
    If Me.Parent.Worksheets(nm).Visible <> xlSheetVisible Then Me.Parent.Worksheets(nm).Visible = xlSheetVisible
    Me.Parent.Worksheets(nm).Activate
    Application.Goto Me.Parent.Worksheets(nm).Range("A1"), True
    Exit Sub
Failed:
    MsgBox "Could not open " & nm & "." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub

Public Function ViewSetting(ByVal key As String, ByVal fallback As Variant) As Variant
    Dim a As Variant, r As Long
    a = Me.Parent.Worksheets("Config_MasterFormatting").Range("A6:B24").Value2
    For r = 1 To UBound(a, 1)
        If UCase$(CStr(a(r, 1))) = UCase$(key) Then
            If Len(CStr(a(r, 2))) > 0 Then ViewSetting = a(r, 2): Exit Function
        End If
    Next r
    ViewSetting = fallback
End Function

Public Sub PrepareStyles()
    Dim f As Worksheet, r As Long, key As String, st As Style, sample As Range, edge As Variant
    Set f = Me.Parent.Worksheets("Config_MasterFormatting")
    mRules = Empty: Set mRuleColumns = Nothing: Set mNumberFormats = Nothing
    For r = 29 To 60
        key = CStr(f.Cells(r, 26).Value2)
        If Len(key) > 0 Then
            Set sample = f.Cells(r, 3)
            Set st = Nothing
            On Error Resume Next
            Set st = Me.Parent.styles("JKB_" & key)
            On Error GoTo 0
            If st Is Nothing Then Set st = Me.Parent.styles.Add(name:="JKB_" & key)
            st.IncludeNumber = False: st.IncludeProtection = False
            st.IncludeFont = True: st.IncludeAlignment = True: st.IncludeBorder = True: st.IncludePatterns = True
            With st.Font
                .name = sample.Font.name: .Size = sample.Font.Size: .Bold = sample.Font.Bold
                .Italic = sample.Font.Italic: .Underline = sample.Font.Underline: .Color = sample.Font.Color
                .Strikethrough = sample.Font.Strikethrough: .Superscript = sample.Font.Superscript: .Subscript = sample.Font.Subscript
            End With
            st.Interior.pattern = sample.Interior.pattern: st.Interior.Color = sample.Interior.Color
            st.Interior.PatternColor = sample.Interior.PatternColor
            st.HorizontalAlignment = sample.HorizontalAlignment: st.VerticalAlignment = sample.VerticalAlignment
            st.WrapText = sample.WrapText: st.orientation = sample.orientation
            st.IndentLevel = sample.IndentLevel: st.ShrinkToFit = sample.ShrinkToFit
            For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlDiagonalUp, xlDiagonalDown)
                With st.Borders(edge)
                    .LineStyle = sample.Borders(edge).LineStyle
                    If .LineStyle <> xlLineStyleNone Then .Color = sample.Borders(edge).Color: .Weight = sample.Borders(edge).Weight
                End With
            Next edge
        End If
    Next r
End Sub

Public Sub ApplyFormatting()
    Dim ws As Worksheet, oldSheet As Worksheet, oldScreen As Boolean, er As String
    oldScreen = Application.ScreenUpdating
    Set oldSheet = ActiveSheet
    On Error GoTo Failed
    ValidateViewSettings
    Application.ScreenUpdating = False
    PrepareStyles
    For Each ws In Me.Parent.Worksheets
        If Left$(CStr(ws.Range("A4").Value2), 14) = "Element type: " Then
            FormatReport ws
            ws.Activate
            ActiveWindow.DisplayGridlines = (UCase$(CStr(ViewSetting("Show gridlines", "No"))) = "YES")
            ActiveWindow.zoom = CLng(ViewSetting("Zoom (%)", 90))
        End If
        If ws.Visible = xlSheetVisible Then EnsureRibbon ws
    Next ws
    oldSheet.Activate
    Application.ScreenUpdating = oldScreen
    MsgBox "Formatting applied. Values and formulas are unchanged.", vbInformation, "JKB Stress Testing"
    Exit Sub
Failed:
    er = Err.description: Application.ScreenUpdating = oldScreen
    MsgBox "Formatting stopped." & vbCrLf & vbCrLf & "Reason: " & er, vbExclamation, "JKB Stress Testing"
End Sub

Public Sub ValidateViewSettings()
    CheckNumber "Label column width", 12, 100
    CheckNumber "Value column width", 8, 60
    CheckNumber "Minimum row height", 12, 100
    CheckNumber "Maximum row height", 20, 409
    CheckNumber "Zoom (%)", 50, 200
    CheckNumber "Print pages wide", 1, 10
    CheckNumber "Difference highlight tolerance", 0, 1000000000
    If CDbl(ViewSetting("Minimum row height", 18)) > CDbl(ViewSetting("Maximum row height", 120)) Then Err.Raise vbObjectError + 571, , "Minimum row height must not exceed maximum row height."
End Sub
Private Sub CheckNumber(ByVal key As String, ByVal low As Double, ByVal high As Double)
    Dim v As Variant: v = ViewSetting(key, low)
    If Not IsNumeric(v) Then Err.Raise vbObjectError + 570, , "Enter a number for " & key & "."
    If CDbl(v) < low Or CDbl(v) > high Then Err.Raise vbObjectError + 570, , key & " must be between " & low & " and " & high & "."
End Sub

Public Sub FormatReport(ByVal ws As Worksheet)
    Dim lo As ListObject, a As Variant, numbers As Object, styles As Object, idx As Object
    Dim f As Worksheet, r As Long, c As Long, k As Variant, label As String, t As String, p As Long, lastRow As Long, rowIx As Long
    Dim sectionCols As Variant, styleCols As Variant, fallback As Variant, sk As String, address As String, nr As String, h As Double, hdr As Variant
    Set f = Me.Parent.Worksheets("Config_MasterFormatting")
    Set lo = Me.Parent.Worksheets("Config_ElementTypeRules").ListObjects("tblElementTypeRules")
    Set idx = CreateObject("Scripting.Dictionary"): idx.CompareMode = vbTextCompare
    Set styles = CreateObject("Scripting.Dictionary"): styles.CompareMode = vbTextCompare
    Set numbers = CreateObject("Scripting.Dictionary"): numbers.CompareMode = vbTextCompare
    t = Mid$(CStr(ws.Range("A4").Value2), 15): p = InStr(1, t, " | ", vbBinaryCompare)
    If p > 0 Then t = Left$(t, p - 1)
    If IsEmpty(mRules) Then
        mRules = lo.DataBodyRange.Value2
        hdr = lo.HeaderRowRange.Value2
        Set mRuleColumns = CreateObject("Scripting.Dictionary"): mRuleColumns.CompareMode = vbTextCompare
        For c = 1 To UBound(hdr, 2): mRuleColumns(CStr(hdr(1, c))) = c: Next c
        Set mNumberFormats = CreateObject("Scripting.Dictionary"): mNumberFormats.CompareMode = vbTextCompare
        For r = 7 To 16: mNumberFormats(CStr(f.Cells(r, 27).Value2)) = CStr(f.Cells(r, 7).NumberFormat): Next r
    End If
    a = mRules
    For r = 1 To UBound(a, 1)
        label = Trim$(CStr(a(r, mRuleColumns("ELEMENT_TYPE"))))
        If StrComp(label, Trim$(t), vbTextCompare) = 0 Then
            label = UCase$(Trim$(CStr(a(r, mRuleColumns("OUTPUT_ROW_LABEL")))))
            idx(label) = r
        End If
    Next r
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 7 Then Exit Sub
    sectionCols = Array("A", "B", "C", "F", "I")
    styleCols = Array(mRuleColumns("LABEL_STYLE_KEY"), mRuleColumns("BASE_STYLE_KEY"), mRuleColumns("SYSTEM_STYLE_KEY"), mRuleColumns("MANUAL_STYLE_KEY"), mRuleColumns("DIFFERENCE_STYLE_KEY"))
    fallback = Array("ROW_LABEL", "BASE_NUMERIC", "SYSTEM_NUMERIC", "MANUAL_NUMERIC", "DIFFERENCE_NUMERIC")
    For r = 7 To lastRow
        If CStr(ws.Cells(r, 12).Value2) <> "SECTION_HEADER" Then
            label = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))
            If idx.Exists(label) Then
                rowIx = idx(label)
                For c = 0 To 4
                    sk = Trim$(CStr(a(rowIx, styleCols(c))))
                    If Len(sk) = 0 Then sk = fallback(c)
                    address = sectionCols(c) & r
                    If c >= 2 Then address = address & ":" & Chr$(Asc(sectionCols(c)) + 2) & r
                    AddArea styles, UCase$(sk), address, ws, True
                Next c
                nr = NumberCode(CStr(a(rowIx, mRuleColumns("VALUE_FORMAT_KEY"))))
                If Len(nr) > 0 Then AddArea numbers, nr, "B" & r & ":H" & r, ws, False
                nr = NumberCode(CStr(a(rowIx, mRuleColumns("DIFFERENCE_FORMAT_KEY"))))
                If Len(nr) = 0 Then nr = NumberCode(CStr(a(rowIx, mRuleColumns("VALUE_FORMAT_KEY"))))
                If Len(nr) > 0 Then AddArea numbers, nr, "I" & r & ":K" & r, ws, False
            End If
        End If
    Next r
    For Each k In styles.keys
        If Len(styles(k)) > 0 Then ws.Range(styles(k)).Style = "JKB_" & CStr(k)
    Next k
    For Each k In numbers.keys
        If Len(numbers(k)) > 0 Then ws.Range(numbers(k)).NumberFormat = CStr(k)
    Next k
    ws.Range("A3:K3").Style = "JKB_TITLE"
    ws.Range("A4:K4").Style = "JKB_INFO_BAND"
    ws.Range("A5:K5").Style = "JKB_GROUP_HEADER"
    ws.Range("A6:K6").Style = "JKB_SUBHEADER"

    ' Reconciliation palette. The three column families carry the meaning, so they are what
    ' gets the colour; everything else stays quiet so the numbers lead.
    '   A-B  identity and base      neutral slate
    '   C-E  system output          white
    '   F-H  independent / manual   cool green
    '   I-K  difference             warm amber
    ws.Cells.Font.name = "Aptos"
    With ws.Range("A7:B" & lastRow)
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(16, 24, 40): .Font.Bold = True
    End With
    ws.Range("A7:A" & lastRow).IndentLevel = 1
    With ws.Range("C7:E" & lastRow)
        .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(16, 24, 40): .Font.Bold = False
    End With
    With ws.Range("F7:H" & lastRow)
        .Interior.Color = RGB(236, 253, 243): .Font.Color = RGB(16, 24, 40): .Font.Bold = False
    End With
    With ws.Range("I7:K" & lastRow)
        .Interior.Color = RGB(255, 250, 235): .Font.Color = RGB(16, 24, 40): .Font.Bold = False
    End With
    ' Hairline separators where one column family ends and the next begins. Literal letters
    ' rather than a helper: this module is copied into generated workbooks, which do not
    ' have the standard modules.
    For Each k In Array("C", "F", "I")
        With ws.Range(CStr(k) & "5:" & CStr(k) & lastRow).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Color = RGB(152, 162, 179): .Weight = xlThin
        End With
    Next k
    ws.Range("A7:K" & lastRow).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    ws.Range("A7:K" & lastRow).Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)

    ' Section band (row 5) and severity band (row 6).
    With ws.Range("A5:B6")
        .Interior.Color = RGB(234, 236, 240): .Font.Color = RGB(16, 24, 40): .Font.Bold = True
    End With
    With ws.Range("C5:E6")
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(16, 24, 40): .Font.Bold = True
    End With
    With ws.Range("F5:H6")
        .Interior.Color = RGB(236, 253, 243): .Font.Color = RGB(6, 118, 71): .Font.Bold = True
    End With
    With ws.Range("I5:K6")
        .Interior.Color = RGB(253, 246, 227): .Font.Color = RGB(176, 138, 46): .Font.Bold = True
    End With
    With ws.Range("A5:K6")
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(152, 162, 179): .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ws.Range("A5:B6").HorizontalAlignment = xlLeft
    ws.Range("A5:B6").IndentLevel = 1

    ' Title block: a deep band with a gold rule underneath, matching the tool's identity.
    With ws.Range("A3:K3")
        .Interior.Color = RGB(14, 34, 64): .Font.Color = vbWhite
        .Font.Bold = True: .Font.Size = 14
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46): .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    With ws.Range("A4:K4")
        .Interior.Color = RGB(245, 247, 250): .Font.Color = RGB(71, 84, 103): .Font.Size = 9
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    ws.columns("A").ColumnWidth = CDbl(ViewSetting("Label column width", 40))
    ws.columns("B:K").ColumnWidth = CDbl(ViewSetting("Value column width", 18))
    ws.columns("L:Z").hidden = True
    ws.rows(3).RowHeight = 32: ws.rows(4).RowHeight = 27: ws.rows("5:6").RowHeight = 21
    ws.Range("A7:A" & lastRow).WrapText = (UCase$(CStr(ViewSetting("Wrap row labels", "Yes"))) = "YES")
    If UCase$(CStr(ViewSetting("Auto-fit row heights", "Yes"))) = "YES" Then ws.rows("7:" & lastRow).AutoFit
    For r = 7 To lastRow
        If CStr(ws.Cells(r, 12).Value2) = "SECTION_HEADER" Then
            ws.rows(r).RowHeight = CDbl(ViewSetting("Section header height", 21))
        ElseIf Not ws.rows(r).hidden Then
            h = ws.rows(r).RowHeight
            If h < CDbl(ViewSetting("Minimum row height", 18)) Then h = CDbl(ViewSetting("Minimum row height", 18))
            If h > CDbl(ViewSetting("Maximum row height", 120)) Then h = CDbl(ViewSetting("Maximum row height", 120))
            ws.rows(r).RowHeight = h
        End If
    Next r
    With ws.Range("I7:K" & lastRow)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=AND(ISNUMBER(I7),ABS(I7)>" & Trim$(Str$(CDbl(ViewSetting("Difference highlight tolerance", 0)))) & ")"
        .FormatConditions(1).Interior.Color = RGB(254, 243, 242)
        .FormatConditions(1).Font.Color = RGB(180, 35, 24)
        .FormatConditions(1).Font.Bold = True
    End With
    FormatDetailGroups ws
    On Error GoTo PrintUnavailable
    With ws.PageSetup
        .PrintArea = ws.Range("A3:K" & lastRow).address
        .PrintTitleRows = "$5:$6": .zoom = False
        .FitToPagesWide = CLng(ViewSetting("Print pages wide", 1)): .FitToPagesTall = False
        .orientation = IIf(UCase$(CStr(ViewSetting("Page orientation", "Landscape"))) = "PORTRAIT", xlPortrait, xlLandscape)
        Select Case UCase$(CStr(ViewSetting("Paper size", "A4")))
            Case "A3": .PaperSize = xlPaperA3
            Case "LETTER": .PaperSize = xlPaperLetter
            Case Else: .PaperSize = xlPaperA4
        End Select
    End With
PrintUnavailable:
    EnsureRibbon ws
End Sub

Private Sub AddArea(ByVal map As Object, ByVal key As String, ByVal address As String, ByVal ws As Worksheet, ByVal isStyle As Boolean)
    If Not map.Exists(key) Then map.Add key, ""
    If Len(map(key)) > 160 Then
        If isStyle Then ws.Range(map(key)).Style = "JKB_" & key Else ws.Range(map(key)).NumberFormat = key
        map(key) = ""
    End If
    If Len(map(key)) > 0 Then map(key) = map(key) & ","
    map(key) = map(key) & address
End Sub
Private Function NumberCode(ByVal key As String) As String
    If mNumberFormats.Exists(key) Then NumberCode = mNumberFormats(key)
End Function

Public Function IsFormulaCell(ByVal cell As Range) As Boolean
    Dim lo As ListObject, col As Long
    If cell.CountLarge <> 1 Then Exit Function
    If Not cell.Parent.Parent Is Me.Parent Then Exit Function
    If cell.Parent.name <> "Config_ElementTypeRules" Then Exit Function
    Set lo = cell.Parent.ListObjects("tblElementTypeRules")
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.Intersect(cell, lo.DataBodyRange) Is Nothing Then Exit Function
    col = cell.Column - lo.Range.Column + 1
    IsFormulaCell = (Left$(CStr(lo.HeaderRowRange.Cells(1, col).Value2), 15) = "MANUAL_FORMULA_")
End Function

Public Sub OpenFormulaEditor()
    Dim cell As Range, lo As ListObject, a As Variant, hdr As Variant, cols As Object
    Dim r As Long, c As Long, idx As Long, elementType As String, label As String, fields As String, formula As String, ctx As String
    Dim fso As Object, stream As Object, path As String, page As String, er As String, headerName As String, modeName As String
    Dim groupKey As String, sourceField As String, fmtKey As String, notesText As String, rowOrder As String, fieldObj As String
    Dim currentGroup As String, variants As String, groupMap As Object
    On Error GoTo Failed
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set cell = ActiveCell
    If Not IsFormulaCell(cell) Then
        MsgBox "Select a formula cell in the Config table first (default or severity-specific)." & vbCrLf & vbCrLf & "Next: click Formula Studio, or double-click the cell.", vbInformation, "JKB Stress Testing"
        Exit Sub
    End If
    Set lo = cell.Parent.ListObjects("tblElementTypeRules")
    a = lo.DataBodyRange.Value2: hdr = lo.HeaderRowRange.Value2
    Set cols = CreateObject("Scripting.Dictionary"): cols.CompareMode = vbTextCompare
    For c = 1 To UBound(hdr, 2): cols(CStr(hdr(1, c))) = c: Next c
    idx = cell.row - lo.DataBodyRange.row + 1
    elementType = Trim$(CStr(a(idx, cols("ELEMENT_TYPE"))))
    ' Groups come from the grouping rules, not only column U, so ask the engine. Late-bound
    ' through Application.Run because this sheet's code also travels into generated books.
    On Error Resume Next
    Set groupMap = Application.Run("'" & Replace(Me.Parent.Name, "'", "''") & "'!GroupMapForElement", elementType)
    Err.Clear
    On Error GoTo Failed

    ' Formula Studio deliberately receives every configured row for this element type,
    ' not only enabled rows. Disabled rows are clearly flagged in the editor so the user
    ' can see what exists without the UI silently hiding a possible reference.
    For r = 1 To UBound(a, 1)
        If StrComp(Trim$(CStr(a(r, cols("ELEMENT_TYPE")))), elementType, vbTextCompare) = 0 Then
            label = Trim$(CStr(a(r, cols("OUTPUT_ROW_LABEL"))))
            If Len(label) > 0 Then
                groupKey = "": If cols.Exists("DETAIL_GROUP_KEY") Then groupKey = Trim$(CStr(a(r, cols("DETAIL_GROUP_KEY"))))
                If Not groupMap Is Nothing Then If groupMap.Exists(label) Then groupKey = CStr(groupMap(label))
                sourceField = Trim$(CStr(a(r, cols("SOURCE_FIELD_NAME"))))
                fmtKey = Trim$(CStr(a(r, cols("VALUE_FORMAT_KEY"))))
                notesText = Trim$(CStr(a(r, cols("NOTES"))))
                rowOrder = "null": If IsNumeric(a(r, cols("OUTPUT_ORDER"))) Then rowOrder = CStr(CLng(a(r, cols("OUTPUT_ORDER"))))
                fieldObj = "{""name"":" & EditorJson(label)
                fieldObj = fieldObj & ",""enabled"":" & EditorBoolJson(a(r, cols("IS_ENABLED")))
                fieldObj = fieldObj & ",""order"":" & rowOrder
                fieldObj = fieldObj & ",""group"":" & EditorJson(groupKey)
                fieldObj = fieldObj & ",""source"":" & EditorJson(sourceField)
                fieldObj = fieldObj & ",""format"":" & EditorJson(fmtKey)
                fieldObj = fieldObj & ",""diffFormat"":" & EditorJson(EditorColText(a, cols, r, "DIFFERENCE_FORMAT_KEY"))
                fieldObj = fieldObj & ",""notes"":" & EditorJson(notesText)
                fieldObj = fieldObj & ",""configRow"":" & CStr(lo.DataBodyRange.row + r - 1)
                fieldObj = fieldObj & ",""showManual"":" & EditorBoolJson(a(r, cols("SHOW_MANUAL")))
                fieldObj = fieldObj & ",""showDifference"":" & EditorBoolJson(a(r, cols("SHOW_DIFFERENCE")))
                fieldObj = fieldObj & ",""hideOnOpen"":" & EditorBoolJson(EditorColText(a, cols, r, "HIDE_ROW_ON_OPEN"))
                fieldObj = fieldObj & ",""formulaDefault"":" & EditorJson(EditorFormulaText(a(r, cols("MANUAL_FORMULA_DEFAULT"))))
                fieldObj = fieldObj & ",""formulaModerate"":" & EditorJson(EditorFormulaText(a(r, cols("MANUAL_FORMULA_MODERATE"))))
                fieldObj = fieldObj & ",""formulaMedium"":" & EditorJson(EditorFormulaText(a(r, cols("MANUAL_FORMULA_MEDIUM"))))
                fieldObj = fieldObj & ",""formulaSevere"":" & EditorJson(EditorFormulaText(a(r, cols("MANUAL_FORMULA_SEVERE"))))
                fieldObj = fieldObj & ",""hasDefault"":" & EditorBoolJson(Len(EditorFormulaText(a(r, cols("MANUAL_FORMULA_DEFAULT")))) > 0)
                fieldObj = fieldObj & ",""hasModerate"":" & EditorBoolJson(Len(EditorFormulaText(a(r, cols("MANUAL_FORMULA_MODERATE")))) > 0)
                fieldObj = fieldObj & ",""hasMedium"":" & EditorBoolJson(Len(EditorFormulaText(a(r, cols("MANUAL_FORMULA_MEDIUM")))) > 0)
                fieldObj = fieldObj & ",""hasSevere"":" & EditorBoolJson(Len(EditorFormulaText(a(r, cols("MANUAL_FORMULA_SEVERE")))) > 0) & "}"
                If Len(fields) > 0 Then fields = fields & ","
                fields = fields & fieldObj
            End If
        End If
    Next r

    formula = EditorFormulaText(cell.Value2)
    c = cell.Column - lo.Range.Column + 1
    headerName = UCase$(Trim$(CStr(hdr(1, c))))
    Select Case headerName
        Case "MANUAL_FORMULA_MODERATE": modeName = "Moderate"
        Case "MANUAL_FORMULA_MEDIUM": modeName = "Medium"
        Case "MANUAL_FORMULA_SEVERE": modeName = "Severe"
        Case Else: modeName = "Default"
    End Select
    currentGroup = "": If cols.Exists("DETAIL_GROUP_KEY") Then currentGroup = Trim$(CStr(a(idx, cols("DETAIL_GROUP_KEY"))))
    If Not groupMap Is Nothing Then If groupMap.Exists(Trim$(CStr(a(idx, cols("OUTPUT_ROW_LABEL"))))) Then currentGroup = CStr(groupMap(Trim$(CStr(a(idx, cols("OUTPUT_ROW_LABEL"))))))
    sourceField = Trim$(CStr(a(idx, cols("SOURCE_FIELD_NAME"))))
    variants = "{""default"":" & EditorJson(EditorFormulaText(a(idx, cols("MANUAL_FORMULA_DEFAULT"))))
    variants = variants & ",""moderate"":" & EditorJson(EditorFormulaText(a(idx, cols("MANUAL_FORMULA_MODERATE"))))
    variants = variants & ",""medium"":" & EditorJson(EditorFormulaText(a(idx, cols("MANUAL_FORMULA_MEDIUM"))))
    variants = variants & ",""severe"":" & EditorJson(EditorFormulaText(a(idx, cols("MANUAL_FORMULA_SEVERE")))) & "}"

    ctx = "{""workbook"":" & EditorJson(Me.Parent.name) & ",""target"":" & EditorJson(cell.Parent.name & "!" & cell.address(False, False))
    ctx = ctx & ",""elementType"":" & EditorJson(elementType) & ",""label"":" & EditorJson(CStr(a(idx, cols("OUTPUT_ROW_LABEL"))))
    ctx = ctx & ",""mode"":" & EditorJson(modeName) & ",""sourceField"":" & EditorJson(sourceField) & ",""group"":" & EditorJson(currentGroup)
    ctx = ctx & ",""formula"":" & EditorJson(formula) & ",""variants"":" & variants & ",""fields"":[" & fields & "]}"
    page = Replace(FormulaEditorHtml(), "__JKB_CONTEXT__", ctx)
    page = Replace(page, "charset=""utf-8""", "charset=""utf-16""")
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(fso.GetSpecialFolder(2), "JKB_Formula_" & fso.GetTempName & ".html")
    Set stream = fso.CreateTextFile(path, False, True)
    stream.Write page: stream.Close: Set stream = Nothing
    Me.Parent.FollowHyperlink address:=path, NewWindow:=True
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    MsgBox "Could not open Formula Studio." & vbCrLf & vbCrLf & "Reason: " & er & vbCrLf & vbCrLf & "Next: you can still edit the selected cell with F2.", vbExclamation, "JKB Stress Testing"
End Sub

Private Function EditorJson(ByVal s As String) As String
    Dim i As Long, n As Long, out As String
    For i = 1 To Len(s)
        n = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case n
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 0 To 31, 60, 62, 38, 8232, 8233: out = out & "\u" & Right$("0000" & Hex$(n), 4)
            Case Else: out = out & Mid$(s, i, 1)
        End Select
    Next i
    EditorJson = """" & out & """"
End Function

Private Function EditorColText(ByRef a As Variant, ByVal cols As Object, ByVal r As Long, ByVal key As String) As String
    If cols.Exists(key) Then EditorColText = EditorFormulaText(a(r, cols(key)))
End Function

Private Function EditorBoolJson(ByVal v As Variant) As String
    Select Case UCase$(Trim$(CStr(v)))
        Case "YES", "Y", "TRUE", "1", "ON": EditorBoolJson = "true"
        Case Else: EditorBoolJson = "false"
    End Select
End Function

Private Function EditorFormulaText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    s = Trim$(CStr(v))
    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    EditorFormulaText = Trim$(Replace(s, ChrW(160), " "))
End Function

Private Function FormulaEditorHtml() As String
    Dim s As String
    s = FormulaEditorHtmlPart1()
    s = s & FormulaEditorHtmlPart2()
    s = s & FormulaEditorHtmlPart3()
    s = s & FormulaEditorHtmlPart4()
    ' The Formula Studio wears the same clothes as the sheets it edits. Appended
    ' rather than merged into the page's own stylesheet, so the page stays
    ' readable on its own and the theme is one block you can lift out.
    FormulaEditorHtml = Replace(s, "</style>", ProStudioCss() & "</style>", 1, 1)
End Function

Private Function FormulaEditorHtmlPart1() As String
    Dim s As String
    s = s & "<!doctype html>"
    s = s & vbLf
    s = s & "<html lang=""en""><head><meta charset=""utf-8"">"
    s = s & vbLf
    s = s & "<meta name=""viewport"" content=""width=device-width,initial-scale=1"">"
    s = s & vbLf
    s = s & "<meta http-equiv=""Content-Security-Policy"" content=""default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"">"
    s = s & vbLf
    s = s & "<title>JKB | Formula Studio</title>"
    s = s & vbLf
    s = s & "<style>"
    s = s & vbLf
    s = s & ":root{--green:#067647;--green2:#067647;--green3:#ECFDF3;--ink:#101828;--muted:#667085;--line:#EAECF0;--panel:#FFFFFF;--bg:#F2F4F7;--soft:#F9FAFB;--amber:#B08A2E;--amberbg:#FFFAEB;--red:#B42318;--redbg:#FEF3F2;--blue:#0E7490;--bluebg:#EEF3FB;--shadow:0 14px 44px rgba(23,75,55,.10)}"
    s = s & vbLf
    s = s & "*{box-sizing:border-box}html,body{min-height:100%}body{margin:0;background:linear-gradient(180deg,#EAECF0 0,#F9FAFB 260px,#F2F4F7 100%);color:var(--ink);font:14px/1.45 ""Segoe UI Variable"",""Segoe UI"",Arial,sans-serif}"
    s = s & vbLf
    s = s & "button,input,textarea{font:inherit}.shell{max-width:1540px;height:100vh;margin:0 auto;padding:18px;display:flex;flex-direction:column;overflow:hidden}.topbar{display:flex;align-items:center;justify-content:space-between;gap:18px;padding:18px 22px;background:linear-gradient(120deg,#344054,#067647);border-radius:16px;color:#FFFFFF;box-shadow:var(--shadow)}"
    s = s & vbLf
    s = s & ".brand{display:flex;align-items:center;gap:14px;min-width:0}.mark{width:48px;height:48px;display:grid;place-items:center;border:1px solid rgba(255,255,255,.25);border-radius:12px;background:rgba(255,255,255,.10);font-weight:800;letter-spacing:.8px}.title h1{font-size:23px;line-height:1.1;margin:0 0 5px}.title p{margin:0;color:#ECFDF3;font-size:13px}.topstats{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end}.stat{padding:7px 10px;border-radius:999px;background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.14);font-size:12px;white-sp"
    s = s & "ace:nowrap}.stat b{font-weight:700;color:#FFFFFF}"
    s = s & vbLf
    s = s & ".contextbar{display:grid;grid-template-columns:1.2fr 1fr 1fr 1fr;gap:10px;margin:14px 0}.contextitem{min-width:0;background:rgba(255,255,255,.9);border:1px solid var(--line);border-radius:11px;padding:10px 12px}.contextitem span{display:block;color:var(--muted);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;margin-bottom:3px}.contextitem b{display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px}"
    s = s & vbLf
    s = s & ".grid{display:grid;grid-template-columns:minmax(0,1fr) 405px;gap:14px;align-items:stretch;min-height:0;flex:1;overflow:hidden}.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;box-shadow:0 4px 18px rgba(23,75,55,.04)}.editor{overflow:hidden;min-height:0;display:flex;flex-direction:column}.cardhead{display:flex;align-items:flex-start;justify-content:space-between;gap:15px;padding:18px 20px 12px}.cardhead h2,.sidehead h3{margin:0;font-size:16px}.sub{color:var(--muted);font-size:12px;margin-top:3px}.dirty{display:none;color:var(-"
    s = s & "-amber);background:var(--amberbg);border:1px solid #FEDF89;border-radius:999px;padding:5px 8px;font-size:11px;font-weight:700}.dirty.on{display:inline-flex}"
    s = s & vbLf
    s = s & ".toolbar{padding:0 20px 11px}.toolrow{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-top:7px}.tooltitle{font-size:10px;text-transform:uppercase;letter-spacing:.08em;font-weight:800;color:#667085;width:76px;flex:0 0 76px}.btn{border:1px solid #D0D5DD;background:#FFFFFF;color:#067647;border-radius:8px;padding:7px 9px;cursor:pointer;font-size:12px;font-weight:600;transition:.12s}.btn:hover,.btn:focus{background:#ECFDF3;border-color:#D0D5DD;outline:none}.btn.ref{background:#F9FAFB}.btn.helper{color:#475467;background:#F9FAFB}.btn.operator{font-fa"
    s = s & "mily:Consolas,monospace;min-width:32px}.btn.template{color:#B08A2E;background:#FFFAEB}.btn.primary{background:var(--green);border-color:var(--green);color:#FFFFFF;padding:9px 13px}.btn.primary:hover,.btn.primary:focus{background:var(--green2);color:#FFFFFF}.btn.ghost{background:#FFFFFF}.btn.small{padding:5px 7px;font-size:11px}.btn[disabled]{opacity:.45;cursor:not-allowed}"
    s = s & vbLf
    s = s & ".canvaswrap{position:relative;padding:0 20px}.canvasbar{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:4px 0 7px}.canvasbar label{font-weight:750}.keyhint{color:var(--muted);font-size:11px}.key{display:inline-block;padding:1px 5px;border:1px solid #D0D5DD;border-bottom-width:2px;border-radius:5px;background:#F9FAFB;color:#475467;font:11px Consolas,monospace;margin-left:3px}"
    s = s & vbLf
    s = s & "textarea{display:block;width:100%;height:clamp(190px,25vh,275px);resize:none;border:1px solid #D0D5DD;border-radius:10px;background:#FFFFFF;color:#101828;padding:16px 17px;font:15px/1.58 Consolas,""Cascadia Mono"",monospace;tab-size:2;caret-color:var(--green)}textarea:focus{outline:3px solid rgba(35,102,74,.15);border-color:#067647}.canvasmeta{display:flex;justify-content:space-between;gap:12px;color:var(--muted);font-size:11px;margin:7px 2px 0}.canvasmeta span{white-space:nowrap}"
    s = s & vbLf
    s = s & ".suggestbox{position:fixed;z-index:80;margin:0;border:1px solid #D0D5DD;border-radius:11px;overflow:hidden;box-shadow:0 12px 28px rgba(23,75,55,.08)}.suggestbox[hidden]{display:none}.suggesthead{display:flex;align-items:center;justify-content:space-between;padding:7px 10px;background:#F9FAFB;border-bottom:1px solid var(--line);font-size:11px;color:var(--muted)}#suggestions{max-height:300px;overflow:auto;margin:0;padding:4px;list-style:none;background:#FFFFFF}.suggestion{width:100%;display:grid;grid-template-columns:30px minmax(0,1fr) auto;gap:9px;align-item"
    s = s & "s:center;text-align:left;border:0;border-radius:7px;background:#FFFFFF;padding:8px 9px;cursor:pointer;color:var(--ink)}.suggestion:hover,.suggestion.selected{background:#ECFDF3}.suggestion.disabled{opacity:.68}.sicon{width:28px;height:28px;display:grid;place-items:center;border-radius:7px;background:#F2F4F7;color:var(--green);font-size:10px;font-weight:800}.sicon.fn{background:var(--bluebg);color:var(--blue)}.sicon.sec{background:#FFFAEB;color:#B08A2E}.sname{font:13px Consolas,""Cascadia Mono"",monospace;font-weight:700;overflow:hidden;text-overflow:ellipsis}"
    s = s & ".sdetail{font-size:10px;color:var(--muted);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.tag{display:inline-flex;align-items:center;border-radius:999px;padding:3px 6px;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;background:#F2F4F7;color:#475467}.tag.off{background:var(--redbg);color:var(--red)}"
    s = s & vbLf
    s = s & ".translatebox{margin:10px 20px 0;border:1px solid #EAECF0;border-radius:10px;background:#F9FAFB;overflow:hidden}.translatehead{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:8px 10px;border-bottom:1px solid #EAECF0}.translatehead b{font-size:11px;text-transform:uppercase;letter-spacing:.05em}.translatehead span{font-size:10px;color:var(--muted)}.translation{margin:0;padding:10px 11px;max-height:72px;overflow:auto;white-space:pre-wrap;word-break:break-word;font:11px/1.45 Consolas,monospace;color:#475467}.transnote{padding:0"
    s = s & " 11px 9px;color:var(--muted);font-size:10px}.editorfoot{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 20px 18px;border-top:1px solid #F2F4F7;margin-top:12px}.actions{display:flex;gap:7px;flex-wrap:wrap}.status{min-height:20px;color:#067647;font-size:12px;text-align:right}.status.bad{color:var(--red)}"
    s = s & vbLf
    s = s & ".side{min-height:0}.sidecard{padding:0;height:100%;display:flex;flex-direction:column;min-height:0;overflow:hidden}.sidetop{padding:14px 14px 8px}.tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;padding:0 10px 10px;border-bottom:1px solid #EAECF0}.tab{border:1px solid #EAECF0;background:#F9FAFB;border-radius:7px;padding:7px 5px;color:#475467;font-size:10px;font-weight:800;cursor:pointer}.tab.active{background:#ECFDF3;color:#067647;border-color:#D0D5DD}.tabpane{display:none;padding:12px 14px 14px;overflow:auto;min-height:0;flex:1}.tabpane.ac"
    s = s & "tive{display:block}.sidehead{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:10px}.sidehead{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:10px}.healthpill{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:5px 8px;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;background:#ECFDF3;color:#067647}.healthpill.review{background:var(--amberbg);color:var(--amber)}.healthpill.blocked{background:var(--redbg);color:var(--red)}.healthline{d"
    s = s & "isplay:flex;align-items:flex-start;gap:8px;padding:7px 0;border-top:1px solid #F2F4F7}.healthline:first-of-type{border-top:0}.hicon{width:18px;height:18px;flex:0 0 18px;border-radius:50%;display:grid;place-items:center;background:#ECFDF3;color:#067647;font-size:10px;font-weight:900}.healthline.warn .hicon{background:var(--amberbg);color:var(--amber)}.healthline.error .hicon{background:var(--redbg);color:var(--red)}.healthtext{font-size:12px}.healthtext b{display:block;font-size:12px}.healthtext span{color:var(--muted)}"
    s = s & vbLf
    s = s & ".searchrow{display:flex;gap:6px}.search{width:100%;border:1px solid #D0D5DD;border-radius:8px;padding:8px 9px;color:var(--ink);background:#FFFFFF}.search:focus{outline:2px solid rgba(35,102,74,.12);border-color:#067647}.toggle{display:flex;align-items:center;gap:5px;color:var(--muted);font-size:10px;white-space:nowrap}.library{overflow:visible;margin-top:8px;border-top:1px solid #F2F4F7}.librow{padding:8px 2px;border-bottom:1px solid #F2F4F7}.libtop{display:flex;justify-content:space-between;align-items:center;gap:8px}.libname{font:12px Consolas,monospace;f"
    s = s & "ont-weight:700;word-break:break-all}.libmeta{color:var(--muted);font-size:10px;margin:3px 0 6px}.libbuttons{display:flex;gap:4px}.libbuttons button{border:1px solid #D0D5DD;border-radius:5px;background:#F9FAFB;color:#067647;padding:3px 6px;font:10px Consolas,monospace;cursor:pointer}.libbuttons button:hover{background:#ECFDF3}.empty{color:var(--muted);font-size:12px;padding:12px 2px}"
    s = s & vbLf
    s = s & ".inspector{background:#FCFCFD;border:1px solid #F2F4F7;border-radius:9px;padding:10px}.inspector h4{font:13px Consolas,monospace;margin:0 0 7px;word-break:break-all}.kv{display:grid;grid-template-columns:88px minmax(0,1fr);gap:5px 8px;font-size:11px}.kv dt{color:var(--muted)}.kv dd{margin:0;overflow-wrap:anywhere}.note{margin-top:8px;padding-top:8px;border-top:1px solid #EAECF0;color:#475467;font-size:11px;white-space:pre-wrap}.refs{display:flex;flex-wrap:wrap;gap:5px}.refchip{border:1px solid #EAECF0;background:#F9FAFB;border-radius:7px;padding:4px 6px;"
    s = s & "font:10px Consolas,monospace;color:#475467;cursor:pointer}.refchip:hover{background:#ECFDF3}.variants{display:grid;gap:6px}.variant{border:1px solid #EAECF0;border-radius:8px;padding:7px 8px;background:#FCFCFD}.varianttop{display:flex;justify-content:space-between;align-items:center;gap:7px}.variant b{font-size:10px;text-transform:uppercase;letter-spacing:.05em;color:#667085}.variant code{display:block;margin-top:4px;font:10px/1.4 Consolas,monospace;color:#344054;max-height:44px;overflow:hidden;white-space:pre-wrap;word-break:break-word}.variant.empty co"
    s = s & "de{color:#98A2B3;font-family:inherit}.docline{font-size:11px;color:var(--muted);margin-top:8px}.docline code{font-family:Consolas,monospace;color:#475467}"
    s = s & vbLf
    s = s & "@media(max-height:850px) and (min-width:1121px){.shell{padding:10px}.topbar{padding:11px 15px;border-radius:12px}.mark{width:40px;height:40px}.title h1{font-size:20px}.title p{font-size:11px}.stat{padding:5px 8px;font-size:10px}.contextbar{margin:8px 0;gap:7px}.contextitem{padding:6px 9px}.contextitem span{font-size:9px;margin-bottom:1px}.contextitem b{font-size:11px}.cardhead{padding:10px 15px 5px}.cardhead h2,.sidehead h3{font-size:14px}.sub{font-size:10px}.toolbar{padding:0 15px 6px}.toolrow{gap:4px;margin-top:4px}.tooltitle{font-size:9px;width:72px;f"
    s = s & "lex-basis:72px}.btn{padding:5px 7px;font-size:10px;border-radius:7px}.canvaswrap{padding:0 15px}.canvasbar{margin:2px 0 4px}.canvasbar label{font-size:12px}.keyhint{font-size:9px}textarea{height:clamp(140px,23vh,185px);padding:11px 12px;font-size:13px;line-height:1.45}.canvasmeta{margin-top:4px;font-size:9px}.translatebox{margin:6px 15px 0}.translatehead{padding:5px 8px}.translation{padding:6px 8px;max-height:44px;font-size:9px}.transnote{padding:0 8px 5px;font-size:9px}.editorfoot{padding:8px 15px 10px;margin-top:6px}.sidetop{padding:10px 11px 5px}.tabs"
    s = s & "{padding:0 8px 6px}.tab{padding:5px 3px}.tabpane{padding:8px 11px 10px}.healthline{padding:5px 0}.docline{font-size:9px}.library{margin-top:5px}.librow{padding:5px 2px}}"
    s = s & vbLf
    s = s & "@media(max-width:1120px){.shell{height:auto;overflow:visible}.grid{grid-template-columns:1fr;overflow:visible}.sidecard{height:auto;min-height:560px}.contextbar{grid-template-columns:1fr 1fr}}@media(max-width:720px){.shell{padding:0}.topbar{border-radius:0}.topstats{display:none}.contextbar{padding:0 10px;grid-template-columns:1fr 1fr}.grid{padding:0 10px 12px}.side{grid-template-columns:1fr}.tooltitle{width:100%;flex-basis:100%}.editorfoot{align-items:flex-start;flex-direction:column}.status{text-align:left}.contextitem:nth-child(n+3){display:none}}"
    s = s & vbLf
    s = s & "</style></head><body><div class=""shell"">"
    s = s & vbLf
    s = s & "<div class=""topbar""><div class=""brand""><div class=""mark"">JKB</div><div class=""title""><h1>Formula Studio</h1><p>Stress-testing reconciliation | intelligent configuration editor | offline</p></div></div><div class=""topstats""><div class=""stat""><b id=""activeCount"">0</b> active fields</div><div class=""stat""><b id=""disabledCount"">0</b> disabled fields</div><div class=""stat""><b id=""modeTop"">Default</b> formula</div></div></div>"
    s = s & vbLf
    s = s & "<div class=""contextbar""><div class=""contextitem""><span>Element type</span><b id=""ctxElement"">-</b></div><div class=""contextitem""><span>Output row</span><b id=""ctxLabel"">-</b></div><div class=""contextitem""><span>Formula slot</span><b id=""ctxMode"">-</b></div><div class=""contextitem""><span>Excel target</span><b id=""ctxTarget"">-</b></div></div>"
    s = s & vbLf
    s = s & "<div class=""grid""><section class=""card editor""><div class=""cardhead""><div><h2>Formula canvas</h2><div class=""sub"" id=""rowMeta"">Build with row references, helpers and Excel functions. Nothing is written to Excel until you paste it back.</div></div><span class=""dirty"" id=""dirty"">Modified</span></div>"
    s = s & vbLf
    s = s & "<div class=""toolbar""><div class=""toolrow""><span class=""tooltitle"">References</span><button class=""btn ref"" data-section=""base"" title=""Pre-shock/base column"">@base.</button><button class=""btn ref"" data-section=""system"" title=""System output for the active severity"">@system.</button><button class=""btn ref"" data-section=""manual"" title=""Independent/manual result for the active severity"">@manual.</button><button class=""btn ref"" data-section=""difference"" title=""System minus manual for the active severity; use carefully"">@difference.</button></div>"
    s = s & vbLf
    s = s & "<div class=""toolrow""><span class=""tooltitle"">Logic</span><button class=""btn helper"" data-template=""ifelse"" title=""Natural IF / THEN / ELSE; the engine translates it to Excel IF()"">IF / ELSE</button><button class=""btn helper"" data-template=""elseif"" title=""Nested ELSE IF chain"">ELSE IF</button><button class=""btn helper"" data-insert="" AND "" title=""Natural AND; translated to Excel AND()"">AND</button><button class=""btn helper"" data-insert="" OR "" title=""Natural OR; translated to Excel OR()"">OR</button><button class=""btn helper"" data-fn=""NOT"">NOT</button><butto"
    s = s & "n class=""btn helper"" data-fn=""IFNULL"">IFNULL</button><button class=""btn helper"" data-fn=""ISNULL"">ISNULL</button><button class=""btn helper"" data-fn=""IFERROR"">IFERROR</button></div>"
    s = s & vbLf
    FormulaEditorHtmlPart1 = s
End Function

Private Function FormulaEditorHtmlPart2() As String
    Dim s As String
    s = s & "<div class=""toolrow""><span class=""tooltitle"">Math</span><button class=""btn helper"" data-fn=""ABS"">ABS</button><button class=""btn helper"" data-fn=""MIN"">MIN</button><button class=""btn helper"" data-fn=""MAX"">MAX</button><button class=""btn helper"" data-fn=""SUM"">SUM</button><button class=""btn helper"" data-fn=""ROUND"">ROUND</button><button class=""btn helper"" data-fn=""ROUNDUP"">ROUNDUP</button><button class=""btn helper"" data-fn=""ROUNDDOWN"">ROUNDDOWN</button><button class=""btn helper"" data-fn=""SIGN"">SIGN</button></div>"
    s = s & vbLf
    s = s & "<div class=""toolrow""><span class=""tooltitle"">Operators</span><button class=""btn operator"" data-insert="" + "">+</button><button class=""btn operator"" data-insert="" - "">-</button><button class=""btn operator"" data-insert="" * "">*</button><button class=""btn operator"" data-insert="" / "">/</button><button class=""btn operator"" data-insert=""("">(</button><button class=""btn operator"" data-insert="")"">)</button><button class=""btn operator"" data-insert=""="">=</button><button class=""btn operator"" data-insert=""&lt;&gt;"">&lt;&gt;</button><button class=""btn operator"" data-ins"
    s = s & "ert=""&lt;="">&lt;=</button><button class=""btn operator"" data-insert=""&gt;="">&gt;=</button><button class=""btn operator"" data-insert=""&lt;"">&lt;</button><button class=""btn operator"" data-insert=""&gt;"">&gt;</button><button class=""btn operator"" data-insert="" ^ "">^</button></div>"
    s = s & vbLf
    s = s & "<div class=""toolrow""><span class=""tooltitle"">Templates</span><button class=""btn template"" data-template=""ifand"">IF + AND</button><button class=""btn template"" data-template=""ifor"">IF + OR</button><button class=""btn template"" data-template=""passthrough"">System value</button><button class=""btn template"" data-template=""baseimpact"">Base + impact</button><button class=""btn template"" data-template=""saferatio"">Safe ratio</button><button class=""btn template"" data-template=""stage3"">3-stage sum</button><button class=""btn template"" data-template=""shockpct"">% shock</"
    s = s & "button><button class=""btn template"" data-template=""nullfallback"">Null fallback</button></div></div>"
    s = s & vbLf
    s = s & "<div class=""canvaswrap""><div class=""canvasbar""><label for=""formula"">Manual formula</label><span class=""keyhint"">Autocomplete <span class=""key"">Ctrl</span><span class=""key"">Space</span> &nbsp; Copy <span class=""key"">Ctrl</span><span class=""key"">Enter</span></span></div><textarea id=""formula"" spellcheck=""false"" autocomplete=""off"" aria-autocomplete=""list"" aria-controls=""suggestions"" wrap=""soft""></textarea><div class=""canvasmeta""><span id=""matches"">Ready</span><span id=""metrics"">0 chars | 1 line | 0 refs</span></div></div>"
    s = s & vbLf
    s = s & "<div class=""suggestbox"" id=""suggestBox"" hidden><div class=""suggesthead""><span id=""suggestTitle"">Suggestions</span><span>Up/Down to move | Enter/Tab to insert | Esc to close</span></div><ul id=""suggestions"" role=""listbox"" aria-label=""Formula suggestions""></ul></div>"
    s = s & vbLf
    s = s & "<div class=""translatebox""><div class=""translatehead""><b>Excel logic preview</b><span><span id=""translationState"">Engine-normalized syntax</span> &nbsp; <button id=""normalizeBtn"" class=""btn small"" type=""button"">Use normalized syntax</button></span></div><pre class=""translation"" id=""translation""></pre><div class=""transnote"">This preview translates IF / THEN / ELSE, infix AND / OR and IFELSE. Row tokens remain symbolic here and are resolved to generated workbook cell references by the VBA engine.</div></div><div class=""editorfoot""><div class=""actions""><butt"
    s = s & "on id=""copy"" class=""btn primary"" type=""button"">Copy formula to Config</button><button id=""reset"" class=""btn ghost"" type=""button"">Restore starting formula</button><button id=""wrap"" class=""btn ghost"" type=""button"">Do not wrap</button></div><div class=""status"" id=""status"" role=""status""></div></div></section>"
    s = s & vbLf
    s = s & "<aside class=""side""><section class=""card sidecard""><div class=""sidetop""><div class=""sidehead""><h3>Formula intelligence</h3><span class=""healthpill"" id=""healthPill"">Ready</span></div><div class=""sub"">Everything configured for this element type is searchable here, including disabled rows.</div></div><div class=""tabs""><button class=""tab active"" data-tab=""healthPane"">Health</button><button class=""tab"" data-tab=""fieldsPane"">Fields</button><button class=""tab"" data-tab=""inspectPane"">Inspector</button><button class=""tab"" data-tab=""variantsPane"">Variants</button>"
    s = s & "</div><div class=""tabpane active"" id=""healthPane""><div class=""sidehead""><h3>Formula health</h3><span class=""sub"" id=""refCount"">0 refs</span></div><div id=""health""></div><div class=""sidehead"" style=""margin-top:14px""><h3>References in formula</h3><span class=""sub"">click to inspect</span></div><div class=""refs"" id=""refs""><span class=""empty"">No row references yet.</span></div><div class=""docline"">Canonical references: <code>@base.FIELD</code>, <code>@system.FIELD</code>, <code>@manual.FIELD</code>, <code>@difference.FIELD</code>. Natural <code>IF ... THEN .."
    s = s & ". ELSE</code>, infix <code>AND</code>/<code>OR</code>, standard Excel functions and the IFNULL/ISNULL helpers are supported. The function autocomplete is curated, not restrictive: other valid Excel functions can still be typed directly.</div></div><div class=""tabpane"" id=""fieldsPane""><div class=""sidehead""><h3>Field library</h3><span class=""sub"" id=""libraryCount""></span></div><div class=""searchrow""><input class=""search"" id=""fieldSearch"" type=""search"" placeholder=""Search label, source field, group or notes..."" autocomplete=""off""><label class=""toggle""><inpu"
    s = s & "t id=""showDisabled"" type=""checkbox"" checked> disabled</label></div><div class=""library"" id=""library""></div></div><div class=""tabpane"" id=""inspectPane""><div class=""sidehead""><h3>Field inspector</h3><span class=""sub"" id=""inspectState"">Select a field</span></div><div class=""inspector"" id=""inspector""><div class=""empty"">Choose a suggestion, reference or field. The inspector shows configuration status, source mapping, grouping, formats, notes and formula variants.</div></div></div><div class=""tabpane"" id=""variantsPane""><div class=""sidehead""><h3>Formula variant"
    s = s & "s</h3><span class=""sub"">same output row</span></div><div class=""variants"" id=""variants""></div></div></section></aside></div></div>"
    s = s & vbLf
    s = s & "<script>"
    s = s & vbLf
    s = s & "const data=__JKB_CONTEXT__;"
    s = s & vbLf
    s = s & "const input=document.getElementById('formula'),list=document.getElementById('suggestions'),suggestBox=document.getElementById('suggestBox'),matches=document.getElementById('matches'),status=document.getElementById('status'),health=document.getElementById('health'),healthPill=document.getElementById('healthPill');"
    s = s & vbLf
    s = s & "const sections=[{name:'base',label:'Base',desc:'Pre-shock/base value'},{name:'system',label:'System',desc:'System output for active severity'},{name:'manual',label:'Manual',desc:'Independent/manual value for active severity'},{name:'difference',label:'Difference',desc:'System minus manual; can create dependency loops'}];"
    s = s & vbLf
    s = s & "const sectionAliases={base:'base',system:'system',systemoutput:'system',manual:'manual',man:'manual',difference:'difference',differences:'difference',diff:'difference'};"
    s = s & vbLf
    s = s & "const functions=["
    s = s & vbLf
    s = s & "{name:'IF',snippet:'IF(condition,value_if_true,value_if_false)',desc:'Excel conditional',placeholder:'condition'},{name:'IFNULL',snippet:'IFNULL(value,fallback)',desc:'Engine helper: use fallback when blank/error',placeholder:'value'},{name:'ISNULL',snippet:'ISNULL(value)',desc:'Engine helper: test blank/error',placeholder:'value'},{name:'IFELSE',snippet:'IFELSE(condition,value_if_true,value_if_false)',desc:'Engine alias converted to IF',placeholder:'condition'},{name:'IFERROR',snippet:'IFERROR(value,value_if_error)',desc:'Excel error fallback',placehold"
    s = s & "er:'value'},{name:'AND',snippet:'AND(logical1,logical2)',desc:'All conditions true',placeholder:'logical1'},{name:'OR',snippet:'OR(logical1,logical2)',desc:'Any condition true',placeholder:'logical1'},{name:'NOT',snippet:'NOT(logical)',desc:'Reverse TRUE/FALSE',placeholder:'logical'},{name:'ABS',snippet:'ABS(number)',desc:'Absolute value',placeholder:'number'},{name:'MIN',snippet:'MIN(number1,number2)',desc:'Minimum',placeholder:'number1'},{name:'MAX',snippet:'MAX(number1,number2)',desc:'Maximum',placeholder:'number1'},{name:'SUM',snippet:'SUM(number1,nu"
    s = s & "mber2)',desc:'Sum values',placeholder:'number1'},{name:'ROUND',snippet:'ROUND(number,num_digits)',desc:'Round to digits',placeholder:'number'},{name:'ROUNDUP',snippet:'ROUNDUP(number,num_digits)',desc:'Round away from zero',placeholder:'number'},{name:'ROUNDDOWN',snippet:'ROUNDDOWN(number,num_digits)',desc:'Round toward zero',placeholder:'number'},{name:'SIGN',snippet:'SIGN(number)',desc:'Return sign of a number',placeholder:'number'},{name:'LEN',snippet:'LEN(text)',desc:'Text length',placeholder:'text'},{name:'AVERAGE',snippet:'AVERAGE(number1,number2)'"
    s = s & ",desc:'Average values',placeholder:'number1'},{name:'MEDIAN',snippet:'MEDIAN(number1,number2)',desc:'Median value',placeholder:'number1'},{name:'COUNT',snippet:'COUNT(value1,value2)',desc:'Count numeric values',placeholder:'value1'},{name:'COUNTA',snippet:'COUNTA(value1,value2)',desc:'Count non-blank values',placeholder:'value1'},{name:'COUNTIF',snippet:'COUNTIF(range,criteria)',desc:'Conditional count',placeholder:'range'},{name:'COUNTIFS',snippet:'COUNTIFS(range1,criteria1,range2,criteria2)',desc:'Multiple-condition count',placeholder:'range1'},{name:'"
    s = s & "SUMIF',snippet:'SUMIF(range,criteria,sum_range)',desc:'Conditional sum',placeholder:'range'},{name:'SUMIFS',snippet:'SUMIFS(sum_range,criteria_range1,criteria1)',desc:'Multiple-condition sum',placeholder:'sum_range'},{name:'AVERAGEIF',snippet:'AVERAGEIF(range,criteria,average_range)',desc:'Conditional average',placeholder:'range'},{name:'MOD',snippet:'MOD(number,divisor)',desc:'Remainder',placeholder:'number'},{name:'POWER',snippet:'POWER(number,power)',desc:'Raise to a power',placeholder:'number'},{name:'SQRT',snippet:'SQRT(number)',desc:'Square root',p"
    s = s & "laceholder:'number'},{name:'INT',snippet:'INT(number)',desc:'Round down to integer',placeholder:'number'},{name:'CEILING',snippet:'CEILING(number,significance)',desc:'Round up to multiple',placeholder:'number'},{name:'FLOOR',snippet:'FLOOR(number,significance)',desc:'Round down to multiple',placeholder:'number'},{name:'LEFT',snippet:'LEFT(text,num_chars)',desc:'Left-side text',placeholder:'text'},{name:'RIGHT',snippet:'RIGHT(text,num_chars)',desc:'Right-side text',placeholder:'text'},{name:'MID',snippet:'MID(text,start_num,num_chars)',desc:'Middle text',"
    s = s & "placeholder:'text'},{name:'TRIM',snippet:'TRIM(text)',desc:'Remove extra spaces',placeholder:'text'},{name:'UPPER',snippet:'UPPER(text)',desc:'Uppercase text',placeholder:'text'},{name:'LOWER',snippet:'LOWER(text)',desc:'Lowercase text',placeholder:'text'},{name:'VALUE',snippet:'VALUE(text)',desc:'Convert text to number',placeholder:'text'},{name:'TEXT',snippet:'TEXT(value,format_text)',desc:'Format value as text',placeholder:'value'},{name:'ISNUMBER',snippet:'ISNUMBER(value)',desc:'Test whether value is numeric',placeholder:'value'},{name:'ISTEXT',snipp"
    s = s & "et:'ISTEXT(value)',desc:'Test whether value is text',placeholder:'value'},{name:'ISBLANK',snippet:'ISBLANK(value)',desc:'Test blank cell',placeholder:'value'},{name:'ISERROR',snippet:'ISERROR(value)',desc:'Test any Excel error',placeholder:'value'},{name:'TRUE',snippet:'TRUE()',desc:'Logical TRUE',placeholder:'TRUE'},{name:'FALSE',snippet:'FALSE()',desc:'Logical FALSE',placeholder:'FALSE'}];"
    s = s & vbLf
    s = s & "const templates={"
    s = s & vbLf
    s = s & "passthrough:{text:'@system.FIELD',select:'FIELD'},"
    s = s & vbLf
    s = s & "baseimpact:{text:'@base.FIELD + @manual.IMPACT_FIELD',select:'FIELD'},"
    s = s & vbLf
    s = s & "saferatio:{text:'IF(@manual.DENOMINATOR=0,@base.RESULT,@manual.NUMERATOR/@manual.DENOMINATOR)',select:'DENOMINATOR'},"
    s = s & vbLf
    s = s & "stage3:{text:'@manual.STAGE1 + @manual.STAGE2 + @manual.STAGE3',select:'STAGE1'},"
    s = s & vbLf
    s = s & "shockpct:{text:'@base.FIELD * @system.PCT_CHANGE * 0.01',select:'FIELD'},"
    s = s & vbLf
    s = s & "nullfallback:{text:'IFNULL(@manual.FIELD,0)',select:'FIELD'},"
    s = s & vbLf
    s = s & "ifelse:{text:'IF condition THEN value_if_true ELSE value_if_false',select:'condition'},"
    s = s & vbLf
    s = s & "elseif:{text:'IF condition1 THEN value1 ELSE IF condition2 THEN value2 ELSE value3',select:'condition1'},"
    s = s & vbLf
    s = s & "ifand:{text:'IF condition1 AND condition2 THEN value_if_true ELSE value_if_false',select:'condition1'},"
    s = s & vbLf
    s = s & "ifor:{text:'IF condition1 OR condition2 THEN value_if_true ELSE value_if_false',select:'condition1'}};"
    s = s & vbLf
    s = s & "const rawFields=(data.fields||[]).map((f,i)=>typeof f==='string'?{name:f,enabled:true,order:i+1}:f).filter(f=>f&&f.name);"
    s = s & vbLf
    s = s & "const fieldMap=new Map();rawFields.forEach(f=>{const k=f.name.toLowerCase();if(!fieldMap.has(k))fieldMap.set(k,f);else if(f.enabled)fieldMap.set(k,f)});"
    s = s & vbLf
    s = s & "const fields=[...fieldMap.values()].sort((a,b)=>(a.enabled===b.enabled?((Number(a.order)||999999)-(Number(b.order)||999999)||a.name.localeCompare(b.name)):(a.enabled?-1:1)));"
    s = s & vbLf
    FormulaEditorHtmlPart2 = s
End Function

Private Function FormulaEditorHtmlPart3() As String
    Dim s As String
    s = s & "let options=[],selected=0,activeToken=null,dismissed=false,lastInspected=null,forceFunctions=false;"
    s = s & vbLf
    s = s & "function isEnabledFlag(v){return v===true||String(v).toUpperCase()==='YES'||String(v).toUpperCase()==='Y'||String(v)==='1'||String(v).toUpperCase()==='ON'}"
    s = s & vbLf
    s = s & "function insideString(s){let q=false;for(let i=0;i<s.length;i++){if(s[i]==='""'){if(q&&s[i+1]==='""')i++;else q=!q}}return q}"
    s = s & vbLf
    s = s & "function normalSection(s){return sectionAliases[String(s||'').replace(/^@/,'').toLowerCase()]||''}"
    s = s & vbLf
    s = s & "function tokenAt(text,position){const before=text.slice(0,position);if(insideString(before))return null;let m=before.match(/(?:^|[^A-Za-z0-9_])(@?[A-Za-z_][A-Za-z0-9_]*\.([A-Za-z0-9_]*)|@[A-Za-z_]*)$/i);if(m){const value=m[1],start=position-value.length,dot=value.indexOf('.');if(dot<0)return {start,end:position,kind:'section',prefix:value.slice(1).toLowerCase()};const sec=normalSection(value.slice(0,dot));if(!sec)return {start,end:position,kind:'badsection',prefix:value.slice(0,dot),fieldPrefix:value.slice(dot+1)};return {start:start+dot+1,end:position,k"
    s = s & "ind:'field',section:sec,prefix:value.slice(dot+1).toLowerCase()}}m=before.match(/(?:^|[,(+\-*/=<>\s])([A-Za-z][A-Za-z0-9_]*)$/);if(m){const word=m[1];return {start:position-word.length,end:position,kind:'function',prefix:word.toLowerCase()}}return null}"
    s = s & vbLf
    s = s & "function compact(s){return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'')}"
    s = s & vbLf
    s = s & "function subseqScore(name,q){const n=compact(name),x=compact(q);if(!x)return 0;if(n===x)return 0;if(n.startsWith(x))return 1;const segs=String(name).toLowerCase().split(/[_\s\-]+/);if(segs.some(v=>v.startsWith(String(q).toLowerCase())))return 5;const at=n.indexOf(x);if(at>=0)return 10+at;let j=0,gap=0,last=-1;for(let i=0;i<n.length&&j<x.length;i++){if(n[i]===x[j]){if(last>=0)gap+=i-last-1;last=i;j++}}return j===x.length?30+gap:9999}"
    s = s & vbLf
    s = s & "function fieldOptions(prefix,includeDisabled=true){return fields.map(f=>{const candidates=[[f.name,0],[f.source,12],[f.group,22],[f.notes,45],[f.format,55]];let score=Math.min(...candidates.filter(x=>x[0]).map(x=>subseqScore(x[0],prefix)+x[1]),9999);return {type:'field',value:f.name,field:f,score}}).filter(o=>o.score<9999&&(includeDisabled||o.field.enabled)).sort((a,b)=>a.score-b.score||(a.field.enabled===b.field.enabled?((Number(a.field.order)||999999)-(Number(b.field.order)||999999)):(a.field.enabled?-1:1))||a.value.localeCompare(b.value))}"
    s = s & vbLf
    s = s & "function suggestionsFor(text,position){const token=tokenAt(text,position);let items=[];if(token){if(token.kind==='section')items=sections.filter(x=>x.name.startsWith(token.prefix)).map(x=>({type:'section',value:x.name,desc:x.desc}));else if(token.kind==='field')items=fieldOptions(token.prefix,true);else if(token.kind==='function')items=functions.filter(x=>subseqScore(x.name,token.prefix)<9999).sort((a,b)=>subseqScore(a.name,token.prefix)-subseqScore(b.name,token.prefix)).map(x=>({type:'function',value:x.name,fn:x,desc:x.desc}));}return {token,items}}"
    s = s & vbLf
    s = s & "function inspectField(f){if(!f)return;lastInspected=f;document.getElementById('inspectState').textContent=f.enabled?'Active field':'Disabled field';const formTags=['default','moderate','medium','severe'].filter(k=>f['has'+k[0].toUpperCase()+k.slice(1)]).map(k=>k[0].toUpperCase()).join(' / ')||'none';const formulaRows=[['Default',f.formulaDefault],['Moderate',f.formulaModerate],['Medium',f.formulaMedium],['Severe',f.formulaSevere]].filter(x=>x[1]).map(x=>'<div class=""variant""><div class=""varianttop""><b>'+x[0]+'</b></div><code>'+esc(x[1])+'</code></div>')."
    s = s & "join('');document.getElementById('inspector').innerHTML='<h4>'+esc(f.name)+'</h4><dl class=""kv""><dt>Status</dt><dd>'+(f.enabled?'<span class=""tag"">Enabled</span>':'<span class=""tag off"">Disabled</span>')+'</dd><dt>Config row</dt><dd>'+esc(f.configRow||'-')+'</dd><dt>Output order</dt><dd>'+esc(f.order||'-')+'</dd><dt>Detail group</dt><dd>'+esc(f.group||'-')+'</dd><dt>Source field</dt><dd>'+esc(f.source||'-')+'</dd><dt>Value format</dt><dd>'+esc(f.format||'-')+'</dd><dt>Diff format</dt><dd>'+esc(f.diffFormat||'-')+'</dd><dt>Manual output</dt><dd>'+(isEnabl"
    s = s & "edFlag(f.showManual)?'Yes':'No')+'</dd><dt>Difference</dt><dd>'+(isEnabledFlag(f.showDifference)?'Yes':'No')+'</dd><dt>Hidden on open</dt><dd>'+(isEnabledFlag(f.hideOnOpen)?'Yes':'No')+'</dd><dt>Formula slots</dt><dd>'+esc(formTags)+'</dd></dl>'+(f.notes?'<div class=""note"">'+esc(f.notes)+'</div>':'')+(formulaRows?'<div class=""note""><b>Configured formulas</b><div class=""variants"" style=""margin-top:7px"">'+formulaRows+'</div></div>':'')+(f.enabled?'':'<div class=""note"" style=""color:var(--red)"">Disabled rows are deliberately shown so the editor does not hide"
    s = s & " available configuration. They are not generated; current engine behavior falls back to 0 when a disabled/absent row is referenced.</div>')}"
    s = s & vbLf
    s = s & "function esc(v){return String(v??'').replace(/[&<>""']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','""':'&quot;',""'"":'&#39;'}[c]))}"
    s = s & vbLf
    s = s & "function positionSuggestionBox(){const r=input.getBoundingClientRect();suggestBox.style.left=Math.max(8,r.left)+'px';suggestBox.style.width=Math.max(280,r.width)+'px';suggestBox.style.top=Math.min(innerHeight-330,r.top+42)+'px'}"
    s = s & vbLf
    s = s & "function paintSuggestions(){list.replaceChildren();const visible=!dismissed&&options.length>0&&activeToken&&activeToken.kind!=='badsection';suggestBox.hidden=!visible;if(visible){positionSuggestionBox();const max=Math.min(options.length,80);for(let i=0;i<max;i++){const o=options[i],li=document.createElement('li'),b=document.createElement('button');b.type='button';b.className='suggestion'+(i===selected?' selected':'')+(o.field&&!o.field.enabled?' disabled':'');const icon=document.createElement('span');icon.className='sicon '+(o.type==='function'?'fn':o.ty"
    s = s & "pe==='section'?'sec':'');icon.textContent=o.type==='field'?'F':o.type==='function'?'FN':'@';const mid=document.createElement('span'),name=document.createElement('div'),detail=document.createElement('div');name.className='sname';name.textContent=o.value;detail.className='sdetail';detail.textContent=o.type==='field'?((o.field.enabled?'Enabled':'Disabled')+(o.field.group?' | '+o.field.group:'')+(o.field.source?' | source: '+o.field.source:'')):(o.desc||'');mid.append(name,detail);const tag=document.createElement('span');tag.className='tag'+(o.field&&!o.fiel"
    s = s & "d.enabled?' off':'');tag.textContent=o.type==='field'?(o.field.enabled?'field':'disabled'):o.type;b.append(icon,mid,tag);b.addEventListener('mousedown',e=>e.preventDefault());b.addEventListener('mouseenter',()=>{selected=i;inspectField(o.field);highlightSuggestion()});b.addEventListener('click',()=>choose(i));li.append(b);list.append(li)}document.getElementById('suggestTitle').textContent=(activeToken.kind==='field'?(activeToken.section+' fields'):activeToken.kind==='function'?'Functions':'Reference sections')+' | '+options.length+' match'+(options.lengt"
    s = s & "h===1?'':'es')}}"
    s = s & vbLf
    s = s & "function highlightSuggestion(){[...list.querySelectorAll('.suggestion')].forEach((b,i)=>b.classList.toggle('selected',i===selected));const o=options[selected];if(o&&o.field)inspectField(o.field)}"
    s = s & vbLf
    s = s & "function choose(i){const o=options[i];if(!activeToken||!o)return;if(o.type==='section'){input.setRangeText('@'+o.value+'.',activeToken.start,activeToken.end,'end')}else if(o.type==='field'){const tail=input.value.slice(activeToken.end).match(/^[A-Za-z0-9_]*/)?.[0]?.length||0;input.setRangeText(o.value,activeToken.start,activeToken.end+tail,'end');inspectField(o.field)}else if(o.type==='function'){insertSnippet(o.fn.snippet,o.fn.placeholder,activeToken.start,activeToken.end)}input.focus();refresh()}"
    s = s & vbLf
    s = s & "function insertSnippet(text,placeholder,start=input.selectionStart,end=input.selectionEnd){input.setRangeText(text,start,end,'end');const base=start,at=text.indexOf(placeholder);if(at>=0)input.setSelectionRange(base+at,base+at+placeholder.length);input.focus();refresh()}"
    s = s & vbLf
    s = s & "function insertText(text){input.setRangeText(text,input.selectionStart,input.selectionEnd,'end');input.focus();refresh()}"
    s = s & vbLf
    s = s & "function scanRefs(text){const refs=[];const add=(section,field,index,syntax)=>{if(insideString(text.slice(0,index)))return;refs.push({section:normalSection(section)||String(section||'system').toLowerCase(),field:String(field||'').trim(),index,syntax})};let m,re=/@?\b([A-Za-z_]+)\.([A-Za-z0-9_]+)/g;while((m=re.exec(text)))add(m[1],m[2],m.index,'dot');re=/@([A-Za-z_]+)\(([^)]+)\)/g;while((m=re.exec(text)))add(m[1],m[2],m.index,'function');re=/\[([^\]]+)\]/g;while((m=re.exec(text))){const p=m[1].indexOf(':');if(p>=0)add(m[1].slice(0,p),m[1].slice(p+1),m.ind"
    s = s & "ex,'bracket');else add('system',m[1],m.index,'bracket')}return refs}"
    s = s & vbLf
    s = s & "function parenState(text){let depth=0,q=false;for(let i=0;i<text.length;i++){const c=text[i];if(c==='""'){if(q&&text[i+1]==='""')i++;else q=!q;continue}if(q)continue;if(c==='(')depth++;else if(c===')'){depth--;if(depth<0)return {ok:false,msg:'Closing parenthesis appears before a matching opening parenthesis.'}}}return depth===0?{ok:!q,msg:q?'A text string is not closed.':''}:{ok:false,msg:depth+' opening parenthesis'+(depth===1?' is':'es are')+' not closed.'}}"
    s = s & vbLf
    s = s & "function keywordAt(s,i,w){if(s.slice(i,i+w.length).toUpperCase()!==w)return false;const a=i? s[i-1]:' ',b=s[i+w.length]||' ';return !/[A-Za-z0-9_]/.test(a)&&!/[A-Za-z0-9_]/.test(b)}"
    s = s & vbLf
    s = s & "function findTopKeyword(s,w,start=0){let q=false,depth=0;for(let i=start;i<s.length;i++){const c=s[i];if(c==='""'){if(q&&s[i+1]==='""')i++;else q=!q;continue}if(q)continue;if(c==='(')depth++;else if(c===')')depth=Math.max(0,depth-1);else if(depth===0&&keywordAt(s,i,w))return i}return -1}"
    s = s & vbLf
    s = s & "function findNaturalElse(s,start){let q=false,depth=0,nested=0;for(let i=start;i<s.length;i++){const c=s[i];if(c==='""'){if(q&&s[i+1]==='""')i++;else q=!q;continue}if(q)continue;if(c==='(')depth++;else if(c===')')depth=Math.max(0,depth-1);if(depth!==0)continue;if(keywordAt(s,i,'IF')){let j=i+2;while(/\s/.test(s[j]||''))j++;if(s[j]!=='('){nested++;i++}}else if(keywordAt(s,i,'ELSE')){if(nested===0)return i;nested--;i+=3}}return -1}"
    s = s & vbLf
    s = s & "function translateNaturalIf(s){const t=String(s||'').trim();if(!/^IF\s+/i.test(t))return t;const then=findTopKeyword(t,'THEN',2);if(then<0)return t;const els=findNaturalElse(t,then+4);if(els<0)return t;const cond=t.slice(2,then).trim(),yes=t.slice(then+4,els).trim(),no=t.slice(els+4).trim();return 'IF('+normalizeLogic(cond)+','+normalizeLogic(translateNaturalIf(yes))+','+normalizeLogic(translateNaturalIf(no))+')'}"
    s = s & vbLf
    s = s & "function splitLogical(s,w){let q=false,depth=0,start=0,parts=[];for(let i=0;i<s.length;i++){const c=s[i];if(c==='""'){if(q&&s[i+1]==='""')i++;else q=!q;continue}if(q)continue;if(c==='(')depth++;else if(c===')')depth--;if(depth===0&&keywordAt(s,i,w)){const prev=s[i-1]||'',next=s[i+w.length]||'';if(/\s/.test(prev)&&/\s/.test(next)){parts.push(s.slice(start,i).trim());start=i+w.length;i+=w.length-1}}}if(parts.length){parts.push(s.slice(start).trim());return parts}return null}"
    s = s & vbLf
    s = s & "function splitArgsJs(s){let q=false,depth=0,start=0,parts=[];for(let i=0;i<s.length;i++){const c=s[i];if(c==='""'){if(q&&s[i+1]==='""')i++;else q=!q;continue}if(q)continue;if(c==='(')depth++;else if(c===')')depth--;else if(c===','&&depth===0){parts.push(s.slice(start,i).trim());start=i+1}}parts.push(s.slice(start).trim());return parts}"
    s = s & vbLf
    s = s & "function normalizeGroups(s){let out='',i=0,q=false;while(i<s.length){if(s[i]==='""'){let j=i+1;while(j<s.length){if(s[j]==='""'){if(s[j+1]==='""'){j+=2;continue}j++;break}j++}out+=s.slice(i,j);i=j;continue}if(s[i]==='('){let depth=1,j=i+1,qq=false;for(;j<s.length&&depth;j++){const c=s[j];if(c==='""'){if(qq&&s[j+1]==='""')j++;else qq=!qq;continue}if(!qq){if(c==='(')depth++;else if(c===')')depth--}}if(depth!==0){out+=s.slice(i);break}const inner=s.slice(i+1,j-1);let k=out.length-1;while(k>=0&&/\s/.test(out[k]))k--;let e=k;while(k>=0&&/[A-Za-z0-9_]/.test(out[k])"
    s = s & ")k--;const fn=out.slice(k+1,e+1);let processed;if(fn){processed=splitArgsJs(inner).map(normalizeLogic).join(',')}else processed=normalizeLogic(inner);out+='('+processed+')';i=j;continue}out+=s[i++]}return out}"
    s = s & vbLf
    s = s & "function normalizeLogic(s){let t=String(s||'').trim();if(!t)return t;t=translateNaturalIf(t);t=t.replace(/\bIFELSE\s*\(/gi,'IF(');let p=splitLogical(t,'OR');if(p)return 'OR('+p.map(normalizeLogic).join(',')+')';p=splitLogical(t,'AND');if(p)return 'AND('+p.map(normalizeLogic).join(',')+')';return normalizeGroups(t)}"
    s = s & vbLf
    s = s & "function translationPreview(){const raw=String(input.value||'').trim().replace(/^'/,'').replace(/^=/,'');let normalized='';try{normalized=normalizeLogic(raw)}catch(e){normalized=raw}const box=document.getElementById('translation');box.textContent=normalized||'(empty formula)';const changed=normalized!==raw;document.getElementById('translationState').textContent=changed?'Natural logic translated for Excel':'Already Excel-style / engine-ready';return normalized}"
    s = s & vbLf
    s = s & "function validate(){const text=input.value,refs=scanRefs(text),problems=[];translationPreview();const add=(level,title,detail)=>problems.push({level,title,detail});if(!text.trim())add('info','Formula is empty','No independent/manual formula will be generated from this cell.');if(text.includes('#REF!'))add('error','Legacy #REF! placeholder detected','The calculation engine intentionally leaves formulas containing #REF! blank.');const ps=parenState(text);if(!ps.ok)add('error','Formula structure needs attention',ps.msg);if(/^\s*IF\s+/i.test(text)&&(!/\bTHEN"
    s = s & "\b/i.test(text)||!/\bELSE\b/i.test(text)))add('error','Incomplete natural IF / ELSE','Natural form must contain IF, THEN and ELSE. You can alternatively use Excel form IF(condition,true,false).');if(text.length>8191)add('error','Formula is too long','Excel formulas are limited to 8,192 characters; current length is '+text.length+'.');else if(text.length>7600)add('warn','Approaching Excel formula limit','Current length is '+text.length+' characters.');const badSecs=[];for(const r of refs){if(!normalSection(r.section))badSecs.push(r.section)}if(badSecs.len"
    s = s & "gth)add('error','Unknown reference section','Check: '+[...new Set(badSecs)].join(', ')+'.');const unknown=[],disabled=[];for(const r of refs){const f=fieldMap.get(r.field.toLowerCase());if(!f)unknown.push(r.field);else if(!f.enabled)disabled.push(r.field);if(f){}const sec=normalSection(r.section);if(sec==='manual'&&r.field.toLowerCase()===String(data.label||'').toLowerCase())add('error','Direct manual self-reference','@manual.'+r.field+' points back to the manual cell being calculated and will create a circular reference.');if(sec==='difference'&&r.field"
    s = s & ".toLowerCase()===String(data.label||'').toLowerCase())add('error','Direct difference self-reference','The current-row difference depends on the manual result, so referring to it from this manual formula is circular.')}if(unknown.length)add('error','Unknown output field'+(new Set(unknown).size>1?'s':''),[...new Set(unknown)].join(', ')+' will use the engine zero fallback because no enabled row can be resolved.');if(disabled.length)add('warn','Disabled field'+(new Set(disabled).size>1?'s':'' )+' referenced',[...new Set(disabled)].join(', ')+' '+(new Set(di"
    s = s & "sabled).size>1?'are':'is')+' present in configuration but disabled; the generated row is absent and the engine falls back to 0.');if(refs.some(r=>normalSection(r.section)==='difference'))add('warn','Difference dependency in manual formula','Difference values are System - Manual. Cross-row difference references can be valid, but review dependency order to avoid circular chains.');if(/@?[A-Za-z_][A-Za-z0-9_]*\.\s*(?:$|[+\-*/,)])/m.test(text))add('error','Incomplete field reference','A section prefix is present without a field name.');if(/[+\-*/^=<>]\s*$/.t"
    FormulaEditorHtmlPart3 = s
End Function

Private Function FormulaEditorHtmlPart4() As String
    Dim s As String
    s = s & "est(text.trim()))add('warn','Formula ends with an operator','Complete the expression before copying it back to Excel.');const refKeys=new Set(refs.map(r=>normalSection(r.section)+'.'+r.field.toLowerCase()));renderHealth(problems);renderRefs(refs);document.getElementById('metrics').textContent=text.length+' chars | '+(text.split(/\r?\n/).length)+' line'+(text.split(/\r?\n/).length===1?'':'s')+' | '+refKeys.size+' ref'+(refKeys.size===1?'':'s');document.getElementById('dirty').classList.toggle('on',text!==String(data.formula||''));return problems}"
    s = s & vbLf
    s = s & "function renderHealth(problems){health.replaceChildren();const errors=problems.filter(x=>x.level==='error').length,warns=problems.filter(x=>x.level==='warn').length;healthPill.className='healthpill '+(errors?'blocked':warns?'review':'');healthPill.textContent=errors?'Blocked':warns?'Review':'Ready';const items=problems.length?problems:[{level:'ok',title:'Formula checks are clear',detail:'All recognized row references resolve to active fields and no structural issue is detected.'}];for(const p of items.slice(0,8)){const row=document.createElement('div');r"
    s = s & "ow.className='healthline '+(p.level==='error'?'error':p.level==='warn'?'warn':'');row.innerHTML='<span class=""hicon"">'+(p.level==='error'?'!':p.level==='warn'?'!':'OK')+'</span><span class=""healthtext""><b>'+esc(p.title)+'</b><span>'+esc(p.detail||'')+'</span></span>';health.append(row)}}"
    s = s & vbLf
    s = s & "function renderRefs(refs){const box=document.getElementById('refs');box.replaceChildren();const uniq=new Map();refs.forEach(r=>uniq.set(normalSection(r.section)+'.'+r.field.toLowerCase(),r));document.getElementById('refCount').textContent=uniq.size+' unique';if(!uniq.size){box.innerHTML='<span class=""empty"">No row references yet.</span>';return}for(const r of uniq.values()){const b=document.createElement('button');b.type='button';b.className='refchip';b.textContent='@'+(normalSection(r.section)||r.section)+'.'+r.field;const f=fieldMap.get(r.field.toLower"
    s = s & "Case());if(f)b.addEventListener('click',()=>{inspectField(f);document.getElementById('fieldSearch').value=f.name;paintLibrary()});box.append(b)}}"
    s = s & vbLf
    s = s & "function refresh(){const s=suggestionsFor(input.value,input.selectionStart);activeToken=s.token;options=s.items;selected=0;dismissed=false;paintSuggestions();const p=validate();matches.textContent=activeToken?(options.length?options.length+' smart match'+(options.length===1?'':'es')+(options.length>80?' | first 80 shown':''):(activeToken.kind==='badsection'?'Unknown reference section':'No matching suggestions')):(fields.filter(f=>f.enabled).length+' active fields available | type @m then . or use the field library');if(options[0]?.field)inspectField(opti"
    s = s & "ons[0].field);status.className='status'+(p.some(x=>x.level==='error')?' bad':'')}"
    s = s & vbLf
    s = s & "function paintLibrary(){const q=document.getElementById('fieldSearch').value.trim(),showOff=document.getElementById('showDisabled').checked,box=document.getElementById('library');box.replaceChildren();let arr=fieldOptions(q,showOff).slice(0,80);document.getElementById('libraryCount').textContent=arr.length+(arr.length===80?'+':'')+' shown';if(!arr.length){box.innerHTML='<div class=""empty"">No fields match this search.</div>';return}for(const o of arr){const f=o.field,row=document.createElement('div');row.className='librow';row.innerHTML='<div class=""libto"
    s = s & "p""><span class=""libname"">'+esc(f.name)+'</span>'+(f.enabled?'<span class=""tag"">active</span>':'<span class=""tag off"">disabled</span>')+'</div><div class=""libmeta"">'+esc((f.group||'Ungrouped')+(f.source?' | source: '+f.source:'')+(f.order?' | order '+f.order:''))+'</div><div class=""libbuttons""><button data-s=""base"">@base</button><button data-s=""system"">@system</button><button data-s=""manual"">@manual</button><button data-s=""difference"">@diff</button></div>';row.querySelector('.libtop').addEventListener('click',()=>inspectField(f));row.querySelectorAll('[da"
    s = s & "ta-s]').forEach(b=>b.addEventListener('click',()=>{insertText('@'+b.dataset.s+'.'+f.name);inspectField(f)}));box.append(row)}}"
    s = s & vbLf
    s = s & "function paintVariants(){const box=document.getElementById('variants');box.replaceChildren();const v=data.variants||{};[['default','Default'],['moderate','Moderate'],['medium','Medium'],['severe','Severe']].forEach(([k,label])=>{const value=String(v[k]||''),d=document.createElement('div');d.className='variant'+(value?'':' empty');d.innerHTML='<div class=""varianttop""><b>'+label+'</b>'+(value?'<button class=""btn small"" type=""button"">Use</button>':'')+'</div><code>'+(value?esc(value):'No formula configured')+'</code>';if(value)d.querySelector('button').addE"
    s = s & "ventListener('click',()=>{input.value=value;input.focus();input.setSelectionRange(input.value.length,input.value.length);status.textContent=label+' variant loaded into the canvas. Original opening formula can still be restored.';refresh()});box.append(d)})}"
    s = s & vbLf
    s = s & "function functionByName(n){return functions.find(f=>f.name===n)}"
    s = s & vbLf
    s = s & "function excelText(s){return ""'""+String(s||'').trim().replace(/^'/,'').replace(/^=/,'')}"
    s = s & vbLf
    s = s & "async function copyFormula(){const problems=validate();if(problems.some(x=>x.level==='error')){status.className='status bad';status.textContent='Formula has blocking checks. You can still copy after reviewing, but resolve errors first where possible.'}const value=excelText(input.value);let ok=false;try{if(navigator.clipboard){await navigator.clipboard.writeText(value);ok=true}}catch(e){}if(!ok){const t=document.createElement('textarea');t.value=value;t.style.position='fixed';t.style.left='-10000px';document.body.append(t);t.select();try{ok=document.execC"
    s = s & "ommand('copy')}catch(e){}t.remove()}if(ok){status.className='status';status.textContent='Copied as Config-safe text. Paste into '+data.target+'. The engine translates it to a normal Excel formula when outputs are generated.';input.focus()}else{input.value=value;input.focus();input.select();status.textContent='Clipboard access was blocked. Press Ctrl+C on the selected text, then paste it into '+data.target+'.'}}"
    s = s & vbLf
    s = s & "input.addEventListener('input',()=>{status.textContent='';refresh()});input.addEventListener('click',refresh);input.addEventListener('keyup',e=>{if(['ArrowLeft','ArrowRight','Home','End','PageUp','PageDown'].includes(e.key))refresh()});input.addEventListener('keydown',e=>{if(e.ctrlKey&&e.key==='Enter'){e.preventDefault();copyFormula();return}if(e.ctrlKey&&e.code==='Space'){e.preventDefault();const s=tokenAt(input.value,input.selectionStart);if(!s){const pos=input.selectionStart;input.setRangeText('@manual.',pos,pos,'end')}refresh();return}if(e.key==='Esc"
    s = s & "ape'){dismissed=true;paintSuggestions();return}if(suggestBox.hidden)return;const n=Math.min(options.length,80);if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();selected=(selected+(e.key==='ArrowDown'?1:-1)+n)%n;highlightSuggestion();list.children[selected]?.scrollIntoView({block:'nearest'})}else if(e.key==='Enter'||e.key==='Tab'){e.preventDefault();choose(selected)}});"
    s = s & vbLf
    s = s & "document.querySelectorAll('[data-section]').forEach(b=>b.addEventListener('click',()=>{insertText('@'+b.dataset.section+'.')}));document.querySelectorAll('[data-insert]').forEach(b=>b.addEventListener('click',()=>insertText(b.dataset.insert)));document.querySelectorAll('[data-fn]').forEach(b=>b.addEventListener('click',()=>{const f=functionByName(b.dataset.fn);insertSnippet(f.snippet,f.placeholder)}));document.querySelectorAll('[data-template]').forEach(b=>b.addEventListener('click',()=>{const t=templates[b.dataset.template];insertSnippet(t.text,t.select"
    s = s & ")}));"
    s = s & vbLf
    s = s & "document.getElementById('normalizeBtn').addEventListener('click',()=>{const n=translationPreview();input.value=n;input.focus();input.setSelectionRange(input.value.length,input.value.length);status.className='status';status.textContent='Natural IF / ELSE and infix AND / OR were converted to engine-ready Excel-style logic.';refresh()});window.addEventListener('resize',()=>{if(!suggestBox.hidden)positionSuggestionBox()});"
    s = s & vbLf
    s = s & "document.getElementById('copy').addEventListener('click',copyFormula);document.getElementById('reset').addEventListener('click',()=>{input.value=String(data.formula||'');input.focus();input.setSelectionRange(input.value.length,input.value.length);status.className='status';status.textContent='Starting formula restored.';refresh()});document.getElementById('wrap').addEventListener('click',e=>{const currentlyOff=input.getAttribute('wrap')==='off';input.setAttribute('wrap',currentlyOff?'soft':'off');e.currentTarget.textContent=currentlyOff?'Do not wrap':'Wra"
    s = s & "p lines';input.focus()});document.getElementById('fieldSearch').addEventListener('input',paintLibrary);document.getElementById('showDisabled').addEventListener('change',paintLibrary);"
    s = s & vbLf
    s = s & "document.querySelectorAll('.tab').forEach(b=>b.addEventListener('click',()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.tabpane').forEach(x=>x.classList.toggle('active',x.id===b.dataset.tab));if(b.dataset.tab==='fieldsPane')document.getElementById('fieldSearch').focus()}));"
    s = s & vbLf
    s = s & "document.getElementById('ctxElement').textContent=data.elementType||'-';document.getElementById('ctxLabel').textContent=data.label||'-';document.getElementById('ctxMode').textContent=data.mode||'-';document.getElementById('ctxTarget').textContent=data.target||'-';document.getElementById('modeTop').textContent=data.mode||'Default';document.getElementById('activeCount').textContent=fields.filter(f=>f.enabled).length;document.getElementById('disabledCount').textContent=fields.filter(f=>!f.enabled).length;document.getElementById('rowMeta').textContent=(data."
    s = s & "sourceField?'Source: '+data.sourceField+' | ':'')+(data.group?'Group: '+data.group+' | ':'')+'Edit the '+(data.mode||'selected')+' formula for this output row.';"
    s = s & vbLf
    s = s & "input.value=String(data.formula||'');input.focus();input.setSelectionRange(input.value.length,input.value.length);paintLibrary();paintVariants();const currentField=fieldMap.get(String(data.label||'').toLowerCase());if(currentField)inspectField(currentField);refresh();dismissed=true;paintSuggestions();"
    s = s & vbLf
    s = s & "window.JKBEditor={tokenAt,suggestionsFor,scanRefs,validate,excelText,subseqScore,fieldOptions,normalSection,normalizeLogic,translateNaturalIf,translationPreview};"
    s = s & vbLf
    s = s & "</script></body></html>"
    s = s & vbLf
    s = s & vbLf
    FormulaEditorHtmlPart4 = s
End Function

Public Sub StyleWorkbookUI()
    Dim ws As Worksheet
    On Error Resume Next
    StyleDashboard
    Set ws = Me.Parent.Worksheets("Config_ElementTypeRules"): If Not ws Is Nothing Then StyleRulesSheet ws
    Set ws = Me.Parent.Worksheets("Config_MasterFormatting"): If Not ws Is Nothing Then StyleFormattingSheet ws
    
    Set ws = Nothing: Set ws = Me.Parent.Worksheets("Scenario Element Output"): If Not ws Is Nothing Then StyleSourceSheet ws
    Set ws = Nothing: Set ws = Me.Parent.Worksheets("Build_Log"): If Not ws Is Nothing Then StyleBuildLog ws
    Set ws = Nothing: Set ws = Me.Parent.Worksheets("Manual_Formula_Guide"): If Not ws Is Nothing Then StyleGuide ws
    GroupSheetsByPurpose Me.Parent
    On Error GoTo 0
End Sub

' Colours and orders the tabs by what each sheet is FOR, so a workbook with a dozen tabs
' reads as four short groups instead of one undifferentiated row:
'
'   navy    start here          Dashboard
'   teal    what goes in        the uploaded testcase output
'   green   independent check   pre-shock setup, and the derived base / pre-shock values
'   amber   configuration       rules, formatting, the formula guide
'   grey    diagnostics         the build log
'
' Element sheets in a generated workbook keep the green "checked" colour, since that is what
' they are: the system's numbers beside the independently rebuilt ones.
Public Sub GroupSheetsByPurpose(ByVal wb As Workbook)
    Dim ws As Worksheet, rank As Long, ordered As Object
    On Error Resume Next
    Set ordered = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            rank = SheetPurposeRank(ws)
            ws.Tab.Color = SheetPurposeColor(rank)
            If Not ordered.Exists(CStr(rank)) Then ordered(CStr(rank)) = ""
            ordered(CStr(rank)) = ordered(CStr(rank)) & ws.name & vbTab
        End If
    Next ws
    ' Move the visible sheets into group order, each one placed directly after the one before
    ' it. "After the previous sheet" is unambiguous; moving to an absolute index is not, once
    ' hidden sheets sit in between - that left the formula guide stranded past the build log.
    Dim prev As Worksheet, nm As Variant
    Set prev = Nothing
    For rank = 0 To 5
        If ordered.Exists(CStr(rank)) Then
            For Each nm In Split(ordered(CStr(rank)), vbTab)
                If Len(CStr(nm)) > 0 Then
                    Set ws = Nothing
                    Set ws = wb.Worksheets(CStr(nm))
                    If Not ws Is Nothing Then
                        If prev Is Nothing Then
                            If ws.index > 1 Then ws.Move Before:=wb.Worksheets(1)
                        Else
                            ws.Move After:=prev
                        End If
                        Set prev = ws
                    End If
                End If
            Next nm
        End If
    Next rank
    Err.Clear
End Sub

Private Function SheetPurposeRank(ByVal ws As Worksheet) As Long
    Dim n As String
    n = ws.name
    If n = "Dashboard" Or n = "Recon_Workbench" Then SheetPurposeRank = 0: Exit Function
    If n = "Scenario Element Output" Then SheetPurposeRank = 1: Exit Function
    If IsPreShockSheet(ws) Then SheetPurposeRank = 2: Exit Function
    If IsElementSheet(ws) Then SheetPurposeRank = 2: Exit Function
    If n = "Config_ElementTypeRules" Or n = "Config_MasterFormatting" Or n = "Config_ValueSources" Then SheetPurposeRank = 3: Exit Function
    If n = "Manual_Formula_Guide" Or n = "Format" Then SheetPurposeRank = 3: Exit Function
    If n = "Build_Log" Then SheetPurposeRank = 5: Exit Function
    SheetPurposeRank = 4
End Function

Private Function SheetPurposeColor(ByVal rank As Long) As Long
    Select Case rank
        Case 0: SheetPurposeColor = RGB(18, 48, 107)      ' start here
        Case 1: SheetPurposeColor = RGB(14, 116, 144)    ' what goes in
        Case 2: SheetPurposeColor = RGB(6, 118, 71)     ' independent check
        Case 3: SheetPurposeColor = RGB(176, 138, 46)    ' configuration
        Case 5: SheetPurposeColor = RGB(102, 112, 133)   ' diagnostics
        Case Else: SheetPurposeColor = RGB(71, 84, 103)
    End Select
End Function

Public Sub StyleDashboard()
    Dim ws As Worksheet, r As Long, isMaster As Boolean, sh As Shape
    Set ws = Me
    isMaster = (UCase$(Trim$(CStr(ws.Range("N1").Value2))) = "MASTER")
    On Error Resume Next
    ws.Range("D3:K3").UnMerge: ws.Range("D4:K4").UnMerge: ws.Range("D6:K6").UnMerge
    ws.Range("D15:K15").UnMerge: ws.Range("D16:K16").UnMerge
    ws.Range("D3:K3").Merge: ws.Range("D4:K4").Merge: ws.Range("D6:K6").Merge
    ws.Range("D15:K15").Merge: ws.Range("D16:K16").Merge
    For r = 18 To 28
        ws.Range("F" & r & ":K" & r).UnMerge
        ws.Range("F" & r & ":K" & r).Merge
    Next r
    On Error GoTo 0
    With ws.Range("A1:K60")
        .Interior.Color = RGB(255, 255, 255)
        .Font.name = "Aptos": .Font.Size = 9.5: .Font.Color = RGB(16, 24, 40)
    End With
    If isMaster Then
        ws.Range("D3").Value2 = "JKB Stress Testing Reconciliation"
        ws.Range("D4").Value2 = "Independent controls for testcase upload, base validation and reconciled output generation"
        ws.Range("B6").Value2 = "WORKSPACE"
        ws.Range("D6").Value2 = "Run a reconciliation"
        ws.Range("D15").Value2 = "Source health and latest activity"
        ws.Range("D16").Value2 = "Current workspace status"
    Else
        ws.Range("D3").Value2 = "JKB Stress Testing Reconciliation"
        ws.Range("D4").Value2 = "Generated review workbook  |  system output vs independent / manual check"
        ws.Range("B6").Value2 = "REVIEW"
        ws.Range("D6").Value2 = "Testcase reconciliation"
        ws.Range("D15").Value2 = "Generated run information"
        ws.Range("D16").Value2 = "Run details"
    End If
    ' Masthead: solid band, white type, gold rule. Reads as a product rather than a form.
    ' B3:C4 is left clear for the brand logo, which floats above the sheet.
    With ws.Range("D3:K3")
        .Interior.Color = RGB(14, 34, 64): .Font.Color = vbWhite
        .Font.Size = 17: .Font.Bold = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With ws.Range("D4:K4")
        .Interior.Color = RGB(14, 34, 64): .Font.Color = RGB(183, 196, 216)
        .Font.Size = 9.5: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 2
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46): .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.rows(3).RowHeight = 30: ws.rows(4).RowHeight = 19: ws.rows(5).RowHeight = 8
    With ws.Range("B6:C15")
        .Interior.Color = RGB(245, 247, 250)
        .Borders.Color = RGB(234, 236, 240): .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    With ws.Range("B6:C6")
        .Interior.Color = RGB(238, 243, 251): .Font.Color = RGB(18, 48, 107)
        .Font.Bold = True: .Font.Size = 8.5: .HorizontalAlignment = xlCenter
    End With
    With ws.Range("D6:K6")
        .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(18, 48, 107): .Font.Size = 13: .Font.Bold = True
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46): .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("D15:K15")
        .Font.Color = RGB(102, 112, 133): .Font.Size = 8.5: .Font.Italic = True
        .HorizontalAlignment = xlLeft
    End With
    With ws.Range("D16:K16")
        .Interior.Color = RGB(238, 243, 251): .Font.Color = RGB(18, 48, 107)
        .Font.Bold = True: .Font.Size = 10: .HorizontalAlignment = xlLeft: .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RGB(214, 226, 245)
    End With
    If isMaster Then
        ws.Range("D18").Value2 = "Source workbook"
        ws.Range("D19").Value2 = "Detected sheet"
        ws.Range("D20").Value2 = "Uploaded at"
        ws.Range("D21").Value2 = "Source rows"
        ws.Range("D22").Value2 = "Test cases"
        ws.Range("D23").Value2 = "As-of dates"
        ws.Range("D24").Value2 = "Entities"
        ws.Range("D25").Value2 = "Status"
        ws.Range("D26").Value2 = "ECL base testing"
        ws.Range("D27").Value2 = "Last output folder"
        ws.Range("D28").Value2 = "Last run result"
    End If
    For r = 18 To 28
        With ws.Range("D" & r & ":E" & r)
            .Interior.Color = RGB(245, 247, 250): .Font.Color = RGB(52, 64, 84)
            .Font.Bold = True: .Font.Size = 8.5: .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RGB(234, 236, 240)
        End With
        With ws.Range("F" & r & ":K" & r)
            .Interior.Color = vbWhite: .Font.Color = RGB(16, 24, 40)
            .Font.Size = 9: .WrapText = True: .VerticalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Color = RGB(234, 236, 240)
        End With
        ws.rows(r).RowHeight = IIf(r = 25 Or r = 27 Or r = 28, 21, 18)
    Next r
    ws.rows("6:16").RowHeight = 18
    ws.rows(8).RowHeight = 20: ws.rows(9).RowHeight = 20: ws.rows(10).RowHeight = 20
    ws.rows(12).RowHeight = 20: ws.rows(13).RowHeight = 20: ws.rows(14).RowHeight = 20
    If Len(CStr(ws.Range("D31").Value2)) > 0 Then
        With ws.Range("D31:G31")
            .Interior.Color = RGB(238, 243, 251): .Font.Color = RGB(18, 48, 107): .Font.Bold = True
            .HorizontalAlignment = xlLeft: .Borders.Color = RGB(214, 226, 245)
        End With
        ws.Range("D32:G200").Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
    End If
    ws.columns("A").ColumnWidth = 2.5: ws.columns("B:C").ColumnWidth = 11.5
    ws.columns("D").ColumnWidth = 22: ws.columns("E").ColumnWidth = 3
    ws.columns("F").ColumnWidth = 23: ws.columns("G:H").ColumnWidth = 14
    ws.columns("I:J").ColumnWidth = 11.5: ws.columns("K").ColumnWidth = 4.5
    On Error Resume Next
    Set sh = ws.Shapes("JKB_BrandLogo")
    If Not sh Is Nothing Then
        sh.LockAspectRatio = msoTrue
        sh.Left = ws.Range("B3").Left + 3: sh.Top = ws.Range("B3").Top + 1
        sh.Height = ws.Range("B3:C4").Height - 2
        If sh.Width > ws.Range("B3:C4").Width - 6 Then sh.Width = ws.Range("B3:C4").Width - 6
        sh.Placement = xlFreeFloating
    End If
    On Error GoTo 0
    If isMaster Then
        On Error Resume Next
        ws.Range("F26").Value2 = PreShockHomeStatus()
        ws.ScrollArea = "A1:K28"
        On Error GoTo 0
    End If
    EnsureRibbon ws
    ProStyleSheet Me.Parent.Worksheets("Dashboard")
End Sub

Public Sub StyleRulesSheet(ByVal ws As Worksheet)
    Dim lo As ListObject
    On Error Resume Next
    ws.Range("A3:U3").UnMerge: ws.Range("A4:U4").UnMerge
    ws.Range("A3:U3").Merge: ws.Range("A4:U4").Merge
    On Error GoTo 0
    With ws.Range("A1:AE1964")
        .Font.name = "Aptos": .Font.Size = 9
    End With
    With ws.Range("A3:U3")
        .Interior.Color = RGB(6, 118, 71): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 16
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    With ws.Range("A4:U4")
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(71, 84, 103): .Font.Size = 9.5
        .WrapText = True: .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    ws.rows(3).RowHeight = 28: ws.rows(4).RowHeight = 34: ws.rows(6).RowHeight = 32
    Set lo = ws.ListObjects("tblElementTypeRules")
    If Not lo Is Nothing Then lo.TableStyle = "TableStyleMedium2"
    ws.columns("A").ColumnWidth = 36: ws.columns("B").ColumnWidth = 11
    ws.columns("C:E").ColumnWidth = 12: ws.columns("F:G").ColumnWidth = 31
    ws.columns("H:I").ColumnWidth = 17: ws.columns("J:M").ColumnWidth = 42
    ws.columns("N:S").ColumnWidth = 17: ws.columns("T").ColumnWidth = 38: ws.columns("U").ColumnWidth = 22
    ws.Range("J7:M1964").WrapText = True: ws.Range("T7:T1964").WrapText = True
    With ws.Range("V4:W4")
        .Interior.Color = RGB(249, 250, 251): .Font.Color = RGB(71, 84, 103): .Font.Bold = True
    End With
    With ws.Range("V6:W6")
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True
    End With
    With ws.Range("Y3:AE3")
        .Interior.Color = RGB(6, 118, 71): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 11
    End With
    With ws.Range("Y4:AE4")
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(71, 84, 103): .Font.Size = 9
    End With
    With ws.Range("Y6:AE6")
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True: .WrapText = True
    End With
    ws.columns("V").ColumnWidth = 34: ws.columns("W").ColumnWidth = 18: ws.columns("X").ColumnWidth = 3
    ws.columns("Y").ColumnWidth = 20: ws.columns("Z").ColumnWidth = 31: ws.columns("AA").ColumnWidth = 12
    ws.columns("AB").ColumnWidth = 20: ws.columns("AC:AD").ColumnWidth = 16: ws.columns("AE").ColumnWidth = 13
    On Error Resume Next
    ws.Range("U7:U1964").Validation.Delete
    ws.Range("U7:U1964").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$Y$7:$Y$20"
    ws.Range("U7:U1964").Validation.InCellDropdown = True
    On Error GoTo 0
    ProStyleSheet ws
End Sub

Public Sub StyleFormattingSheet(ByVal ws As Worksheet)
    Dim r As Long
    ' Professional output palette: neutral labels/base, white system, soft green manual, soft yellow difference.
    With ws.Range("A1:AB60")
        .Font.name = "Aptos": .Font.Size = 9.5
    End With
    With ws.Range("A3:G3")
        .Interior.Color = RGB(6, 118, 71): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 16
    End With
    With ws.Range("A4:G4")
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(71, 84, 103): .WrapText = True
    End With
    With ws.Range("A5:B5")
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True
    End With
    With ws.Range("F5:G6")
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True
    End With
    ' Accounting 2 is the standard amount format: comma separators, no decimals, red bracketed negatives.
    ws.Range("G14").NumberFormat = "#,##0;[Red](#,##0);-"
    ' Keep variance 2 visually consistent when used for amount-style comparisons.
    ws.Range("G15").NumberFormat = "#,##0;[Red](#,##0);-"
    With ws.Range("C33:C34,C39,C43")
        .Interior.Color = RGB(234, 236, 240): .Font.Color = RGB(16, 24, 40): .Font.Bold = True
    End With
    With ws.Range("C35,C39,C43")
        .Interior.Color = RGB(234, 236, 240): .Font.Color = RGB(16, 24, 40): .Font.Bold = True
    End With
    With ws.Range("C36,C40,C44")
        .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(16, 24, 40): .Font.Bold = False
    End With
    With ws.Range("C37,C41,C45")
        .Interior.Color = RGB(236, 253, 243): .Font.Color = RGB(16, 24, 40): .Font.Bold = False
    End With
    With ws.Range("C38,C42,C46")
        .Interior.Color = RGB(253, 246, 227): .Font.Color = RGB(16, 24, 40): .Font.Bold = False
    End With
    ws.columns("A").ColumnWidth = 34: ws.columns("B").ColumnWidth = 18
    ws.columns("C").ColumnWidth = 47: ws.columns("F").ColumnWidth = 26: ws.columns("G").ColumnWidth = 20
    If CStr(ws.Range("AB1").Value2) <> "JKB_UI_V5" Then
        SectionSample ws.Range("C48"), "Context & identifiers", RGB(242, 244, 247), RGB(71, 84, 103): ws.Range("Z48").Value2 = "SECTION_CONTEXT"
        SectionSample ws.Range("C49"), "Shock & scenario inputs", RGB(242, 244, 247), RGB(6, 118, 71): ws.Range("Z49").Value2 = "SECTION_SHOCK_INPUT"
        SectionSample ws.Range("C50"), "Pre-shock IFRS 9", RGB(238, 243, 251), RGB(52, 64, 84): ws.Range("Z50").Value2 = "SECTION_PRE_IFRS9"
        SectionSample ws.Range("C51"), "Ratios & drivers", RGB(236, 253, 255), RGB(71, 84, 103): ws.Range("Z51").Value2 = "SECTION_RATIOS"
        SectionSample ws.Range("C52"), "Shock / impact", RGB(253, 246, 227), RGB(176, 138, 46): ws.Range("Z52").Value2 = "SECTION_SHOCK"
        SectionSample ws.Range("C53"), "Post-shock results", RGB(236, 253, 243), RGB(6, 118, 71): ws.Range("Z53").Value2 = "SECTION_POST_SHOCK"
        SectionSample ws.Range("C54"), "Pre-shock capital / RWA", RGB(238, 243, 251), RGB(18, 48, 107): ws.Range("Z54").Value2 = "SECTION_CAPITAL"
        SectionSample ws.Range("C55"), "Neutral section", RGB(249, 250, 251), RGB(71, 84, 103): ws.Range("Z55").Value2 = "SECTION_NEUTRAL"
        ws.Range("AB1").Value2 = "JKB_UI_V6"
    End If
    For r = 48 To 55: ws.rows(r).RowHeight = 22: Next r
    ProStyleSheet ws
End Sub


Private Sub SectionSample(ByVal c As Range, ByVal caption As String, ByVal fillColor As Long, ByVal fontColor As Long)
    c.Value2 = caption
    c.Interior.Color = fillColor: c.Font.Color = fontColor: c.Font.Bold = True: c.Font.Size = 10
    c.HorizontalAlignment = xlLeft: c.VerticalAlignment = xlCenter: c.IndentLevel = 1
    c.Borders(xlEdgeBottom).LineStyle = xlContinuous: c.Borders(xlEdgeBottom).Color = RGB(208, 213, 221)
End Sub




Public Sub StyleSourceSheet(ByVal ws As Worksheet)
    Dim r As Long, c As Long, lastCol As Long, foundA As Boolean, foundE As Boolean, hr As Long
    On Error GoTo Done
    For r = 1 To WorksheetFunction.Min(50, ws.UsedRange.rows.count + ws.UsedRange.row - 1)
        lastCol = ws.Cells(r, ws.columns.count).End(xlToLeft).Column
        foundA = False: foundE = False
        For c = 1 To lastCol
            If UCase$(Trim$(CStr(ws.Cells(r, c).Value2))) = "AS_OF_DATE" Then foundA = True
            If UCase$(Trim$(CStr(ws.Cells(r, c).Value2))) = "ELEMENT_TYPE" Then foundE = True
        Next c
        If foundA And foundE Then hr = r: Exit For
    Next r
    If hr = 0 Then Exit Sub
    With ws.rows(hr)
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 28
    End With
    ws.rows(hr).WrapText = True
    If lastCol > 0 Then ws.Range(ws.Cells(hr, 1), ws.Cells(hr, lastCol)).Borders(xlEdgeBottom).Color = RGB(6, 118, 71)
Done:
    ProStyleSheet ws
End Sub

Public Sub StyleBuildLog(ByVal ws As Worksheet)
    Dim hr As Long, r As Long, lastRow As Long
    If UCase$(Trim$(CStr(ws.Range("A3").Value2))) = "LEVEL" Then
        hr = 3
    ElseIf UCase$(Trim$(CStr(ws.Range("A1").Value2))) = "LEVEL" Then
        hr = 1
    End If
    If hr = 0 Then Exit Sub
    With ws.Range("A" & hr & ":E" & hr)
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 24
    End With
    ws.columns("A").ColumnWidth = 10: ws.columns("B").ColumnWidth = 24: ws.columns("C").ColumnWidth = 52
    ws.columns("D").ColumnWidth = 78: ws.columns("E").ColumnWidth = 22
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow > hr Then ws.Range("C" & (hr + 1) & ":D" & lastRow).WrapText = True
    ProStyleSheet ws
End Sub

Public Sub StyleGuide(ByVal ws As Worksheet)
    With ws.Range("A1:E1")
        .Interior.Color = RGB(6, 118, 71): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 15
    End With
    With ws.Range("A3:E3")
        .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True: .WrapText = True
    End With
    ws.Range("A4:E100").Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
    ws.Range("A1:E100").Font.name = "Aptos"
    ProStyleSheet ws
End Sub

Public Sub StyleSelectionSheet(ByVal ws As Worksheet)
    Dim col As Variant, lastRow As Long, flags As Range
    With ws.Range("A3:L3")
        .Interior.Color = RGB(6, 118, 71): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 16
    End With
    With ws.Range("A4:L4")
        .Interior.Color = RGB(242, 244, 247): .Font.Color = RGB(71, 84, 103): .WrapText = True
    End With
    For Each col In Array(1, 4, 7, 10)
        With ws.Cells(6, CLng(col)).Resize(1, 2)
            .Interior.Color = RGB(71, 84, 103): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 24
        End With
        lastRow = ws.Cells(ws.rows.count, CLng(col) + 1).End(xlUp).row
        If lastRow >= 7 Then
            Set flags = ws.Cells(7, CLng(col)).Resize(lastRow - 6, 1)
            flags.FormatConditions.Delete
            flags.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Yes"""
            flags.FormatConditions(1).Interior.Color = RGB(236, 253, 243): flags.FormatConditions(1).Font.Color = RGB(6, 118, 71)
            flags.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""No"""
            flags.FormatConditions(2).Interior.Color = RGB(242, 244, 247): flags.FormatConditions(2).Font.Color = RGB(102, 112, 133)
        End If
    Next col
End Sub

Public Sub FormatDetailGroups(ByVal ws As Worksheet)
    Dim defs As Object, gd As Object, lastRow As Long, r As Long, startRow As Long, endRow As Long, headerRow As Long
    Dim gkey As String, nextKey As String, lvl As Long, i As Long
    Set defs = DetailGroupDefinitions()
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 7 Then Exit Sub
    On Error Resume Next: ws.Cells.ClearOutline: On Error GoTo 0
    ws.Outline.SummaryRow = xlAbove
    r = 7
    Do While r <= lastRow
        headerRow = 0
        If CStr(ws.Cells(r, 12).Value2) = "SECTION_HEADER" Then
            headerRow = r: gkey = UCase$(Trim$(CStr(ws.Cells(r, 13).Value2)))
            r = r + 1
        Else
            gkey = UCase$(Trim$(CStr(ws.Cells(r, 13).Value2)))
        End If
        If Len(gkey) = 0 Then gkey = "UNASSIGNED"
        startRow = r
        Do While r <= lastRow
            If CStr(ws.Cells(r, 12).Value2) = "SECTION_HEADER" Then Exit Do
            nextKey = UCase$(Trim$(CStr(ws.Cells(r, 13).Value2)))
            If Len(nextKey) = 0 Then nextKey = "UNASSIGNED"
            If nextKey <> gkey Then Exit Do
            r = r + 1
        Loop
        endRow = r - 1
        If defs.Exists(gkey) Then Set gd = defs(gkey) Else Set gd = defs("UNASSIGNED")
        If headerRow > 0 Then
            With ws.Range("A" & headerRow & ":K" & headerRow)
                .Style = "JKB_" & CStr(gd("StyleKey"))
                .Font.Bold = True: .Font.Size = 10: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
                .IndentLevel = 1
            End With
            ws.rows(headerRow).RowHeight = CDbl(ViewSetting("Section header height", 21))
        End If
        If startRow <= endRow And UCase$(CStr(ViewSetting("Apply outline groups", "Yes"))) = "YES" Then
            lvl = CLng(gd("OutlineLevel")): If lvl < 1 Then lvl = 1: If lvl > 3 Then lvl = 3
            For i = 1 To lvl: ws.rows(startRow & ":" & endRow).rows.Group: Next i
            If CBool(gd("Collapsed")) Then ws.rows(startRow & ":" & endRow).hidden = True
        End If
    Loop
End Sub

Private Function DetailGroupDefinitions() As Object
    Dim d As Object, item As Object, ws As Worksheet, r As Long, lastRow As Long, key As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = Me.Parent.Worksheets("Config_ElementTypeRules")
    lastRow = ws.Cells(ws.rows.count, 25).End(xlUp).row
    For r = 7 To lastRow
        key = UCase$(Trim$(CStr(ws.Cells(r, 25).Value2)))
        If Len(key) > 0 Then
            Set item = CreateObject("Scripting.Dictionary"): item.CompareMode = vbTextCompare
            item("Label") = CStr(ws.Cells(r, 26).Value2)
            If IsNumeric(ws.Cells(r, 27).Value2) Then item("Order") = CLng(ws.Cells(r, 27).Value2) Else item("Order") = 900
            item("StyleKey") = CStr(ws.Cells(r, 28).Value2): If Len(item("StyleKey")) = 0 Then item("StyleKey") = "SECTION_NEUTRAL"
            item("Collapsed") = IsTrueValue(ws.Cells(r, 29).Value2)
            item("ShowHeader") = Not (UCase$(Trim$(CStr(ws.Cells(r, 30).Value2))) = "NO")
            If IsNumeric(ws.Cells(r, 31).Value2) Then item("OutlineLevel") = CLng(ws.Cells(r, 31).Value2) Else item("OutlineLevel") = 1
            If d.Exists(key) Then d.Remove key
            d.Add key, item
        End If
    Next r
    ' No catch-all "Other" group: a row only reaches this fallback when the run was allowed
    ' to continue with unassigned rows, and it is labelled so nobody mistakes it for a group.
    If d.Exists("UNASSIGNED") Then d.Remove "UNASSIGNED"
    Set item = CreateObject("Scripting.Dictionary")
    item("Label") = "Needs a group - set it in Config": item("Order") = 99999: item("StyleKey") = "SECTION_NEUTRAL"
    item("Collapsed") = False: item("ShowHeader") = True: item("OutlineLevel") = 1
    d.Add "UNASSIGNED", item
    Set DetailGroupDefinitions = d
End Function

Private Function IsTrueValue(ByVal v As Variant) As Boolean
    Select Case UCase$(Trim$(CStr(v)))
        Case "Y", "YES", "TRUE", "1", "ON": IsTrueValue = True
    End Select
End Function

Private Function CountDetailGroups(ByVal ws As Worksheet, ByRef anyHidden As Boolean) As Long
    Dim lastRow As Long, r As Long, g As String, prev As String
    anyHidden = False
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 7 Then Exit Function
    For r = 7 To lastRow
        g = UCase$(Trim$(CStr(ws.Cells(r, 13).Value2)))
        If Len(g) > 0 And g <> prev Then CountDetailGroups = CountDetailGroups + 1: prev = g
        If Len(g) > 0 And CStr(ws.Cells(r, 12).Value2) <> "SECTION_HEADER" Then If ws.rows(r).hidden Then anyHidden = True
    Next r
End Function

Public Sub ToggleDetailGroups()
    Dim ws As Worksheet, n As Long, anyHidden As Boolean
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set ws = ActiveSheet
    n = CountDetailGroups(ws, anyHidden): If n = 0 Then Exit Sub
    On Error GoTo Failed
    If anyHidden Then ws.Outline.ShowLevels RowLevels:=8 Else ws.Outline.ShowLevels RowLevels:=1
    EnsureRibbon ws
    Exit Sub
Failed:
    MsgBox "Could not change detail-group visibility." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub

' ============================================================================
'  THE VISUAL DESIGN.
'
'  Per-sheet layout: exact column widths and row heights, merged label and value
'  pairs, a tinted side panel, hairline rules, gridlines off. One palette, one
'  typeface, thirteen sheets that look like one product rather than thirteen
'  spreadsheets that happen to share a workbook.
'
'      primary   #193F89   ink     #172B4D   muted  #63728A
'      tint      #EBF1FB   panel   #F4F7FB   soft   #F8FAFD
'      hairline  #DFE6F0   amber   #936A1E on #FFF5DC
'
'  It arrived as a design pass over an earlier build of this tool and is kept
'  here deliberately, because it is better than what it replaced. What it is NOT
'  is a place to put behaviour: everything below formats, positions and titles.
'  The one exception is the headline text on three sheets, which is part of the
'  design rather than part of the data.
'
'  ProStyleSheet is the only entry point. Every styler calls it last, so a sheet
'  is built, filled, and then dressed - in that order, always.
' ============================================================================
Public Sub ProStyleSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If ws.name = "Dashboard" Then
        If UCase$(Trim$(CStr(ws.Range("N1").Value2))) <> "MASTER" Then Exit Sub
    End If
    On Error GoTo ProDone
    Select Case ws.name
        Case "Dashboard": ProLayout1 ws
        Case "Pre_Shock": ProLayout2 ws
        Case "Pre_Shock_Sources": ProLayout3 ws
        Case "Pre_Shock_Metrics": ProLayout4 ws
        Case "Pre_Shock_Fields": ProLayout5 ws
        Case "Pre_Shock_Cases": ProLayout6 ws
        Case "Pre_Shock_Results": ProLayout7 ws
        Case "Derived_Values": ProLayout8 ws
        Case "Config_ElementTypeRules": ProLayout9 ws
        Case "Config_MasterFormatting": ProLayout10 ws
        Case "Scenario Element Output": ProLayout11 ws
        Case "Build_Log": ProLayout12 ws
        Case "Manual_Formula_Guide": ProLayout13 ws
        Case "Pre_Shock_Breakdown": ProLayout14 ws
    End Select
    If ActiveSheet Is ws Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = (ws.name <> "Dashboard")
    End If
ProDone:
End Sub

Private Sub ProLayout1(ByVal ws As Worksheet)
    Dim r As Long, stepNames As Variant
    On Error GoTo Done
    ' Preserve D18:K28: import and build routines write the run status here.
    ws.Range("A1:K60").Borders.LineStyle = xlNone
    ws.Range("B7:C15").ClearContents
    ws.Range("B18:C28").UnMerge: ws.Range("B18:C28").ClearContents
    ws.Range("D8:K14").UnMerge: ws.Range("D8:K14").ClearContents
    ws.Range("D3:K3").UnMerge: ws.Range("D3:K3").Merge
    ws.Range("D4:K4").UnMerge: ws.Range("D4:K4").Merge
    ws.Range("D6:K6").UnMerge: ws.Range("D6:K6").Merge
    ws.Range("D15:K15").UnMerge: ws.Range("D15:K15").Merge
    ws.Range("D16:K16").UnMerge: ws.Range("D16:K16").Merge
    ws.Range("B6:C6").UnMerge: ws.Range("B6:C6").Merge
    For r = 18 To 28
        ws.Range("D" & r & ":E" & r).UnMerge: ws.Range("D" & r & ":E" & r).Merge
        ws.Range("F" & r & ":K" & r).UnMerge: ws.Range("F" & r & ":K" & r).Merge
    Next r
    ws.columns("A").ColumnWidth = 3: ws.columns("B:C").ColumnWidth = 13
    ws.columns("D").ColumnWidth = 22: ws.columns("E").ColumnWidth = 5
    ws.columns("F").ColumnWidth = 22: ws.columns("G:H").ColumnWidth = 15
    ws.columns("I:J").ColumnWidth = 13: ws.columns("K").ColumnWidth = 5
    ws.rows(2).RowHeight = 12
    ws.rows(3).RowHeight = 38: ws.rows(4).RowHeight = 26
    ws.rows(5).RowHeight = 18: ws.rows(6).RowHeight = 26
    ws.rows(7).RowHeight = 8: ws.rows("8:10").RowHeight = 20
    ws.rows(11).RowHeight = 10: ws.rows("12:14").RowHeight = 18
    ws.rows(15).RowHeight = 24: ws.rows(16).RowHeight = 28
    ws.rows(17).RowHeight = 8: ws.rows("18:24").RowHeight = 23
    ws.rows("25:28").RowHeight = 29: ws.rows("29:30").RowHeight = 8
    With ws.Range("A1:K60")
        .Font.name = "Aptos": .Font.Size = 10: .Font.Bold = False
        .Font.Color = RGB(16, 24, 40): .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255): .HorizontalAlignment = xlLeft
        .WrapText = False: .IndentLevel = 0
    End With
    With ws.Range("B6:C28")
        .Interior.Color = RGB(245, 247, 250): .Font.Color = RGB(102, 112, 133)
    End With
    With ws.Range("B6:C6")
        .Font.Size = 9: .Font.Bold = True: .IndentLevel = 1: .Value2 = "WORKSPACE"
    End With
    With ws.Range("D3:K3")
        .Font.Size = 25: .Font.Bold = True: .Font.Color = RGB(14, 34, 64)
        .NumberFormat = "@": .Value2 = "Stress testing reconciliation"
    End With
    With ws.Range("D4:K4")
        .Font.Size = 11: .Font.Color = RGB(102, 112, 133)
        .NumberFormat = "@": .Value2 = "Jordan Kuwait Bank  /  Independent validation  /  V14"
    End With
    With ws.Range("D6:K6")
        .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(14, 34, 64)
        .NumberFormat = "@": .Value2 = "End-to-end control workspace"
    End With
    With ws.Range("D15:K15")
        .Font.Size = 10: .Font.Color = RGB(102, 112, 133): .Font.Italic = False
        .NumberFormat = "@": .Value2 = "Validate inputs, reconcile each stage, investigate exceptions."
    End With
    With ws.Range("D16:K16")
        .Font.Size = 12: .Font.Bold = True: .Font.Color = RGB(14, 34, 64)
        .Interior.Color = RGB(245, 247, 250): .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(234, 236, 240)
        .NumberFormat = "@": .Value2 = "Current run"
    End With
    For r = 18 To 28
        With ws.Range("D" & r & ":E" & r)
            .Font.Size = 10: .Font.Color = RGB(102, 112, 133): .IndentLevel = 1
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(242, 244, 247)
        End With
        With ws.Range("F" & r & ":K" & r)
            .Font.Size = 10: .Font.Color = RGB(16, 24, 40): .WrapText = True: .IndentLevel = 1
            .Interior.Color = IIf(r Mod 2 = 0, RGB(245, 247, 250), vbWhite)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(242, 244, 247)
        End With
    Next r
    ws.Range("F21:K24").Font.Bold = True: ws.Range("D26").Value2 = "Base / pre-shock"
    ws.Range("B18:C18").Merge
    With ws.Range("B18:C18")
        .Value2 = "CONTROL SEQUENCE": .Font.Size = 9: .Font.Bold = True: .IndentLevel = 1
    End With
    stepNames = Array("01  Input validation", "02  Base", "03  Filter / pre-shock", "04  Shock", "05  Post-shock", "06  Exceptions")
    For r = 0 To 5
        With ws.Range("B" & (20 + r) & ":C" & (20 + r))
            .Merge: .Value2 = stepNames(r): .Font.Size = 10
            .Font.Color = RGB(71, 84, 103): .IndentLevel = 1
        End With
    Next r
    Button ws, "JKB_ReconStart", "Open Reconciliation", "OpenReconWorkbench", 0, 0, 100, 40, False
    On Error Resume Next
    With ws.Shapes("JKB_BrandLogo")
        .LockAspectRatio = msoTrue: .Width = 132
        .Left = ws.Range("B3").Left + 8: .Top = ws.Range("B3").Top + 3
    End With
    ws.Tab.Color = RGB(25, 63, 137): ws.ScrollArea = "A1:K28"
    On Error GoTo Done
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout2(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:L25").Borders.LineStyle = xlNone
    ws.Range("A3:L3").UnMerge: ws.Range("A3:L3").Merge
    ws.Range("A4:L4").UnMerge: ws.Range("A4:L4").Merge
    ws.Range("A6:L6").UnMerge: ws.Range("A6:L6").Merge
    ws.Range("A7:L7").UnMerge: ws.Range("A7:L7").Merge
    ws.Range("B9:F9").UnMerge: ws.Range("B9:F9").Merge
    ws.Range("B10:F10").UnMerge: ws.Range("B10:F10").Merge
    ws.Range("A16:L16").UnMerge: ws.Range("A16:L16").Merge
    ws.Range("B17:L17").UnMerge: ws.Range("B17:L17").Merge
    ws.Range("B18:L18").UnMerge: ws.Range("B18:L18").Merge
    ws.Range("B19:L19").UnMerge: ws.Range("B19:L19").Merge
    ws.Range("B20:L20").UnMerge: ws.Range("B20:L20").Merge
    ws.Range("B21:L21").UnMerge: ws.Range("B21:L21").Merge
    ws.Range("B22:L22").UnMerge: ws.Range("B22:L22").Merge
    ws.columns("A").ColumnWidth = 26
    ws.columns("B").ColumnWidth = 36
    ws.columns("C").ColumnWidth = 3
    ws.columns("D").ColumnWidth = 10
    ws.columns("E").ColumnWidth = 24
    ws.columns("F").ColumnWidth = 24
    ws.columns("G").ColumnWidth = 3
    ws.columns("H").ColumnWidth = 26
    ws.columns("I:L").ColumnWidth = 5
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 42
    ws.rows("5").RowHeight = 12
    ws.rows("6").RowHeight = 30
    ws.rows("7").RowHeight = 24
    ws.rows("8").RowHeight = 8
    ws.rows("9:11").RowHeight = 26
    ws.rows("12").RowHeight = 30
    ws.rows("13:14").RowHeight = 23
    ws.rows("15").RowHeight = 12
    ws.rows("16").RowHeight = 24
    ws.rows("17:22").RowHeight = 28
    With ws.Range("A1:L25")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:L3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:L4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A6:L6")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(176, 138, 46)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A7:L7")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(234, 236, 240)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A9:A11")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(238, 243, 251)
        .IndentLevel = 1
    End With
    With ws.Range("B9:F10")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A12:B12")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
        .WrapText = True
    End With
    With ws.Range("E11:F12")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
        .WrapText = True
    End With
    With ws.Range("B12")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F11")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F12")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("H11:L14")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A16:L16")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(234, 236, 240)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A17:A22")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With ws.Range("B17:L22")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A17:L17")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    With ws.Range("A19:L19")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    With ws.Range("A21:L21")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Range("A3").NumberFormat = "@": ws.Range("A3").Value2 = "Base and pre-shock checks"
    ws.Range("A4").NumberFormat = "@": ws.Range("A4").Value2 = "Rebuild BASE across the whole portfolio, then apply each test case filter to derive PRE-SHOCK. Review the differences against the system before using the derived values in your reconciliation."
    ws.Range("A7").NumberFormat = "@": ws.Range("A7").Value2 = "Set up the independent check"
    ws.Range("A16").NumberFormat = "@": ws.Range("A16").Value2 = "Explore the workspace"
    ws.Range("B9").NumberFormat = "@": ws.Range("B9").Value2 = "Upload ECL and CAPRWA extracts, or load a folder."
    ws.Range("B10").NumberFormat = "@": ws.Range("B10").Value2 = "Enable independent checking and select date matching."
    ws.Range("B11").NumberFormat = "@": ws.Range("B11").Value2 = "Run tests and review differences."
    ws.Tab.Color = RGB(59, 99, 176)
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout3(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:H25").Borders.LineStyle = xlNone
    ws.Range("A3:H3").UnMerge: ws.Range("A3:H3").Merge
    ws.Range("A4:H4").UnMerge: ws.Range("A4:H4").Merge
    ws.columns("A").ColumnWidth = 14
    ws.columns("B").ColumnWidth = 34
    ws.columns("C").ColumnWidth = 40
    ws.columns("D").ColumnWidth = 17
    ws.columns("E").ColumnWidth = 12
    ws.columns("F").ColumnWidth = 27
    ws.columns("G").ColumnWidth = 22
    ws.columns("H").ColumnWidth = 24
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("7").RowHeight = 34
    ws.rows("8:9").RowHeight = 52
    With ws.Range("A1:H25")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:H3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:H4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A7:H9")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A7:H7")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 8 To 9 Step 2
        ws.Range("A" & r & ":H" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("B8:C9")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 7
        ActiveWindow.SplitColumn = 1
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout4(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:J36").Borders.LineStyle = xlNone
    ws.Range("A3:J3").UnMerge: ws.Range("A3:J3").Merge
    ws.Range("A4:J4").UnMerge: ws.Range("A4:J4").Merge
    ws.columns("A").ColumnWidth = 36
    ws.columns("B").ColumnWidth = 10
    ws.columns("C").ColumnWidth = 10
    ws.columns("D").ColumnWidth = 54
    ws.columns("E").ColumnWidth = 8
    ws.columns("F").ColumnWidth = 30
    ws.columns("G").ColumnWidth = 25
    ws.columns("H").ColumnWidth = 35
    ws.columns("I").ColumnWidth = 32
    ws.columns("J").ColumnWidth = 22
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("7").RowHeight = 34
    ws.rows("8:36").RowHeight = 40
    With ws.Range("A1:J36")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:J3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:J4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A7:J36")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A7:J7")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 8 To 36 Step 2
        ws.Range("A" & r & ":J" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("C8:C36")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("D8:D36")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(238, 243, 251)
        .WrapText = True
    End With
    With ws.Range("J8:J36")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 7
        ActiveWindow.SplitColumn = 1
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout5(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:G27").Borders.LineStyle = xlNone
    ws.Range("A3:G3").UnMerge: ws.Range("A3:G3").Merge
    ws.Range("A4:G4").UnMerge: ws.Range("A4:G4").Merge
    ws.columns("A").ColumnWidth = 32
    ws.columns("B").ColumnWidth = 54
    ws.columns("C").ColumnWidth = 17
    ws.columns("D").ColumnWidth = 25
    ws.columns("E").ColumnWidth = 16
    ws.columns("F").ColumnWidth = 40
    ws.columns("G").ColumnWidth = 16
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("7").RowHeight = 34
    ws.rows("8:27").RowHeight = 38
    With ws.Range("A1:G27")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:G3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:G4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A7:G27")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A7:G7")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 8 To 27 Step 2
        ws.Range("A" & r & ":F" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("B8:C27")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(238, 243, 251)
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 7
        ActiveWindow.SplitColumn = 1
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout6(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:L57").Borders.LineStyle = xlNone
    ws.Range("A3:L3").UnMerge: ws.Range("A3:L3").Merge
    ws.Range("A4:L4").UnMerge: ws.Range("A4:L4").Merge
    ws.columns("A").ColumnWidth = 9
    ws.columns("B").ColumnWidth = 15
    ws.columns("C").ColumnWidth = 19
    ws.columns("D").ColumnWidth = 36
    ws.columns("E").ColumnWidth = 22
    ws.columns("F:H").ColumnWidth = 55
    ws.columns("I").ColumnWidth = 28
    ws.columns("J").ColumnWidth = 13
    ws.columns("K").ColumnWidth = 30
    ws.columns("L").ColumnWidth = 32
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("7").RowHeight = 34
    ws.rows("8:57").RowHeight = 66
    With ws.Range("A1:L57")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:L3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:L4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A7:L57")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A7:L7")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 8 To 57 Step 2
        ws.Range("A" & r & ":L" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("A8:A57")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("G8:G57")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(238, 243, 251)
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 7
        ActiveWindow.SplitColumn = 3
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout7(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:L25").Borders.LineStyle = xlNone
    ws.Range("A3:L3").UnMerge: ws.Range("A3:L3").Merge
    ws.Range("A4:L4").UnMerge: ws.Range("A4:L4").Merge
    ws.columns("A").ColumnWidth = 14
    ws.columns("B").ColumnWidth = 20
    ws.columns("C").ColumnWidth = 16
    ws.columns("D").ColumnWidth = 20
    ws.columns("E").ColumnWidth = 36
    ws.columns("F:K").ColumnWidth = 19
    ws.columns("L").ColumnWidth = 30
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("7").RowHeight = 34
    ws.rows("8:8").RowHeight = 32
    With ws.Range("A1:L25")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:L3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:L4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A7:L8")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A7:L7")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 8 To 8 Step 2
        ws.Range("A" & r & ":L" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("G8:G8")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(236, 253, 243)
    End With
    With ws.Range("J8:J8")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(236, 253, 243)
    End With
    With ws.Range("H8:H8")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
    End With
    With ws.Range("K8:K8")
        .Font.name = "Aptos"
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 7
        ActiveWindow.SplitColumn = 5
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout8(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:P25").Borders.LineStyle = xlNone
    ws.Range("A3:P3").UnMerge: ws.Range("A3:P3").Merge
    ws.Range("A4:P4").UnMerge: ws.Range("A4:P4").Merge
    ws.columns("A").ColumnWidth = 14
    ws.columns("B").ColumnWidth = 20
    ws.columns("C").ColumnWidth = 16
    ws.columns("D").ColumnWidth = 20
    ws.columns("E").ColumnWidth = 36
    ws.columns("F").ColumnWidth = 25
    ws.columns("G").ColumnWidth = 14
    ws.columns("H:J").ColumnWidth = 20
    ws.columns("K").ColumnWidth = 60
    ws.columns("L").ColumnWidth = 36
    ws.columns("M:O").ColumnWidth = 20
    ws.columns("P").ColumnWidth = 32
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("7").RowHeight = 34
    ws.rows("8:25").RowHeight = 30
    With ws.Range("A1:P25")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:P3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:P4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    If Len(CStr(ws.Range("A3").Value2)) = 0 Or CStr(ws.Range("A3").Value2) = "Derived values" Or CStr(ws.Range("A3").Value2) = "Base & pre-shock" Then
    ws.Range("A3").NumberFormat = "@": ws.Range("A3").Value2 = "Base & pre-shock"
    ws.Range("A4").NumberFormat = "@": ws.Range("A4").Value2 = "Base and pre-shock rebuilt from the input files beside the system values, with the difference and whether they reconcile. Click Run on the home screen to refresh."
    End If
    If Len(CStr(ws.Range("A7").Value2)) > 0 Then
        With ws.Range("A7:P7")
            .Interior.Color = RGB(25, 63, 137): .Font.Color = vbWhite: .Font.Bold = True
            .WrapText = True: .HorizontalAlignment = xlCenter
        End With
        r = ws.Cells(ws.rows.count, 1).End(xlUp).row
        If r >= 8 Then
            ws.Range("A8:P" & r).WrapText = True
            ws.Range("H8:O" & r).HorizontalAlignment = xlRight
            ws.Range("K8:L" & r).HorizontalAlignment = xlLeft
            ws.Range("H8:I" & r).Interior.Color = RGB(236, 253, 243)
            ws.Range("O8:O" & r).Interior.Color = RGB(238, 243, 251)
        End If
    End If
    ws.Tab.Color = RGB(59, 99, 176)
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout9(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:AE1964").Borders.LineStyle = xlNone
    ws.Range("A3:U3").UnMerge: ws.Range("A3:U3").Merge
    ws.Range("A4:U4").UnMerge: ws.Range("A4:U4").Merge
    ws.Range("Y3:AE3").UnMerge: ws.Range("Y3:AE3").Merge
    ws.Range("Y4:AE4").UnMerge: ws.Range("Y4:AE4").Merge
    ws.columns("A").ColumnWidth = 36
    ws.columns("B").ColumnWidth = 12
    ws.columns("C:E").ColumnWidth = 13
    ws.columns("F:G").ColumnWidth = 35
    ws.columns("H:I").ColumnWidth = 18
    ws.columns("J:M").ColumnWidth = 48
    ws.columns("N:S").ColumnWidth = 18
    ws.columns("T").ColumnWidth = 38
    ws.columns("U").ColumnWidth = 24
    ws.columns("V").ColumnWidth = 35
    ws.columns("W").ColumnWidth = 19
    ws.columns("X").ColumnWidth = 3
    ws.columns("Y").ColumnWidth = 22
    ws.columns("Z").ColumnWidth = 34
    ws.columns("AA").ColumnWidth = 13
    ws.columns("AB").ColumnWidth = 23
    ws.columns("AC:AD").ColumnWidth = 17
    ws.columns("AE").ColumnWidth = 14
    ws.rows("1").RowHeight = 48
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("6").RowHeight = 43
    ws.rows("7:1964").RowHeight = 36
    With ws.Range("A1:AE1964")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:U3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:U4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A6:U1964")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A6:U6")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 8 To 1964 Step 2
        ws.Range("A" & r & ":U" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("J7:M1964")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(238, 243, 251)
        .WrapText = True
    End With
    With ws.Range("C7:E1964")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F7:F1964")
        .Font.name = "Aptos"
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("N7:S1964")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("Y3:AE3")
        .Font.name = "Aptos"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("Y4:AE4")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("V6:W6")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
    End With
    With ws.Range("Y6:AE6")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("Y7:AE20")
        .Font.name = "Aptos"
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
        .WrapText = True
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(176, 138, 46)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 6
        ActiveWindow.SplitColumn = 1
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout10(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:G28").Borders.LineStyle = xlNone
    ws.Range("A3:G3").UnMerge: ws.Range("A3:G3").Merge
    ws.Range("A4:G4").UnMerge: ws.Range("A4:G4").Merge
    ws.Range("A22:G22").UnMerge: ws.Range("A22:G22").Merge
    ws.Range("A23:G23").UnMerge: ws.Range("A23:G23").Merge
    ws.Range("A26:G26").UnMerge: ws.Range("A26:G26").Merge
    ws.Range("A28:B28").UnMerge: ws.Range("A28:B28").Merge
    ws.Range("C28:G28").UnMerge: ws.Range("C28:G28").Merge
    ws.Range("F5:G5").UnMerge: ws.Range("F5:G5").Merge
    ws.columns("A").ColumnWidth = 34
    ws.columns("B").ColumnWidth = 18
    ws.columns("C").ColumnWidth = 47
    ws.columns("D:E").ColumnWidth = 3
    ws.columns("F").ColumnWidth = 25
    ws.columns("G").ColumnWidth = 23
    ws.rows("1").RowHeight = 48
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("5:21").RowHeight = 25
    ws.rows("22:23").RowHeight = 26
    ws.rows("24:25").RowHeight = 10
    ws.rows("26").RowHeight = 28
    ws.rows("27").RowHeight = 10
    ws.rows("28").RowHeight = 26
    ws.rows("29:55").RowHeight = 28
    With ws.Range("A1:G28")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:G3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:G4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A5:B5")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
    End With
    With ws.Range("F5:G6")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
    End With
    With ws.Range("A6:A21")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    With ws.Range("B6:B21")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(253, 246, 227)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F7:F16")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    With ws.Range("G7:G16")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("A22:G23")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A26:G26")
        .Font.name = "Aptos"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A28:G28")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
    End With
    With ws.Range("A29:B55")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(176, 138, 46)
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout11(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A3:EV145").Borders.LineStyle = xlNone
    ws.columns("A").ColumnWidth = 14
    ws.columns("B").ColumnWidth = 12
    ws.columns("C").ColumnWidth = 21
    ws.columns("D:G").ColumnWidth = 14
    ws.columns("H").ColumnWidth = 20
    ws.columns("I").ColumnWidth = 38
    ws.columns("J").ColumnWidth = 16
    ws.columns("K").ColumnWidth = 23
    ws.columns("L:M").ColumnWidth = 40
    ws.columns("N").ColumnWidth = 58
    ws.columns("O:EV").ColumnWidth = 22
    ws.rows("3").RowHeight = 58
    ws.rows("4:145").RowHeight = 30
    ws.rows("1").RowHeight = 48
    ws.rows("2").RowHeight = 8
    With ws.Range("A3:EV145")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:EV145")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A3:EV3")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 4 To 145 Step 2
        ws.Range("A" & r & ":EV" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("N4:N145")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 3
        ActiveWindow.SplitColumn = 3
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout12(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A3:E340").Borders.LineStyle = xlNone
    ws.columns("A").ColumnWidth = 11
    ws.columns("B").ColumnWidth = 27
    ws.columns("C").ColumnWidth = 56
    ws.columns("D").ColumnWidth = 80
    ws.columns("E").ColumnWidth = 24
    ws.rows("3").RowHeight = 30
    ws.rows("4:340").RowHeight = 43
    ws.rows("1").RowHeight = 48
    ws.rows("2").RowHeight = 8
    With ws.Range("A3:E340")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:E340")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A3:E3")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 4 To 340 Step 2
        ws.Range("A" & r & ":E" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("A4:A340")
        .Font.name = "Aptos"
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("E4:E340")
        .Font.name = "Aptos"
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(152, 162, 179)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 3
        ActiveWindow.SplitColumn = 1
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

Private Sub ProLayout13(ByVal ws As Worksheet)
    Dim r As Long
    On Error GoTo Done
    ws.Range("A1:E102").Borders.LineStyle = xlNone
    ws.Range("A3:E3").UnMerge: ws.Range("A3:E3").Merge
    ws.Range("A4:E4").UnMerge: ws.Range("A4:E4").Merge
    ws.columns("A").ColumnWidth = 26
    ws.columns("B").ColumnWidth = 54
    ws.columns("C").ColumnWidth = 52
    ws.columns("D").ColumnWidth = 45
    ws.columns("E").ColumnWidth = 45
    ws.rows("1").RowHeight = 48
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 30
    ws.rows("5").RowHeight = 30
    ws.rows("6:102").RowHeight = 48
    With ws.Range("A1:E102")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:E3")
        .Font.name = "Aptos"
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:E4")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A5:E102")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With ws.Range("A5:E5")
        .Font.name = "Aptos"
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    For r = 6 To 102 Step 2
        ws.Range("A" & r & ":E" & r).Interior.Color = RGB(245, 247, 250)
    Next r
    With ws.Range("B6:B102")
        .Font.name = "Aptos"
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(238, 243, 251)
    End With
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(245, 247, 250)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(176, 138, 46)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 5
        ActiveWindow.SplitColumn = 1
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub


' Pre_Shock_Breakdown - why each figure is what it is.
'
' In the same language as ProLayout7 and ProLayout8, which are the other two
' wide tables: navy header band, banded rows, the gold rule under the title.
' Two things are specific to this sheet. The pair of columns holding what the
' filter SELECTED gets the soft green the design uses for a derived figure, and
' the pair holding what the whole PORTFOLIO holds gets the neutral tint - so the
' comparison the sheet exists to make is visible before anything is read. The
' share between them is the one column with a rule under it.
Private Sub ProLayout14(ByVal ws As Worksheet)
    Dim lastR As Long
    On Error GoTo Done
    ws.Range("A1:Q2000").Borders.LineStyle = xlNone
    ws.Range("A3:Q3").UnMerge: ws.Range("A3:Q3").Merge
    ws.Range("A4:Q4").UnMerge: ws.Range("A4:Q4").Merge
    ws.columns("A").ColumnWidth = 13
    ws.columns("B").ColumnWidth = 14
    ws.columns("C").ColumnWidth = 15
    ws.columns("D").ColumnWidth = 15
    ws.columns("E").ColumnWidth = 11
    ws.columns("F").ColumnWidth = 20
    ws.columns("G").ColumnWidth = 34
    ws.columns("H:I").ColumnWidth = 15
    ws.columns("J:K").ColumnWidth = 19
    ws.columns("L").ColumnWidth = 14
    ws.columns("M:N").ColumnWidth = 17
    ws.columns("O:P").ColumnWidth = 22
    ws.columns("Q").ColumnWidth = 34
    ws.rows("1").RowHeight = 70
    ws.rows("2").RowHeight = 9
    ws.rows("3").RowHeight = 36
    ws.rows("4").RowHeight = 32
    ws.rows("5").RowHeight = 12
    ws.rows("6").RowHeight = 18
    ws.rows("7").RowHeight = 34

    lastR = ws.Cells(ws.rows.count, 6).End(xlUp).row
    If lastR < 8 Then lastR = 8

    ' The typeface over a generous block, not just the rows that happen to hold
    ' data: an empty sheet is the FIRST thing anyone sees here, and it should look
    ' like the rest of the workbook before it has been run rather than after.
    With ws.Range("A1:Q400")
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Color = RGB(14, 34, 64)
        .Interior.Color = RGB(255, 255, 255)
    End With
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 17))
        .Font.name = "Aptos"
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(14, 34, 64)
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .WrapText = False
    End With
    With ws.Range("A3:Q3")
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A4:Q4")
        .Font.Color = RGB(102, 112, 133)
        .WrapText = True
        .IndentLevel = 1
    End With
    With ws.Range("A6:Q6")
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(102, 112, 133)
        .IndentLevel = 1
    End With
    With ws.Range("A7:Q7")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    If lastR >= 8 Then
        With ws.Range(ws.Cells(8, 1), ws.Cells(lastR, 17))
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
            .Borders(xlInsideHorizontal).Weight = xlHairline
        End With
        ' What the filter selected, and what the portfolio holds of the same
        ' thing. Colouring the pair is the whole point of the sheet.
        ws.Range(ws.Cells(8, 8), ws.Cells(lastR, 8)).Interior.Color = RGB(236, 253, 243)
        ws.Range(ws.Cells(8, 10), ws.Cells(lastR, 10)).Interior.Color = RGB(236, 253, 243)
        ws.Range(ws.Cells(8, 9), ws.Cells(lastR, 9)).Interior.Color = RGB(245, 247, 250)
        ws.Range(ws.Cells(8, 11), ws.Cells(lastR, 11)).Interior.Color = RGB(245, 247, 250)
        With ws.Range(ws.Cells(8, 12), ws.Cells(lastR, 12))
            .Interior.Color = RGB(253, 246, 227)
            .Font.Bold = True
            .Font.Color = RGB(176, 138, 46)
            .HorizontalAlignment = xlCenter
        End With
        ws.Range(ws.Cells(8, 17), ws.Cells(lastR, 17)).Font.Color = RGB(102, 112, 133)
    End If
    With ws.Range("A1:BZ1")
        .Font.name = "Aptos"
        .Font.Size = 9
        .Font.Color = RGB(245, 247, 250)
        .Interior.Color = RGB(245, 247, 250)
    End With
    ws.Tab.Color = RGB(59, 99, 176)
    If ActiveSheet Is ws Then
        ActiveWindow.FreezePanes = False
        ActiveWindow.splitRow = 7
        ActiveWindow.SplitColumn = 7
        ActiveWindow.FreezePanes = True
    End If
    ProRibbon ws, True
Done:
End Sub

' Packs a button into the toolbar. Width follows the caption, so every button
' has the same padding and the row has one rhythm, whatever the captions are.
Private Sub ProMove(ByVal ws As Worksheet, ByVal nm As String, ByRef x As Double, ByVal y As Double, ByVal w As Double)
    Dim sh As Shape, n As Long
    On Error Resume Next
    Set sh = ws.Shapes(nm)
    If sh Is Nothing Then Exit Sub
    If sh.Visible = msoFalse Then Exit Sub
    n = Len(sh.TextFrame.Characters.text)
    If n > 0 Then w = n * 5.3 + 24
    If w < 56 Then w = 56
    sh.Left = x: sh.Top = y: sh.Width = w: sh.Height = 22
    sh.TextFrame.Characters.Font.Size = 9.5
    x = x + w + 6
End Sub

' The toolbar, design system v2. One white band on every sheet:
'
'   row 1   Home . Review console  |  what this sheet does
'   row 1b  (pre-shock sheets only) one tab per pre-shock sheet, current tinted
'
' Everything that used to sit in a second "view" strip is Excel's own zoom.
Private Sub ProRibbon(ByVal ws As Worksheet, ByVal isMaster As Boolean)
    Dim x As Double, sh As Shape, nm As Variant, rg As String, twoRows As Boolean
    On Error Resume Next
    twoRows = (IsPreShockSheet(ws) And isMaster)
    ws.rows(1).RowHeight = IIf(twoRows, 64, 36)
    ws.rows(2).RowHeight = 8
    With ws.Range("A1:BZ1")
        .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(255, 255, 255)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(234, 236, 240): .Borders(xlEdgeBottom).Weight = xlThin
    End With
    If isMaster And Not IsPreShockSheet(ws) Then Button ws, "JKB_ConsoleHome", "Review console", "OpenStressConsole", 0, 0, 110, 22, False
    If Not isMaster Then ws.Shapes("JKB_ConsoleHome").Delete
    If isMaster And ws.name = "Recon_Workbench" Then
        ' The home carries its own actions; a toolbar above it would repeat them.
        For Each nm In Array("JKB_Dashboard", "JKB_Config", "JKB_Formula", "JKB_GroupToggle", "JKB_ConsoleHome")
            ws.Shapes(CStr(nm)).Visible = msoFalse
        Next nm
    End If
    x = 12
    ProMove ws, "JKB_Dashboard", x, 7, 0
    ProMove ws, "JKB_ConsoleHome", x, 7, 0
    ProMove ws, "JKB_PsConsole", x, 7, 0
    x = x + 14
    For Each nm In Array("JKB_Config", "JKB_Editor", "JKB_GroupSetup", "JKB_Formula", "JKB_GroupToggle", _
                         "JKB_SevMod", "JKB_SevMed", "JKB_SevSev", "JKB_BackElement", "JKB_OutFormat", "JKB_OutApply", _
                         "JKB_PsUpAll", "JKB_PsUpOne", "JKB_PsEntity", "JKB_PsJoin", "JKB_PsBuilder", "JKB_PsTest", _
                         "JKB_PsTrace", "JKB_PsClearTr", "JKB_PsRules", "JKB_PsManual", "JKB_PsUseDerived", "JKB_PsRun")
        ProMove ws, CStr(nm), x, 7, 0
    Next nm
    If twoRows Then
        x = 12
        For Each nm In Array("JKB_PsNavHome", "JKB_PsNavSrc", "JKB_PsNavMetrics", "JKB_PsNavFields", _
                             "JKB_PsNavCases", "JKB_PsNavVS", "JKB_PsNavResults", "JKB_PsNavDerived")
            ProMove ws, CStr(nm), x, 36, 0
        Next nm
    End If
    If ws.name = "Dashboard" And isMaster Then
        For Each nm In Array("JKB_ReconStart", "JKB_Upload", "JKB_Generate", "JKB_SideConfig", "JKB_SideFormat", "JKB_SidePreShock")
            rg = ""
            Select Case CStr(nm)
                Case "JKB_ReconStart": rg = "D8:K10"
                Case "JKB_Upload": rg = "D12:G14"
                Case "JKB_Generate": rg = "H12:K14"
                Case "JKB_SideConfig": rg = "B8:C9"
                Case "JKB_SideFormat": rg = "B11:C12"
                Case "JKB_SidePreShock": rg = "B14:C15"
            End Select
            Set sh = Nothing: Set sh = ws.Shapes(CStr(nm))
            If Not sh Is Nothing And Len(rg) > 0 Then
                sh.Left = ws.Range(rg).Left + 5: sh.Top = ws.Range(rg).Top + 3
                sh.Width = ws.Range(rg).Width - 10: sh.Height = ws.Range(rg).Height - 6
                sh.TextFrame.Characters.Font.Size = IIf(CStr(nm) = "JKB_ReconStart", 15, 11)
            End If
        Next nm
    End If
    If ws.name = "Config_MasterFormatting" Then
        Set sh = Nothing: Set sh = ws.Shapes("JKB_ApplyFormat")
        If Not sh Is Nothing Then
            sh.Left = ws.Range("F3").Left: sh.Top = ws.Range("F3").Top + 4
            sh.Width = ws.Range("F3:G3").Width: sh.Height = 26
        End If
    End If
    Err.Clear
End Sub

Private Function ProStudioCss() As String
    Dim s As String
    s = s & ":root{--green:#193F89;--green2:#3B63B0;--green3:#EEF3FB;--ink:#0E2240;--muted:#667085;--line:#EAECF0;--panel:#FFFFFF;--bg:#F5F7FA;--soft:#F5F7FA;--amber:#B08A2E;--amberbg:#FDF6E3;--red:#B42318;--redbg:#FEF3F2;--blue:#3B63B0;--bluebg:#EEF3FB;--shadow:0 2px 12px rgba(25,63,137,.04);}"
    s = s & vbLf
    s = s & "body{background:#F5F7FA;font-family:""Segoe UI"",Arial,sans-serif;}"
    s = s & vbLf
    s = s & ".shell{max-width:none;padding:16px 20px;}"
    s = s & vbLf
    s = s & ".topbar{background:#FFFFFF;color:#193F89;border:1px solid #EAECF0;border-top:3px solid #193F89;border-radius:7px;padding:18px 22px;box-shadow:var(--shadow);}"
    s = s & vbLf
    s = s & ".mark{background:#193F89;color:#FFFFFF;border:0;border-radius:5px;width:43px;height:43px;font-size:12px;font-weight:700;}"
    s = s & vbLf
    s = s & ".title h1{font-size:23px;letter-spacing:-.5px;}.title p{color:#667085;font-size:12px;margin-top:6px;}.stat{background:#EEF3FB;border:1px solid #EAECF0;color:#667085;border-radius:4px;font-size:11px;padding:6px 9px;}.stat b{color:#193F89;}"
    s = s & vbLf
    s = s & ".contextbar{gap:10px;margin:12px 0;grid-template-columns:1.3fr 1fr 1fr 1fr;}.contextitem{border-radius:5px;padding:10px 13px;box-shadow:none;}.contextitem span{font-size:9px;letter-spacing:.8px;color:#667085;}.contextitem b{color:#344054;font-size:12px;}"
    s = s & vbLf
    s = s & ".grid{gap:14px;grid-template-columns:minmax(0,1fr) 355px;}.card{border-radius:7px;box-shadow:var(--shadow);}.cardhead{padding:19px 20px 12px;}.cardhead h2,.sidehead h3{font-size:15px;color:#193F89;}.sub{font-size:11px;line-height:1.6;}"
    s = s & vbLf
    s = s & ".btn{border-color:#D6E2F5;color:#344054;border-radius:5px;padding:6px 9px;font-size:11px;}.btn:hover,.btn:focus{background:#EEF3FB;border-color:#98A2B3;}.btn.ref{background:#EEF3FB;color:#193F89;}.btn.helper{background:#F5F7FA;color:#475467;}.btn.template{background:#FFFAEB;color:#B08A2E;}.btn.primary{background:#193F89;border-color:#193F89;color:#FFFFFF;padding:9px 14px;}.btn.primary:hover{background:#3B63B0;}"
    s = s & vbLf
    s = s & ".tooltitle{color:#667085;font-size:9px;letter-spacing:.7px;}.toolbar{padding-bottom:8px;}.toolrow{gap:5px;margin-top:6px;}"
    s = s & vbLf
    s = s & "textarea{border-color:#D6E2F5;border-radius:6px;background:#FCFCFD;color:#193F89;caret-color:#193F89;height:clamp(145px,22vh,245px);font-size:14px;line-height:1.6;}"
    s = s & vbLf
    s = s & "textarea:focus{outline:2px solid #D6E2F5;border-color:#84ADFF;}.canvasbar label{color:#193F89;font-weight:700;font-size:12px;}.key{border-color:#D6E2F5;background:#F5F7FA;color:#667085;border-radius:3px;}"
    s = s & vbLf
    s = s & ".suggestbox{border-color:#B7C4D8;border-radius:6px;box-shadow:0 10px 25px rgba(25,63,137,.13);}.suggesthead{background:#EEF3FB;}.suggestion{border-radius:4px;}.suggestion:hover,.suggestion.selected{background:#EEF3FB;}.sicon{background:#EEF3FB;color:#193F89;border-radius:4px;}.sname{color:#193F89;}.tag{background:#F2F4F7;color:#667085;border-radius:4px;}"
    s = s & vbLf
    s = s & ".translatebox{border-color:#EAECF0;background:#F5F7FA;border-radius:6px;}.translatehead{border-color:#EEF3FB;}.translation{color:#344054;max-height:60px;}.transnote{color:#667085;}"
    s = s & vbLf
    s = s & ".editor{overflow-y:auto;}.editorfoot{margin-top:auto;padding:12px 20px;border-color:#EAECF0;background:#FFFFFF;position:sticky;bottom:0;z-index:8;}.status{color:#067647;font-size:11px;}.actions{gap:6px;}"
    s = s & vbLf
    s = s & ".tabs{gap:4px;border-color:#EAECF0;}.tab{background:#F5F7FA;border-color:#D6E2F5;color:#667085;border-radius:4px;}.tab.active{background:#EEF3FB;border-color:#B7C4D8;color:#193F89;}"
    s = s & vbLf
    s = s & ".healthpill{border-radius:4px;letter-spacing:.4px;}.search{border-color:#D6E2F5;border-radius:5px;}.search:focus{outline:2px solid #D6E2F5;border-color:#84ADFF;}.librow{border-color:#EEF3FB;}.libname{color:#193F89;}.libbuttons button{border-color:#D6E2F5;color:#344054;background:#F9FAFB;border-radius:4px;}.libbuttons button:hover{background:#EEF3FB;}"
    s = s & vbLf
    s = s & ".inspector{background:#F5F7FA;border-color:#EAECF0;border-radius:5px;}.inspector h4{color:#193F89;}.note{border-color:#EAECF0;color:#667085;}.refchip{background:#EEF3FB;border-color:#D6E2F5;color:#193F89;border-radius:4px;}"
    s = s & vbLf
    s = s & "button:focus-visible{outline:2px solid #84ADFF;outline-offset:2px;}"
    s = s & vbLf
    s = s & "@media(max-height:760px){.shell{padding:10px 14px;}.topbar{padding:13px 18px;}.title h1{font-size:20px;}.contextbar{margin:9px 0;}.contextitem{padding:8px 11px;}.cardhead{padding-top:13px;}.toolrow{margin-top:4px;}.btn{padding:5px 8px;}textarea{height:145px;}.editorfoot{padding:10px 20px;}}"
    s = s & vbLf
    s = s & "@media(max-width:1100px){.grid{grid-template-columns:minmax(0,1fr) 300px;}.contextbar{grid-template-columns:1.3fr 1fr 1fr 1fr;}.topstats{max-width:320px;}.title p{max-width:430px;}.tooltitle{width:64px;flex-basis:64px;}}"
    s = s & vbLf
    s = s & "@media(min-width:901px) and (max-width:1120px){.shell{height:100vh;overflow:hidden;}.grid{grid-template-columns:minmax(0,1fr) 300px;overflow:hidden;}.sidecard{height:100%;min-height:0;}}"
    s = s & vbLf
    s = s & "@media(max-width:900px){.shell{height:auto;overflow:visible;}.grid{grid-template-columns:minmax(0,1fr);overflow:visible;}.contextbar{grid-template-columns:1fr 1fr;}.sidecard{height:auto;min-height:420px;}.editorfoot{position:static;}.topbar{align-items:flex-start;}.topstats{display:none;}}"
    s = s & vbLf
    ProStudioCss = s
End Function


