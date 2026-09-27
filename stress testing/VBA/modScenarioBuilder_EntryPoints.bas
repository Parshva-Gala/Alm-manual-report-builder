Attribute VB_Name = "modScenarioBuilder_EntryPoints"
Option Explicit

Public Sub InitializeJKBTool()
    Dim ws As Worksheet, state As Object, er As String
    If JKB_Busy Then Exit Sub
    On Error GoTo Failed
    Set state = CaptureState()
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayStatusBar = True
    Application.StatusBar = "JKB: preparing reconciliation workspace..."
    EnsurePreShockSheet
    ' Opening only rebinds the buttons and shows Home. The slow, data-changing work
    ' (metric defaults, test-case sync, restyling every sheet) runs once per layout
    ' stamp - see modWorkbookCare - or on demand with JKB_UpgradeWorkbook.
    If WorkbookNeedsUpgrade() Then
        CleanMasterTabs
        RunWorkbookUpgrade "open"
    End If
    Sheet1.RebindAll
    EnsureReconWorkbench
    FinishWorkbookUpgrade
    OpenReconWorkbench
    RestoreState state
    Exit Sub
Failed:
    er = Err.description: RestoreState state
    UiProblem "Open the tool", "The tool did not finish opening.", er, "Close and reopen the workbook. If it happens again, check Build_Log."
End Sub

Public Sub BuildAllScenarioSheets()
    StartScenarioCategorySelection
End Sub
' Optional native Excel check. It uses synthetic records and deletes its own temporary sheet.
Public Sub JKB_RunSmokeTests()
    Dim a As Variant, h As Object, idx As Object, stats As Object, rm As Object, tc As Object, k As Variant
    Dim headers As Variant, r As Long, c As Long, d As Long, e As Long, baseRow As Long, er As String, errNo As Long
    Dim ws As Worksheet, previous As Object, state As Object, f As String
    If JKB_Busy Then Exit Sub
    On Error GoTo Failed
    Set previous = ActiveSheet: Set state = CaptureState(): JKB_Busy = True
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.DisplayAlerts = False
    headers = Array("AS_OF_DATE", "ENTITY_ID", "ENTITY_CODE", "RUN_ID", "SCENARIO_TEST_CASE_CODE", "SCENARIO_TEST_CASE_NAME", "SCENARIO_ELEMENT_CODE", "SCENARIO_ELEMENT_NAME", "ELEMENT_TYPE", "SEVERITY_CODE", "AMOUNT_CHANGE", "SEVERITY_ID", "CHECK_VALUE")
    ReDim a(1 To 10, 1 To 13): Set h = NewMap()
    For c = 1 To 13: a(1, c) = headers(c - 1): h(headers(c - 1)) = c: Next c
    r = 2
    For d = 0 To 1
        For e = 1 To 2
            a(r, 1) = CDbl(DateSerial(2025, 12 + d, 31)): a(r, 2) = e: a(r, 3) = "JKB_TEST": a(r, 4) = 1
            a(r, 10) = "BASE": a(r, 13) = 1000 * (d + 1) + e
            For c = 1 To 13: a(r + 1, c) = a(r, c): Next c
            a(r + 1, 5) = "TEST_A": a(r + 1, 6) = "Synthetic check": a(r + 1, 7) = "ELEMENT_A"
            a(r + 1, 8) = "Synthetic element": a(r + 1, 9) = "Test type": a(r + 1, 10) = "MODERATE"
            a(r + 1, 11) = 0: a(r + 1, 12) = 1: a(r + 1, 13) = 0
            r = r + 2
        Next e
    Next d
    Set idx = IndexSource(a, h, 1, False, stats)
    JKB_Assert idx.count = 4, "Four independent date/entity combinations"
    JKB_Assert stats("Dates").count = 2 And stats("Entities").count = 2, "Date/entity lists"
    For Each k In idx.keys
        Set tc = idx(k): baseRow = tc("BaseRow")
        JKB_Assert CStr(a(baseRow, 2)) = CStr(tc("EntityID")), "Matching entity base row"
        JKB_Assert DateKey(a(baseRow, 1)) = tc("Date"), "Matching date base row"
        JKB_Assert tc("Elements")("ELEMENT_A")("SeverityRows").Exists("MODERATE"), "Zero is retained"
    Next k
    JKB_Assert DateKey(44560, True) = "2025-12-31", "1904 date conversion"
    For c = 1 To 13: a(10, c) = a(9, c): Next c
    Set idx = IndexSource(a, h, 1, False, stats)
    JKB_Assert idx.count = 4 And stats("Duplicates") = 1, "Exact duplicates are deduplicated"
    a(10, 13) = 99
    On Error Resume Next
    Set idx = IndexSource(a, h, 1, False, stats): errNo = Err.Number: Err.Clear
    On Error GoTo Failed
    JKB_Assert errNo = vbObjectError + 527, "Conflicting duplicates are rejected"
    For c = 1 To 13: a(10, c) = Empty: a(8, c) = Empty: Next c
    On Error Resume Next
    Set idx = IndexSource(a, h, 1, False, stats): errNo = Err.Number: Err.Clear
    On Error GoTo Failed
    JKB_Assert errNo = vbObjectError + 528, "Missing matching base is rejected"
    Set rm = NewMap(): rm("FOO") = 7: rm("FOOBAR") = 8: ResetFormulaCache
    f = ConvertManualFormula("@manual.FOO + manual.FOOBAR + ""manual.FOO""", rm, 0, "Native check")
    JKB_Assert f = "=$F$7 + $F$8 + ""manual.FOO""", "Formula prefixes and quoted text"
    f = ConvertManualFormula("@man.FOO + @differences.FOOBAR", rm, 0, "Reference alias check")
    JKB_Assert f = "=$F$7 + $I$8", "Manual/difference shorthand aliases resolve consistently"
    f = ConvertManualFormula("IFNULL(IFNULL(@manual.FOO,0),1)", rm, 0, "Native helper check")
    JKB_Assert InStr(1, f, "IFNULL", vbTextCompare) = 0, "Nested null helpers expand"
    f = ConvertManualFormula("IF @manual.FOO > 0 AND @manual.FOOBAR < 10 THEN @manual.FOO ELSE @manual.FOOBAR", rm, 0, "Natural IF AND check")
    JKB_Assert f = "=IF(AND($F$7 > 0,$F$8 < 10),$F$7,$F$8)", "Natural IF/THEN/ELSE and AND translate to Excel"
    f = ConvertManualFormula("IF @manual.FOO > 0 OR @manual.FOOBAR < 10 THEN 1 ELSE 0", rm, 0, "Natural IF OR check")
    JKB_Assert f = "=IF(OR($F$7 > 0,$F$8 < 10),1,0)", "Natural OR translates to Excel"
    f = ConvertManualFormula("IF @manual.FOO > 0 THEN IF @manual.FOOBAR > 0 THEN 1 ELSE 2 ELSE 3", rm, 0, "Nested natural IF check")
    JKB_Assert f = "=IF($F$7 > 0,IF($F$8 > 0,1,2),3)", "Nested natural ELSE IF logic translates"
    f = ConvertManualFormula("IF(@manual.FOO>0 AND @manual.FOOBAR<10,1,0)", rm, 0, "Excel IF plus infix AND check")
    JKB_Assert f = "=IF(AND($F$7>0,$F$8<10),1,0)", "Infix AND inside Excel IF translates"
    f = ConvertManualFormula("ROUND(IF @manual.FOO > 0 THEN 1 ELSE 2,0)", rm, 0, "Nested natural IF function-argument check")
    JKB_Assert f = "=ROUND(IF($F$7 > 0,1,2),0)", "Natural IF translates inside Excel function arguments"
    Set ws = ThisWorkbook.Worksheets.Add
    ws.name = UniqueSheetName(ThisWorkbook, "__JKB_Control_Check")
    ws.Range("A1").Value2 = "JKB": ws.Range("A7").Value2 = "Formula row": ws.Range("L7").Value2 = MARKER_TOGGLE_ROW
    ws.Range("A8").Value2 = "Ordinary row": ws.Range("F7").formula = "=1+1": ws.Calculate
    Sheet1.EnsureRibbon ws
    JKB_Assert ws.Shapes("JKB_Dashboard").Height <= 18, "Slim controls"
    ws.Activate: Sheet1.ToggleFormulaRows
    JKB_Assert ws.rows(7).hidden And Not ws.rows(8).hidden And ws.Range("F7").Value2 = 2, "Hide only marked formula rows"
    Sheet1.ToggleFormulaRows
    JKB_Assert Not ws.rows(7).hidden, "Show formula rows again"
    Sheet1.GoDashboard: JKB_Assert ActiveSheet.name = SHEET_DASHBOARD, "Dashboard navigation"
    Sheet1.GoConfig: JKB_Assert ActiveSheet.name = SHEET_RULES, "Config navigation"
    Sheet1.GoFormat: JKB_Assert ActiveSheet.name = SHEET_FORMATTING, "Formatting navigation"
    Sheet1.GoGroups: JKB_Assert ActiveSheet.name = SHEET_RULES And ActiveCell.Column = 25, "Group setup navigation"
    Set rm = NewMap(): rm("LCR_HQLA") = 1: rm("LCR_INFLOW") = 2: rm("LCR_OUTFLOW") = 3
    rm("LEG_LIQ_TOTAL_ASSETS") = 4: rm("LEG_LIQ_TOTAL_LIAB") = 5
    JKB_Assert ResolveSourceFieldKey(rm, "PRE_LCR_HQLA") = "LCR_HQLA", "Legacy HQLA source alias"
    JKB_Assert ResolveSourceFieldKey(rm, "PRE_LCR_INFLOW") = "LCR_INFLOW", "Legacy LCR inflow source alias"
    JKB_Assert ResolveSourceFieldKey(rm, "PRE_LCR_OUTFLOW") = "LCR_OUTFLOW", "Legacy LCR outflow source alias"
    JKB_Assert ResolveSourceFieldKey(rm, "PRE_LEG_LIQ_TOTAL_ASSETS") = "LEG_LIQ_TOTAL_ASSETS", "Legacy legal-liquidity asset alias"
    JKB_Assert ResolveSourceFieldKey(rm, "PRE_LEG_LIQ_TOTAL_LIAB") = "LEG_LIQ_TOTAL_LIAB", "Legacy legal-liquidity liability alias"
    JKB_Assert ResolveSourceFieldKey(rm, "UNKNOWN_FIELD") = "", "Unknown source field is not guessed"
    ws.Delete: Set ws = Nothing: previous.Activate
    RestoreState state: JKB_Busy = False: ResetFormulaCache
    UiNotice "Self-check", "All built-in checks passed.", "Date and entity matching, duplicates, missing bases, zero values, formula references, IF/ELSE/AND/OR translation, null helpers, source aliases, hide and show, and navigation."
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Delete
    If Not previous Is Nothing Then previous.Activate
    RestoreState state: JKB_Busy = False: ResetFormulaCache
    UiProblem "Self-check", "A built-in check failed.", er
End Sub
Private Sub JKB_Assert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 699, , message
End Sub
Public Sub ToggleFormulaRows()
    Sheet1.ToggleFormulaRows
End Sub
Public Sub ApplyExternalOutputUpgrade(Optional ByVal showConfirmation As Boolean = True)
    JKB_UpgradeWorkbook
End Sub

Public Sub StartScenarioCategorySelection()
    Dim a As Variant, h As Object, hr As Long, idx As Object, stats As Object, ws As Worksheet, state As Object, er As String
    If JKB_Busy Then Exit Sub
    On Error GoTo Failed
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    Set state = CaptureState(): Application.EnableEvents = False: Application.ScreenUpdating = False
    On Error Resume Next: Application.DefaultSheetDirection = xlLTR: On Error GoTo Failed
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_SELECTION)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SHEET_DASHBOARD)): ws.name = SHEET_SELECTION
    ws.Cells.UnMerge: ws.Cells.Clear
    Dim sh As Shape
    For Each sh In ws.Shapes: sh.Delete: Next sh
    ws.Range("A1").Value2 = "JKB"
    ws.Range("A3").Value2 = "Choose output coverage"
    ws.Range("A3").Font.Size = 16: ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value2 = "Select Yes for one or many items in each list. Double-click a Yes/No cell to switch it. Only available combinations are generated."
    ws.Range("A4").Font.Size = 10
    ws.columns("A").ColumnWidth = 9: ws.columns("B").ColumnWidth = 24: ws.columns("C").ColumnWidth = 3
    ws.columns("D").ColumnWidth = 9: ws.columns("E").ColumnWidth = 17: ws.columns("F").ColumnWidth = 3
    ws.columns("G").ColumnWidth = 9: ws.columns("H").ColumnWidth = 31: ws.columns("I").ColumnWidth = 3
    ws.columns("J").ColumnWidth = 9: ws.columns("K").ColumnWidth = 28: ws.columns("L").ColumnWidth = 48
    ws.columns("N:Q").hidden = True
    FillChoices ws, stats("Categories"), 1, "Category", 14
    FillChoices ws, stats("Dates"), 4, "As-of date", 15
    FillChoices ws, stats("Entities"), 7, "Entity", 16
    FillChoices ws, stats("Cases"), 10, "Test case", 17
    ws.UsedRange.Font.name = UI_FONT
    ws.UsedRange.rows.RowHeight = 20: ws.rows(4).RowHeight = 25: ws.rows(5).RowHeight = 25
    ws.rows(1).RowHeight = 22: ws.rows(2).RowHeight = 6
    Dim choiceCol As Variant
    For Each choiceCol In Array(1, 4, 7, 10)
        AddMasterButton ws, "All_" & choiceCol, "All", "SelectChoiceList", ws.Cells(5, choiceCol).Left, ws.rows(5).Top + 1, 48, 20
        AddMasterButton ws, "None_" & choiceCol, "None", "SelectChoiceList", ws.Cells(5, choiceCol).Left + 53, ws.rows(5).Top + 1, 48, 20
    Next choiceCol
    AddMasterButton ws, "JKB_Run", "Generate selected outputs", "ConfirmScenarioCategorySelection", 760, 44, 190, 25
    Sheet1.StyleSelectionSheet ws
    Sheet1.EnsureRibbon ws
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False: ActiveWindow.splitRow = 6: ActiveWindow.SplitColumn = 0: ActiveWindow.FreezePanes = True
    ActiveWindow.zoom = 85
    RestoreState state
    Exit Sub
Failed:
    er = Err.description: RestoreState state
    UiProblem "Choose outputs", "The output selection could not be opened.", er
End Sub

Private Sub FillChoices(ByVal ws As Worksheet, ByVal map As Object, ByVal col As Long, ByVal title As String, ByVal keyCol As Long)
    Dim keys As Variant, i As Long, j As Long, tmp As Variant, out As Variant, flags As Range
    keys = map.keys
    ' The lists are small; sorting here keeps selection stable between uploads.
    For i = 0 To map.count - 2
        For j = i + 1 To map.count - 1
            If StrComp(CStr(keys(i)), CStr(keys(j)), vbTextCompare) > 0 Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i
    ws.Cells(6, col).Value2 = "Include": ws.Cells(6, col + 1).Value2 = title
    ReDim out(1 To map.count, 1 To 2)
    For i = 0 To map.count - 1
        out(i + 1, 1) = "Yes"
        If col = 10 Then out(i + 1, 2) = keys(i) Else out(i + 1, 2) = map(keys(i))
        ws.Cells(i + 7, keyCol).NumberFormat = "@": ws.Cells(i + 7, keyCol).Value2 = keys(i)
        If col = 10 Then ws.Cells(i + 7, 12).Value2 = map(keys(i))
    Next i
    ws.Cells(7, col).Resize(map.count, 2).Value2 = out
    Set flags = ws.Cells(7, col).Resize(map.count, 1)
    flags.Interior.Color = RGB(253, 246, 227)
    flags.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
    flags.Validation.InCellDropdown = True: flags.Validation.ShowError = True
    ws.Cells(6, col).Resize(1, 2).Interior.Color = RGB(6, 118, 71)
    ws.Cells(6, col).Resize(1, 2).Font.Color = vbWhite
End Sub

Public Sub RebindSelectionActions()
    Dim ws As Worksheet, sh As Shape, proc As String
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_SELECTION)
    If ws Is Nothing Then Exit Sub
    For Each sh In ws.Shapes
        proc = ""
        If sh.name Like "All_*" Or sh.name Like "None_*" Then proc = "SelectChoiceList"
        If sh.name = "JKB_Run" Then proc = "ConfirmScenarioCategorySelection"
        If Len(proc) > 0 Then sh.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & proc
    Next sh
End Sub

Public Sub AddMasterButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal macro As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(5, x, y, w, h)
    s.name = nm: s.Placement = xlFreeFloating
    s.fill.ForeColor.RGB = RGB(6, 118, 71): s.line.Visible = msoTrue
    s.line.ForeColor.RGB = RGB(6, 118, 71): s.line.Weight = 0.5
    With s.TextFrame
        .Characters.text = caption: .Characters.Font.name = UI_FONT: .Characters.Font.Size = 9.5
        .Characters.Font.Color = vbWhite: .Characters.Font.Bold = True
        .MarginTop = 0: .MarginBottom = 0: .HorizontalAlignment = xlHAlignCenter: .VerticalAlignment = xlVAlignCenter
    End With
    s.OnAction = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & macro
End Sub

Public Sub SelectChoiceList()
    Dim parts As Variant, col As Long, lastRow As Long, ws As Worksheet
    If JKB_Busy Then Exit Sub
    If TypeName(Application.Caller) <> "String" Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SHEET_SELECTION)
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) <> 1 Then Exit Sub
    col = CLng(parts(1)): lastRow = ws.Cells(ws.rows.count, col + 1).End(xlUp).row
    If lastRow >= 7 Then ws.Cells(7, col).Resize(lastRow - 6, 1).Value2 = IIf(parts(0) = "All", "Yes", "No")
End Sub
Public Sub ClearRunSelection()
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_SELECTION)
    If ws Is Nothing Then Exit Sub
    ws.Delete
End Sub
Public Function ReadSelections() As Object
    Dim result As Object, map As Object, ws As Worksheet, cols As Variant, names As Variant
    Dim i As Long, r As Long, lastRow As Long, Flag As String, key As String, displayText As String
    Dim errNo As Long, er As String, sectionName As String
    On Error GoTo Failed
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_SELECTION)
    If ws Is Nothing Then Err.Raise vbObjectError + 592, , "The output-selection sheet is not available. Reopen Choose and generate outputs."
    Set result = NewMap()
    cols = Array(1, 4, 7, 10): names = Array("Categories", "Dates", "Entities", "Cases")
    For i = 0 To 3
        sectionName = CStr(names(i))
        Set map = NewMap()
        ' Use the visible description column to determine the list length. Hidden key columns are only identifiers.
        lastRow = ws.Cells(ws.rows.count, CLng(cols(i)) + 1).End(xlUp).row
        If lastRow < 7 Then Err.Raise vbObjectError + 593, , "No choices are available in " & sectionName & ". Reopen the selection screen."
        For r = 7 To lastRow
            Flag = SafeUpperText(ws.Cells(r, CLng(cols(i))).Value2)
            displayText = SafeText(ws.Cells(r, CLng(cols(i)) + 1).Value2)
            key = SafeText(ws.Cells(r, 14 + i).Value2)
            If Flag <> "YES" And Flag <> "NO" Then _
                Err.Raise vbObjectError + 591, , "Choose Yes or No in " & ws.Cells(r, CLng(cols(i))).address(False, False) & " (" & sectionName & ")."
            If Flag = "YES" Then
                If Len(key) = 0 Then _
                    Err.Raise vbObjectError + 594, , "The internal selection key is missing for " & sectionName & " row " & r & IIf(Len(displayText) > 0, " (" & displayText & ")", "") & ". Reopen the selection screen."
                map(key) = True
            End If
        Next r
        If map.count = 0 Then Err.Raise vbObjectError + 590, , "Choose at least one item in " & sectionName & "."
        If result.Exists(sectionName) Then result.Remove sectionName
        result.Add sectionName, map
    Next i
    Set ReadSelections = result
    Exit Function
Failed:
    errNo = Err.Number: er = Err.description
    Err.Raise errNo, "ReadSelections", IIf(Len(sectionName) > 0, sectionName & ": ", "") & er
End Function
Public Function IsSelected(ByVal tc As Object, ByVal selections As Object) As Boolean
    IsSelected = selections("Categories").Exists(tc("Category")) And selections("Dates").Exists(tc("Date")) And selections("Entities").Exists(tc("EntityKey")) And selections("Cases").Exists(tc("TestCaseCode"))
End Function
Public Sub ConfirmScenarioCategorySelection()
    Dim selections As Object, stage As String, errNo As Long, er As String
    If JKB_Busy Then Exit Sub
    On Error GoTo Failed
    stage = "reading the output selections"
    Set selections = ReadSelections()
    If selections Is Nothing Then Err.Raise vbObjectError + 595, , "The output selection could not be read."
    stage = "starting the generator"
    RunExternalScenarioOutput selections
    Exit Sub
Failed:
    errNo = Err.Number: er = Err.description
    On Error Resume Next
    LogIssue "ERROR", "GenerateSelection", "Step: " & stage & " | Error " & errNo & " | " & er, ""
    On Error GoTo 0
    UiProblem "Generate outputs", "Generation could not start while " & IIf(Len(stage) > 0, stage, "checking the selection") & ".", _
              er & " (error " & errNo & ")", "If the selection sheet was edited or partly cleared, reopen Choose and generate outputs."
End Sub
Public Sub ValidateScenarioBuilderSetup()
    Dim a As Variant, h As Object, hr As Long, stats As Object, idx As Object, rules As Object
    On Error GoTo Failed
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    Set rules = LoadRulesByElementType(ThisWorkbook)
    Sheet1.ValidateViewSettings
    UiNotice "Validate", "Everything checked is valid.", "Source structure, date and entity keys, and formatting settings."
    Exit Sub
Failed:
    UiProblem "Validate", "Validation found a problem.", Err.description
End Sub

' The sheets that make up the tool itself, as opposed to a generated element sheet
' or an engine sheet. The pre-shock workspace is six of them now, so the list is a
' function rather than a very long ElseIf.
Private Function IsWorkspaceSheet(ByVal nm As String) As Boolean
    Dim k As Variant
    Select Case nm
        Case SHEET_DASHBOARD, SHEET_RULES, SHEET_FORMATTING, SHEET_SOURCE, _
             SHEET_LOG, SHEET_SELECTION, "Manual_Formula_Guide"
            IsWorkspaceSheet = True
            Exit Function
    End Select
    For Each k In PsSheetNames()
        If nm = CStr(k) Then IsWorkspaceSheet = True: Exit Function
    Next k
End Function

Private Sub CleanMasterTabs()
    Const DETAIL_TEMPLATE As String = "_Detail_Template"
    Dim ws As Worksheet, marker As String
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = DETAIL_TEMPLATE Then
            ws.Visible = xlSheetVeryHidden
        ElseIf IsWorkspaceSheet(ws.name) Then
            ws.Visible = xlSheetVisible
        Else
            marker = CStr(ws.Range("A4").Value2)
            If Left$(marker, 14) = "Element type: " Then ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    On Error GoTo 0
End Sub

