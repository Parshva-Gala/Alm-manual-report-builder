Attribute VB_Name = "modScenarioBuilder_Utils"
Option Explicit

Public Function SafeText(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeText = Trim$(CStr(v))
End Function
Public Function SafeUpperText(ByVal v As Variant) As String
    SafeUpperText = UCase$(SafeText(v))
End Function
Public Sub WriteLiteralValues(ByVal target As Range, ByRef values As Variant)
    Dim safe As Variant, r As Long, c As Long
    safe = values
    For r = 1 To UBound(safe, 1)
        For c = 1 To UBound(safe, 2)
            If VarType(safe(r, c)) = vbString Then
                If Left$(safe(r, c), 1) = "=" Then safe(r, c) = "'" & safe(r, c)
            End If
        Next c
    Next r
    target.Value2 = safe
End Sub
Public Function ParseBool(ByVal v As Variant) As Boolean
    Select Case SafeUpperText(v)
        Case "Y", "YES", "TRUE", "1", "ON": ParseBool = True
    End Select
End Function
Public Function GetWorksheetSafe(ByVal wb As Workbook, ByVal n As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetSafe = wb.Worksheets(n)
    On Error GoTo 0
End Function
Public Function SeverityList() As Variant
    SeverityList = Array("MODERATE", "MEDIUM", "SEVERE")
End Function
Public Function SectionColumn(ByVal s As String, ByVal sev As Long) As Long
    Select Case UCase$(s)
        Case "BASE": SectionColumn = 2
        Case "MAN", "MANUAL": SectionColumn = 6 + sev
        Case "DIFF", "DIFFERENCE", "DIFFERENCES": SectionColumn = 9 + sev
        Case Else: SectionColumn = 3 + sev
    End Select
End Function
Public Function colLetter(ByVal c As Long) As String
    Do While c > 0
        colLetter = Chr$(65 + (c - 1) Mod 26) & colLetter: c = (c - 1) \ 26
    Loop
End Function
Public Function NormalizeConfigFormulaText(ByVal v As Variant) As String
    Dim s As String
    s = SafeText(v)
    If Left$(s, 1) = "'" Then s = Mid$(s, 2)
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    s = Replace(s, ChrW(160), " ")
    ' Config cells are wrapped for readability, so a formula can carry hard line breaks.
    ' The translator treats them as ordinary separators; runs of spaces are left alone so
    ' quoted literals inside a formula keep their exact text.
    s = Replace(s, "_x000D_", " ")
    s = Replace(s, vbCrLf, " "): s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    NormalizeConfigFormulaText = Trim$(s)
End Function
Public Function RegexExecute(ByVal patternText As String, ByVal sourceText As String, Optional ByVal ignoreCase As Boolean = True) As Object
    Static cache As Object
    Dim re As Object, key As String
    If cache Is Nothing Then Set cache = NewMap()
    key = CStr(ignoreCase) & patternText
    If Not cache.Exists(key) Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True: re.ignoreCase = ignoreCase: re.pattern = patternText: cache.Add key, re
    End If
    Set RegexExecute = cache(key).Execute(sourceText)
End Function
Public Function CaptureState() As Object
    Dim s As Object: Set s = NewMap()
    s("Screen") = Application.ScreenUpdating: s("Events") = Application.EnableEvents
    s("Alerts") = Application.DisplayAlerts: s("Calc") = Application.Calculation
    s("Security") = Application.AutomationSecurity: s("Status") = Application.StatusBar
    s("Cancel") = Application.EnableCancelKey
    s("Direction") = Application.DefaultSheetDirection
    s("Interactive") = Application.interactive
    Set CaptureState = s
End Function
Public Sub RestoreState(ByVal s As Object)
    If s Is Nothing Then Exit Sub
    On Error Resume Next
    Application.Calculation = s("Calc"): Application.EnableEvents = s("Events")
    Application.ScreenUpdating = s("Screen"): Application.DisplayAlerts = s("Alerts")
    Application.AutomationSecurity = s("Security"): Application.StatusBar = s("Status")
    Application.EnableCancelKey = s("Cancel"): Application.CutCopyMode = False
    Application.DefaultSheetDirection = s("Direction")
    Application.interactive = s("Interactive")
    On Error GoTo 0
End Sub
Public Function UniqueSheetName(ByVal wb As Workbook, ByVal raw As String) As String
    Dim s As String, t As Variant, n As Long, candidate As String
    s = Trim$(raw)
    For Each t In Array("/", "\", "[", "]", ":", "*", "?", "'")
        s = Replace(s, CStr(t), "_")
    Next t
    If Len(s) = 0 Then s = "Scenario"
    candidate = Left$(s, 31)
    Do While Not GetWorksheetSafe(wb, candidate) Is Nothing
        n = n + 1: candidate = Left$(s, 26) & "_" & format$(n, "0000")
    Loop
    UniqueSheetName = candidate
End Function
Public Function FileComponent(ByVal raw As String) As String
    Dim s As String, i As Long, c As String
    For i = 1 To Len(raw)
        c = Mid$(raw, i, 1)
        If AscW(c) >= 32 And InStr(1, "\/:*?""<>|", c, vbBinaryCompare) = 0 Then s = s & c Else s = s & "_"
    Next i
    s = Trim$(s)
    Do While Len(s) > 0
        If Right$(s, 1) <> "." And Right$(s, 1) <> " " Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "Output"
    FileComponent = Left$(s, 60)
End Function
Public Function JoinPath(ByVal p As String, ByVal leaf As String) As String
    If Right$(p, 1) <> Application.PathSeparator Then p = p & Application.PathSeparator
    JoinPath = p & leaf
End Function
Public Sub EnsureFolder(ByVal p As String)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then fso.CreateFolder p
End Sub
Public Function CategoryMap() As Object
    Dim result As Object, ws As Worksheet, a As Variant, r As Long, n As Long
    Set result = NewMap(): Set ws = ThisWorkbook.Worksheets(SHEET_RULES)
    n = ws.Cells(ws.rows.count, 22).End(xlUp).row
    If n >= 7 Then
        a = ws.Range("V7:W" & n).Value2
        For r = 1 To UBound(a, 1)
            If Len(SafeText(a(r, 1))) > 0 Then result(SafeUpperText(a(r, 1))) = SafeText(a(r, 2))
        Next r
    End If
    Set CategoryMap = result
End Function
Public Function CategoryFor(ByVal code As String, ByVal map As Object) As String
    Dim p As String
    If map.Exists(code) Then CategoryFor = map(code): Exit Function
    p = SafeUpperText(Split(code & "_", "_")(0))
    Select Case p
        Case "CR": CategoryFor = "Credit Risk"
        Case "COR": CategoryFor = "Concentration"
        Case "LR": CategoryFor = "Liquidity"
        Case "MR": CategoryFor = "Market"
        Case "OR": CategoryFor = "Operational"
        Case "CC": CategoryFor = "Climate"
        Case "GP": CategoryFor = "Geopolitical"
        Case "MEF": CategoryFor = "MEF"
        Case Else: CategoryFor = "Additional"
    End Select
End Function
Public Function SettingValue(ByVal key As String, ByVal fallback As Variant) As Variant
    Dim a As Variant, r As Long
    a = ThisWorkbook.Worksheets(SHEET_FORMATTING).Range("A6:B24").Value2
    For r = 1 To UBound(a, 1)
        If SafeUpperText(a(r, 1)) = UCase$(key) Then
            If Len(SafeText(a(r, 2))) > 0 Then SettingValue = a(r, 2): Exit Function
        End If
    Next r
    SettingValue = fallback
End Function
Public Sub ResetBuildLog(ByVal wb As Workbook)
    Dim ws As Worksheet: Set ws = GetWorksheetSafe(wb, SHEET_LOG)
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = SHEET_LOG
    ws.Cells.UnMerge: ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB"
    ws.Range("A3:E3").value = Array("Level", "Procedure", "Message", "Context", "Logged at")
    ws.Range("A3:E3").Font.Bold = True
    ws.columns("A:B").ColumnWidth = 22: ws.columns("C:D").ColumnWidth = 60: ws.columns("E").ColumnWidth = 22
    Sheet1.EnsureRibbon ws
End Sub
Public Sub LogIssue(ByVal level As String, ByVal procedure As String, ByVal message As String, Optional ByVal context As String = "")
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    r = ws.Cells(ws.rows.count, 1).End(xlUp).row + 1
    ws.Cells(r, 1).Resize(1, 5).value = Array(level, procedure, message, context, Now)
    ws.Cells(r, 5).NumberFormat = "dd-mmm-yyyy hh:mm:ss"
End Sub
Public Function BuildLogHasErrors(ByVal wb As Workbook) As Boolean
    BuildLogHasErrors = (Application.CountIf(wb.Worksheets(SHEET_LOG).Range("A4:A" & wb.Worksheets(SHEET_LOG).Cells(wb.Worksheets(SHEET_LOG).rows.count, 1).End(xlUp).row), "ERROR") > 0)
End Function
