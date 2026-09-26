Option Explicit

' ============================================================================
'  The small shared things, in one place so no module carries a private copy -
'  which is how two copies come to disagree.
' ============================================================================

' The log keeps this many entries; older ones fall off the bottom. A log
' nobody can scroll to the end of is not read.
Private Const LOG_KEEP As Long = 2000

Public Function NewMap() As Object
    Set NewMap = CreateObject("Scripting.Dictionary")
    NewMap.CompareMode = 1          ' TextCompare: keys differing only in case are one key
End Function

Public Function SafeText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeText = Trim$(CStr(v))
    ' The exports write a literal "NULL" for an empty cell, and carrying that
    ' through puts the word NULL on a pivot row label.
    If StrComp(SafeText, "NULL", vbTextCompare) = 0 Then SafeText = ""
    Err.Clear
End Function

Public Function SafeUpper(ByVal v As Variant) As String
    SafeUpper = UCase$(SafeText(v))
End Function

Public Function SafeNum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsNumeric(v) Then SafeNum = CDbl(v)
    Err.Clear
End Function

' A heading reduced to the one spelling everything compares against: upper case,
' every run of anything that is not a letter or digit becoming one underscore.
Public Function NormKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, lastU As Boolean
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch: lastU = False
        ElseIf Not lastU And Len(out) > 0 Then
            out = out & "_": lastU = True
        End If
    Next i
    If Right$(out, 1) = "_" Then out = Left$(out, Len(out) - 1)
    NormKey = out
End Function

Public Function Fmt(ByVal v As Double) As String
    Fmt = Format$(v, "#,##0")
End Function

Public Function FileLeaf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then FileLeaf = Mid$(p, i + 1) Else FileLeaf = p
End Function

Public Function FolderOf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then FolderOf = Left$(p, i - 1)
End Function

' The as-of date as a person writes it, whatever shape the cell handed over:
' a serial number (what .Value2 gives for a date), a Date, or text already.
Public Function AsOfText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then AsOfText = Format$(v, "d mmm yyyy"): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then
            AsOfText = Format$(CDate(CDbl(v)), "d mmm yyyy")
            Exit Function
        End If
    End If
    AsOfText = SafeText(v)
    Err.Clear
End Function

Public Function PathJoin(ByVal folder As String, ByVal leaf As String) As String
    If Right$(folder, 1) = "\" Then PathJoin = folder & leaf Else PathJoin = folder & "\" & leaf
End Function

' OneDrive hands out https:// paths for synced folders, and Workbooks.Open on
' one of those fails with "cannot access the file" rather than anything useful.
Public Function IsLocalPath(ByVal p As String) As Boolean
    If Len(p) < 3 Then Exit Function
    If LCase$(Left$(p, 4)) = "http" Then Exit Function
    IsLocalPath = True
End Function

Public Function GetSheet(ByVal nm As String, Optional ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set GetSheet = wb.Worksheets(nm)
    Err.Clear
End Function

Public Function EnsureSheet(ByVal nm As String, Optional ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = GetSheet(nm, wb)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        On Error Resume Next
        ws.Name = nm
        Err.Clear
    End If
    Set EnsureSheet = ws
End Function

Public Sub KillSheet(ByVal nm As String, Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = GetSheet(nm, wb)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Err.Clear
End Sub

Public Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    On Error Resume Next
    LastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    Err.Clear
End Function

' A sheet name Excel will accept: 31 characters, none of the seven it forbids.
Public Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, o As String
    o = s
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each b In bad
        o = Replace(o, CStr(b), " ")
    Next b
    Do While InStr(o, "  ") > 0
        o = Replace(o, "  ", " ")
    Loop
    o = Trim$(o)
    If Len(o) > 31 Then o = Trim$(Left$(o, 31))
    If Len(o) = 0 Then o = "sheet"
    SafeSheetName = o
End Function

' A name Excel will accept AND that is free in this workbook, remembering what
' it gave out so two callers cannot be handed the same one.
Public Function FreeSheetName(ByVal wanted As String, ByVal wb As Workbook) As String
    Dim stem As String, nm As String, i As Long
    stem = SafeSheetName(wanted)
    nm = stem
    i = 1
    Do While Not GetSheet(nm, wb) Is Nothing
        i = i + 1
        nm = SafeSheetName(Left$(stem, 28 - Len(CStr(i))) & " " & i)
        If i > 500 Then Exit Do
    Loop
    FreeSheetName = nm
End Function

' ===================== the activity log =====================================

Public Sub LogIt(ByVal level As String, ByVal stage As String, ByVal msg As String, _
                 Optional ByVal ctx As String = "")
    Dim ws As Worksheet, r As Long, lastR As Long
    On Error Resume Next
    Set ws = GetSheet(SH_LOG)
    If ws Is Nothing Then Exit Sub
    ' Newest first: the thing that just happened is the thing being looked for.
    '
    ' Formatted from BELOW. The default takes the row above - which is the
    ' black table header - and every entry came out looking like a heading.
    ws.Rows(modPD_Theme.R_FIRST).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    r = modPD_Theme.R_FIRST
    ws.Cells(r, 1).Value2 = Format$(Now, "dd mmm  hh:nn:ss")
    ws.Cells(r, 2).Value2 = level
    ws.Cells(r, 3).Value2 = stage
    ws.Cells(r, 4).Value2 = msg
    ws.Cells(r, 5).Value2 = ctx
    modPD_Theme.DressLogRow ws, r
    lastR = LastRow(ws, 4)
    If lastR > modPD_Theme.R_FIRST + LOG_KEEP Then
        ws.Range(ws.Rows(modPD_Theme.R_FIRST + LOG_KEEP), ws.Rows(lastR)).Delete
    End If
    Err.Clear
End Sub

' ===================== settings =============================================
'
' What the Desk remembers between sessions - which frameworks are switched on,
' where the last build went, what the last reconciliation found - lives on a
' very hidden sheet as plain key / value rows. A sheet rather than document
' properties or the registry: it travels with the workbook, a person can read
' it, and nothing about it needs permission.

Private Function SettingsSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(SH_SETTINGS)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SH_SETTINGS
        ws.Cells(1, 1).Value2 = "Key"
        ws.Cells(1, 2).Value2 = "Value"
        ws.visible = xlSheetVeryHidden
        Err.Clear
        On Error GoTo 0
    End If
    Set SettingsSheet = ws
End Function

Private Function SettingRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For r = 2 To lastR
        If StrComp(SafeText(ws.Cells(r, 1).Value2), key, vbTextCompare) = 0 Then SettingRow = r: Exit Function
    Next r
End Function

Public Function SettingGet(ByVal key As String, Optional ByVal dflt As String = "") As String
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    SettingGet = dflt
    Set ws = SettingsSheet()
    If ws Is Nothing Then Exit Function
    r = SettingRow(ws, key)
    If r > 0 Then SettingGet = CStr(ws.Cells(r, 2).Value2)
    Err.Clear
End Function

Public Sub SettingSet(ByVal key As String, ByVal value As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SettingsSheet()
    If ws Is Nothing Then Exit Sub
    r = SettingRow(ws, key)
    If r = 0 Then r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value2 = key
    ' Text, always: a folder called 2025 or a verdict called "1 of 2" must come
    ' back exactly as it went in, not as a number or a date.
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value2 = value
    Err.Clear
End Sub

' Every key starting with a prefix, gone - how a reset forgets results
' without forgetting preferences.
Public Sub SettingClear(ByVal prefix As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SettingsSheet()
    If ws Is Nothing Then Exit Sub
    For r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row To 2 Step -1
        If StrComp(Left$(SafeText(ws.Cells(r, 1).Value2), Len(prefix)), prefix, vbTextCompare) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r
    Err.Clear
End Sub

' Small numbers as words, the way a sentence reads them.
Public Function Words(ByVal n As Long) As String
    Select Case n
        Case 0: Words = "No"
        Case 1: Words = "One"
        Case 2: Words = "Two"
        Case 3: Words = "Three"
        Case 4: Words = "Four"
        Case 5: Words = "Five"
        Case Else: Words = CStr(n)
    End Select
End Function

' 412806 -> "412.8 k", 30077375161 -> "30.1 bn". For tiles, never tables.
Public Function Compact(ByVal v As Double) As String
    Dim a As Double
    a = Abs(v)
    Select Case True
        Case a >= 1000000000#: Compact = Format$(v / 1000000000#, "0.0") & " bn"
        Case a >= 1000000#: Compact = Format$(v / 1000000#, "0.0") & " m"
        Case a >= 10000#: Compact = Format$(v / 1000#, "0.0") & " k"
        Case Else: Compact = Format$(v, "#,##0")
    End Select
End Function

' A long file name shortened in the middle, so the start and the extension -
' the two parts a person recognises - both survive.
Public Function MidTrim(ByVal s As String, ByVal maxLen As Long) As String
    Dim keepR As Long, keepL As Long
    If Len(s) <= maxLen Or maxLen < 8 Then MidTrim = s: Exit Function
    keepR = 9
    keepL = maxLen - keepR - 1
    MidTrim = Left$(s, keepL) & ChrW(8230) & Right$(s, keepR)
End Function

' ===================== application state ====================================

Public Function CaptureState() As Object
    Dim s As Object
    Set s = NewMap()
    On Error Resume Next
    s("Calc") = Application.Calculation
    s("Screen") = Application.ScreenUpdating
    s("Events") = Application.EnableEvents
    s("Alerts") = Application.DisplayAlerts
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Err.Clear
    Set CaptureState = s
End Function

Public Sub RestoreState(ByVal s As Object)
    On Error Resume Next
    If s Is Nothing Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.StatusBar = False
        Exit Sub
    End If
    Application.Calculation = s("Calc")
    Application.ScreenUpdating = CBool(s("Screen"))
    Application.EnableEvents = CBool(s("Events"))
    Application.DisplayAlerts = CBool(s("Alerts"))
    Application.StatusBar = False
    Err.Clear
End Sub
