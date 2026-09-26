Option Explicit

' ============================================================================
'  The small shared things, in one place so no module carries a private copy -
'  which is how two copies come to disagree.
' ============================================================================

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

Public Function GetSheet(ByVal nm As String, Optional ByVal wb As Workbook = Nothing) As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set GetSheet = wb.Worksheets(nm)
    Err.Clear
End Function

Public Function EnsureSheet(ByVal nm As String, Optional ByVal wb As Workbook = Nothing) As Worksheet
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

Public Sub KillSheet(ByVal nm As String, Optional ByVal wb As Workbook = Nothing)
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
    Dim base As String, nm As String, i As Long
    base = SafeSheetName(wanted)
    nm = base
    i = 1
    Do While Not GetSheet(nm, wb) Is Nothing
        i = i + 1
        nm = SafeSheetName(Left$(base, 28 - Len(CStr(i))) & " " & i)
        If i > 500 Then Exit Do
    Loop
    FreeSheetName = nm
End Function

' ===================== the activity log =====================================

Public Sub LogIt(ByVal level As String, ByVal stage As String, ByVal msg As String, _
                 Optional ByVal ctx As String = "")
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetSheet(SH_LOG)
    If ws Is Nothing Then Exit Sub
    ' Newest first: the thing that just happened is the thing being looked for.
    ws.Rows(modPD_Theme.R_FIRST).Insert Shift:=xlDown
    r = modPD_Theme.R_FIRST
    ws.Cells(r, 1).Value2 = Format$(Now, "dd mmm  hh:nn:ss")
    ws.Cells(r, 2).Value2 = level
    ws.Cells(r, 3).Value2 = stage
    ws.Cells(r, 4).Value2 = msg
    ws.Cells(r, 5).Value2 = ctx
    modPD_Theme.PaintVerdict ws.Cells(r, 2)
    Err.Clear
End Sub

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
