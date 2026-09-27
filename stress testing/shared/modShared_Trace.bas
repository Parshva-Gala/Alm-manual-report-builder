Attribute VB_Name = "modShared_Trace"
Option Explicit

' ============================================================================
'  TRACE - click a number, see what made it.
'
'  Every figure either tool produces is an aggregate of source rows. A trace puts
'  those rows on a sheet and builds a PivotTable over them, so the number stops
'  being something you take on faith.
'
'  On PivotTables, having removed them everywhere else: the objection was never
'  to pivots, it was to a pivot cache the size of a 480,000-row output, rebuilt
'  on every open. A TRACE is a few hundred to a few thousand rows, built when you
'  ask for it and thrown away when you are done. That is what a PivotTable is
'  good at, and here it earns its place - you can drag a field, expand a level,
'  and interrogate the evidence yourself.
'
'  A trace is two sheets:
'     Trace_<key>   what this number is, how it was derived, and the pivot
'     _TD_<key>     the evidence rows, very hidden until you ask for them
'
'  Shared by both tools, because "which rows made this number" is the same
'  question in both.
' ============================================================================

Public Const TRACE_PREFIX As String = "Trace_"
Public Const TRACE_DATA_PREFIX As String = "_TD_"
Public Const TRACE_MAX_ROWS As Long = 60000

Private Const T_TITLE_ROW As Long = 3
Private Const T_ABOUT_ROW As Long = 4
Private Const T_STATUS_ROW As Long = 6
Private Const T_PIVOT_ROW As Long = 8

' ===================== building a trace =====================================

' Opens a trace and hands back the DATA sheet to write evidence into: headers in
' row 1, rows from row 2. Call TraceEnd when the rows are written.
Public Function TraceBegin(ByVal key As String, ByVal title As String, ByVal about As String, _
                           ByVal narrative As String) As Worksheet
    Dim wsT As Worksheet, wsD As Worksheet
    Set wsT = ResetSheet(TraceName(key))
    Set wsD = ResetSheet(TraceDataName(key))

    wsT.Range("A1").Value2 = "TRACE"
    wsT.Cells(T_TITLE_ROW, 1).Value2 = title
    wsT.Cells(T_ABOUT_ROW, 1).Value2 = about
    wsT.Cells(T_STATUS_ROW, 1).Value2 = narrative
    wsD.Visible = xlSheetVeryHidden
    Set TraceBegin = wsD
End Function

' Builds the pivot over whatever was written and returns the trace sheet.
'
' rowFields / colFields are pipe-separated header names; either may be empty.
Public Function TraceEnd(ByVal key As String, ByVal rowFields As String, ByVal colFields As String, _
                         ByVal dataField As String, ByVal dataCaption As String) As Worksheet
    Dim wsT As Worksheet, wsD As Worksheet, lastR As Long, lastC As Long
    Dim pc As PivotCache, pt As PivotTable, er As String, n As Long

    Set wsT = GetSheetSafe(TraceName(key))
    Set wsD = GetSheetSafe(TraceDataName(key))
    If wsT Is Nothing Or wsD Is Nothing Then Exit Function

    lastR = wsD.Cells(wsD.rows.count, 1).End(xlUp).row
    lastC = wsD.Cells(1, wsD.columns.count).End(xlToLeft).Column
    n = lastR - 1

    If n <= 0 Then
        wsT.Cells(T_PIVOT_ROW, 1).Value2 = "No source row matched. Either the filter selects nothing, or the extract that would answer this is not loaded."
        StyleTrace wsT, 0
        Set TraceEnd = wsT
        Exit Function
    End If

    wsT.Cells(T_STATUS_ROW, 2).Value2 = format$(n, "#,##0") & " row(s) behind this figure"

    On Error GoTo NoPivot
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, _
             wsD.Range(wsD.Cells(1, 1), wsD.Cells(lastR, lastC)))
    Set pt = pc.CreatePivotTable(wsT.Cells(T_PIVOT_ROW, 1), "PT_" & SafeKey(key))
    AddFields pt, rowFields, xlRowField
    AddFields pt, colFields, xlColumnField
    If Len(dataField) > 0 Then
        If FieldExists(pt, dataField) Then
            With pt.AddDataField(pt.PivotFields(dataField), _
                 IIf(Len(dataCaption) > 0, dataCaption, "Sum of " & dataField), xlSum)
                .NumberFormat = "#,##0.00"
            End With
        End If
    End If
    On Error Resume Next
    pt.RowAxisLayout xlTabularRow
    pt.ShowDrillIndicators = True
    pt.TableStyle2 = "PivotStyleLight16"
    Err.Clear
    On Error GoTo 0

NoPivot:
    If Err.Number <> 0 Then
        er = Err.description
        Err.Clear
        On Error Resume Next
        wsT.Cells(T_PIVOT_ROW, 1).Value2 = "The pivot could not be built (" & er & "). The evidence rows are still there - click Show the rows."
        Err.Clear
        On Error GoTo 0
    End If
    StyleTrace wsT, n
    Set TraceEnd = wsT
End Function

Private Sub AddFields(ByVal pt As PivotTable, ByVal fields As String, ByVal orientation As Long)
    Dim p As Variant, pf As PivotField
    If Len(Trim$(fields)) = 0 Then Exit Sub
    For Each p In Split(fields, "|")
        If Len(Trim$(CStr(p))) > 0 Then
            If FieldExists(pt, Trim$(CStr(p))) Then
                On Error Resume Next
                Set pf = pt.PivotFields(Trim$(CStr(p)))
                pf.orientation = orientation
                ' Position AFTER Orientation, and to the CURRENT count - not count+1.
                ' Excel has already inserted the field by the time Orientation returns,
                ' so count+1 is outside the valid range and raises 1004. Position is
                ' presentation only, so a refusal must never stop the trace.
                If orientation = xlRowField Then
                    pf.position = pt.rowFields.count
                Else
                    pf.position = pt.ColumnFields.count
                End If
                pf.Subtotals(1) = True
                pf.Subtotals(1) = False
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next p
End Sub

Private Function FieldExists(ByVal pt As PivotTable, ByVal nm As String) As Boolean
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(nm)
    FieldExists = Not pf Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

' ===================== navigation and housekeeping ==========================

Public Sub TraceGo(ByVal key As String)
    Dim ws As Worksheet
    Set ws = GetSheetSafe(TraceName(key))
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range("A1").Select
    Err.Clear
End Sub

' The evidence rows, when the pivot is not enough.
Public Sub ShowTraceRows()
    Dim ws As Worksheet, key As String, wsD As Worksheet
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set ws = ActiveSheet
    If Left$(ws.name, Len(TRACE_PREFIX)) <> TRACE_PREFIX Then Exit Sub
    key = Mid$(ws.name, Len(TRACE_PREFIX) + 1)
    Set wsD = GetSheetSafe(TRACE_DATA_PREFIX & key)
    If wsD Is Nothing Then Exit Sub
    On Error Resume Next
    wsD.Visible = xlSheetVisible
    wsD.Activate
    wsD.Range("A1").Select
    Err.Clear
End Sub

' Traces are disposable by design. This removes every one of them.
Public Sub ClearAllTraces()
    Dim ws As Worksheet, n As Long, doomed As Collection, item As Variant
    Set doomed = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.name, Len(TRACE_PREFIX)) = TRACE_PREFIX Or _
           Left$(ws.name, Len(TRACE_DATA_PREFIX)) = TRACE_DATA_PREFIX Then
            doomed.Add ws.name
        End If
    Next ws
    If doomed.count = 0 Then
        MsgBox "There are no traces to clear.", vbInformation, "JKB Stress Testing"
        Exit Sub
    End If
    On Error Resume Next
    Application.DisplayAlerts = False
    For Each item In doomed
        ThisWorkbook.Worksheets(CStr(item)).Delete
        n = n + 1
    Next item
    Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0
    MsgBox n & " trace sheet(s) removed.", vbInformation, "JKB Stress Testing"
End Sub

Public Function TraceCount() As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.name, Len(TRACE_PREFIX)) = TRACE_PREFIX Then TraceCount = TraceCount + 1
    Next ws
End Function

' ===================== names and sheets =====================================

' Sheet names cap at 31 characters, and a trace key is built from codes that can
' be long. The key is trimmed and made unique by a short hash of the whole thing,
' so two long keys that share a prefix do not collide.
Public Function TraceName(ByVal key As String) As String
    TraceName = TRACE_PREFIX & ShortKey(key, 31 - Len(TRACE_PREFIX))
End Function

Public Function TraceDataName(ByVal key As String) As String
    TraceDataName = TRACE_DATA_PREFIX & ShortKey(key, 31 - Len(TRACE_DATA_PREFIX))
End Function

Private Function ShortKey(ByVal key As String, ByVal maxLen As Long) As String
    Dim s As String, h As Long, i As Long, ch As String
    For i = 1 To Len(key)
        ch = Mid$(key, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            s = s & ch
        Else
            If Right$(s, 1) <> "_" And Len(s) > 0 Then s = s & "_"
        End If
    Next i
    If Len(s) <= maxLen Then ShortKey = s: Exit Function
    For i = 1 To Len(key)
        h = ((h * 31) + AscW(Mid$(key, i, 1))) And &HFFFFF
    Next i
    ShortKey = Left$(s, maxLen - 5) & "_" & Right$("0000" & Hex$(h), 4)
End Function

Private Function SafeKey(ByVal key As String) As String
    SafeKey = ShortKey(key, 24)
End Function

Private Function GetSheetSafe(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set GetSheetSafe = ThisWorkbook.Worksheets(nm)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ResetSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheetSafe(nm)
    On Error Resume Next
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = nm
    Err.Clear
    On Error GoTo 0
    Set ResetSheet = ws
End Function

' ===================== appearance ===========================================

Private Sub StyleTrace(ByVal ws As Worksheet, ByVal n As Long)
    On Error Resume Next
    ws.Cells.Font.name = "Aptos"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(249, 250, 251)
    ws.rows(1).RowHeight = 6
    ws.rows(1).Font.Color = RGB(249, 250, 251)
    With ws.Cells(T_TITLE_ROW, 1)
        .Font.Size = 15
        .Font.Bold = True
        .Font.Color = RGB(18, 48, 107)
    End With
    ws.Cells(T_ABOUT_ROW, 1).Font.Color = RGB(102, 112, 133)
    ws.Cells(T_ABOUT_ROW, 1).WrapText = False
    With ws.Cells(T_STATUS_ROW, 1)
        .Font.Bold = True
        .Font.Size = 9.5
        .Font.Color = RGB(102, 112, 133)
    End With
    ws.Cells(T_STATUS_ROW, 2).Font.Color = RGB(102, 112, 133)
    ws.rows(T_TITLE_ROW).RowHeight = 26
    ws.rows(5).RowHeight = 7
    ws.columns("A").ColumnWidth = 46
    ws.columns("B:H").ColumnWidth = 18
    ws.Tab.Color = RGB(59, 99, 176)
    InstallTraceButtons ws
    Err.Clear
End Sub

Private Sub InstallTraceButtons(ByVal ws As Worksheet)
    TraceButton ws, "TR_Rows", "Show the rows", "ShowTraceRows", 6
    TraceButton ws, "TR_Clear", "Clear all traces", "ClearAllTraces", 108
End Sub

Private Sub TraceButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, _
                        ByVal proc As String, ByVal x As Double)
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes(nm)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, 2, 96, 17)
        sh.name = nm
    End If
    On Error Resume Next
    sh.Left = x: sh.Top = 2: sh.Width = 96: sh.Height = 17
    sh.Placement = xlFreeFloating
    sh.fill.ForeColor.RGB = RGB(255, 255, 255)
    sh.line.ForeColor.RGB = RGB(208, 213, 221)
    sh.line.Weight = 0.75
    sh.Shadow.Visible = msoFalse
    sh.TextFrame2.textRange.text = caption
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .textRange.ParagraphFormat.Alignment = msoAlignCenter
        .textRange.Font.Size = 8.5
        .textRange.Font.Bold = msoTrue
        .textRange.Font.fill.ForeColor.RGB = RGB(16, 24, 40)
    End With
    sh.OnAction = "'" & Replace(ws.Parent.name, "'", "''") & "'!" & proc
    Err.Clear
End Sub

' A number that can be traced says so, and clicking it goes there.
Public Sub MarkTraceable(ByVal cell As Range, ByVal key As String, ByVal tip As String)
    On Error Resume Next
    cell.Parent.Hyperlinks.Add anchor:=cell, address:="", _
        SubAddress:="'" & TraceName(key) & "'!A1", ScreenTip:=tip
    cell.Font.Underline = xlUnderlineStyleNone
    Err.Clear
End Sub
