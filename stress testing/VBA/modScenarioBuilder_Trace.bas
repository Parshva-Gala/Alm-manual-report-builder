Attribute VB_Name = "modScenarioBuilder_Trace"
Option Explicit

' ============================================================================
'  Tracing a derived figure back to the extract rows behind it.
'
'  A pre-shock number is a filtered sum. Until you can see the rows it summed,
'  it is a number you have to believe. Select a row on the results sheet or the
'  derived-values store and press Trace: the tool re-reads the extract, keeps the
'  rows the test case's filter selects, and builds a pivot over them by stage and
'  COA - which is the shape the ECL measures actually live in.
'
'  BASE and PRE-SHOCK trace to the same place with one difference: base applies
'  no filter, because that is what base means.
' ============================================================================

Public Sub TraceSelectedValue()
    Dim ws As Worksheet, r As Long, state As Object
    If JKB_Busy Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set ws = ActiveSheet
    r = ActiveCell.row
    If r < PS_FIRST_ROW Then
        UiNotice "Trace", "Select a row in the table first.", , "Then click Trace."
        Exit Sub
    End If

    On Error GoTo Failed
    Set state = CaptureState(): JKB_Busy = True
    Application.ScreenUpdating = False

    Select Case ws.name
        Case PS_RESULTS_SHEET: TraceFromResults ws, r
        Case DERIVED_SHEET:    TraceFromStore ws, r
        Case PS_CASES_SHEET:   TraceFromCase ws, r
        Case Else
            RestoreState state: JKB_Busy = False
            UiNotice "Trace", "There is nothing to trace on this sheet.", _
                     "Trace works from Base & pre-shock (" & DERIVED_SHEET & ") and from Test cases (" & PS_CASES_SHEET & ")."
            Exit Sub
    End Select

    RestoreState state: JKB_Busy = False
    Exit Sub
Failed:
    RestoreState state: JKB_Busy = False
    LogIssue LOG_LEVEL_ERROR, "Trace", Err.description, ws.name & " row " & r
    UiProblem "Trace", "The trace could not be built.", Err.description
End Sub

' Results: As-of | Entity | Test case | Element | Metric | ...
Private Sub TraceFromResults(ByVal ws As Worksheet, ByVal r As Long)
    RunValueTrace SafeText(ws.Cells(r, 3).Value2), SafeText(ws.Cells(r, 4).Value2), _
                  SafeUpperText(ws.Cells(r, 5).Value2), AskBaseOrPre()
End Sub

' The store: As-of | Entity | Entity ID | Test case | Element | Metric | Source | ...
Private Sub TraceFromStore(ByVal ws As Worksheet, ByVal r As Long)
    RunValueTrace SafeText(ws.Cells(r, 4).Value2), SafeText(ws.Cells(r, 5).Value2), _
                  SafeUpperText(ws.Cells(r, 6).Value2), AskBaseOrPre()
End Sub

' A test case row traces its whole pre-shock population, across every metric.
Private Sub TraceFromCase(ByVal ws As Worksheet, ByVal r As Long)
    RunValueTrace SafeText(ws.Cells(r, 2).Value2), SafeText(ws.Cells(r, 3).Value2), "", "PRE"
End Sub

Private Function AskBaseOrPre() As String
    Dim a As VbMsgBoxResult
    a = MsgBox("Trace the PRE-SHOCK population - the rows this test case's filter selects?" & vbCrLf & vbCrLf & _
               "Yes  = pre-shock (filtered)" & vbCrLf & _
               "No   = base (the whole portfolio, unfiltered)", _
               vbQuestion + vbYesNoCancel, "What should this trace show?")
    If a = vbYes Then AskBaseOrPre = "PRE"
    If a = vbNo Then AskBaseOrPre = "BASE"
End Function

' ===================== the collector ========================================

Private Sub RunValueTrace(ByVal testCase As String, ByVal element As String, ByVal metric As String, _
                          ByVal mode As String)
    Dim expr As String, sourceKey As String, key As String, wsD As Worksheet
    Dim title As String, subtitle As String, narrative As String, n As Long

    If Len(mode) = 0 Then Exit Sub
    If Len(testCase) = 0 Then Exit Sub

    ' Which extract answers this? The metric's own rule says so; without a metric
    ' the ECL extract is the one that carries the filter dimensions.
    sourceKey = "ECL"
    If Len(metric) > 0 Then sourceKey = SourceForMetric(metric)

    expr = ""
    If mode = "PRE" Then
        expr = FilterForCase(testCase, element)
        If Len(expr) = 0 Then
            UiNotice "Trace", "That test case has no filter condition, so its pre-shock population is the whole portfolio.", , "Trace the base instead."
            Exit Sub
        End If
    End If

    key = testCase & "_" & element & "_" & IIf(Len(metric) > 0, metric, "ALL") & "_" & mode
    title = IIf(mode = "BASE", "Base", "Pre-shock") & IIf(Len(metric) > 0, " - " & metric, "")
    subtitle = testCase & " / " & element & "   |   source: " & sourceKey
    If mode = "BASE" Then
        narrative = "The bank's whole portfolio from the " & sourceKey & " extract, with no filter applied. " & _
                    "That is what base means, and it is the same for every test case at this reporting date."
    Else
        narrative = "The rows this test case's filter selects from the " & sourceKey & " extract. " & _
                    "Identical across Moderate, Medium and Severe - the filter does not vary by severity.    FILTER:  " & expr
    End If

    Set wsD = TraceBegin(key, title, subtitle, narrative)
    n = CollectRows(sourceKey, expr, wsD)

    TraceEnd key, "STAGE|COA_CODE", "", MeasureForMetric(metric), _
             IIf(Len(metric) > 0, metric, "Outstanding")
    TraceGo key
    If n = 0 Then
        UiNotice "Trace", "No row of the " & sourceKey & " input file matched.", "Either the filter selects nothing, or a different file is loaded."
    End If
End Sub

' Re-reads the loaded extract and writes the matching rows out. The extract is
' already in memory when it has been uploaded this session; when it has not, the
' loader reopens it from the recorded path.
Private Function CollectRows(ByVal sourceKey As String, ByVal expr As String, ByVal wsD As Worksheet) As Long
    Dim data As Variant, headers As Object, node As Object, needs As String
    Dim r As Long, c As Long, nCols As Long, kept As Long, buf As Long, outRow As Long
    Dim block As Variant, k As Variant, cols() As Long, names() As String, i As Long

    If Not ActivateSourceForTrace(sourceKey) Then Exit Function
    data = CurrentSourceData()
    Set headers = CurrentSourceHeaders()
    If Not IsArray(data) Then Exit Function
    If headers Is Nothing Then Exit Function

    ' Every column the extract carries, so the pivot can be re-arranged into
    ' anything the reviewer wants without going back to the file.
    nCols = WorksheetFunction.Min(UBound(data, 2), 60)
    ReDim cols(1 To nCols)
    ReDim names(1 To nCols)
    For c = 1 To nCols
        cols(c) = c
        names(c) = SafeUpperText(data(1, c))
        If Len(names(c)) = 0 Then names(c) = "COL" & c
        wsD.Cells(1, c).Value2 = names(c)
    Next c
    wsD.rows(1).Font.Bold = True

    Set node = Nothing
    If Len(expr) > 0 Then
        needs = ""
        Set node = CompileTraceFilter(expr, needs)
        If Len(needs) > 0 Then
            wsD.Cells(2, 1).Value2 = "This filter needs a field the extract does not carry: " & needs
            Exit Function
        End If
    End If

    outRow = 2
    ReDim block(1 To 2000, 1 To nCols)
    For r = 2 To UBound(data, 1)
        If node Is Nothing Or EvalTraceFilter(node, data, r) Then
            If kept < TRACE_MAX_ROWS Then
                buf = buf + 1
                For c = 1 To nCols
                    block(buf, c) = data(r, c)
                Next c
                If buf = 2000 Then
                    wsD.Cells(outRow, 1).Resize(buf, nCols).Value2 = block
                    outRow = outRow + buf
                    buf = 0
                End If
            End If
            kept = kept + 1
        End If
    Next r
    If buf > 0 Then
        Dim part As Variant
        ReDim part(1 To buf, 1 To nCols)
        For i = 1 To buf
            For c = 1 To nCols
                part(i, c) = block(i, c)
            Next c
        Next i
        wsD.Cells(outRow, 1).Resize(buf, nCols).Value2 = part
    End If

    On Error Resume Next
    wsD.columns("A:BH").ColumnWidth = 20
    wsD.rows(1).AutoFilter
    Err.Clear
    On Error GoTo 0
    CollectRows = kept
End Function
