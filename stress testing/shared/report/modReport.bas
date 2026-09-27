Attribute VB_Name = "modReport"
Option Explicit

' The first row of the pivot, and the freeze line. Three control strips and a
' slicer band sit above it. Must agree with modShared_Pivot.RPT_PIVOT_ROW.
Public Const RPT_FIRST_ROW As Long = 12

' Where the generators leave the toggle field names on the Data sheet.
Private Const TOG_ROW_COL As Long = 23      ' W - one row field per line
Private Const TOG_COL_COL As Long = 24      ' X - one column field per line

' ============================================================================
'  The controls that live INSIDE a generated report workbook.
'
'  A report workbook is not part of either tool - it is handed to a reviewer who
'  will open it on their own machine, months later, with no folder of extracts
'  and no interest in where it came from. So everything it needs travels with
'  it: this module is baked into the template the workbooks are born from.
'
'  It knows NOTHING about ALM or stress testing. Each generator writes its views
'  onto the Data sheet, in column V, one per line:
'
'      V1   rows=Type|Line|Subline|COA
'      V2   cols=Bucket
'      V3   values=Pre-factor|Post-factor
'
'  and draws buttons that call RptView1, RptView2, RptView3. So the same routines
'  serve a liquidity balance sheet and a pre-shock reconciliation, and adding a
'  view to either tool is a line of data rather than a routine. Any field the
'  pivot does not have is skipped, which is what lets one strip be drawn over
'  reports of different shapes.
'
'  The indirection is not decoration. Shape.OnAction silently REFUSES a macro
'  name with an argument - the property reads back empty and the button does
'  nothing, with no error anywhere - so the argument has to travel some other
'  way, and the sheet the report is already built on is the obvious one.
'
'  If the template is missing the workbooks are still produced - without the
'  button strip, but with their slicers, their field list and their drill-down
'  intact, because those are Excel's own and need no macro at all.
' ============================================================================

' ===================== the views ============================================
'
' Twelve thin wrappers because a button cannot pass an argument. Each one applies
' the view written on line N of column V of the Data sheet.

Public Sub RptView1()
    ApplyView 1
End Sub
Public Sub RptView2()
    ApplyView 2
End Sub
Public Sub RptView3()
    ApplyView 3
End Sub
Public Sub RptView4()
    ApplyView 4
End Sub
Public Sub RptView5()
    ApplyView 5
End Sub
Public Sub RptView6()
    ApplyView 6
End Sub
Public Sub RptView7()
    ApplyView 7
End Sub
Public Sub RptView8()
    ApplyView 8
End Sub
Public Sub RptView9()
    ApplyView 9
End Sub
Public Sub RptView10()
    ApplyView 10
End Sub
Public Sub RptView11()
    ApplyView 11
End Sub
Public Sub RptView12()
    ApplyView 12
End Sub

' "rows=A|B|C", "cols=A|B" or "values=A|B" on line n of column V.
Private Sub ApplyView(ByVal n As Long)
    Dim ws As Worksheet, pt As PivotTable, line As String, eq As Long
    Dim axis As String, spec As String
    Set pt = ActivePivot()
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    If ws Is Nothing Then Exit Sub
    line = CStr(ws.Cells(n, 22).Value2)
    Err.Clear
    eq = InStr(line, "=")
    If eq < 2 Then Exit Sub
    axis = LCase$(Trim$(Left$(line, eq - 1)))
    spec = Mid$(line, eq + 1)
    Select Case axis
        Case "rows"
            SetAxis pt, xlRowField, Split(spec, "|")
            OrderFromSheet pt
        Case "cols"
            SetAxis pt, xlColumnField, Split(spec, "|")
            OrderFromSheet pt
        Case "values"
            SetValues pt, Split(spec, "|")
    End Select
    ' A preset rearranges the very axes the toggles report on, so their captions
    ' are now stale unless they are restated here.
    SyncToggleCaptions ActiveSheet
    Err.Clear
End Sub

' ===================== one field at a time ==================================
'
' The views above are whole arrangements: press one and the rows become type,
' line, subline, COA. That is the right control when you know which arrangement
' you want, and the wrong one when you want the arrangement you are looking at
' plus the subline column - which is most of the time, and which previously meant
' opening the field list and dragging.
'
' So each of these adds ONE field to an axis, or takes it away again, leaving
' everything else where it is. The names come from column W (rows) and column X
' (columns) of the Data sheet, so which fields are offered is the generator's
' choice and not this module's business.

Public Sub RptTogRow1()
    ToggleField 1, xlRowField
End Sub
Public Sub RptTogRow2()
    ToggleField 2, xlRowField
End Sub
Public Sub RptTogRow3()
    ToggleField 3, xlRowField
End Sub
Public Sub RptTogRow4()
    ToggleField 4, xlRowField
End Sub
Public Sub RptTogRow5()
    ToggleField 5, xlRowField
End Sub
Public Sub RptTogRow6()
    ToggleField 6, xlRowField
End Sub
Public Sub RptTogRow7()
    ToggleField 7, xlRowField
End Sub
Public Sub RptTogRow8()
    ToggleField 8, xlRowField
End Sub
Public Sub RptTogCol1()
    ToggleField 1, xlColumnField
End Sub
Public Sub RptTogCol2()
    ToggleField 2, xlColumnField
End Sub
Public Sub RptTogCol3()
    ToggleField 3, xlColumnField
End Sub
Public Sub RptTogCol4()
    ToggleField 4, xlColumnField
End Sub

' Puts every toggle back to nothing, so a rearrangement starts from a clean axis
' rather than from whatever the last eight presses left behind.
Public Sub RptClearRows()
    ClearAxis xlRowField
End Sub

Public Sub RptClearCols()
    ClearAxis xlColumnField
End Sub

Private Sub ClearAxis(ByVal orientation As Long)
    Dim pt As PivotTable, doomed As Collection, pf As PivotField, item As Variant
    Set pt = ActivePivot()
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    pt.ManualUpdate = True
    Set doomed = New Collection
    If orientation = xlRowField Then
        For Each pf In pt.RowFields
            doomed.Add pf.Name
        Next pf
    Else
        For Each pf In pt.ColumnFields
            doomed.Add pf.Name
        Next pf
    End If
    For Each item In doomed
        pt.PivotFields(CStr(item)).orientation = xlHidden
    Next item
    pt.ManualUpdate = False
    Application.ScreenUpdating = True
    SyncToggleCaptions ActiveSheet
    Err.Clear
End Sub

' On becomes off, off becomes on - and on means "at the end of the axis", because
' a field added in the middle of a hierarchy changes what every subtotal beneath
' it means.
Private Sub ToggleField(ByVal n As Long, ByVal orientation As Long)
    Dim pt As PivotTable, nm As String, pf As PivotField
    Set pt = ActivePivot()
    If pt Is Nothing Then Exit Sub
    nm = ToggleName(n, orientation)
    If Len(nm) = 0 Then Exit Sub

    On Error Resume Next
    Set pf = Nothing
    Set pf = pt.PivotFields(nm)
    ' Not every report is built on every field. Saying nothing is better than
    ' raising an error at somebody who only pressed a button that was offered.
    If pf Is Nothing Then Err.Clear: Exit Sub

    Application.ScreenUpdating = False
    pt.ManualUpdate = True
    If OnAxis(pt, nm, orientation) Then
        pf.orientation = xlHidden
    Else
        pf.orientation = orientation
        If orientation = xlRowField Then
            pf.Position = pt.RowFields.Count
        Else
            pf.Position = pt.ColumnFields.Count
        End If
        ' Subtotals off for the field just added, the same as the views do, or
        ' one toggle puts a subtotal row under every line of the table.
        pf.Subtotals(1) = True
        pf.Subtotals(1) = False
    End If
    pt.ManualUpdate = False
    OrderFromSheet pt
    Application.ScreenUpdating = True
    SyncToggleCaptions ActiveSheet
    Err.Clear
End Sub

Private Function OnAxis(ByVal pt As PivotTable, ByVal nm As String, ByVal orientation As Long) As Boolean
    Dim pf As PivotField
    On Error Resume Next
    If orientation = xlRowField Then
        For Each pf In pt.RowFields
            If StrComp(pf.Name, nm, vbTextCompare) = 0 Then OnAxis = True: Exit Function
        Next pf
    Else
        For Each pf In pt.ColumnFields
            If StrComp(pf.Name, nm, vbTextCompare) = 0 Then OnAxis = True: Exit Function
        Next pf
    End If
    Err.Clear
End Function

Private Function ToggleName(ByVal n As Long, ByVal orientation As Long) As String
    Dim ws As Worksheet, col As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    If ws Is Nothing Then Exit Function
    col = IIf(orientation = xlRowField, TOG_ROW_COL, TOG_COL_COL)
    ToggleName = Trim$(CStr(ws.Cells(n, col).Value2))
    Err.Clear
End Function

' A toggle whose button does not say whether it is on is a button you have to
' press to find out - so the caption carries the state, and it is restated after
' anything that could have changed it, including the preset views.
Public Sub SyncToggleCaptions(ByVal ws As Worksheet)
    Dim pt As PivotTable, i As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If ws.PivotTables.Count = 0 Then Exit Sub
    Set pt = ws.PivotTables(1)
    For i = 1 To 8
        MarkToggle ws, pt, "tog_R" & i, ToggleName(i, xlRowField), xlRowField
    Next i
    For i = 1 To 4
        MarkToggle ws, pt, "tog_C" & i, ToggleName(i, xlColumnField), xlColumnField
    Next i
    Err.Clear
End Sub

Private Sub MarkToggle(ByVal ws As Worksheet, ByVal pt As PivotTable, ByVal shapeName As String, _
                       ByVal fieldName As String, ByVal orientation As Long)
    Dim sh As Shape, on_ As Boolean
    On Error Resume Next
    Set sh = Nothing
    Set sh = ws.Shapes(shapeName)
    If sh Is Nothing Then Err.Clear: Exit Sub
    If Len(fieldName) = 0 Then sh.Visible = msoFalse: Err.Clear: Exit Sub
    sh.Visible = msoTrue
    on_ = OnAxis(pt, fieldName, orientation)
    sh.TextFrame2.TextRange.Text = IIf(on_, ChrW(10003) & " ", "+ ") & fieldName
    ' On is filled, off is outlined. Colour carries it at a glance; the tick
    ' carries it for anyone who cannot rely on the colour.
    If on_ Then
        sh.Fill.ForeColor.RGB = RGB(25, 63, 137)
        sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        sh.Fill.ForeColor.RGB = RGB(245, 247, 250)
        sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(25, 63, 137)
    End If
    Err.Clear
End Sub

' ===================== drilling =============================================

Public Sub RptExpand()
    Dim pt As PivotTable, i As Long
    Set pt = ActivePivot()
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    pt.ManualUpdate = True
    For i = 1 To pt.RowFields.Count
        pt.RowFields(i).ShowDetail = True
    Next i
    pt.ManualUpdate = False
    Application.ScreenUpdating = True
    Err.Clear
End Sub

Public Sub RptCollapse()
    Dim pt As PivotTable, i As Long
    Set pt = ActivePivot()
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    pt.ManualUpdate = True
    For i = 1 To pt.RowFields.Count
        pt.RowFields(i).ShowDetail = (i = 1)
    Next i
    pt.ManualUpdate = False
    Application.ScreenUpdating = True
    Err.Clear
End Sub

' Totals per outer level, nothing under a leaf - the shape a return has.
Public Sub RptSubtotals()
    Dim pt As PivotTable, i As Long, wanted As Boolean
    Set pt = ActivePivot()
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    wanted = Not HasSubtotals(pt)
    For i = 1 To pt.RowFields.Count
        pt.RowFields(i).Subtotals(1) = True
        pt.RowFields(i).Subtotals(1) = (wanted And i < pt.RowFields.Count - 1)
    Next i
    Err.Clear
End Sub

Private Function HasSubtotals(ByVal pt As PivotTable) As Boolean
    On Error Resume Next
    If pt.RowFields.Count > 0 Then HasSubtotals = pt.RowFields(1).Subtotals(1)
    Err.Clear
End Function

' ===================== zoom =================================================
'
' Excel's own zoom is a slider in a corner and a modifier-plus-wheel nobody
' remembers. These are the same thing where the eye already is.

Public Sub RptZoomIn()
    ZoomBy 10
End Sub

Public Sub RptZoomOut()
    ZoomBy -10
End Sub

' Fit the table to the window - the useful zoom, and the one Excel hides.
Public Sub RptZoomFit()
    Dim ws As Worksheet, keep As Range
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Set keep = Selection
    ws.UsedRange.Select
    ActiveWindow.Zoom = True
    If ActiveWindow.Zoom > 130 Then ActiveWindow.Zoom = 130
    If ActiveWindow.Zoom < 40 Then ActiveWindow.Zoom = 40
    If Not keep Is Nothing Then keep.Select
    Err.Clear
End Sub

Public Sub RptZoom100()
    On Error Resume Next
    ActiveWindow.Zoom = 100
    Err.Clear
End Sub

Private Sub ZoomBy(ByVal delta As Long)
    Dim z As Long
    On Error Resume Next
    z = ActiveWindow.Zoom + delta
    If z < 30 Then z = 30
    If z > 250 Then z = 250
    ActiveWindow.Zoom = z
    Err.Clear
End Sub

' ===================== reading and getting about ============================

Public Sub RptIndex()
    GoSheet "Index"
End Sub

Public Sub RptData()
    GoSheet "Data"
End Sub

Private Sub GoSheet(ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVisible
    ws.Activate
    ActiveSheet.Range("A1").Select
    Err.Clear
End Sub

' Every pivot in the workbook shares one cache, so one refresh is all of them.
Public Sub RptRefresh()
    On Error Resume Next
    ThisWorkbook.RefreshAll
    Err.Clear
End Sub

' Columns sized to what is in them, with a ceiling: one very long condition
' otherwise makes a column wider than the screen and every other one useless.
Public Sub RptFitColumns()
    Dim ws As Worksheet, c As Long, lastC As Long
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lastC > 60 Then lastC = 60
    For c = 1 To lastC
        If Not ws.Columns(c).Hidden Then
            ws.Columns(c).AutoFit
            If ws.Columns(c).ColumnWidth > 70 Then ws.Columns(c).ColumnWidth = 70
            If ws.Columns(c).ColumnWidth < 7 Then ws.Columns(c).ColumnWidth = 7
        End If
    Next c
    Application.ScreenUpdating = True
    Err.Clear
End Sub

' Rows sized to what is in them - the other half of fitting a table, and the
' half that matters once a column has been narrowed and its text has wrapped.
'
' Only the table's own rows. The band at the top is stated geometry and autofit
' would collapse it, taking the buttons and slicers out of alignment with it.
Public Sub RptFitRows()
    Dim ws As Worksheet, lastR As Long, rng As Range
    On Error Resume Next
    Set ws = ActiveSheet
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastR > RPT_FIRST_ROW Then
        Set rng = ws.Range(ws.Rows(RPT_FIRST_ROW), ws.Rows(lastR))
        rng.AutoFit
        ' AutoFit on a wrapped cell can produce a row several inches tall; cap it.
        For lastR = lastR To RPT_FIRST_ROW Step -1
            If ws.Rows(lastR).RowHeight > 60 Then ws.Rows(lastR).RowHeight = 60
        Next lastR
    End If
    LayoutBand ws
    Application.ScreenUpdating = True
    Err.Clear
End Sub

Public Sub RptFitBoth()
    RptFitColumns
    RptFitRows
End Sub

' The stated geometry of the band at the top of every report sheet. It is
' restated whenever anything might have disturbed it, because the buttons and
' slicers are positioned in points and the band is what those points mean.
' Must agree with modShared_Pivot.LayoutReportSheet in the tool that wrote this
' workbook. Three strips, then a gap, then the slicer band, then the pivot at
' row 12 - 141pt in total, which is what the reader loses off the top of the
' screen before the first figure appears.
Public Sub LayoutBand(ByVal ws As Worksheet)
    Dim r As Long
    On Error Resume Next
    ws.Rows(1).RowHeight = 18
    ws.Rows(2).RowHeight = 12
    ws.Rows(3).RowHeight = 16
    ws.Rows(4).RowHeight = 16
    ws.Rows(5).RowHeight = 16
    ws.Rows(6).RowHeight = 3
    For r = 7 To 10
        ws.Rows(r).RowHeight = 14
    Next r
    ws.Rows(11).RowHeight = 4
    Err.Clear
End Sub

' Freezes the control band on every report sheet, so using a button or a slicer
' never means scrolling up to reach it and back down to see what it did.
'
' Called from Workbook_Open because freezing needs a real window: the workbook is
' generated by an invisible Excel, which has none, so it does this to itself the
' first time a person opens it.
Public Sub RptFreezeAll()
    Dim ws As Worksheet, keep As Worksheet
    On Error Resume Next
    If Not Application.Visible Then Exit Sub
    If ThisWorkbook.Windows.Count = 0 Then Exit Sub
    Set keep = ActiveSheet
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible And ws.PivotTables.Count > 0 Then
            LayoutBand ws
            ws.Activate
            ws.Range("A1").Select
            ActiveWindow.FreezePanes = False
            ws.Cells(RPT_FIRST_ROW, 1).Select
            ActiveWindow.FreezePanes = True
            ws.Range("A1").Select
            SyncToggleCaptions ws
        End If
    Next ws
    If Not keep Is Nothing Then keep.Activate
    Application.ScreenUpdating = True
    Err.Clear
End Sub

' What is on screen, on the clipboard - which is how a finding reaches an email.
Public Sub RptCopyView()
    On Error Resume Next
    ActiveSheet.UsedRange.Copy
    Err.Clear
End Sub

' ===================== the machinery ========================================

Private Function ActivePivot() As PivotTable
    On Error Resume Next
    If ActiveSheet.PivotTables.Count > 0 Then Set ActivePivot = ActiveSheet.PivotTables(1)
    Err.Clear
End Function

' Clears one axis and lays the named fields back onto it in order.
'
' Position is set AFTER Orientation and to the CURRENT count: Excel has already
' inserted the field by the time Orientation returns, so count + 1 is out of
' range and raises 1004.
Private Sub SetAxis(ByVal pt As PivotTable, ByVal orientation As Long, ByVal fields As Variant)
    Dim pf As PivotField, f As Variant, doomed As Collection, item As Variant
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    pt.ManualUpdate = True

    Set doomed = New Collection
    If orientation = xlRowField Then
        For Each pf In pt.RowFields
            doomed.Add pf.Name
        Next pf
    Else
        For Each pf In pt.ColumnFields
            doomed.Add pf.Name
        Next pf
    End If
    For Each item In doomed
        pt.PivotFields(CStr(item)).orientation = xlHidden
    Next item

    For Each f In fields
        If Len(Trim$(CStr(f))) > 0 Then
            Set pf = Nothing
            Set pf = pt.PivotFields(Trim$(CStr(f)))
            If Not pf Is Nothing Then
                pf.orientation = orientation
                If orientation = xlRowField Then
                    pf.Position = pt.RowFields.Count
                Else
                    pf.Position = pt.ColumnFields.Count
                End If
                pf.Subtotals(1) = True
                pf.Subtotals(1) = False
            End If
        End If
    Next f

    pt.ManualUpdate = False
    Application.ScreenUpdating = True
    Err.Clear
End Sub

Private Sub SetValues(ByVal pt As PivotTable, ByVal fields As Variant)
    Dim df As PivotField, f As Variant, doomed As Collection, item As Variant, nm As String
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    pt.ManualUpdate = True
    Set doomed = New Collection
    For Each df In pt.DataFields
        doomed.Add df.Name
    Next df
    For Each item In doomed
        pt.PivotFields(CStr(item)).orientation = xlHidden
    Next item
    For Each f In fields
        nm = Trim$(CStr(f))
        If Len(nm) > 0 Then
            Set df = Nothing
            ' A data field's caption may not equal a source field's name, or Excel
            ' raises 1004 and the value is never added. One trailing space is the
            ' smallest caption that satisfies it.
            Set df = pt.AddDataField(pt.PivotFields(nm), nm & " ", xlSum)
            If Not df Is Nothing Then df.NumberFormat = "#,##0"
        End If
    Next f
    pt.ManualUpdate = False
    Application.ScreenUpdating = True
    Err.Clear
End Sub

' A ladder sorted alphabetically reads "1 month, 1 year, 3 months" and is worse
' than useless. The generator leaves the intended order of each such field on the
' Data sheet, as "<field>=<item>|<item>|..." lines starting at S1; this puts the
' pivot items back into it after a rearrangement.
Public Sub OrderFromSheet(ByVal pt As PivotTable)
    Dim ws As Worksheet, r As Long, line As String, eq As Long
    If pt Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Data")
    If ws Is Nothing Then Exit Sub
    For r = 1 To 12
        line = CStr(ws.Cells(r, 20).Value2)
        If Len(line) > 0 Then
            eq = InStr(line, "=")
            If eq > 1 Then OrderOne pt, Left$(line, eq - 1), Mid$(line, eq + 1)
        End If
    Next r
    Err.Clear
End Sub

Private Sub OrderOne(ByVal pt As PivotTable, ByVal fieldName As String, ByVal pipedOrder As String)
    Dim pf As PivotField, parts As Variant, i As Long, pos As Long, nm As String
    On Error Resume Next
    Set pf = Nothing
    Set pf = pt.PivotFields(fieldName)
    If pf Is Nothing Then Exit Sub
    pt.ManualUpdate = True
    pf.AutoSort xlManual, pf.Name
    parts = Split(pipedOrder, "|")
    pos = 0
    For i = LBound(parts) To UBound(parts)
        nm = Trim$(CStr(parts(i)))
        If Len(nm) > 0 Then
            Err.Clear
            pf.PivotItems(nm).Position = pos + 1
            If Err.Number = 0 Then pos = pos + 1
        End If
    Next i
    pt.ManualUpdate = False
    Err.Clear
End Sub
