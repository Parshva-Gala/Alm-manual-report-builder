Attribute VB_Name = "modShared_Pivot"
Option Explicit

' ============================================================================
'  Building a real PivotTable, described rather than constructed.
'
'  Both tools now hand their reports to this. A report is stated as four field
'  lists and nothing else:
'
'      rows     "Type|Line|Subline|COA"
'      columns  "Bucket"
'      filters  "Framework|Currency"
'      values   "Pre>Pre-factor>#,##0|Post>Post-factor>#,##0"
'
'  Everything that used to be a hand-computed cross-tab is one of those. The
'  reader gets the field list, drill-down, expand and collapse, and a slicer per
'  filter - none of which a range of written cells can offer however carefully it
'  is formatted.
'
'  Three things here are hard-won rather than decorative:
'
'    * Position must be set AFTER Orientation and to the CURRENT count, not
'      count + 1. Excel has already inserted the field by the time Orientation
'      returns, so count + 1 is out of range and raises 1004.
'    * Layout, styling and slicers are all attempted defensively. A pivot that
'      is built but unstyled is a report; a pivot that raised on a cosmetic call
'      is a blank sheet.
'    * The cache is created on the workbook that will HOLD the pivot. Creating it
'      on the wrong workbook silently produces an external data reference that
'      breaks the moment the file is moved.
' ============================================================================

' Why a slicer or a field did not appear, kept until the caller reads it. A
' cosmetic failure must never stop a report, and must never be invisible either.
Private mProblem As String

' ===================== the report sheet's geometry ==========================
'
' EVERY report sheet in both tools has the same band at the top, and it is stated
' in points here rather than inferred from wherever the rows happen to land.
'
' Shapes are positioned in points and rows are sized in points, so a sheet whose
' row heights are left at Excel's defaults puts the controls in a different place
' from one where something has autofitted a row - and a pivot anchored at "row
' 16" lands somewhere different again. That drift is why the buttons and slicers
' looked unstable and why the pivot needed scrolling to reach.
'
'   row 1      18pt   title
'   row 2      12pt   what this sheet is
'   row 3      16pt   the actions strip          buttons at y = 30
'   row 4      16pt   the view strip             buttons at y = 46
'   row 5      16pt   the fields strip           buttons at y = 62
'   row 6       3pt   gap
'   rows 7-10 14pt    the slicer band            slicers at y = 81, 54 tall
'   row 11      4pt   gap
'   row 12            THE PIVOT, and the freeze line
'
' Every one of these is derived from the one above it, so a strip can be added
' without anybody having to re-add the offsets by hand: the last time these were
' separate literals, a third row of buttons would have been drawn straight on top
' of the slicers.
'
' The whole band is 141pt, against 257pt before. The band is what a reader loses
' off the top of the screen before the first figure appears, and at 257pt it cost
' them a third of a laptop screen on every one of ninety report sheets. Most of
' the saving is the slicers: a one-column slicer 116pt tall shows six items down
' a narrow strip, and the same slicer in two columns shows the same six in 54pt.
Public Const RPT_TITLE_H As Double = 18
Public Const RPT_ABOUT_H As Double = 12
Public Const RPT_STRIP_H As Double = 16
Public Const RPT_BTN_H As Double = 15
Public Const RPT_BTN1_Y As Double = RPT_TITLE_H + RPT_ABOUT_H
Public Const RPT_BTN2_Y As Double = RPT_BTN1_Y + RPT_STRIP_H
Public Const RPT_BTN3_Y As Double = RPT_BTN2_Y + RPT_STRIP_H
Public Const RPT_SLICER_Y As Double = RPT_BTN3_Y + RPT_STRIP_H + 3
Public Const RPT_SLICER_H As Double = 54
Public Const RPT_SLICER_W As Double = 230
Public Const RPT_SLICER_COLS As Long = 2
Public Const RPT_PIVOT_ROW As Long = 12

' Applies that band to a sheet. Called before anything is drawn on it.
Public Sub LayoutReportSheet(ByVal ws As Worksheet)
    Dim r As Long
    On Error Resume Next
    ws.rows(1).RowHeight = RPT_TITLE_H
    ws.rows(2).RowHeight = RPT_ABOUT_H
    ws.rows(3).RowHeight = RPT_STRIP_H
    ws.rows(4).RowHeight = RPT_STRIP_H
    ws.rows(5).RowHeight = RPT_STRIP_H
    ws.rows(6).RowHeight = 3
    For r = 7 To 10
        ws.rows(r).RowHeight = 14
    Next r
    ws.rows(11).RowHeight = 4
    Err.Clear
End Sub

' Pins the band so it cannot scroll away.
'
' This is the other half of "it needs a lot of scrolling": the controls were at
' the top of a sheet whose table is two thousand rows long, so using them meant
' scrolling up, clicking, and scrolling back down to see what changed. Frozen,
' they stay in front of the reader while the numbers move underneath.
'
' Freezing needs a visible window, which a headless generation does not have, so
' the workbook ALSO freezes itself when a person opens it. Between the two, it is
' frozen however it was made.
Public Sub FreezeReportSheet(ByVal ws As Worksheet)
    On Error Resume Next
    If Not Application.Visible Then Exit Sub
    If ws.Parent.Windows.count = 0 Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.FreezePanes = False
    ws.Cells(RPT_PIVOT_ROW, 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    Err.Clear
End Sub

' ===================== the one entry point ==================================

' src is the whole data block INCLUDING its header row.
Public Function AddPivot(ByVal src As Range, ByVal anchor As Range, ByVal ptName As String, _
                         ByVal rowFields As String, ByVal colFields As String, _
                         ByVal filterFields As String, ByVal dataFields As String) As PivotTable
    Dim wb As Workbook, pc As PivotCache
    On Error GoTo Failed
    Set wb = anchor.Parent.Parent
    Set pc = wb.PivotCaches.Create(xlDatabase, sourceData:=src)
    Set AddPivot = AddPivotFrom(pc, anchor, ptName, rowFields, colFields, filterFields, dataFields)
    Exit Function
Failed:
    Err.Clear
End Function

' The same, over a cache the caller already holds. Ninety pivots over ONE cache
' is a few megabytes of layout; ninety caches of the same rows is half a
' gigabyte, and they can disagree with each other, which is worse.
Public Function AddPivotFrom(ByVal pc As PivotCache, ByVal anchor As Range, ByVal ptName As String, _
                             ByVal rowFields As String, ByVal colFields As String, _
                             ByVal filterFields As String, ByVal dataFields As String, _
                             Optional ByVal deferUpdate As Boolean = False) As PivotTable
    Dim wb As Workbook, pt As PivotTable

    On Error GoTo Failed
    Set wb = anchor.Parent.Parent
    Set pt = pc.CreatePivotTable(TableDestination:=anchor, TableName:=UniquePivotName(wb, ptName))
    pt.ManualUpdate = True

    AddFields pt, filterFields, xlPageField
    AddFields pt, rowFields, xlRowField
    AddFields pt, colFields, xlColumnField
    AddValues pt, dataFields

    On Error Resume Next
    ' Tabular with repeated labels: every row of the pivot then carries its full
    ' type / line / subline / COA path, which is what makes it filterable and
    ' copyable into anything else. Compact form looks tidier and is useless the
    ' moment somebody wants the rows somewhere other than this sheet.
    pt.RowAxisLayout xlTabularRow
    pt.RepeatAllLabels xlRepeatLabels
    pt.ShowDrillIndicators = True
    pt.HasAutoFormat = False            ' keep column widths across refreshes
    pt.TableStyle2 = "PivotStyleMedium2"
    pt.ColumnGrand = True
    pt.RowGrand = True
    pt.DisplayImmediateItems = False
    Err.Clear
    On Error GoTo Failed
    ' Existing callers receive a complete pivot. A caller adding scope filters
    ' can defer this update until those filters and its layout are in place.
    If Not deferUpdate Then pt.ManualUpdate = False

    Set AddPivotFrom = pt
    Exit Function
Failed:
    Err.Clear
End Function

' Empty optional display filters can collapse a saved Excel pivot when its scope
' changes. Only these report-only fields may be removed; required scope stays.
Public Function RemoveEmptyOptionalPageFields(ByVal pt As PivotTable, Optional ByVal bridgeScope As Boolean = False) As Long
    Dim field As Variant, pf As PivotField, item As PivotItem
    Dim visibleCount As Long, removeField As Boolean
    If pt Is Nothing Then Exit Function
    For Each field In Array("ALM_FRAMEWORK_NAME", "SOURCE_CODE", "STAGE_ID")
        Set pf = Nothing
        On Error Resume Next
        Set pf = pt.PivotFields(CStr(field))
        Err.Clear
        On Error GoTo 0
        If Not pf Is Nothing Then
            If pf.orientation = xlPageField Then
                removeField = bridgeScope
                If Not removeField Then
                    visibleCount = 0
                    For Each item In pf.PivotItems
                        If item.Visible Then visibleCount = visibleCount + 1
                    Next item
                    If visibleCount = 0 Then removeField = (pf.PivotItems.count = 0 Or CStr(pf.CurrentPage) = "(All)")
                End If
                If removeField Then
                    pf.orientation = xlHidden
                    RemoveEmptyOptionalPageFields = RemoveEmptyOptionalPageFields + 1
                End If
            End If
        End If
    Next field
End Function

' Excel orders pivot items alphabetically unless told otherwise, and alphabetical
' order is wrong for most dimensions worth reporting on. A maturity ladder sorted
' that way reads "1 month, 1 year, 3 months"; a set of report lines sorted that
' way loses the order of the return entirely. Both have a real order somewhere -
' the ladder, the rule set - and this applies it.
'
' Items the caller did not name keep their relative position at the end, so an
' order list that has gone stale degrades to a partial sort rather than an error.
Public Sub OrderItems(ByVal pt As PivotTable, ByVal fieldName As String, ByVal pipedOrder As String)
    Dim pf As PivotField, parts As Variant, i As Long, pos As Long, nm As String
    If pt Is Nothing Then Exit Sub
    If Len(Trim$(pipedOrder)) = 0 Then Exit Sub
    If Not FieldExists(pt, fieldName) Then Exit Sub
    On Error Resume Next
    Set pf = pt.PivotFields(fieldName)
    pf.AutoSort xlManual, pf.name
    parts = Split(pipedOrder, "|")
    pos = 0
    For i = LBound(parts) To UBound(parts)
        nm = Trim$(CStr(parts(i)))
        If Len(nm) > 0 Then
            Err.Clear
            pf.PivotItems(nm).position = pos + 1
            If Err.Number = 0 Then pos = pos + 1
        End If
    Next i
    Err.Clear
End Sub

' A ratio that stays correct when the pivot is collapsed.
'
' "31% of this sector" cannot be a column of the source table: a pivot would SUM
' the percentages, and a subtotal of four rows reading 154% is worse than having
' no ratio at all. A calculated field is evaluated on the aggregates instead, so
' it is right at the leaf, right at the subtotal and right at the grand total.
Public Function AddRatio(ByVal pt As PivotTable, ByVal caption As String, _
                         ByVal numerator As String, ByVal denominator As String) As Boolean
    Dim cf As Object, df As PivotField
    If pt Is Nothing Then Exit Function
    If Not FieldExists(pt, numerator) Then Note "ratio numerator '" & numerator & "' is not in the pivot": Exit Function
    If Not FieldExists(pt, denominator) Then Note "ratio denominator '" & denominator & "' is not in the pivot": Exit Function
    On Error Resume Next
    Err.Clear
    Set cf = pt.CalculatedFields.Add(caption, "='" & numerator & "'/'" & denominator & "'")
    If cf Is Nothing Then
        Note "CalculatedFields.Add(" & caption & "): " & Err.description
        Err.Clear
        Exit Function
    End If
    Set df = pt.AddDataField(pt.PivotFields(caption), caption & " ", xlSum)
    If df Is Nothing Then
        Note "AddDataField(" & caption & "): " & Err.description
    Else
        df.NumberFormat = "0.0%"
        AddRatio = True
    End If
    Err.Clear
End Function

' Subtotals on the outer levels only, which is how a report of this shape reads:
' a total per type, a total per line, and no total under a leaf.
Public Sub SubtotalOuterOnly(ByVal pt As PivotTable)
    Dim i As Long, pf As PivotField
    On Error Resume Next
    For i = 1 To pt.rowFields.count
        Set pf = pt.rowFields(i)
        SetSubtotals pf, (i < pt.rowFields.count - 1)
    Next i
    Err.Clear
End Sub

Public Sub CollapseTo(ByVal pt As PivotTable, ByVal level As Long)
    Dim i As Long
    On Error Resume Next
    For i = 1 To pt.rowFields.count
        pt.rowFields(i).ShowDetail = (i < level)
    Next i
    Err.Clear
End Sub

' ===================== slicers ==============================================

' A slicer is the control the reader actually wants: click a type, click a
' bucket, click a currency, and every figure follows. It needs no macro, which
' is what makes it work in a generated .xlsx as well as in the tool itself.
'
' Returns how many were added, and leaves the reason in LastPivotProblem when
' fewer arrived than were asked for - a slicer that silently fails to appear is
' otherwise indistinguishable from one that was never asked for.
Public Function AddSlicers(ByVal pt As PivotTable, ByVal fields As String, _
                           ByVal leftPt As Double, ByVal topPt As Double, _
                           Optional ByVal widthPt As Double = RPT_SLICER_W, _
                           Optional ByVal heightPt As Double = RPT_SLICER_H, _
                           Optional ByVal columns As Long = RPT_SLICER_COLS) As Long
    ' EARLY bound, and that is the whole fix.
    '
    ' Slicers.Add's second argument (Level) is meaningful only for an OLAP cache
    ' and has to be OMITTED for an ordinary one. A late-bound call - which is what
    ' `Dim sc As Object` makes this - cannot omit a middle argument: VBA passes a
    ' Missing marker through IDispatch, which rejects it as "Invalid procedure
    ' call or argument". Every slicer in both tools failed this way, silently,
    ' because the failure was swallowed by the On Error that protects the report.
    Dim wb As Workbook, ws As Worksheet, f As Variant, x As Double, nm As String
    Dim sc As SlicerCache, sl As Slicer, n As Long
    If pt Is Nothing Then Exit Function
    If Len(Trim$(fields)) = 0 Then Exit Function
    Set ws = pt.Parent
    Set wb = ws.Parent
    x = leftPt
    For Each f In Split(fields, "|")
        nm = Trim$(CStr(f))
        If Len(nm) > 0 Then
            If Not FieldExists(pt, nm) Then
                Note "slicer field '" & nm & "' is not in the pivot"
            Else
                On Error Resume Next
                Set sc = Nothing
                Err.Clear
                Set sc = wb.SlicerCaches.Add2(pt, nm)
                If sc Is Nothing Then Note "SlicerCaches.Add2(" & nm & "): " & Err.description
                If Not sc Is Nothing Then
                    Err.Clear
                    Set sl = Nothing
                    Set sl = sc.Slicers.Add(ws, , UniqueSlicerName(wb, nm), nm, topPt, x, widthPt, heightPt)
                    If sl Is Nothing Then
                        Note "Slicers.Add(" & nm & "): " & Err.description
                    Else
                        sl.Style = "SlicerStyleLight2"
                        ' Two columns rather than one. The same items, half the
                        ' height, and the height is what the reader pays for on
                        ' every sheet before a single figure appears.
                        If columns < 1 Then columns = 1
                        sl.NumberOfColumns = columns
                        ' A slicer arrives set to "move and size with cells", so the
                        ' first column a reader widens, row a pivot adds or filter
                        ' they apply drags every slicer somewhere else and resizes
                        ' it. That is the whole of the "placement is unstable"
                        ' problem: the controls were anchored to the data they
                        ' control. Free-floating pins them to the sheet instead.
                        sl.Shape.Placement = xlFreeFloating
                        sl.Shape.LockAspectRatio = msoFalse
                        sl.Top = topPt
                        sl.Left = x
                        sl.Width = widthPt
                        sl.Height = heightPt
                        n = n + 1
                        x = x + widthPt + 8
                    End If
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next f
    AddSlicers = n
End Function

Private Sub Note(ByVal s As String)
    If Len(mProblem) < 400 Then
        If Len(mProblem) > 0 Then mProblem = mProblem & "; "
        mProblem = mProblem & s
    End If
End Sub

Public Function LastPivotProblem() As String
    LastPivotProblem = mProblem
End Function

Public Sub ClearPivotProblem()
    mProblem = ""
End Sub

' ===================== field plumbing =======================================

Private Sub AddFields(ByVal pt As PivotTable, ByVal fields As String, ByVal orientation As Long)
    Dim p As Variant, pf As PivotField, nm As String
    If Len(Trim$(fields)) = 0 Then Exit Sub
    For Each p In Split(fields, "|")
        nm = Trim$(CStr(p))
        If Len(nm) > 0 Then
            If FieldExists(pt, nm) Then
                On Error Resume Next
                Set pf = pt.PivotFields(nm)
                pf.orientation = orientation
                ' AFTER Orientation, and to the CURRENT count. See the header.
                Select Case orientation
                    Case xlRowField:    pf.position = pt.rowFields.count
                    Case xlColumnField: pf.position = pt.ColumnFields.count
                    Case xlPageField:   pf.position = pt.PageFields.count
                End Select
                SetSubtotals pf, False
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next p
End Sub

' "Field>Caption>Format" per entry; caption and format are optional.
Private Sub AddValues(ByVal pt As PivotTable, ByVal spec As String)
    Dim p As Variant, parts As Variant, nm As String, cap As String, fmt As String, df As PivotField
    If Len(Trim$(spec)) = 0 Then Exit Sub
    For Each p In Split(spec, "|")
        If Len(Trim$(CStr(p))) > 0 Then
            parts = Split(CStr(p) & ">>", ">")
            nm = Trim$(CStr(parts(0)))
            cap = Trim$(CStr(parts(1)))
            fmt = Trim$(CStr(parts(2)))
            If Len(cap) = 0 Then cap = nm
            If Len(fmt) = 0 Then fmt = "#,##0"
            If FieldExists(pt, nm) Then
                On Error Resume Next
                ' A data field's caption may not equal a source field's name, or
                ' Excel raises 1004 and the value is never added.
                If StrComp(cap, nm, vbTextCompare) = 0 Then cap = cap & " "
                Set df = pt.AddDataField(pt.PivotFields(nm), cap, xlSum)
                df.NumberFormat = fmt
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next p
End Sub

Private Sub SetSubtotals(ByVal pf As PivotField, ByVal wanted As Boolean)
    On Error Resume Next
    ' Subtotals(1) is "Automatic". Turning it on and off again is the documented
    ' way to clear the other twelve, which otherwise reappear one at a time.
    pf.Subtotals(1) = True
    pf.Subtotals(1) = wanted
    Err.Clear
End Sub

Public Function FieldExists(ByVal pt As PivotTable, ByVal nm As String) As Boolean
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(nm)
    FieldExists = Not pf Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function UniquePivotName(ByVal wb As Workbook, ByVal stem As String) As String
    Dim s As String, n As Long, ws As Worksheet, pt As PivotTable, taken As Boolean
    s = CleanName(stem)
    Do
        taken = False
        For Each ws In wb.Worksheets
            For Each pt In ws.PivotTables
                If StrComp(pt.name, s, vbTextCompare) = 0 Then taken = True
            Next pt
        Next ws
        If Not taken Then Exit Do
        n = n + 1
        s = CleanName(stem) & "_" & n
    Loop
    UniquePivotName = s
End Function

Private Function UniqueSlicerName(ByVal wb As Workbook, ByVal stem As String) As String
    Static seq As Long
    seq = seq + 1
    UniqueSlicerName = CleanName(stem) & "_" & seq
End Function

' A pivot or slicer name must start with a letter and hold no spaces or symbols.
Private Function CleanName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        ElseIf Len(out) > 0 Then
            If Right$(out, 1) <> "_" Then out = out & "_"
        End If
    Next i
    If Len(out) = 0 Then out = "PT"
    If Not ((Left$(out, 1) >= "A" And Left$(out, 1) <= "Z") Or (Left$(out, 1) >= "a" And Left$(out, 1) <= "z")) Then out = "PT_" & out
    If Len(out) > 40 Then out = Left$(out, 40)
    CleanName = out
End Function


