Option Explicit

' ============================================================================
'  The pivots.
'
'  Every pivot in a framework's workbook is built on ONE PivotCache over the
'  staging table. That is what makes the workbook feel like a desk rather than
'  a report: one refresh refreshes everything, a slicer can drive every sheet at
'  once, and adding a fortieth sheet costs a layout, not another copy of the
'  data.
'
'  Three kinds of sheet, and they are the three asked for:
'
'    OUTPUT         rule order / category / rule name / factor, against
'                   pre- and post-factor, split local against foreign.
'    BALANCE SHEET  type / line / subline / COA name, pre-factor, same split.
'    ONE PER RULE   type / line / subline / COA name across the buckets, for a
'                   single rule and a single currency side.
'
'  Names, never codes. The staging table carries no code column at all, so a
'  pivot cannot accidentally show one.
'
'  Nothing here writes a number into a cell. Everything is a live PivotTable
'  with its field list on, drill-down enabled and slicers wired to the same
'  cache - so when the analysis needs to change, it is dragged, not rebuilt.
' ============================================================================

Private Const PT_STYLE As String = "PivotStyleMedium2"

' Where the pivot starts on a sheet: under the masthead and the slicer strip.
Private Const PIVOT_ROW As Long = 9
Private Const SLICER_TOP As Double = 74
Private Const SLICER_H As Double = 58

Private mCache As PivotCache
Private mSeq As Long
Private mMade As Collection       ' Array(sheetName, what, rows, pre)

Public Function MadeSheets() As Collection
    If mMade Is Nothing Then Set mMade = New Collection
    Set MadeSheets = mMade
End Function

Public Sub ResetPivots()
    Set mCache = Nothing
    Set mMade = New Collection
    mSeq = 0
End Sub

Public Sub UseCache(ByVal wb As Workbook, ByVal lo As ListObject)
    Set mCache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=lo.Range)
End Sub

' ===================== the three sheet kinds ================================

' Rule order / category / rule name / factor, against pre- and post-factor,
' split local against foreign, with the bucket as a filter that excludes blanks.
Public Function BuildOutputSheet(ByVal wb As Workbook, ByVal fw As String) As Worksheet
    Dim ws As Worksheet, pt As PivotTable
    Set ws = NewPivotSheet(wb, FwLabel(fw) & " Output", _
        "Every rule, in the order the engine evaluates them, against what it read and what it kept.", _
        "The factor is post divided by pre, so it always agrees with the two figures beside it.")
    Set pt = NewPivot(ws, "pt_output")
    If pt Is Nothing Then Exit Function

    pt.ManualUpdate = True
    RowField pt, H_RULE_ORDER, 1
    RowField pt, H_RULE_CAT, 2
    RowField pt, H_RULE_NAME, 3
    RowField pt, H_FACTOR, 4

    ' Values BEFORE the currency split in the column area, so the order reads
    ' pre-LCY, pre-FCY, post-LCY, post-FCY rather than interleaving the two
    ' measures under each currency.
    AddSum pt, H_PRE, CAP_PRE
    AddSum pt, H_POST, CAP_POST
    ColField pt, H_CCYCLASS, 2
    DataFirst pt

    PageField pt, H_BUCKET
    DropBlanks pt, H_BUCKET, "(no bucket)"
    Finish pt, ws
    NoteSheet ws, "Rule-level output"
    Slicers ws, pt, Array(H_RULE_CAT, H_CCYCLASS, H_BUCKET)
    Set BuildOutputSheet = ws
End Function

' Type / line / subline / COA name, pre-factor, split local against foreign.
Public Function BuildBalanceSheet(ByVal wb As Workbook, ByVal fw As String) As Worksheet
    Dim ws As Worksheet, pt As PivotTable
    Set ws = NewPivotSheet(wb, SH_BALSHEET, _
        "The same balances as the balance sheet reads them, down to the COA.", _
        "Pre-factor only - this is what the engine took in, before any weighting.")
    Set pt = NewPivot(ws, "pt_bs")
    If pt Is Nothing Then Exit Function

    pt.ManualUpdate = True
    RowField pt, H_TYPE, 1
    RowField pt, H_LINE, 2
    RowField pt, H_SUBLINE, 3
    RowField pt, H_COA_NAME, 4
    AddSum pt, H_PRE, CAP_PRE
    ColField pt, H_CCYCLASS, 1
    PageField pt, H_BUCKET
    DropBlanks pt, H_BUCKET, "(no bucket)"
    Finish pt, ws
    NoteSheet ws, "Balance sheet"
    Slicers ws, pt, Array(H_TYPE, H_CCYCLASS, H_RULE_CAT)
    Set BuildBalanceSheet = ws
End Function

' One rule, one currency side: type / line / subline / COA name across buckets.
'
' filterField is "LCY / FCY" for LCR and NSFR and "Currency" for the ladder;
' filterValue is the side or the currency. Everything else is identical, which
' is why there is one function rather than two.
Public Function BuildRuleSheet(ByVal wb As Workbook, ByVal ruleName As String, _
                               ByVal filterField As String, ByVal filterValue As String, _
                               ByVal tabName As String) As Worksheet
    Dim ws As Worksheet, pt As PivotTable
    Set ws = NewPivotSheet(wb, tabName, ruleName, _
        filterField & ": " & filterValue & "   -   balances across the maturity buckets.")
    Set pt = NewPivot(ws, "pt_" & mSeq)
    If pt Is Nothing Then Exit Function

    pt.ManualUpdate = True
    RowField pt, H_TYPE, 1
    RowField pt, H_LINE, 2
    RowField pt, H_SUBLINE, 3
    RowField pt, H_COA_NAME, 4
    ColField pt, H_BUCKET, 1
    AddSum pt, H_PRE, CAP_PRE

    PageField pt, H_RULE_NAME
    PickOne pt, H_RULE_NAME, ruleName
    PageField pt, filterField
    PickOne pt, filterField, filterValue

    ' Blanks out of the bucket axis here too, not only where the bucket is a
    ' filter. Otherwise "(no bucket)" is the FIRST column on every sheet and the
    ' maturity profile starts one column to the right of where the eye lands.
    DropBlanks pt, H_BUCKET, "(no bucket)"
    Finish pt, ws
    ' No slicers here, deliberately - see the note on Slicers. This sheet is
    ' already filtered to one rule and one currency side, and a slicer would
    ' share its cache with every other sheet in the book.
    NoteSheet ws, ruleName & "  -  " & filterValue
    Set BuildRuleSheet = ws
End Function

' One currency, every rule: the ladder's per-currency worksheet.
'
' The ladder is not split local against foreign - it is split by the actual
' currency, one worksheet each - so the rule name moves into the rows instead of
' the filter and the currency takes its place.
Public Function BuildCurrencySheet(ByVal wb As Workbook, ByVal ccy As String, _
                                   ByVal tabName As String) As Worksheet
    Dim ws As Worksheet, pt As PivotTable
    Set ws = NewPivotSheet(wb, tabName, ccy, _
        "Every rule for this currency, across the maturity buckets.")
    Set pt = NewPivot(ws, "pt_ccy" & mSeq)
    If pt Is Nothing Then Exit Function

    pt.ManualUpdate = True
    RowField pt, H_RULE_NAME, 1
    RowField pt, H_TYPE, 2
    RowField pt, H_LINE, 3
    RowField pt, H_SUBLINE, 4
    RowField pt, H_COA_NAME, 5
    ColField pt, H_BUCKET, 1
    AddSum pt, H_PRE, CAP_PRE

    PageField pt, H_CURRENCY
    PickOne pt, H_CURRENCY, ccy

    DropBlanks pt, H_BUCKET, "(no bucket)"
    Finish pt, ws
    NoteSheet ws, ccy
    Set BuildCurrencySheet = ws
End Function

' ===================== the mechanics ========================================

Private Function NewPivotSheet(ByVal wb As Workbook, ByVal wanted As String, _
                               ByVal title As String, ByVal about As String) As Worksheet
    Dim ws As Worksheet, nm As String
    nm = FreeSheetName(wanted, wb)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    On Error Resume Next
    ws.Name = nm
    Err.Clear
    On Error GoTo 0
    modPD_Theme.Dress ws, title, about
    ws.Rows(modPD_Theme.R_STATUS).RowHeight = 6
    Set NewPivotSheet = ws
End Function

Private Function NewPivot(ByVal ws As Worksheet, ByVal nm As String) As PivotTable
    On Error Resume Next
    mSeq = mSeq + 1
    Set NewPivot = mCache.CreatePivotTable(TableDestination:=ws.Cells(PIVOT_ROW, 1), _
                                           TableName:=nm & "_" & mSeq)
    Err.Clear
End Function

Private Sub RowField(ByVal pt As PivotTable, ByVal nm As String, ByVal pos As Long)
    On Error Resume Next
    With pt.PivotFields(nm)
        .Orientation = xlRowField
        .Position = pos
        .Subtotals = NoSubtotals()
    End With
    Err.Clear
End Sub

Private Sub ColField(ByVal pt As PivotTable, ByVal nm As String, ByVal pos As Long)
    On Error Resume Next
    With pt.PivotFields(nm)
        .Orientation = xlColumnField
        .Position = pos
        .Subtotals = NoSubtotals()
    End With
    Err.Clear
End Sub

Private Sub PageField(ByVal pt As PivotTable, ByVal nm As String)
    On Error Resume Next
    With pt.PivotFields(nm)
        .Orientation = xlPageField
        .EnableMultiplePageItems = True
    End With
    Err.Clear
End Sub

' A data field that did not get added is not a cosmetic loss - it is a pivot
' with nothing in it, and the first thing to notice was a confusing error three
' calls later. So this one raises rather than swallowing.
Private Sub AddSum(ByVal pt As PivotTable, ByVal srcName As String, ByVal caption As String)
    Dim df As PivotField
    On Error Resume Next
    Set df = pt.AddDataField(pt.PivotFields(srcName), caption, xlSum)
    On Error GoTo 0
    If df Is Nothing Then
        Err.Raise vbObjectError + 513, "AddSum", _
            "Could not add " & srcName & " as a data field called " & Chr$(34) & caption & Chr$(34) & _
            ". Excel refuses a caption that matches an existing field name."
    End If
    ' Function is NOT set again here. AddDataField already made it a sum, and
    ' assigning .Function afterwards makes Excel regenerate the caption - which
    ' turned "Pre-factor" back into "Sum of Pre factor amount" on every sheet.
    df.NumberFormat = NUM_FMT
    df.caption = caption
End Sub

' The measures sit OUTSIDE the currency split, so the columns read pre-LCY,
' pre-FCY, post-LCY, post-FCY rather than interleaving both measures under each
' currency. Guarded: with a single data field Excel owns this field and refuses
' to be told where to put it.
Private Sub DataFirst(ByVal pt As PivotTable)
    On Error Resume Next
    pt.DataPivotField.Orientation = xlColumnField
    pt.DataPivotField.Position = 1
    Err.Clear
End Sub

' Twelve Falses: the Subtotals property wants the whole array, and a row field
' with subtotals at every level buries the numbers being read.
Private Function NoSubtotals() As Variant
    NoSubtotals = Array(False, False, False, False, False, False, _
                        False, False, False, False, False, False)
End Function

' Blanks unselected by default, without turning the filter off. The staging pass
' writes "(no bucket)" instead of an empty cell, so there is a real item to hide
' - hiding a genuinely blank PivotItem is not reliable across Excel versions.
Private Sub DropBlanks(ByVal pt As PivotTable, ByVal nm As String, ByVal blankLabel As String)
    Dim pi As PivotItem, visible As Long
    On Error Resume Next
    ' At least one item must stay visible or Excel refuses the whole filter.
    For Each pi In pt.PivotFields(nm).PivotItems
        If StrComp(pi.Name, blankLabel, vbTextCompare) <> 0 Then visible = visible + 1
    Next pi
    If visible = 0 Then Exit Sub
    For Each pi In pt.PivotFields(nm).PivotItems
        If StrComp(pi.Name, blankLabel, vbTextCompare) = 0 Then pi.visible = False
    Next pi
    Err.Clear
End Sub

Private Sub PickOne(ByVal pt As PivotTable, ByVal nm As String, ByVal value As String)
    Dim pf As PivotField
    On Error Resume Next
    Set pf = pt.PivotFields(nm)
    If pf Is Nothing Then Exit Sub
    pf.EnableMultiplePageItems = False
    pf.CurrentPage = value
    Err.Clear
End Sub

' Layout, format and the settings that decide whether this feels like a desk.
Private Sub Finish(ByVal pt As PivotTable, ByVal ws As Worksheet)
    On Error Resume Next
    With pt
        .TableStyle2 = PT_STYLE
        .RowAxisLayout xlTabularRow          ' one field per column, not nested
        .RepeatAllLabels xlDoNotRepeatLabels
        .ShowDrillIndicators = True
        .EnableDrilldown = True
        .EnableFieldList = True
        .EnableWizard = True
        .DisplayFieldCaptions = True
        .ColumnGrand = True
        .RowGrand = True
        .HasAutoFormat = False               ' stop autofit fighting the widths
        .PreserveFormatting = True
        .NullString = "-"                    ' an empty cell reads as nothing
        .DisplayNullString = True
        .ManualUpdate = False
    End With
    FitColumns ws, pt
    Err.Clear
End Sub

' "Fit it on one screen without adjusting it every time."
'
' Row-label columns get width by what they hold - a COA name needs room, a type
' does not - and every value column gets one narrow width, because a figure with
' no decimals and a thousands separator has a known size. Then the sheet's own
' zoom is set from how wide the result came out, so it lands on screen whatever
' the pivot turned out to be.
Private Sub FitColumns(ByVal ws As Worksheet, ByVal pt As PivotTable)
    Dim rng As Range, c As Long, firstData As Long, lastCol As Long, total As Double

    On Error Resume Next
    Set rng = pt.TableRange1
    If rng Is Nothing Then Exit Sub
    lastCol = rng.Columns.count
    firstData = pt.RowFields.count
    If firstData < 1 Then firstData = 1

    For c = 1 To lastCol
        If c <= firstData Then
            ws.Columns(c).ColumnWidth = RowLabelWidth(pt, c)
        Else
            ws.Columns(c).ColumnWidth = 14
        End If
        total = total + ws.Columns(c).ColumnWidth
    Next c

    ws.Rows(modPD_Theme.R_HDR).RowHeight = 6
    ' A rough character-width budget for a laptop screen. Clamped so it never
    ' goes so small the figures cannot be read.
    ws.Activate
    ActiveWindow.Zoom = ZoomFor(total)
    ActiveWindow.FreezePanes = False
    ws.Cells(PIVOT_ROW + 2, firstData + 1).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    Err.Clear
End Sub

Private Function RowLabelWidth(ByVal pt As PivotTable, ByVal c As Long) As Double
    Dim nm As String
    On Error Resume Next
    nm = pt.RowFields(c).Name
    Select Case True
        Case InStr(1, nm, H_COA_NAME, vbTextCompare) > 0: RowLabelWidth = 38
        Case InStr(1, nm, H_SUBLINE, vbTextCompare) > 0: RowLabelWidth = 34
        Case InStr(1, nm, H_LINE, vbTextCompare) > 0: RowLabelWidth = 34
        Case InStr(1, nm, H_RULE_NAME, vbTextCompare) > 0: RowLabelWidth = 40
        Case InStr(1, nm, H_RULE_CAT, vbTextCompare) > 0: RowLabelWidth = 18
        Case InStr(1, nm, H_TYPE, vbTextCompare) > 0: RowLabelWidth = 14
        Case InStr(1, nm, H_FACTOR, vbTextCompare) > 0: RowLabelWidth = 9
        Case InStr(1, nm, H_RULE_ORDER, vbTextCompare) > 0: RowLabelWidth = 8
        Case Else: RowLabelWidth = 16
    End Select
    If RowLabelWidth = 0 Then RowLabelWidth = 16
    Err.Clear
End Function

Private Function ZoomFor(ByVal totalWidth As Double) As Long
    Dim z As Long
    If totalWidth <= 0 Then ZoomFor = 100: Exit Function
    z = CLng(150 * 120 / totalWidth)
    If z > 100 Then z = 100
    If z < 55 Then z = 55
    ZoomFor = z
End Function

' ===================== slicers ==============================================

' Slicers are how this stays a desk without needing macros in the file you are
' handed. They filter every pivot wired to the same cache, they multi-select,
' and they survive being emailed - which a button that calls a macro does not.
' ============================================================================
'  Slicers, and why there are only two sheets' worth.
'
'  Every pivot here shares ONE PivotCache, and Excel allows exactly one slicer
'  cache per field per workbook. So the first sheet's slicers are created and
'  every later Add2 for the same field fails with "invalid procedure call".
'
'  The fix is not to reuse the cache on all forty-four sheets. That would work,
'  and it would mean clicking "Assets" on one rule sheet silently re-filtered
'  the other forty-three - a shared filter nobody asked for and nobody would
'  find. So the slicers live on the two OVERVIEW sheets, where a filter is meant
'  to move the whole view, and the rule sheets keep their own report filters,
'  their field list and drill-down.
'
'  Slicers.Add is (Destination, Level, Name, Caption, Top, Left, WIDTH, HEIGHT):
'  width and height last, in that order.
' ============================================================================
Private Sub Slicers(ByVal ws As Worksheet, ByVal pt As PivotTable, ByRef fields As Variant)
    Dim i As Long, x As Double, sc As SlicerCache, sl As Slicer, nm As String, why As String
    x = 4
    For i = 0 To UBound(fields)
        nm = CStr(fields(i))
        Set sc = Nothing
        Set sl = Nothing
        why = ""

        Set sc = CacheFor(ws.Parent, pt, nm, why)
        If Not sc Is Nothing Then
            On Error Resume Next
            Set sl = sc.Slicers.Add(ws, , SlicerName(nm, i), nm, SLICER_TOP, x, 156, SLICER_H)
            If sl Is Nothing Then why = "slicer: " & Err.Number & " " & Err.Description
            Err.Clear
            On Error GoTo 0
        End If

        If sl Is Nothing Then
            LogIt V_CHECK, "Slicer", "No slicer for " & Chr$(34) & nm & Chr$(34) & " - " & why, ws.Name
        Else
            On Error Resume Next
            sl.style = "SlicerStyleDark2"
            sl.NumberOfColumns = 2
            sl.RowHeight = 15
            sl.caption = nm
            Err.Clear
            On Error GoTo 0
            x = x + 162
        End If
    Next i
End Sub

' The slicer cache for a field, made if this is the first time and reused - with
' this pivot connected to it - if another sheet already made one.
Private Function CacheFor(ByVal wb As Workbook, ByVal pt As PivotTable, ByVal nm As String, _
                          ByRef why As String) As SlicerCache
    Dim sc As SlicerCache
    On Error Resume Next
    For Each sc In wb.SlicerCaches
        If StrComp(sc.SourceName, nm, vbTextCompare) = 0 Then
            sc.PivotTables.AddPivotTable pt
            Err.Clear
            Set CacheFor = sc
            Exit Function
        End If
    Next sc
    Err.Clear

    Set CacheFor = wb.SlicerCaches.Add2(pt, nm)
    If CacheFor Is Nothing Then why = "cache: " & Err.Number & " " & Err.Description
    Err.Clear
End Function

' A slicer name must be a valid defined name: no spaces, no punctuation. The
' CAPTION is what anyone reads, so it keeps the real wording.
Private Function SlicerName(ByVal nm As String, ByVal i As Long) As String
    Dim j As Long, ch As String, out As String
    For j = 1 To Len(nm)
        ch = Mid$(nm, j, 1)
        If (UCase$(ch) >= "A" And UCase$(ch) <= "Z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next j
    If Len(out) = 0 Then out = "F"
    SlicerName = "sl_" & out & "_" & mSeq & "_" & i
End Function

Private Sub NoteSheet(ByVal ws As Worksheet, ByVal what As String)
    MadeSheets.Add Array(ws.Name, what)
End Sub
