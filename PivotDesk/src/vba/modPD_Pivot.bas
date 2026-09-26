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

' The fallback if the workbook will not take the Avati style. With the
' brand theme applied, Medium2 is itself emerald.
Private Const PT_STYLE As String = "PivotStyleMedium2"
Private Const PT_CUSTOM As String = "Avati"
Private Const SLICER_STYLE As String = "SlicerStyleDark1"

' A built sheet, row by row:
'
'   1  the bar: the Avati mark, the book, Start here, previous and next
'   2  title          3  what it is
'   4  tiles: each value's grand total, live off the pivot
'   5  spacer         6  slicers, when the sheet has them
'   7  spacer         8  the pivot's filters, then the pivot itself
'
' The pivot is placed as many rows below row 8 as it has filters, so its
' filters land on rows of their own and never on the chrome's.
Private Const TILE_ROW As Long = 4
Private Const SLICER_BAND_ROW As Long = 6
Private Const FILTER_TOP As Long = 8
Private Const SLICER_H As Double = 62
' Where each tile's live figure is worked out: a cell far off to the right,
' in the sheet's own colour, that the tile's text is linked to.
Private Const TILE_HELPER_COL As Long = 240
Private Const SLICER_CUSTOM As String = "Avati Slicer"

Private mCache As PivotCache
Private mSeq As Long
Private mBook As String            ' "LCR  ·  MIDBANK CAIRO", on every sheet's bar
Private mMade As Collection       ' Array(sheetName, what, rows, pre)

Public Function MadeSheets() As Collection
    If mMade Is Nothing Then Set mMade = New Collection
    Set MadeSheets = mMade
End Function

Public Sub SetBook(ByVal crumb As String)
    mBook = crumb
End Sub

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
        "The factor is post divided by pre, so it always agrees with the two figures beside it.", _
        UCase$(FwLabel(fw)) & "  " & ChrW(183) & "  OUTPUT")
    Set pt = NewPivot(ws, "pt_output", 1)
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
        "Pre-factor only - this is what the engine took in, before any weighting.", _
        UCase$(FwLabel(fw)) & "  " & ChrW(183) & "  BALANCE SHEET")
    Set pt = NewPivot(ws, "pt_bs", 1)
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
        filterField & ": " & filterValue & "   -   balances across the maturity buckets.", _
        UCase$(filterValue) & "  " & ChrW(183) & "  ONE RULE")
    Set pt = NewPivot(ws, "pt_" & mSeq, 2)
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
        "Every rule for this currency, across the maturity buckets.", "ONE CURRENCY")
    Set pt = NewPivot(ws, "pt_ccy" & mSeq, 1)
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

' ===================== a sheet from a Pivot config recipe ===================
'
' One recipe row, made into one sheet: the same mechanics the four built-in
' kinds use - one cache, tabular rows, blanks as named items so they can be
' hidden, captions that cannot collide - driven by what the row says rather
' than by code.
'
' splitVals is Empty for a single pivot, or the values of the "one sheet per"
' fields for this sheet, in the order the recipe lists them.
Public Function BuildRecipeSheet(ByVal wb As Workbook, ByVal rc As Object, ByVal fw As String, _
                                 ByVal splitVals As Variant, ByVal tabName As String) As Worksheet
    Dim ws As Worksheet, pt As PivotTable, x As Variant, pos As Long, v As Object, flt As Object
    Dim title As String, about As String, fl As Object, i As Long, sf As Collection, what As String

    Set fl = modPD_Config.Fields()
    Set sf = rc("Split")
    If sf.count > 0 Then
        title = CStr(splitVals(0))
        what = title
        For i = 2 To sf.count
            about = about & CStr(sf(i)) & ": " & CStr(splitVals(i - 1)) & "   -   "
            what = what & "  -  " & CStr(splitVals(i - 1))
        Next i
        about = about & CStr(rc("Desc"))
    Else
        title = Replace(CStr(rc("Name")), "{fw}", FwLabel(fw))
        about = CStr(rc("Desc"))
        what = IIf(Len(about) > 0, about, title)
    End If

    Set ws = NewPivotSheet(wb, tabName, title, about, _
        UCase$(FwLabel(fw)) & "  " & ChrW(183) & "  " & IIf(sf.count > 0, "ONE SHEET PER " & UCase$(JoinC(sf)), "PIVOT"))
    Set pt = NewPivot(ws, "pt_r", sf.count + PageFilterCount(rc))
    If pt Is Nothing Then Exit Function
    pt.ManualUpdate = True

    pos = 0
    For Each x In rc("Rows")
        pos = pos + 1
        RowField pt, CStr(x), pos
    Next x
    pos = 0
    For Each x In rc("Cols")
        pos = pos + 1
        ColField pt, CStr(x), pos
    Next x
    For Each v In rc("Values")
        AddValue pt, v, CStr(rc("Format"))
    Next v
    If rc("Values").count > 1 And rc("Cols").count > 0 Then DataFirst pt

    ' One sheet per: its fields become filters fixed to this sheet's values.
    For i = 1 To sf.count
        PageField pt, CStr(sf(i))
        PickOne pt, CStr(sf(i)), CStr(splitVals(i - 1))
    Next i
    ' Show only / hide: on whichever axis the field is on, or - if it is on
    ' neither - as a report filter of its own.
    For Each flt In rc("Filters")
        If Not modPD_Config.InCollection(rc("Rows"), CStr(flt("Field"))) And _
           Not modPD_Config.InCollection(rc("Cols"), CStr(flt("Field"))) Then
            PageField pt, CStr(flt("Field"))
        End If
        ShowItems pt, CStr(flt("Field")), CBool(flt("Include")), flt("Items")
    Next flt
    RecipeSubtotals pt, rc

    FinishRecipe pt, ws, rc, fl
    NoteSheet ws, what
    If sf.count = 0 And rc("Slicers").count > 0 Then Slicers ws, pt, ToArray(rc("Slicers"))
    RecipeTab ws, CStr(rc("Tab"))
    Set BuildRecipeSheet = ws
End Function

' A value, as the recipe spelled it: which field, how to aggregate, and what
' to call it. Raises if Excel refuses it - a pivot without its value is
' empty, and saying so beats building it.
Private Sub AddValue(ByVal pt As PivotTable, ByVal v As Object, ByVal fmt As String)
    Dim df As PivotField, fn As Long, calc As Long, nf As String
    Select Case CStr(v("Agg"))
        Case "count": fn = xlCount
        Case "average", "avg", "mean": fn = xlAverage
        Case "max": fn = xlMax
        Case "min": fn = xlMin
        Case Else: fn = xlSum
    End Select
    Select Case CStr(v("Agg"))
        Case "%row": calc = xlPercentOfRow
        Case "%col", "%column": calc = xlPercentOfColumn
        Case "%total": calc = xlPercentOfTotal
    End Select
    On Error Resume Next
    Set df = pt.AddDataField(pt.PivotFields(CStr(v("Field"))), CStr(v("Caption")), fn)
    On Error GoTo 0
    If df Is Nothing Then
        Err.Raise vbObjectError + 514, "AddValue", _
            "Could not add " & CStr(v("Field")) & " as a value called " & Chr$(34) & CStr(v("Caption")) & Chr$(34) & "."
    End If
    On Error Resume Next
    If calc <> 0 Then
        df.Calculation = calc
        nf = "0.0%"
    ElseIf fn = xlCount Then
        nf = "#,##0"
    Else
        nf = fmt
    End If
    df.NumberFormat = nf
    ' Set again: changing the calculation can make Excel regenerate it.
    df.caption = CStr(v("Caption"))
    Err.Clear
End Sub

' Which items of a field show. Excel refuses a filter that hides every item,
' so a rule that would leave nothing is left unapplied rather than failing
' the sheet.
Private Sub ShowItems(ByVal pt As PivotTable, ByVal nm As String, ByVal include As Boolean, _
                      ByVal items As Collection)
    Dim pi As PivotItem, pf As PivotField, keep As Long, listed As Boolean
    On Error Resume Next
    Set pf = pt.PivotFields(nm)
    If pf Is Nothing Then Exit Sub
    If pf.Orientation = xlPageField Then pf.EnableMultiplePageItems = True
    For Each pi In pf.PivotItems
        listed = modPD_Config.InCollection(items, pi.Name)
        If listed = include Then keep = keep + 1
    Next pi
    If keep = 0 Then Exit Sub
    For Each pi In pf.PivotItems
        listed = modPD_Config.InCollection(items, pi.Name)
        If listed = include Then pi.visible = True
    Next pi
    For Each pi In pf.PivotItems
        listed = modPD_Config.InCollection(items, pi.Name)
        If listed <> include Then pi.visible = False
    Next pi
    Err.Clear
End Sub

Private Sub RecipeSubtotals(ByVal pt As PivotTable, ByVal rc As Object)
    Dim x As Variant, subs As Object
    Set subs = rc("SubFields")
    On Error Resume Next
    For Each x In rc("Rows")
        If CBool(rc("SubAll")) Or subs.Exists(CStr(x)) Then
            pt.PivotFields(CStr(x)).Subtotals(1) = True       ' automatic
        End If
    Next x
    Err.Clear
End Sub

Private Sub FinishRecipe(ByVal pt As PivotTable, ByVal ws As Worksheet, ByVal rc As Object, ByVal fl As Object)
    On Error Resume Next
    With pt
        .TableStyle2 = PivotStyleFor(ws.Parent)
        .ShowTableStyleRowStripes = True
        .ShowTableStyleColumnHeaders = True
        .ShowTableStyleRowHeaders = True
        .RowAxisLayout CLng(rc("Layout"))
        If CBool(rc("Repeat")) Then
            .RepeatAllLabels xlRepeatLabels
        Else
            .RepeatAllLabels xlDoNotRepeatLabels
        End If
        .ShowDrillIndicators = False
        .EnableDrilldown = True
        .EnableFieldList = True
        .EnableWizard = True
        .DisplayFieldCaptions = True
        .ColumnGrand = CBool(rc("ColGrand"))
        .RowGrand = CBool(rc("RowGrand"))
        .GrandTotalName = "Total"
        .HasAutoFormat = False
        .PreserveFormatting = True
        .NullString = "-"
        .DisplayNullString = True
        .ManualUpdate = False
    End With
    OrderBuckets pt
    SortRecipe pt, rc
    FitPivot ws, pt, rc
    If Not rc.Exists("Tiles") Then
        PivotTiles ws, pt
    ElseIf CBool(rc("Tiles")) Then
        PivotTiles ws, pt
    End If
    PrintPivot ws, pt
    Err.Clear
End Sub

Private Sub SortRecipe(ByVal pt As PivotTable, ByVal rc As Object)
    Dim x As Variant, ord As Long, sortBy As String, v As Object
    sortBy = CStr(rc("SortBy"))
    If Len(sortBy) = 0 Then Exit Sub
    If CBool(rc("SortDesc")) Then ord = xlDescending Else ord = xlAscending
    ' the caption as Excel holds it - it may carry the anti-collision space
    If sortBy <> "label" Then
        For Each v In rc("Values")
            If StrComp(Trim$(CStr(v("Caption"))), sortBy, vbTextCompare) = 0 Then sortBy = CStr(v("Caption"))
        Next v
    End If
    On Error Resume Next
    For Each x In rc("Rows")
        If sortBy = "label" Then
            pt.PivotFields(CStr(x)).AutoSort ord, CStr(x)
        Else
            pt.PivotFields(CStr(x)).AutoSort ord, sortBy
        End If
    Next x
    Err.Clear
End Sub

' Every label column as wide as its longest label, every figure column as wide
' as its widest figure - a recipe's own Widths win - then the zoom that shows
' the table on a laptop, the bar frozen at the top, and the cursor on the first
' figure rather than on the chrome.
Private Sub FitPivot(ByVal ws As Worksheet, ByVal pt As PivotTable, ByVal rc As Object)
    Dim rng As Range, c As Long, i As Long, nLab As Long, lastCol As Long, total As Double, w As Double
    Dim compact As Boolean, pf As PivotField, wd As Object, nm As String, t As String
    On Error Resume Next
    Set rng = pt.TableRange1
    If rng Is Nothing Then Exit Sub
    If Not rc Is Nothing Then
        compact = (CLng(rc("Layout")) = 0)
        Set wd = rc("Widths")
    End If
    nLab = pt.RowFields.count
    If compact And nLab > 1 Then nLab = 1
    lastCol = rng.Column + rng.Columns.count - 1
    For c = rng.Column To lastCol
        i = c - rng.Column + 1
        w = 0
        If i <= nLab And pt.RowFields.count > 0 Then
            If compact Then
                For Each pf In pt.RowFields
                    If LabelChars(pf) + 2 * (pf.Position - 1) > w Then w = LabelChars(pf) + 2 * (pf.Position - 1)
                Next pf
                nm = ""
            Else
                Set pf = pt.RowFields(i)
                nm = pf.Name
                w = LabelChars(pf)
            End If
            If Not wd Is Nothing And Len(nm) > 0 Then
                If wd.Exists(nm) Then w = CDbl(wd(nm))
            End If
            w = w + 3
            If w > 62 Then w = 62
            If w < 10 Then w = 10
        Else
            ' The widest figure in a column is its total; the header may be wider.
            t = ws.Cells(rng.Row + rng.Rows.count - 1, c).Text
            w = Len(t)
            If Len(ws.Cells(rng.Row, c).Text) > w Then w = Len(ws.Cells(rng.Row, c).Text)
            If Len(ws.Cells(rng.Row + 1, c).Text) > w Then w = Len(ws.Cells(rng.Row + 1, c).Text)
            w = w + 4
            If Not rc Is Nothing Then
                If CDbl(rc("ValueWidth")) <> 14 Then w = CDbl(rc("ValueWidth"))
            End If
            If w < 12 Then w = 12
            If w > 30 Then w = 30
        End If
        ws.Columns(c).ColumnWidth = w
        total = total + w
    Next c
    ' Filters and body on even rows, tall enough to read.
    ws.Range(ws.Rows(FILTER_TOP), ws.Rows(rng.Row + rng.Rows.count + 400)).RowHeight = 20
    ws.Rows(rng.Row - 1).RowHeight = 10
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = ZoomFor(total)
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    rng.Cells(IIf(pt.ColumnFields.count > 0, 3, 2), nLab + 1).Select
    Err.Clear
End Sub

' The longest item of a row field, and its heading, in characters.
Private Function LabelChars(ByVal pf As PivotField) As Double
    Dim pi As PivotItem, n As Long
    On Error Resume Next
    n = Len(pf.caption) + 3                 ' the heading and its filter button
    For Each pi In pf.PivotItems
        If pi.visible Then
            If Len(pi.caption) > n Then n = Len(pi.caption)
        End If
    Next pi
    LabelChars = n
    Err.Clear
End Function


Private Sub RecipeTab(ByVal ws As Worksheet, ByVal tabWord As String)
    On Error Resume Next
    Select Case LCase$(tabWord)
        Case "emerald": ws.Tab.Color = modPD_Theme.C_BRAND
        Case "deep": ws.Tab.Color = modPD_Theme.C_BRAND_DEEP
        Case "slate": ws.Tab.Color = modPD_Theme.HX("3E5A50")
        Case "black": ws.Tab.Color = modPD_Theme.C_INK
    End Select
    Err.Clear
End Sub

Private Function ToArray(ByVal c As Collection) As Variant
    Dim a() As String, i As Long
    If c.count = 0 Then ToArray = Array(): Exit Function
    ReDim a(0 To c.count - 1)
    For i = 1 To c.count
        a(i - 1) = CStr(c(i))
    Next i
    ToArray = a
End Function

' ===================== the mechanics ========================================

Private Function NewPivotSheet(ByVal wb As Workbook, ByVal wanted As String, _
                               ByVal title As String, ByVal about As String, _
                               Optional ByVal overline As String = "") As Worksheet
    Dim ws As Worksheet, nm As String
    nm = FreeSheetName(wanted, wb)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    On Error Resume Next
    ws.Name = nm
    Err.Clear
    On Error GoTo 0
    modPD_Theme.Dress ws, title, about, overline
    modPD_Theme.BookBar ws, mBook, True
    ' Until tiles or slicers arrive, their rows are only breathing room.
    ws.Rows(TILE_ROW).RowHeight = 6
    ws.Rows(5).RowHeight = 6
    ws.Rows(SLICER_BAND_ROW).RowHeight = 4
    ws.Rows(7).RowHeight = 8
    ' Tabs say what kind of sheet they are before they are opened: the
    ' overviews in emerald, the local-currency sheets deep green, foreign
    ' currency slate.
    Select Case True
        Case Right$(nm, 4) = " FCY": ws.Tab.Color = modPD_Theme.HX("3E5A50")
        Case Right$(nm, 4) = " LCY": ws.Tab.Color = modPD_Theme.C_BRAND_DEEP
        Case Else: ws.Tab.Color = modPD_Theme.C_BRAND
    End Select
    Set NewPivotSheet = ws
End Function

' A pivot placed so its filters - there will be nPages of them - sit on rows
' of their own from row 8, with one blank row between them and the table.
Private Function NewPivot(ByVal ws As Worksheet, ByVal nm As String, ByVal nPages As Long) As PivotTable
    On Error Resume Next
    mSeq = mSeq + 1
    Set NewPivot = mCache.CreatePivotTable(TableDestination:=ws.Cells(FILTER_TOP + nPages + 1, 1), _
                                           TableName:=nm & "_" & mSeq)
    Err.Clear
End Function

' How many report filters a recipe's pivot will carry beyond its one-sheet-
' per fields: the Show only / hide rules on fields that are not on an axis.
Private Function PageFilterCount(ByVal rc As Object) As Long
    Dim flt As Object, seen As Object
    Set seen = NewMap()
    For Each flt In rc("Filters")
        If Not modPD_Config.InCollection(rc("Rows"), CStr(flt("Field"))) And _
           Not modPD_Config.InCollection(rc("Cols"), CStr(flt("Field"))) Then seen(CStr(flt("Field"))) = True
    Next flt
    PageFilterCount = seen.count
End Function

Private Function JoinC(ByVal c As Collection) As String
    Dim i As Long
    For i = 1 To c.count
        If i > 1 Then JoinC = JoinC & " and "
        JoinC = JoinC & CStr(c(i))
    Next i
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
        .TableStyle2 = PivotStyleFor(ws.Parent)
        .ShowTableStyleRowStripes = True
        .ShowTableStyleColumnHeaders = True
        .ShowTableStyleRowHeaders = True
        .RowAxisLayout xlTabularRow          ' one field per column, not nested
        .RepeatAllLabels xlDoNotRepeatLabels
        .ShowDrillIndicators = False         ' no +/- boxes down the labels
        .EnableDrilldown = True
        .EnableFieldList = True
        .EnableWizard = True
        .DisplayFieldCaptions = True
        .ColumnGrand = True
        .RowGrand = True
        .GrandTotalName = "Total"
        .HasAutoFormat = False               ' stop autofit fighting the widths
        .PreserveFormatting = True
        .NullString = "-"                    ' an empty cell reads as nothing
        .DisplayNullString = True
        .ManualUpdate = False
    End With
    OrderBuckets pt
    FitPivot ws, pt, Nothing
    PivotTiles ws, pt
    PrintPivot ws, pt
    Err.Clear
End Sub

' Buckets in tenor order - UPTO 1 MONTH, 1 - 3 MONTHS, ... OVER 5 YEARS,
' NON MATURITY - rather than alphabetical, which puts OVER 5 YEARS before UPTO
' 1 MONTH. The order is the staging pass's (modPD_Stage.BucketOrder), set item
' by item on this pivot; a Bucket field that is only a filter is left alone.
Private Sub OrderBuckets(ByVal pt As PivotTable)
    Dim pf As PivotField, order As Variant, i As Long, pos As Long, pi As PivotItem
    On Error Resume Next
    Set pf = pt.PivotFields(H_BUCKET)
    If pf Is Nothing Then Exit Sub
    If pf.Orientation <> xlRowField And pf.Orientation <> xlColumnField Then Exit Sub
    order = modPD_Stage.BucketOrder()
    If UBound(order) < 1 Then Exit Sub
    pt.ManualUpdate = True
    For i = 0 To UBound(order)
        Set pi = Nothing
        Set pi = pf.PivotItems(CStr(order(i)))
        If Not pi Is Nothing Then
            If pi.visible Then
                pos = pos + 1
                pi.Position = pos
            End If
        End If
    Next i
    pt.ManualUpdate = False
    Err.Clear
End Sub

' Landscape, one page wide, the pivot's own headers repeated on every page,
' and a footer that says which book and which sheet. PrintCommunication off
' while it is set: with a hundred sheets, talking to the printer driver for
' each property is the difference between a second and a minute.
Private Sub PrintPivot(ByVal ws As Worksheet, ByVal pt As PivotTable)
    On Error Resume Next
    pt.PrintTitles = True
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .BlackAndWhite = True
        .CenterHeader = "&""Segoe UI Semibold,Regular""&11&A"
        .LeftFooter = "&8" & TOOL_NAME & " ALM Desk  " & ChrW(183) & "  &F"
        .CenterFooter = ""
        .RightFooter = "&8Page &P of &N"
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub

' ===================== tiles ================================================
'
' Up to four tiles over a pivot, one per value: its grand total. Each is kept
' live by GETPIVOTDATA in a cell far to the right, and the tile's text is
' linked to that cell - so a slicer, a filter or a refresh moves the pivot and
' the tile follows. Values shown as a share of something have no total worth
' a tile and are passed over.
Private Sub PivotTiles(ByVal ws As Worksheet, ByVal pt As PivotTable)
    Dim df As PivotField, n As Long, x As Double, cell As Range, mag As Double, anchor As String, cap As String
    On Error Resume Next
    anchor = pt.TableRange1.Cells(1, 1).Address
    x = 14
    For Each df In pt.DataFields
        If n >= 4 Then Exit For
        If df.Calculation = xlNoAdditionalCalculation Then
            n = n + 1
            cap = df.caption
            Set cell = ws.Cells(TILE_ROW, TILE_HELPER_COL + n)
            cell.Formula = "=IFERROR(GETPIVOTDATA(""" & Replace(cap, """", """""") & """," & anchor & "),""-"")"
            mag = Abs(SafeNum(pt.GetPivotData(cap).value))
            cell.NumberFormat = CompactFormat(mag, df.Function)
            cell.HorizontalAlignment = xlLeft
            cell.Font.Color = modPD_Theme.C_SHEET
            modPD_Theme.Tile ws, "pdb_tile" & n, x, ws.Rows(TILE_ROW).Top + 8, 196, 50, _
                             UCase$(Trim$(cap)) & IIf(df.Function = xlCount, "", "  " & ChrW(183) & "  TOTAL"), "", cell
            x = x + 206
        End If
    Next df
    If n > 0 Then ws.Rows(TILE_ROW).RowHeight = 66
    Err.Clear
End Sub

' The unit a figure of this size is best read in.
Private Function CompactFormat(ByVal mag As Double, ByVal fn As Long) As String
    If fn = xlCount Then
        CompactFormat = "#,##0"
    ElseIf mag >= 1000000000# Then
        CompactFormat = "#,##0.00,,,"" bn"";-#,##0.00,,,"" bn"";""-"""
    ElseIf mag >= 1000000# Then
        CompactFormat = "#,##0.0,,"" m"";-#,##0.0,,"" m"";""-"""
    Else
        CompactFormat = NUM_FMT
    End If
End Function

' ===================== between sheets =======================================
'
' Previous and next on every pivot sheet's bar, in the order the index lists
' them - a family of forty rule sheets is read one after another, not by
' going back to the index each time. Links, not macros: the book has none.
Public Sub LinkSiblings(ByVal wb As Workbook)
    Dim made As Collection, i As Long, ws As Worksheet, sh As Shape, nm As String
    On Error Resume Next
    Set made = MadeSheets()
    For i = 1 To made.count
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(made(i)(0)))
        If Not ws Is Nothing Then
            If i > 1 Then
                nm = CStr(made(i - 1)(0))
                Set sh = modPD_Theme.Pill(ws, "pdb_prev", ChrW(8249) & "  Previous", "", 0, 8, 88, 24, 0)
                PlaceAfterBack ws, sh, 0
                ws.Hyperlinks.Add Anchor:=sh, Address:="", SubAddress:="'" & Replace(nm, "'", "''") & "'!A1", _
                                  ScreenTip:=nm
            End If
            If i < made.count Then
                nm = CStr(made(i + 1)(0))
                Set sh = modPD_Theme.Pill(ws, "pdb_next", "Next  " & ChrW(8250), "", 0, 8, 72, 24, 0)
                PlaceAfterBack ws, sh, 94
                ws.Hyperlinks.Add Anchor:=sh, Address:="", SubAddress:="'" & Replace(nm, "'", "''") & "'!A1", _
                                  ScreenTip:=nm
            End If
        End If
    Next i
    Err.Clear
End Sub

Private Sub PlaceAfterBack(ByVal ws As Worksheet, ByVal sh As Shape, ByVal offset As Double)
    Dim back As Shape
    On Error Resume Next
    Set back = ws.Shapes("pdb_back")
    If back Is Nothing Then Exit Sub
    sh.Left = back.Left + back.Width + 8 + offset
    sh.Top = back.Top
    ' Quiet pills on the black bar, with a hairline so they read as buttons.
    sh.Line.visible = msoTrue
    sh.Line.ForeColor.RGB = modPD_Theme.C_LINE_2
    sh.Line.Weight = 0.75
    Err.Clear
End Sub

' "Fit it on one screen without adjusting it every time."
'
' Row-label columns get width by what they hold - a COA name needs room, a type
' does not - and every value column gets one narrow width, because a figure with
' no decimals and a thousands separator has a known size. Then the sheet's own
' zoom is set from how wide the result came out, so it lands on screen whatever
' the pivot turned out to be.


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
    Dim yTop As Double
    ' The slicers get a band of their own between the masthead and the pivot,
    ' so they sit over nothing.
    On Error Resume Next
    ws.Rows(SLICER_BAND_ROW).RowHeight = SLICER_H + 12
    yTop = ws.Rows(SLICER_BAND_ROW).Top + 6
    Err.Clear
    On Error GoTo 0
    x = 14
    For i = 0 To UBound(fields)
        nm = CStr(fields(i))
        Set sc = Nothing
        Set sl = Nothing
        why = ""

        Set sc = CacheFor(ws.Parent, pt, nm, why)
        If Not sc Is Nothing Then
            On Error Resume Next
            Set sl = sc.Slicers.Add(ws, , SlicerName(nm, i), nm, yTop, x, 156, SLICER_H)
            If sl Is Nothing Then why = "slicer: " & Err.Number & " " & Err.Description
            Err.Clear
            On Error GoTo 0
        End If

        If sl Is Nothing Then
            LogIt V_CHECK, "Slicer", "No slicer for " & Chr$(34) & nm & Chr$(34) & " - " & why, ws.Name
        Else
            On Error Resume Next
            sl.style = SlicerStyleFor(ws.Parent)
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


' The Avati pivot style, made once per workbook: the Desk's rows - near-black,
' every other one a shade lighter, hairlines between - a black header with
' mint type over an emerald rule, subtotals a step brighter, and the total on
' the deep emerald band with a rule above it. The grand-total column's heading
' is set explicitly: left to inherit, it was pale type on a pale fill.
'
' Element indexes are the XlTableStyleElementType values, written as numbers
' so a name this Excel does not know is a skipped line, not a module that
' will not compile.
Private Function PivotStyleFor(ByVal wb As Workbook) As String
    Dim ts As TableStyle
    On Error Resume Next
    Set ts = wb.TableStyles(PT_CUSTOM)
    If ts Is Nothing Then
        Set ts = wb.TableStyles.Add(PT_CUSTOM)
        If ts Is Nothing Then
            PivotStyleFor = PT_STYLE
            Exit Function
        End If
        ts.ShowAsAvailablePivotTableStyle = True
        ts.ShowAsAvailableTableStyle = False
        With ts.TableStyleElements(0)                  ' whole table
            .Interior.Color = modPD_Theme.C_ROW
            .Font.Color = modPD_Theme.C_TEXT
            .Borders(12).LineStyle = xlContinuous      ' inside horizontal
            .Borders(12).Color = modPD_Theme.C_LINE
            .Borders(9).LineStyle = xlContinuous       ' bottom edge
            .Borders(9).Color = modPD_Theme.C_LINE_2
        End With
        With ts.TableStyleElements(1)                  ' header row
            .Interior.Color = modPD_Theme.C_INK
            .Font.Color = modPD_Theme.C_BRAND_SOFT
            .Font.Bold = True
            .Borders(9).LineStyle = xlContinuous
            .Borders(9).Color = modPD_Theme.C_BRAND
            .Borders(9).Weight = xlMedium
        End With
        With ts.TableStyleElements(9)                  ' first header cell
            .Interior.Color = modPD_Theme.C_INK
            .Font.Color = modPD_Theme.C_BRAND_SOFT
            .Font.Bold = True
        End With
        With ts.TableStyleElements(10)                 ' last header cell: the Total column's heading
            .Interior.Color = modPD_Theme.C_INK
            .Font.Color = modPD_Theme.C_TEXT
            .Font.Bold = True
        End With
        With ts.TableStyleElements(2)                  ' grand total row
            .Interior.Color = modPD_Theme.C_TOTAL
            .Font.Bold = True
            .Font.Color = modPD_Theme.C_TEXT
            .Borders(8).LineStyle = xlContinuous       ' top edge
            .Borders(8).Color = modPD_Theme.C_BRAND
            .Borders(8).Weight = xlMedium
        End With
        With ts.TableStyleElements(4)                  ' grand total column
            .Interior.Color = modPD_Theme.HX("16271F")
            .Font.Color = modPD_Theme.C_TEXT
            .Font.Bold = True
        End With
        With ts.TableStyleElements(5)                  ' row stripe 1
            .Interior.Color = modPD_Theme.C_ROW_ALT
        End With
        With ts.TableStyleElements(16)                 ' subtotal row 1
            .Interior.Color = modPD_Theme.HX("16271F")
            .Font.Bold = True
            .Font.Color = modPD_Theme.C_TEXT
        End With
        With ts.TableStyleElements(17)                 ' subtotal row 2
            .Interior.Color = modPD_Theme.C_ROW_ALT
            .Font.Bold = True
        End With
        With ts.TableStyleElements(20)                 ' column subheading 1
            .Interior.Color = modPD_Theme.C_INK
            .Font.Color = modPD_Theme.C_BRAND_SOFT
            .Font.Bold = True
        End With
        With ts.TableStyleElements(23)                 ' row subheading 1
            .Font.Bold = True
            .Font.Color = modPD_Theme.C_TEXT
        End With
        With ts.TableStyleElements(26)                 ' report filter labels
            .Interior.Color = modPD_Theme.C_ROW_ALT
            .Font.Bold = True
            .Font.Color = modPD_Theme.C_LINK
        End With
        With ts.TableStyleElements(27)                 ' report filter values
            .Interior.Color = modPD_Theme.C_ROW
            .Font.Color = modPD_Theme.C_TEXT
        End With
    End If
    If ts Is Nothing Then PivotStyleFor = PT_STYLE Else PivotStyleFor = PT_CUSTOM
    Err.Clear
End Function

' The slicers in the same room: dark tiles, the chosen items in emerald, the
' ones with no data dimmed.
Private Function SlicerStyleFor(ByVal wb As Workbook) As String
    Dim ts As TableStyle
    On Error Resume Next
    Set ts = wb.TableStyles(SLICER_CUSTOM)
    If ts Is Nothing Then
        Set ts = wb.TableStyles.Add(SLICER_CUSTOM)
        If ts Is Nothing Then
            SlicerStyleFor = SLICER_STYLE
            Exit Function
        End If
        ts.ShowAsAvailableSlicerStyle = True
        ts.ShowAsAvailablePivotTableStyle = False
        ts.ShowAsAvailableTableStyle = False
        With ts.TableStyleElements(0)                  ' whole slicer
            .Interior.Color = modPD_Theme.C_ROW
            .Font.Color = modPD_Theme.C_TEXT_2
            .Borders(xlEdgeLeft).Color = modPD_Theme.C_LINE_2
            .Borders(xlEdgeTop).Color = modPD_Theme.C_LINE_2
            .Borders(xlEdgeRight).Color = modPD_Theme.C_LINE_2
            .Borders(xlEdgeBottom).Color = modPD_Theme.C_LINE_2
        End With
        With ts.TableStyleElements(1)                  ' header
            .Font.Color = modPD_Theme.C_TEXT
            .Font.Bold = True
        End With
        With ts.TableStyleElements(28)                 ' unselected, with data
            .Interior.Color = modPD_Theme.C_ROW_ALT
            .Font.Color = modPD_Theme.C_TEXT_2
        End With
        With ts.TableStyleElements(29)                 ' unselected, no data
            .Interior.Color = modPD_Theme.C_ROW
            .Font.Color = modPD_Theme.C_TX4
        End With
        With ts.TableStyleElements(30)                 ' selected, with data
            .Interior.Color = modPD_Theme.C_BRAND_DEEP
            .Font.Color = modPD_Theme.C_TEXT
        End With
        With ts.TableStyleElements(31)                 ' selected, no data
            .Interior.Color = modPD_Theme.C_BRAND_900
            .Font.Color = modPD_Theme.C_TEXT_3
        End With
        With ts.TableStyleElements(32)                 ' hovered, unselected
            .Interior.Color = modPD_Theme.HX("16271F")
            .Font.Color = modPD_Theme.C_TEXT
        End With
        With ts.TableStyleElements(33)                 ' hovered, selected
            .Interior.Color = modPD_Theme.C_BRAND
            .Font.Color = modPD_Theme.C_TEXT
        End With
    End If
    If ts Is Nothing Then SlicerStyleFor = SLICER_STYLE Else SlicerStyleFor = SLICER_CUSTOM
    Err.Clear
End Function
