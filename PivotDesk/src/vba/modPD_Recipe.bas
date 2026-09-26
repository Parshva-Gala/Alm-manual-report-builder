Option Explicit

' ============================================================================
'  Reading a recipe - one row of Pivot config - and checking it.
'
'  Each column is read into a Dictionary the pivot builder works from: the
'  fields for each axis, the values with how each is aggregated and shown,
'  the item, label and value filters, the groups, and every setting of how
'  the sheet looks. Nothing is guessed: a cell that does not say something
'  the builder can do is a problem, reported in words beside the row.
'
'  The rules, in the order Check applies them, are in Validate. Its messages
'  point at the column to change.
' ============================================================================

Private Function Cell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Cell = SafeText(ws.Cells(r, c).Value2)
End Function

' Every row of the table, parsed and checked. Rows that are off are parsed
' too, so Check can say what is wrong with them before anyone switches them on.
Public Function AllRecipes() As Collection
    Dim ws As Worksheet, out As Collection, r As Long, lastR As Long, fl As Object
    Dim seenNames As Object, rc As Object
    Set out = New Collection
    Set AllRecipes = out
    Set ws = GetSheet(SH_CONFIG)
    If ws Is Nothing Then Exit Function
    Set fl = modPD_Config.Fields()
    Set seenNames = NewMap()
    lastR = modPD_Config.RecipeLastRow(ws)
    For r = modPD_Theme.R_FIRST To lastR
        If Len(SafeText(ws.Cells(r, K_NAME).Value2)) > 0 Or Len(SafeText(ws.Cells(r, K_ROWS).Value2)) > 0 Then
            Set rc = ParseRow(ws, r, fl)
            If CBool(rc("On")) And Len(rc("Problem")) = 0 Then
                If seenNames.Exists(rc("Name") & "|" & rc("Frameworks")) Then
                    rc("Problem") = "Another row that is on has the same pivot name and frameworks."
                Else
                    seenNames(rc("Name") & "|" & rc("Frameworks")) = True
                End If
            End If
            out.Add rc
        End If
    Next r
End Function

' The recipes a build of this framework will use: on, and for it.
Public Function RecipesFor(ByVal fw As String) As Collection
    Dim out As Collection, rc As Object
    Set out = New Collection
    For Each rc In AllRecipes()
        If CBool(rc("On")) And ForFramework(rc, fw) Then out.Add rc
    Next rc
    Set RecipesFor = out
End Function

Private Function ForFramework(ByVal rc As Object, ByVal fw As String) As Boolean
    Dim t As Variant
    For Each t In SplitList(CStr(rc("Frameworks")), ",")
        Select Case NormFw(CStr(t))
            Case "ALL": ForFramework = True: Exit Function
            Case UCase$(fw): ForFramework = True: Exit Function
        End Select
    Next t
End Function

Public Function NormFw(ByVal s As String) As String
    Select Case UCase$(Replace(Replace(Trim$(s), " ", ""), "_", ""))
        Case "ALL", "": NormFw = "ALL"
        Case "LCR": NormFw = FW_LCR
        Case "NSFR": NormFw = FW_NSFR
        Case "MATURITYLADDER", "LADDER", "ML": NormFw = FW_ML
        Case Else: NormFw = "?" & s
    End Select
End Function

Private Function ParseRow(ByVal ws As Worksheet, ByVal r As Long, ByVal fl As Object) As Object
    Dim rc As Object, s As String, prob As String, x As Variant, subs As Object, g As Object, gOf As Object
    Set rc = NewMap()
    Set ParseRow = rc
    rc("Row") = r
    rc("On") = (StrComp(Cell(ws, r, K_ON), "No", vbTextCompare) <> 0 And Len(Cell(ws, r, K_ON)) > 0)
    rc("Name") = Cell(ws, r, K_NAME)
    rc("Frameworks") = IIf(Len(Cell(ws, r, K_FW)) = 0, "All", Cell(ws, r, K_FW))
    Set rc("Split") = SplitList(Cell(ws, r, K_SPLIT), ",")
    Set rc("Rows") = SplitList(Cell(ws, r, K_ROWS), ",")
    Set rc("Cols") = SplitList(Cell(ws, r, K_COLS), ",")
    Set rc("Slicers") = SplitList(Cell(ws, r, K_SLICERS), ",")
    ' Every column is read even after one fails, so every key exists and the
    ' first problem is the one reported.
    Set rc("Values") = ParseValues(Cell(ws, r, K_VALUES), fl, prob)
    Set rc("Filters") = ParseFilters(Cell(ws, r, K_FILTERS), prob)
    Set rc("VFilters") = ParseVFilters(Cell(ws, r, K_VFILTER), prob)
    Set rc("Groups") = ParseGroups(Cell(ws, r, K_GROUP), prob)
    Set gOf = NewMap()
    For Each g In rc("Groups")
        gOf(CStr(g("Field"))) = CStr(g("Name"))
    Next g
    Set rc("GroupOf") = gOf
    rc("Desc") = Cell(ws, r, K_DESC)

    ' how the figures read
    s = LCase$(Cell(ws, r, K_UNITS))
    Select Case s
        Case "", "as is": rc("Units") = ""
        Case "thousands", "millions", "billions": rc("Units") = s
        Case Else
            rc("Units") = ""
            Fail prob, "Units must be As is, Thousands, Millions or Billions."
    End Select
    rc("Format") = Cell(ws, r, K_FORMAT)
    rc("FormatSet") = (Len(rc("Format")) > 0)
    If Not CBool(rc("FormatSet")) Then rc("Format") = UnitFormat(CStr(rc("Units")))
    ParseHighlight Cell(ws, r, K_HILITE), rc, prob

    Select Case LCase$(Cell(ws, r, K_LAYOUT))
        Case "", "tabular": rc("Layout") = 1          ' xlTabularRow
        Case "outline": rc("Layout") = 2              ' xlOutlineRow
        Case "compact": rc("Layout") = 0              ' xlCompactRow
        Case Else
            rc("Layout") = 1
            Fail prob, "Layout must be Tabular, Outline or Compact."
    End Select

    s = LCase$(Cell(ws, r, K_SUBTOT))
    Set subs = NewMap()
    Set rc("SubFields") = subs
    rc("SubAll") = False
    If s = "all" Then
        rc("SubAll") = True
    ElseIf s <> "" And s <> "none" Then
        For Each x In SplitList(Cell(ws, r, K_SUBTOT), ",")
            subs(CStr(x)) = True
        Next x
    End If
    Select Case LCase$(Cell(ws, r, K_SUBAT))
        Case "": rc("SubAt") = 0
        Case "top": rc("SubAt") = 1                   ' xlAtTop
        Case "bottom": rc("SubAt") = 2                ' xlAtBottom
        Case Else
            rc("SubAt") = 0
            Fail prob, "Subtotals at must be Top or Bottom."
    End Select

    Select Case LCase$(Cell(ws, r, K_GRAND))
        Case "", "both": rc("ColGrand") = True: rc("RowGrand") = True
        Case "bottom row": rc("ColGrand") = True: rc("RowGrand") = False
        Case "right column": rc("ColGrand") = False: rc("RowGrand") = True
        Case "none": rc("ColGrand") = False: rc("RowGrand") = False
        Case Else
            rc("ColGrand") = True: rc("RowGrand") = True
            Fail prob, "Grand totals must be Both, Bottom row, Right column or None."
    End Select
    rc("TotalLabel") = Cell(ws, r, K_TOTAL)
    If Len(rc("TotalLabel")) = 0 Then rc("TotalLabel") = "Total"
    rc("Repeat") = YesNo(Cell(ws, r, K_REPEAT), False, "Repeat labels", prob)
    rc("BlankLine") = YesNo(Cell(ws, r, K_BLANKLN), False, "Blank line", prob)
    rc("Tiles") = YesNo(Cell(ws, r, K_TILES), True, "Tiles", prob)
    Select Case LCase$(Cell(ws, r, K_VALIN))
        Case "", "columns", "column": rc("ValuesInRows") = False
        Case "rows", "row": rc("ValuesInRows") = True
        Case Else
            rc("ValuesInRows") = False
            Fail prob, "Values in must be Columns or Rows."
    End Select
    rc("ExpandTo") = Cell(ws, r, K_EXPAND)

    ParseSort Cell(ws, r, K_SORT), rc
    Set rc("Widths") = ParseWidths(Cell(ws, r, K_WIDTHS), rc, prob)
    rc("Tab") = Cell(ws, r, K_TAB)
    s = Cell(ws, r, K_MAX)
    If Len(s) = 0 Then
        rc("Max") = MAX_SHEETS_DEFAULT
    ElseIf IsNumeric(s) Then
        rc("Max") = CLng(Val(s))
        If CLng(rc("Max")) < 1 Or CLng(rc("Max")) > 250 Then Fail prob, "Max sheets must be between 1 and 250."
    Else
        rc("Max") = MAX_SHEETS_DEFAULT
        Fail prob, "Max sheets must be a number."
    End If

    If Len(prob) = 0 Then prob = Validate(rc, fl)
    rc("Problem") = prob
End Function

' The first problem found is the one reported: fix it, and Check says the next.
Private Sub Fail(ByRef prob As String, ByVal msg As String)
    If Len(prob) = 0 Then prob = msg
End Sub

Private Function YesNo(ByVal s As String, ByVal dflt As Boolean, ByVal what As String, ByRef prob As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "": YesNo = dflt
        Case "yes", "y", "true": YesNo = True
        Case "no", "n", "false": YesNo = False
        Case Else
            YesNo = dflt
            Fail prob, what & " must be Yes or No."
    End Select
End Function

' The desk's own format, read in thousands, millions or billions. The data
' is not changed: a filter or a total still works on the full amounts.
Public Function UnitFormat(ByVal units As String) As String
    Select Case units
        Case "thousands": UnitFormat = "#,##0,;[Red](#,##0,);-"
        Case "millions": UnitFormat = "#,##0.0,,;[Red](#,##0.0,,);-"
        Case "billions": UnitFormat = "#,##0.00,,,;[Red](#,##0.00,,,);-"
        Case Else: UnitFormat = NUM_FMT
    End Select
End Function

' "Data bars", "Heatmap on Exposure", "Top 5 on Share"
Private Sub ParseHighlight(ByVal s As String, ByVal rc As Object, ByRef prob As String)
    Dim p As Long, w As String
    rc("Hilite") = ""
    rc("HiliteOn") = ""
    rc("HiliteN") = 10
    s = Trim$(s)
    p = InStr(1, LCase$(s), " on ")
    If p > 0 Then
        rc("HiliteOn") = Trim$(Mid$(s, p + 4))
        s = Trim$(Left$(s, p - 1))
    End If
    w = LCase$(s)
    If w = "" Or w = "none" Then
        rc("HiliteOn") = ""
    ElseIf w = "data bars" Or w = "data bar" Or w = "bars" Then
        rc("Hilite") = "bars"
    ElseIf w = "heatmap" Or w = "heat map" Or w = "colour scale" Or w = "color scale" Then
        rc("Hilite") = "heat"
    ElseIf w = "negatives" Or w = "negative" Then
        rc("Hilite") = "neg"
    ElseIf w = "top" Or Left$(w, 4) = "top " Then
        rc("Hilite") = "top"
        If Len(w) > 4 Then
            If IsNumeric(Mid$(w, 5)) Then
                rc("HiliteN") = CLng(Val(Mid$(w, 5)))
                If CLng(rc("HiliteN")) < 1 Or CLng(rc("HiliteN")) > 1000 Then Fail prob, "Highlight: Top takes 1 to 1000."
            Else
                Fail prob, "Highlight: Top is followed by how many - e.g. Top 10."
            End If
        End If
    Else
        Fail prob, "Highlight must be None, Data bars, Heatmap, Negatives or Top 10 - add on and a caption " & _
                   "to pick one value."
    End If
End Sub

Public Function SplitList(ByVal s As String, ByVal sep As String) As Collection
    Dim c As Collection, p As Variant
    Set c = New Collection
    If Len(Trim$(s)) > 0 Then
        For Each p In Split(s, sep)
            If Len(Trim$(CStr(p))) > 0 Then c.Add Trim$(CStr(p))
        Next p
    End If
    Set SplitList = c
End Function

' "Pre factor amount sum as Pre-factor; Gross pre-factor %running in Counterparty as Cumulative"
'
'   field  [sum|count|average|max|min]  [a calculation]  [in field]  [as caption]
'
' The words are peeled off the end only while what is left is not already a
' field, so a field called "Price index" or "Days in arrears" reads as itself.
Public Function ParseValues(ByVal s As String, ByVal fl As Object, ByRef prob As String) As Collection
    Dim out As Collection, it As Variant, item As String, cap As String, p As Long, w As String
    Dim v As Object, seenCap As Object, along As String
    Set out = New Collection
    Set ParseValues = out
    Set seenCap = NewMap()
    For Each it In SplitList(s, ";")
        item = CStr(it)
        Set v = NewMap()
        v("Agg") = "sum"
        v("Calc") = ""
        v("Base") = ""
        v("RankDesc") = True
        cap = ""
        p = InStr(1, LCase$(item), " as ")
        If p > 0 Then
            cap = Trim$(Mid$(item, p + 4))
            item = Trim$(Left$(item, p - 1))
        End If
        ' "... in Counterparty": what a running total, a rank or a share runs along.
        p = InStrRev(LCase$(item), " in ")
        If p > 0 And Not fl.Exists(item) Then
            along = Trim$(Mid$(item, p + 4))
            If fl.Exists(along) Then
                v("Base") = along
                item = Trim$(Left$(item, p - 1))
            End If
        End If
        w = LastWord(item)
        If Not fl.Exists(item) And (w = "asc" Or w = "desc") Then
            If LastWord(DropLast(item)) = "rank" Then
                v("RankDesc") = (w = "desc")
                item = DropLast(item)
                w = LastWord(item)
            End If
        End If
        If Not fl.Exists(item) And IsCalc(w) And Len(DropLast(item)) > 0 Then
            v("Calc") = CalcKey(w)
            item = DropLast(item)
            w = LastWord(item)
        End If
        If Not fl.Exists(item) And IsAgg(w) And Len(DropLast(item)) > 0 Then
            v("Agg") = AggKey(w)
            item = DropLast(item)
        ElseIf Not fl.Exists(item) And (IsAgg(w) Or IsCalc(w)) And Len(DropLast(item)) = 0 Then
            Fail prob, "A value is missing its field: " & Chr$(34) & CStr(it) & Chr$(34) & "."
        End If
        v("Field") = item
        If Len(cap) = 0 Then cap = DefaultCaption(item, CStr(v("Agg")), CStr(v("Calc")))
        ' Excel refuses a data field caption that equals a field name - the
        ' refusal surfaces three calls later as a pivot with no data. A
        ' trailing space is a different name to Excel and the same to a reader.
        If fl.Exists(cap) Then cap = cap & " "
        Do While seenCap.Exists(cap)
            cap = cap & " "
        Loop
        seenCap(cap) = True
        v("Caption") = cap
        out.Add v
    Next it
End Function

Private Function LastWord(ByVal s As String) As String
    s = Trim$(s)
    LastWord = LCase$(Mid$(s, InStrRev(s, " ") + 1))
End Function

Private Function DropLast(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStrRev(s, " ")
    If p > 0 Then DropLast = Trim$(Left$(s, p - 1))
End Function

Private Function IsAgg(ByVal w As String) As Boolean
    Select Case LCase$(Trim$(w))
        Case "sum", "count", "average", "avg", "mean", "max", "min"
            IsAgg = True
    End Select
End Function

Private Function AggKey(ByVal w As String) As String
    Select Case LCase$(Trim$(w))
        Case "avg", "mean": AggKey = "average"
        Case Else: AggKey = LCase$(Trim$(w))
    End Select
End Function

Private Function IsCalc(ByVal w As String) As Boolean
    IsCalc = (Len(CalcKey(w)) > 0)
End Function

' Every way a value can be shown, under one name each.
Private Function CalcKey(ByVal w As String) As String
    Select Case LCase$(Trim$(w))
        Case "%row": CalcKey = "%row"
        Case "%col", "%column": CalcKey = "%col"
        Case "%total", "%grand": CalcKey = "%total"
        Case "%parent": CalcKey = "%parent"
        Case "%parentrow": CalcKey = "%parentrow"
        Case "%parentcol", "%parentcolumn": CalcKey = "%parentcol"
        Case "running": CalcKey = "running"
        Case "%running": CalcKey = "%running"
        Case "rank": CalcKey = "rank"
        Case "diff": CalcKey = "diff"
        Case "%diff": CalcKey = "%diff"
        Case "index": CalcKey = "index"
    End Select
End Function

Private Function DefaultCaption(ByVal fld As String, ByVal agg As String, ByVal calc As String) As String
    Dim w As String
    Select Case calc
        Case "%row": w = "% of row"
        Case "%col": w = "% of column"
        Case "%total": w = "% of total"
        Case "%parent": w = "% of parent"
        Case "%parentrow": w = "% of parent row"
        Case "%parentcol": w = "% of parent column"
        Case "running": w = "Running total"
        Case "%running": w = "% running total"
        Case "rank": w = "Rank"
        Case "diff": w = "Change"
        Case "%diff": w = "% change"
        Case "index": w = "Index"
        Case Else
            Select Case agg
                Case "count": w = "Count"
                Case "average": w = "Average"
                Case "max": w = "Max"
                Case "min": w = "Min"
                Case Else: w = "Sum"
            End Select
    End Select
    DefaultCaption = w & " of " & fld
End Function

' "Bucket <> (no bucket); LCY / FCY = LCY; Counterparty contains BANK"
Public Function ParseFilters(ByVal s As String, ByRef prob As String) As Collection
    Dim out As Collection, it As Variant, item As String, p As Long, f As Object, rest As String
    Dim pEq As Long, pNe As Long, pOp As Long, op As Variant, q As Long, opFound As String
    Set out = New Collection
    Set ParseFilters = out
    For Each it In SplitList(s, ";")
        item = CStr(it)
        Set f = NewMap()
        pOp = 0
        opFound = ""
        For Each op In LabelOps()
            q = InStr(1, LCase$(item), " " & CStr(op) & " ")
            If q > 0 Then
                If pOp = 0 Or q < pOp Then
                    pOp = q
                    opFound = CStr(op)
                End If
            End If
        Next op
        pNe = InStr(item, "<>")
        pEq = InStr(item, "=")
        ' A label rule when its words come before any = or <>.
        If pOp > 0 And (pNe = 0 Or pOp < pNe) And (pEq = 0 Or pOp < pEq) Then
            f("Kind") = "label"
            f("Field") = Trim$(Left$(item, pOp - 1))
            f("Op") = opFound
            f("Text") = Trim$(Mid$(item, pOp + Len(opFound) + 2))
            f("Include") = True
            Set f("Items") = New Collection
            If Len(f("Text")) = 0 Then
                Fail prob, "A show-only / hide rule says " & opFound & " what: " & Chr$(34) & item & Chr$(34) & "."
                Exit Function
            End If
        Else
            f("Kind") = "items"
            If pNe > 0 And (pEq = 0 Or pNe < pEq) Then
                p = pNe
                f("Include") = False
                rest = Mid$(item, p + 2)
            ElseIf pEq > 0 Then
                p = pEq
                f("Include") = True
                rest = Mid$(item, p + 1)
            Else
                Fail prob, "A show-only / hide rule needs =, <> or contains: " & Chr$(34) & item & Chr$(34) & "."
                Exit Function
            End If
            f("Field") = Trim$(Left$(item, p - 1))
            Set f("Items") = SplitList(rest, "|")
            If f("Items").count = 0 Then
                Fail prob, "A show-only / hide rule names no items: " & Chr$(34) & item & Chr$(34) & "."
                Exit Function
            End If
        End If
        out.Add f
    Next it
End Function

' Longest first where one is inside another's words.
Private Function LabelOps() As Variant
    LabelOps = Array("does not contain", "contains", "does not begin with", "begins with", _
                     "does not end with", "ends with")
End Function

' "Top 25 by Exposure; Top 10 Counterparty by Exposure; Top 5% by Share;
'  Exposure > 1m; Counterparty: Exposure between 1m and 5m"
Public Function ParseVFilters(ByVal s As String, ByRef prob As String) As Collection
    Dim out As Collection, it As Variant, item As String, low As String, p As Long, f As Object
    Dim rest As String, num As String, ok As Boolean, op As Variant, q As Long, rhs As String
    Set out = New Collection
    Set ParseVFilters = out
    For Each it In SplitList(s, ";")
        item = Trim$(CStr(it))
        Set f = NewMap()
        f("Field") = ""
        f("Pct") = False
        f("N") = 0#
        f("V1") = 0#
        f("V2") = 0#
        f("By") = ""
        p = InStr(item, ":")
        If p > 0 Then
            f("Field") = Trim$(Left$(item, p - 1))
            item = Trim$(Mid$(item, p + 1))
        End If
        low = LCase$(item)
        If Left$(low, 4) = "top " Or Left$(low, 7) = "bottom " Then
            If Left$(low, 4) = "top " Then
                f("Kind") = "top"
            Else
                f("Kind") = "bottom"
            End If
            rest = Trim$(Mid$(item, InStr(item, " ") + 1))
            p = InStr(1, LCase$(rest), " by ")
            If p = 0 Then
                Fail prob, "A top / bottom filter says by which value - e.g. Top 25 by Exposure."
                Exit Function
            End If
            f("By") = Trim$(Mid$(rest, p + 4))
            rest = Trim$(Left$(rest, p - 1))
            q = InStr(rest, " ")
            If q > 0 Then
                num = Left$(rest, q - 1)
                If Len(f("Field")) = 0 Then f("Field") = Trim$(Mid$(rest, q + 1))
            Else
                num = rest
            End If
            If Right$(num, 1) = "%" Then
                f("Pct") = True
                num = Left$(num, Len(num) - 1)
            End If
            f("N") = NumOf(num, ok)
            If Not ok Then
                Fail prob, "A top / bottom filter needs how many - e.g. Top 25 by Exposure."
                Exit Function
            End If
            If CDbl(f("N")) <= 0 Or (CBool(f("Pct")) And CDbl(f("N")) > 100) Then
                Fail prob, "Top / value filter: " & Chr$(34) & CStr(it) & Chr$(34) & " asks for an impossible count."
                Exit Function
            End If
        Else
            f("Kind") = ""
            For Each op In Array(" not between ", " between ", ">=", "<=", "<>", ">", "<", "=")
                q = InStr(1, low, CStr(op))
                If q > 0 Then
                    f("Kind") = CompareKind(Trim$(CStr(op)))
                    f("By") = Trim$(Left$(item, q - 1))
                    rhs = Trim$(Mid$(item, q + Len(CStr(op))))
                    Exit For
                End If
            Next op
            If Len(f("Kind")) = 0 Or Len(f("By")) = 0 Then
                Fail prob, "A value filter is Top / Bottom N by a caption, or a caption compared with a number " & _
                           "- e.g. Exposure > 1m."
                Exit Function
            End If
            If f("Kind") = "between" Or f("Kind") = "notbetween" Then
                p = InStr(1, LCase$(rhs), " and ")
                If p = 0 Then
                    Fail prob, "Between takes two numbers joined by and - e.g. Exposure between 1m and 5m."
                    Exit Function
                End If
                f("V1") = NumOf(Left$(rhs, p - 1), ok)
                If ok Then f("V2") = NumOf(Mid$(rhs, p + 5), ok)
            Else
                f("V1") = NumOf(rhs, ok)
            End If
            If Not ok Then
                Fail prob, "A value filter compares with a number - 1500000, 1.5m, 2bn or 750k: " & _
                           Chr$(34) & CStr(it) & Chr$(34) & "."
                Exit Function
            End If
        End If
        out.Add f
    Next it
End Function

Private Function CompareKind(ByVal op As String) As String
    Select Case LCase$(op)
        Case ">": CompareKind = "gt"
        Case ">=": CompareKind = "ge"
        Case "<": CompareKind = "lt"
        Case "<=": CompareKind = "le"
        Case "=": CompareKind = "eq"
        Case "<>": CompareKind = "ne"
        Case "between": CompareKind = "between"
        Case "not between": CompareKind = "notbetween"
    End Select
End Function

' "1,500,000", "1.5m", "2bn", "750k", "5%" - the number, and whether it was one.
Public Function NumOf(ByVal s As String, ByRef ok As Boolean) As Double
    Dim mult As Double, i As Long
    s = LCase$(Replace(Replace(Trim$(s), ",", ""), " ", ""))
    mult = 1
    If Right$(s, 1) = "%" Then
        mult = 0.01
        s = Left$(s, Len(s) - 1)
    ElseIf Right$(s, 2) = "bn" Then
        mult = 1000000000#
        s = Left$(s, Len(s) - 2)
    ElseIf Right$(s, 1) = "m" Then
        mult = 1000000#
        s = Left$(s, Len(s) - 1)
    ElseIf Right$(s, 1) = "k" Then
        mult = 1000#
        s = Left$(s, Len(s) - 1)
    End If
    ' Digits, one point and a leading minus only: IsNumeric alone would take
    ' "(5)", "$5" and "&H10" as well.
    ok = (Len(s) > 0)
    For i = 1 To Len(s)
        Select Case Mid$(s, i, 1)
            Case "0" To "9", "."
            Case "-"
                If i > 1 Then ok = False
            Case Else
                ok = False
        End Select
    Next i
    If ok Then ok = IsNumeric(s)
    If ok Then NumOf = Val(s) * mult
End Function

' "Maturity date by year; Interest rate by 0.5"
'
' Grouped while staging, into a column of its own - "Maturity date by year" -
' rather than by Excel's grouping, which groups the cache every sheet shares
' and gives up on a column with one blank in it.
Public Function ParseGroups(ByVal s As String, ByRef prob As String) As Collection
    Dim out As Collection, it As Variant, p As Long, g As Object, per As String, ok As Boolean
    Set out = New Collection
    Set ParseGroups = out
    For Each it In SplitList(s, ";")
        p = InStrRev(LCase$(CStr(it)), " by ")
        If p = 0 Then
            Fail prob, "Group is a field, by, and a period or a step - e.g. Maturity date by year."
            Exit Function
        End If
        Set g = NewMap()
        g("Field") = Trim$(Left$(CStr(it), p - 1))
        per = Trim$(Mid$(CStr(it), p + 4))
        g("Size") = 0#
        Select Case LCase$(per)
            Case "year", "years": g("By") = "year"
            Case "quarter", "quarters": g("By") = "quarter"
            Case "month", "months": g("By") = "month"
            Case "day", "days", "date": g("By") = "day"
            Case Else
                g("By") = "step"
                g("Size") = NumOf(per, ok)
                If Not ok Then
                    Fail prob, "Group: " & Chr$(34) & per & Chr$(34) & " is not a period (year, quarter, month, " & _
                               "day) or a step (a number)."
                    Exit Function
                End If
                If CDbl(g("Size")) <= 0 Then
                    Fail prob, "Group: the step must be more than zero."
                    Exit Function
                End If
        End Select
        If g("By") = "step" Then
            g("Name") = g("Field") & " by " & per
        Else
            g("Name") = g("Field") & " by " & g("By")
        End If
        out.Add g
    Next it
End Function

' The field a recipe's pivot uses for nm: its grouped column when the recipe
' groups it, else nm itself.
Public Function PivotFieldName(ByVal rc As Object, ByVal nm As String) As String
    PivotFieldName = nm
    If rc("GroupOf").Exists(nm) Then PivotFieldName = CStr(rc("GroupOf")(nm))
End Function

' A value caption as Excel holds it - with the trailing spaces that keep it
' apart from a field name.
Public Function HeldCaption(ByVal rc As Object, ByVal cap As String) As String
    Dim v As Object
    HeldCaption = cap
    For Each v In rc("Values")
        If StrComp(Trim$(CStr(v("Caption"))), Trim$(cap), vbTextCompare) = 0 Then
            HeldCaption = CStr(v("Caption"))
            Exit Function
        End If
    Next v
End Function

Public Sub ParseSort(ByVal s As String, ByVal rc As Object)
    Dim p As Long, dirWord As String
    rc("SortBy") = ""
    rc("SortDesc") = True
    s = Trim$(s)
    If Len(s) = 0 Or LCase$(s) = "none" Then Exit Sub
    p = InStrRev(s, " ")
    If p > 0 Then
        dirWord = LCase$(Mid$(s, p + 1))
        If dirWord = "asc" Or dirWord = "desc" Then
            rc("SortDesc") = (dirWord = "desc")
            s = Trim$(Left$(s, p - 1))
        End If
    End If
    If LCase$(s) = "label" Then
        rc("SortBy") = "label"
        If dirWord <> "desc" Then rc("SortDesc") = False
    Else
        rc("SortBy") = s
    End If
End Sub

Private Function ParseWidths(ByVal s As String, ByVal rc As Object, ByRef prob As String) As Object
    Dim d As Object, it As Variant, p As Long, nm As String, w As String
    Set d = NewMap()
    Set ParseWidths = d
    rc("ValueWidth") = 14
    For Each it In SplitList(s, ";")
        p = InStr(CStr(it), "=")
        If p = 0 Then
            If Len(prob) = 0 Then prob = "Widths are field=width, separated by ; - " & Chr$(34) & CStr(it) & Chr$(34) & "."
        Else
            nm = Trim$(Left$(CStr(it), p - 1))
            w = Trim$(Mid$(CStr(it), p + 1))
            If Not IsNumeric(w) Then
                If Len(prob) = 0 Then prob = "A width must be a number: " & Chr$(34) & CStr(it) & Chr$(34) & "."
            ElseIf LCase$(nm) = "values" Then
                rc("ValueWidth") = CDbl(Val(w))
            Else
                d(nm) = CDbl(Val(w))
            End If
        End If
    Next it
End Function

' Everything that would stop Excel building the pivot, or make it build the
' wrong one, said in words that point at the cell to change.
Public Function Validate(ByVal rc As Object, ByVal fl As Object) As String
    Dim place As Object, x As Variant, v As Object, t As Variant, caps As Object, f As Object, g As Object
    Dim k As String, seen As Object, nPlain As Long
    If Len(rc("Name")) = 0 Then Validate = "Give the pivot a name - it becomes the sheet's name.": Exit Function
    ' measured with the longest framework label, which {fw} can become
    ' A chart on Start here is not a sheet, and its title can be as long as it likes.
    If Not rc.Exists("Chart") Then
        If Len(Replace(Replace(rc("Name"), "{fw}", "Maturity Ladder"), "{split}", "")) > 31 Then
            Validate = "The pivot name is longer than a sheet name can be (31 characters).": Exit Function
        End If
    End If
    For Each t In SplitList(CStr(rc("Frameworks")), ",")
        If Left$(NormFw(CStr(t)), 1) = "?" Then
            Validate = "Frameworks: " & Qt(t) & " is not All, LCR, NSFR or Maturity ladder."
            Exit Function
        End If
    Next t
    If rc("Values").count = 0 Then Validate = "Add at least one value - a pivot with none is empty.": Exit Function
    If rc("Rows").count = 0 And rc("Cols").count = 0 Then
        Validate = "Put at least one field in Rows or Columns.": Exit Function
    End If
    If rc("Split").count > 2 Then Validate = "One sheet per takes one or two fields.": Exit Function
    If rc("Split").count > 0 And InStr(1, rc("Name"), "{split}", vbTextCompare) = 0 Then
        Validate = "With One sheet per, name the pivot {split} so each sheet can take its value.": Exit Function
    End If

    ' every field named, known
    For Each x In FieldNamesOf(rc)
        If Not fl.Exists(CStr(x)) Then
            Validate = Qt(x) & " is not on the Pivot fields sheet - check the spelling, or add it there."
            Exit Function
        End If
    Next x
    For Each x In rc("Split")
        If fl(CStr(x))("Kind") <> "Text" Then
            Validate = "One sheet per takes text fields - " & Qt(x) & " is " & LCase$(fl(CStr(x))("Kind")) & "."
            Exit Function
        End If
    Next x
    For Each x In rc("Widths").keys
        If Not fl.Exists(CStr(x)) Then
            Validate = "Widths: " & Qt(x) & " is not a field.": Exit Function
        End If
    Next x
    For Each x In rc("SubFields").keys
        If Not InCollection(rc("Rows"), CStr(x)) Then
            Validate = "Subtotals: " & Qt(x) & " is not one of the Rows.": Exit Function
        End If
    Next x
    ' A calculated field is worked out from the sums - it can be a value, and
    ' nothing else.
    For Each x In PlacedNames(rc)
        If fl(CStr(x))("Kind") = "Calculated" Then
            Validate = Qt(x) & " is calculated from other values, so it can only go in Values.": Exit Function
        End If
    Next x

    ' groups: a date by a period, a number by a step, on an axis
    Set seen = NewMap()
    For Each g In rc("Groups")
        k = fl(CStr(g("Field")))("Kind")
        If Not OnAxis(rc, CStr(g("Field"))) Then
            Validate = "Group: " & Qt(g("Field")) & " is not in Rows or Columns.": Exit Function
        End If
        If seen.Exists(CStr(g("Field"))) Then
            Validate = "Group: " & Qt(g("Field")) & " is grouped twice.": Exit Function
        End If
        seen(CStr(g("Field"))) = True
        If g("By") = "step" Then
            If k <> "Number" Then
                Validate = "Group: " & Qt(g("Field")) & " is " & LCase$(k) & " - a step groups numbers; a date " & _
                           "groups by year, quarter, month or day."
                Exit Function
            End If
        ElseIf k <> "Date" Then
            Validate = "Group: " & Qt(g("Field")) & " is " & LCase$(k) & " - only a date groups by " & g("By") & "."
            Exit Function
        End If
    Next g

    ' a field can sit in one place only
    Set place = NewMap()
    For Each x In rc("Split")
        If Not Claim(place, CStr(x), "One sheet per") Then Validate = Twice(place, CStr(x), "One sheet per"): Exit Function
    Next x
    For Each x In rc("Rows")
        If Not Claim(place, CStr(x), "Rows") Then Validate = Twice(place, CStr(x), "Rows"): Exit Function
    Next x
    For Each x In rc("Cols")
        If Not Claim(place, CStr(x), "Columns") Then Validate = Twice(place, CStr(x), "Columns"): Exit Function
    Next x
    For Each f In rc("Filters")
        If place.Exists(f("Field")) Then
            If place(f("Field")) = "One sheet per" Then
                Validate = Qt(f("Field")) & " is already one sheet per value; it cannot be filtered as well."
                Exit Function
            End If
        End If
        If f("Kind") = "label" And Not OnAxis(rc, CStr(f("Field"))) Then
            Validate = "Show only / hide: " & Qt(f("Field")) & " " & f("Op") & " works on a field in Rows or Columns."
            Exit Function
        End If
    Next f

    ' values: numbers can be summed, dates and text counted
    Set caps = NewMap()
    For Each v In rc("Values")
        caps(Trim$(v("Caption"))) = True
        Select Case fl(CStr(v("Field")))("Kind")
            Case "Text"
                If v("Agg") <> "count" Then
                    Validate = Qt(v("Field")) & " is text, so it can only be counted - write " & v("Field") & " count."
                    Exit Function
                End If
            Case "Date"
                If v("Agg") <> "count" And v("Agg") <> "min" And v("Agg") <> "max" Then
                    Validate = Qt(v("Field")) & " is a date - count it, or take its min or max."
                    Exit Function
                End If
            Case "Calculated"
                If v("Agg") <> "sum" Then
                    Validate = Qt(v("Field")) & " is calculated from sums, so it can only be summed."
                    Exit Function
                End If
        End Select
        If Len(v("Calc")) = 0 Then nPlain = nPlain + 1
        If Len(v("Base")) = 0 Then v("Base") = DefaultAlong(rc, CStr(v("Calc")))
        If Len(v("Base")) > 0 Then
            If Not OnAxis(rc, CStr(v("Base"))) Then
                Validate = Qt(Trim$(v("Caption"))) & " runs along " & Qt(v("Base")) & ", which is not in Rows or " & _
                           "Columns."
                Exit Function
            End If
        End If
    Next v
    If rc("SortBy") <> "" And rc("SortBy") <> "label" Then
        If Not caps.Exists(rc("SortBy")) Then
            Validate = "Sort: " & Qt(rc("SortBy")) & " is not one of the value captions, or label."
            Exit Function
        End If
    End If

    ' top / value filters: on an axis field, by a value, one per field
    Set seen = NewMap()
    For Each f In rc("VFilters")
        If Len(f("Field")) = 0 Then f("Field") = FirstAxisField(rc)
        If Not OnAxis(rc, CStr(f("Field"))) Then
            Validate = "Top / value filter: " & Qt(f("Field")) & " is not in Rows or Columns.": Exit Function
        End If
        If Not caps.Exists(Trim$(f("By"))) Then
            Validate = "Top / value filter: " & Qt(f("By")) & " is not one of the value captions.": Exit Function
        End If
        If seen.Exists(CStr(f("Field"))) Then
            Validate = "Top / value filter: one per field, and " & Qt(f("Field")) & " has two.": Exit Function
        End If
        seen(CStr(f("Field"))) = True
    Next f

    If Len(rc("ExpandTo")) > 0 Then
        If Not InCollection(rc("Rows"), CStr(rc("ExpandTo"))) Then
            Validate = "Expand to: " & Qt(rc("ExpandTo")) & " is not one of the Rows.": Exit Function
        End If
        If StrComp(CStr(rc("Rows")(rc("Rows").count)), CStr(rc("ExpandTo")), vbTextCompare) = 0 Then
            Validate = "Expand to: " & Qt(rc("ExpandTo")) & " is the last of the Rows - there is nothing under " & _
                       "it to fold."
            Exit Function
        End If
    End If
    If Len(rc("HiliteOn")) > 0 Then
        If Not caps.Exists(rc("HiliteOn")) Then
            Validate = "Highlight: " & Qt(rc("HiliteOn")) & " is not one of the value captions.": Exit Function
        End If
    ElseIf Len(rc("Hilite")) > 0 And nPlain = 0 Then
        Validate = "Highlight marks the plain values, and every value here is a calculation - add on and a " & _
                   "caption to pick one."
        Exit Function
    End If
    Select Case LCase$(CStr(rc("Tab")))
        Case "", "auto", "emerald", "deep", "slate", "black"
        Case Else: Validate = "Tab must be Auto, Emerald, Deep, Slate or Black.": Exit Function
    End Select
End Function

Private Function Qt(ByVal s As Variant) As String
    Qt = Chr$(34) & CStr(s) & Chr$(34)
End Function

Private Function OnAxis(ByVal rc As Object, ByVal nm As String) As Boolean
    OnAxis = InCollection(rc("Rows"), nm) Or InCollection(rc("Cols"), nm)
End Function

Private Function FirstAxisField(ByVal rc As Object) As String
    If rc("Rows").count > 0 Then
        FirstAxisField = CStr(rc("Rows")(1))
    ElseIf rc("Cols").count > 0 Then
        FirstAxisField = CStr(rc("Cols")(1))
    End If
End Function

' What a running total, a rank or a share runs along when the value does not
' say: a running total or a rank down the innermost row, a share of the
' outermost, a change across the columns.
Private Function DefaultAlong(ByVal rc As Object, ByVal calc As String) As String
    Dim rw As Collection, cl As Collection
    Set rw = rc("Rows")
    Set cl = rc("Cols")
    Select Case calc
        Case "running", "%running", "rank"
            If rw.count > 0 Then
                DefaultAlong = CStr(rw(rw.count))
            Else
                DefaultAlong = CStr(cl(cl.count))
            End If
        Case "%parent"
            DefaultAlong = FirstAxisField(rc)
        Case "diff", "%diff"
            If cl.count > 0 Then
                DefaultAlong = CStr(cl(1))
            Else
                DefaultAlong = CStr(rw(rw.count))
            End If
    End Select
End Function

' Every field a recipe puts somewhere other than Values.
Private Function PlacedNames(ByVal rc As Object) As Collection
    Dim c As Collection, x As Variant, f As Object
    Set c = New Collection
    For Each x In rc("Split"): c.Add x: Next x
    For Each x In rc("Rows"): c.Add x: Next x
    For Each x In rc("Cols"): c.Add x: Next x
    For Each x In rc("Slicers"): c.Add x: Next x
    For Each f In rc("Filters"): c.Add f("Field"): Next f
    For Each f In rc("Groups"): c.Add f("Field"): Next f
    For Each f In rc("VFilters")
        If Len(f("Field")) > 0 Then c.Add f("Field")
    Next f
    Set PlacedNames = c
End Function

Private Function Claim(ByVal place As Object, ByVal nm As String, ByVal where As String) As Boolean
    If place.Exists(nm) Then Exit Function
    place(nm) = where
    Claim = True
End Function

Private Function Twice(ByVal place As Object, ByVal nm As String, ByVal where As String) As String
    Twice = Chr$(34) & nm & Chr$(34) & " is in " & place(nm) & " and in " & where & " - a field goes in one place."
End Function

Public Function InCollection(ByVal c As Collection, ByVal s As String) As Boolean
    Dim x As Variant
    For Each x In c
        If StrComp(CStr(x), s, vbTextCompare) = 0 Then InCollection = True: Exit Function
    Next x
End Function

' ===================== what staging must add ================================

' Fields named by the recipes that 1.0 did not stage - what the stager has to
' add for this build: the extra columns, the columns a calculated field reads,
' and a column for every group.
'
' alsoField: one more the build needs staged - the field a framework is split
' into workbooks by.
Public Function ExtraFields(ByVal recipes As Collection, Optional ByVal alsoField As String = "") As Collection
    Dim out As Collection, seen As Object, fl As Object, rc As Object, nm As Variant, g As Object
    Set out = New Collection
    Set seen = NewMap()
    Set fl = modPD_Config.Fields()
    If Len(alsoField) > 0 Then NeedField out, seen, fl, alsoField
    For Each rc In recipes
        For Each nm In FieldNamesOf(rc)
            NeedField out, seen, fl, CStr(nm)
        Next nm
        For Each g In rc("Groups")
            If fl.Exists(CStr(g("Field"))) And Not seen.Exists(CStr(g("Name"))) Then
                seen(CStr(g("Name"))) = True
                out.Add GroupField(g, fl(CStr(g("Field"))))
            End If
        Next g
    Next rc
    Set ExtraFields = out
End Function

' One field the stager must add - or, for a calculated one, the fields its
' formula reads, since the pivot works it out from them.
Private Sub NeedField(ByVal out As Collection, ByVal seen As Object, ByVal fl As Object, ByVal nm As String)
    Dim f As Object, other As Variant
    If Not fl.Exists(nm) Then Exit Sub
    If seen.Exists(nm) Then Exit Sub
    seen(nm) = True
    Set f = fl(nm)
    If CBool(f("Builtin")) Then Exit Sub
    If f("Kind") = "Calculated" Then
        For Each other In fl.keys
            If StrComp(CStr(other), nm, vbTextCompare) <> 0 Then
                If InStr(1, CStr(f("Source")), CStr(other), vbTextCompare) > 0 Then NeedField out, seen, fl, CStr(other)
            End If
        Next other
        Exit Sub
    End If
    out.Add f
End Sub

' The column a group is staged into: its field's source, cut to the period.
Private Function GroupField(ByVal g As Object, ByVal fld As Object) As Object
    Dim f As Object
    Set f = NewMap()
    f("Name") = g("Name")
    f("Source") = fld("Source")
    f("Kind") = "Group"
    f("BaseKind") = fld("Kind")
    f("By") = g("By")
    f("Size") = g("Size")
    f("Blank") = fld("Blank")
    If Len(f("Blank")) = 0 Then f("Blank") = "(none)"
    f("Format") = fld("Format")
    f("Width") = fld("Width")
    f("Builtin") = False
    f("Labels") = 0
    Set GroupField = f
End Function

Private Function FieldNamesOf(ByVal rc As Object) As Collection
    Dim c As Collection, x As Variant, v As Object
    Set c = PlacedNames(rc)
    For Each v In rc("Values")
        c.Add v("Field")
        If Len(v("Base")) > 0 Then c.Add v("Base")
    Next v
    Set FieldNamesOf = c
End Function
