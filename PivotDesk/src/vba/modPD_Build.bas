Option Explicit

' When one framework is split into several workbooks, the value this one is for.
Private mPart As String

' ============================================================================
'  Building one framework's workbook.
'
'  One file per framework, written into the output folder. Not one file with
'  everything in it: three staged extracts is a million and a quarter rows and a
'  workbook nobody can open, and the ladder was asked for as its own thing
'  anyway.
'
'  Inside each: a guide, the rule-level Output, the Balance sheet, and then one
'  sheet per rule - times the currency split for LCR and NSFR, or one sheet per
'  currency for the ladder.
' ============================================================================

Private Const GUIDE_COLS As Long = 5

Public Function BuildFramework(ByVal fw As String, ByVal outFolder As String, ByRef errOut As String) As String
    Dim wb As Workbook, lo As ListObject, nSheets As Long, capped As Boolean
    Dim Path As String, t0 As Single, recipes As Collection, fromConfig As Boolean

    t0 = Timer
    fromConfig = (modPD_Config.Engine() = "recipes")
    If fromConfig Then
        Set recipes = modPD_Config.RecipesFor(fw)
        If recipes.count = 0 Then
            errOut = "no pivot on the Pivot config sheet is switched on for " & FwLabel(fw)
            Exit Function
        End If
    End If

    Step_ FwLabel(fw) & " - opening a new workbook"
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Brand wb, fw

    modPD_Pivot.ResetPivots
    Step_ FwLabel(fw) & " - staging the output"
    If fromConfig Then
        Set lo = modPD_Stage.StageFramework(fw, wb, errOut, modPD_Config.ExtraFields(recipes), Signatures(recipes))
    Else
        Set lo = modPD_Stage.StageFramework(fw, wb, errOut)
    End If
    If lo Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        Err.Clear
        Exit Function
    End If

    modPD_Pivot.UseCache wb, lo
    modPD_Pivot.SetBook BookCrumb(fw)

    If fromConfig Then
        nSheets = RecipeSheets(wb, fw, recipes, capped)
    Else
        Step_ FwLabel(fw) & " - the Output pivot"
        modPD_Pivot.BuildOutputSheet wb, fw
        Step_ FwLabel(fw) & " - the Balance sheet pivot"
        modPD_Pivot.BuildBalanceSheet wb, fw
        If SplitsByCurrency(fw) Then
            nSheets = CurrencySheets(wb, capped)
        Else
            nSheets = RuleSheets(wb, capped)
        End If
        nSheets = nSheets + 2
    End If

    Guide wb, fw, nSheets, capped, fromConfig
    modPD_Pivot.LinkSiblings wb
    Tidy wb

    Path = PathJoin(outFolder, SafeFileName(FwLabel(fw)) & ".xlsx")
    Step_ FwLabel(fw) & " - saving"
    On Error GoTo SaveFailed
    Application.DisplayAlerts = False
    wb.SaveAs Path, 51            ' xlOpenXMLWorkbook - no macros, opens anywhere
    modPD_Desk.NoteWorkbook Path, fw, nSheets
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    On Error GoTo 0

    LogIt V_OK, "Pivots", Fmt(modPD_Stage.StagedRows()) & " row(s) staged, " & _
          nSheets & " pivot sheet(s), " & Format$(Timer - t0, "0.0") & "s" & _
          IIf(fromConfig, ", from Pivot config.", ", 1.0 layout."), FwLabel(fw)
    BuildFramework = Path
    Exit Function

SaveFailed:
    errOut = "could not save: " & Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    Err.Clear
End Function

' ===================== from the Pivot config ===============================

' Every "one sheet per" family the recipes ask for, as the staging pass needs
' them: the field names joined, once each.
Private Function Signatures(ByVal recipes As Collection) As Collection
    Dim out As Collection, seen As Object, rc As Object, sig As String
    Set out = New Collection
    Set seen = NewMap()
    For Each rc In recipes
        If rc("Split").count > 0 Then
            sig = SigOf(rc)
            If Not seen.Exists(sig) Then seen(sig) = True: out.Add sig
        End If
    Next rc
    Set Signatures = out
End Function

Private Function SigOf(ByVal rc As Object) As String
    Dim x As Variant, s As String
    For Each x In rc("Split")
        If Len(s) > 0 Then s = s & Chr$(30)
        s = s & CStr(x)
    Next x
    SigOf = s
End Function

' The recipes, in the order the sheet lists them. A single pivot is one
' sheet; a "one sheet per" pivot is a family, biggest first, capped.
Private Function RecipeSheets(ByVal wb As Workbook, ByVal fw As String, ByVal recipes As Collection, _
                              ByRef capped As Boolean) As Long
    Dim rc As Object, n As Long, combos As Variant, i As Long, vals As Variant, nm As String
    Dim made As Long, limit As Long, fl As Object, nRc As Long, k As Long

    Set fl = modPD_Config.Fields()
    For Each rc In recipes
        nRc = nRc + 1
        If rc("Split").count = 0 Then
            nm = Replace(CStr(rc("Name")), "{fw}", FwLabel(fw))
            Progress_ FwLabel(fw) & " - " & nm, nRc / (recipes.count + 1)
            If TryRecipe(wb, rc, fw, Empty, nm) Then n = n + 1
        Else
            combos = BySize(modPD_Stage.SplitWeightsFor(SigOf(rc)))
            limit = CLng(rc("Max"))
            made = 0
            If IsArray(combos) Then
                For i = 0 To UBound(combos)
                    vals = Split(CStr(combos(i)), Chr$(30))
                    If Not HasBlank(vals, rc("Split"), fl) Then
                        If made >= limit Or n >= MAX_BOOK_SHEETS Then capped = True: Exit For
                        nm = SplitTabName(vals)
                        Progress_ FwLabel(fw) & " - sheet " & (n + 1) & ": " & nm, _
                                  (nRc - 1 + (i + 1) / (UBound(combos) + 1)) / (recipes.count + 1)
                        If TryRecipe(wb, rc, fw, vals, nm) Then
                            n = n + 1
                            made = made + 1
                        End If
                    End If
                Next i
            End If
        End If
    Next rc
    RecipeSheets = n
End Function

' One recipe sheet, fenced: a recipe Excel refuses is logged and skipped,
' and the rest of the workbook is still built.
Private Function TryRecipe(ByVal wb As Workbook, ByVal rc As Object, ByVal fw As String, _
                           ByVal vals As Variant, ByVal nm As String) As Boolean
    On Error GoTo Failed
    TryRecipe = Not (modPD_Pivot.BuildRecipeSheet(wb, rc, fw, vals, nm) Is Nothing)
    Exit Function
Failed:
    LogIt V_BREAK, "Pivots", "Pivot " & Chr$(34) & nm & Chr$(34) & " (Pivot config row " & rc("Row") & _
          ") was not built - " & Err.Description, FwLabel(fw)
    Err.Clear
End Function

' Rows with no value for a "one sheet per" field get no sheet of their own -
' as 1.0 never made a sheet for "(no rule)". They are in the overview pivots.
Private Function HasBlank(ByVal vals As Variant, ByVal sf As Collection, ByVal fl As Object) As Boolean
    Dim i As Long, blank As String
    For i = 0 To UBound(vals)
        If Len(CStr(vals(i))) = 0 Then HasBlank = True: Exit Function
        blank = ""
        If fl.Exists(CStr(sf(i + 1))) Then blank = CStr(fl(CStr(sf(i + 1)))("Blank"))
        If Len(blank) > 0 And StrComp(CStr(vals(i)), blank, vbTextCompare) = 0 Then HasBlank = True: Exit Function
    Next i
End Function

' Biggest first - by the first field's total, then within it by the
' combination - so a rule's LCY and FCY sheets sit side by side, as in 1.0.
Private Function BySize(ByVal d As Object) As Variant
    Dim keys() As String, w() As Double, g() As Double, n As Long, i As Long, j As Long, k As Variant
    Dim firstW As Object, head As String, td As Double, ts As String, tg As Double, swap As Boolean
    n = d.count
    If n = 0 Then BySize = Array(): Exit Function
    Set firstW = NewMap()
    For Each k In d.keys
        head = Split(CStr(k) & Chr$(30), Chr$(30))(0)
        firstW(head) = SafeNum(firstW(head)) + SafeNum(d(k))
    Next k
    ReDim keys(0 To n - 1)
    ReDim w(0 To n - 1)
    ReDim g(0 To n - 1)
    i = 0
    For Each k In d.keys
        keys(i) = CStr(k)
        w(i) = SafeNum(d(k))
        g(i) = SafeNum(firstW(Split(CStr(k) & Chr$(30), Chr$(30))(0)))
        i = i + 1
    Next k
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            swap = False
            If g(j) < g(j + 1) Then
                swap = True
            ElseIf g(j) = g(j + 1) Then
                If Split(keys(j) & Chr$(30), Chr$(30))(0) = Split(keys(j + 1) & Chr$(30), Chr$(30))(0) Then
                    If w(j) < w(j + 1) Then swap = True
                ElseIf Split(keys(j) & Chr$(30), Chr$(30))(0) > Split(keys(j + 1) & Chr$(30), Chr$(30))(0) Then
                    swap = True
                End If
            End If
            If swap Then
                td = w(j): w(j) = w(j + 1): w(j + 1) = td
                tg = g(j): g(j) = g(j + 1): g(j + 1) = tg
                ts = keys(j): keys(j) = keys(j + 1): keys(j + 1) = ts
            End If
        Next j
    Next i
    BySize = keys
End Function

' The first value cut to fit and the rest appended after it, so the values
' that tell two sheets apart (LCY, FCY) survive a long first one.
Private Function SplitTabName(ByVal vals As Variant) As String
    Dim tail As String, body As String, i As Long
    If UBound(vals) = 0 Then SplitTabName = SafeSheetName(CStr(vals(0))): Exit Function
    For i = 1 To UBound(vals)
        tail = tail & " " & CStr(vals(i))
    Next i
    body = Trim$(CStr(vals(0)))
    If Len(body) > 31 - Len(tail) Then body = Trim$(Left$(body, 31 - Len(tail)))
    SplitTabName = SafeSheetName(body & tail)
End Function

' ===================== one sheet per rule, per side =========================

Private Function RuleSheets(ByVal wb As Workbook, ByRef capped As Boolean) As Long
    Dim names As Variant, i As Long, n As Long, side As Variant, nm As String

    names = RulesBySize()
    If Not IsArray(names) Then Exit Function
    If UBound(names) < 0 Then Exit Function

    For i = 0 To UBound(names)
        For Each side In Array("LCY", "FCY")
            If n >= MAX_RULE_SHEETS Then capped = True: RuleSheets = n: Exit Function
            nm = tabName(CStr(names(i)), CStr(side))
            Step_ "sheet " & (n + 1) & ": " & nm
            If Not modPD_Pivot.BuildRuleSheet(wb, CStr(names(i)), H_CCYCLASS, CStr(side), nm) Is Nothing Then
                n = n + 1
            End If
        Next side
    Next i
    RuleSheets = n
End Function

' ===================== one sheet per currency (the ladder) ==================

Private Function CurrencySheets(ByVal wb As Workbook, ByRef capped As Boolean) As Long
    Dim ccys As Variant, i As Long, n As Long, nm As String
    ccys = SortedByWeight(modPD_Stage.Currencies())
    If Not IsArray(ccys) Then Exit Function
    If UBound(ccys) < 0 Then Exit Function
    For i = 0 To UBound(ccys)
        If n >= MAX_RULE_SHEETS Then capped = True: CurrencySheets = n: Exit Function
        nm = SafeSheetName(CStr(ccys(i)))
        Step_ "sheet " & (n + 1) & ": " & nm
        If Not modPD_Pivot.BuildCurrencySheet(wb, CStr(ccys(i)), nm) Is Nothing Then n = n + 1
    Next i
    CurrencySheets = n
End Function

' ===================== ordering and naming ==================================

' Biggest first. A workbook of forty sheets is only usable if the ones that
' carry the money are near the front.
Private Function RulesBySize() As Variant
    RulesBySize = SortedByWeight(modPD_Stage.RuleNames())
End Function

Private Function SortedByWeight(ByVal d As Object) As Variant
    Dim keys() As String, w() As Double, n As Long, i As Long, j As Long, k As Variant
    Dim td As Double, ts As String
    n = d.count
    If n = 0 Then SortedByWeight = Array(): Exit Function
    ReDim keys(0 To n - 1)
    ReDim w(0 To n - 1)
    i = 0
    For Each k In d.keys
        keys(i) = CStr(k): w(i) = SafeNum(d(k)): i = i + 1
    Next k
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If w(j) < w(j + 1) Then
                td = w(j): w(j) = w(j + 1): w(j + 1) = td
                ts = keys(j): keys(j) = keys(j + 1): keys(j + 1) = ts
            End If
        Next j
    Next i
    SortedByWeight = keys
End Function

' A rule name is long and a tab is thirty-one characters. The side has to
' survive the truncation, because the whole point of two sheets is telling them
' apart, so it is appended AFTER the name is cut to fit.
Private Function tabName(ByVal ruleName As String, ByVal side As String) As String
    Dim body As String
    body = Trim$(ruleName)
    If Len(body) > 31 - Len(side) - 1 Then body = Trim$(Left$(body, 31 - Len(side) - 1))
    tabName = SafeSheetName(body & " " & side)
End Function

Private Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, o As String
    o = s
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each b In bad
        o = Replace(o, CStr(b), " ")
    Next b
    SafeFileName = Trim$(o)
End Function

' ===================== the guide ============================================

Private Sub Guide(ByVal wb As Workbook, ByVal fw As String, ByVal nSheets As Long, ByVal capped As Boolean, _
                  ByVal fromConfig As Boolean)
    Dim ws As Worksheet, made As Collection, i As Long, r As Long, e As Variant, first As Long, wide As Double

    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    On Error Resume Next
    ws.Name = SH_GUIDE
    Err.Clear
    On Error GoTo 0

    modPD_Theme.Dress ws, BookTitle(fw), _
        "Built by " & TOOL_NAME & " ALM Desk " & TOOL_VERSION & " on " & Format$(Now, "d mmm yyyy, hh:nn") & _
        IIf(Len(modPD_Stage.StagedAsOf()) > 0, "   " & ChrW(183) & "   data as of " & modPD_Stage.StagedAsOf(), ""), _
        "START HERE  " & ChrW(183) & "  " & UCase$(BANK_NAME)
    modPD_Theme.BookBar ws, BookCrumb(fw), False
    ws.Columns(1).ColumnWidth = 38
    ws.Columns(2).ColumnWidth = 110
    wide = ws.Cells(1, 3).Left - ws.Cells(1, 1).Left - 14
    ws.Tab.Color = modPD_Theme.C_INK

    ' --- at a glance: the figures wanted before any pivot ----------------------
    ws.Rows(4).RowHeight = 68
    Glance ws, ws.Rows(4).Top + 8, wide
    modPD_Theme.SetStatus ws, Fmt(modPD_Stage.StagedRows()) & " rows staged into one pivot cache.  " & _
        "Every sheet below is a live PivotTable over it - drag a field, drop a slicer, drill a total.", "OK"
    ws.Rows(6).RowHeight = 12

    ' --- the index -----------------------------------------------------------------
    r = 7
    Section ws, r, "IN THIS BOOK  " & ChrW(183) & "  " & modPD_Pivot.MadeSheets().count & " SHEETS"
    modPD_Theme.Head ws, Array("Sheet", "What is on it"), Array(38, 110), r + 1
    first = r + 2
    r = first
    Set made = modPD_Pivot.MadeSheets()
    For i = 1 To made.count
        e = made(i)
        ws.Cells(r, 1).Value2 = CStr(e(0))
        ws.Cells(r, 2).Value2 = CStr(e(1))
        On Error Resume Next
        ' An apostrophe in a sheet name ("Customer's deposits") is doubled in a
        ' link, or the link goes nowhere.
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="", _
                          SubAddress:="'" & Replace(CStr(e(0)), "'", "''") & "'!A1", ScreenTip:="Open " & CStr(e(0))
        Err.Clear
        On Error GoTo 0
        r = r + 1
    Next i
    modPD_Theme.DressTable ws, 2, r - 1, 0, first
    ' Links in the Desk's emerald, not the default blue underline - they are
    ' the index of the book and should read as part of it.
    If r - 1 >= first Then
        With ws.Range(ws.Cells(first, 1), ws.Cells(r - 1, 1)).Font
            .Name = modPD_Theme.UI_SEMI
            .Underline = xlUnderlineStyleNone
            .Color = modPD_Theme.C_LINK
        End With
        ws.Range(ws.Cells(first, 2), ws.Cells(r - 1, 2)).Font.Color = modPD_Theme.C_TEXT_2
    End If

    ' What the tool had to decide for itself goes here, where it is read, not
    ' into a log nobody opens.
    r = r + 1
    Section ws, r, "HOW TO READ THIS"
    r = r + 1
    Note ws, r, "Amounts", "Pre-factor and post-factor from " & modPD_Stage.AmountFieldNote() & _
        IIf(modPD_Stage.UsedNativeAmounts(), ".", _
        ".  THERE IS NO NATIVE-CURRENCY AMOUNT COLUMN IN THIS EXTRACT - these are the converted " & _
        "figures. Every spelling of a CCY/ACY column was tried and none was present.")
    r = r + 1
    Note ws, r, "Local currency", Chr$(34) & modPD_Stage.LocalCurrency() & Chr$(34) & _
        " - the currency on the most rows. Nothing in the extract says which is local, so " & _
        "check this before relying on the LCY / FCY split."
    r = r + 1
    Note ws, r, "Factor", "Post-factor divided by pre-factor, so it always agrees with the two " & _
        "figures beside it. The source's own factor column has no stated units."
    r = r + 1
    Note ws, r, "Labels", "Ledger names are shown without their codes and in title case where Pivot fields " & _
        "says so - 1.07.00.MBGL.1360.LOANS TO CUSTOMERS reads Loans to Customers."
    r = r + 1
    Note ws, r, "Blanks", "Rows with no bucket are excluded from the bucket filter by default. " & _
        "They are still in the data - clear the filter to see them."
    r = r + 1
    If fromConfig Then
        Note ws, r, "Built from", "Reports in " & ThisWorkbook.Name & " - every pivot here is a row on Pivot " & _
            "config, and changing the row changes the next build."
    Else
        Note ws, r, "Built from", "the 1.0 layout (Pivot config's Engine is set to it)."
    End If
    If Len(modPD_Stage.MissingFields()) > 0 Then
        r = r + 1
        Note ws, r, "Not in this file", "These fields' columns are not in this output, so they are blank here: " & _
            modPD_Stage.MissingFields() & "."
    End If
    If capped Then
        r = r + 1
        Note ws, r, "Sheet cap", "Stopped at the most sheets a pivot is allowed (Max sheets on Pivot config). " & _
            "The rest are not here; everything is still in the overview pivots."
    End If

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 100
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Cells(first, 1).Select
    PrintSetupGuide ws
    Err.Clear
End Sub

' The book as its bar and title name it: the framework, and the value it was
' split by when it is one of several workbooks.
Private Function BookTitle(ByVal fw As String) As String
    BookTitle = FwLabel(fw)
    If Len(mPart) > 0 Then BookTitle = BookTitle & "  " & ChrW(183) & "  " & mPart
End Function

Private Function BookCrumb(ByVal fw As String) As String
    BookCrumb = BookTitle(fw) & "  " & ChrW(183) & "  " & BANK_NAME
End Function

Private Sub PrintSetupGuide(ByVal ws As Worksheet)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .BlackAndWhite = True
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub

' A section of the Start here sheet: small capitals in emerald over a rule.
Private Sub Section(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String)
    On Error Resume Next
    With ws.Cells(r, 1)
        .Value2 = title
        .Font.Name = modPD_Theme.UI_SEMI
        .Font.Size = 8
        .Font.Color = modPD_Theme.C_LINK
        .IndentLevel = 1
        .VerticalAlignment = xlBottom
    End With
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = modPD_Theme.C_BRAND_DEEP
        .Weight = xlThin
    End With
    ws.Rows(r).RowHeight = 30
    Err.Clear
End Sub

' Six tiles across the top: how much was read, how much money before and
' after the factors, how many sheets, which currency and which date. Drawn at
' build time - the numbers are this build's, as the title says.
Private Sub Glance(ByVal ws As Worksheet, ByVal top As Double, ByVal wide As Double)
    Dim items As Variant, i As Long, n As Long, w As Double, gp As Double
    items = Array(Array("ROWS STAGED", Fmt(modPD_Stage.StagedRows())), _
                  Array("PRE-FACTOR", Compact(modPD_Stage.GrossPre())), _
                  Array("POST-FACTOR", Compact(modPD_Stage.GrossPost())), _
                  Array("SHEETS", CStr(modPD_Pivot.MadeSheets().count)), _
                  Array("LOCAL CURRENCY", modPD_Stage.LocalCurrency()), _
                  Array("DATA AS OF", modPD_Stage.StagedAsOf()))
    n = UBound(items) + 1
    gp = 10
    w = (wide - (n - 1) * gp) / n
    For i = 0 To n - 1
        modPD_Theme.Tile ws, "pdb_glance" & (i + 1), 14 + i * (w + gp), top, w, 52, _
                         CStr(items(i)(0)), IIf(Len(CStr(items(i)(1))) > 0, CStr(items(i)(1)), ChrW(8212))
    Next i
End Sub

Private Sub Note(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal body As String)
    With ws.Cells(r, 1)
        .Value2 = label
        .Font.Name = modPD_Theme.UI_SEMI
        .Font.Color = modPD_Theme.C_TEXT
        .IndentLevel = 1
        .VerticalAlignment = xlTop
    End With
    With ws.Cells(r, 2)
        .Value2 = body
        .Font.Color = modPD_Theme.C_TEXT_2
        .WrapText = True
        .VerticalAlignment = xlTop
        .IndentLevel = 1
    End With
    ws.Rows(r).AutoFit
    If ws.Rows(r).RowHeight < 22 Then ws.Rows(r).RowHeight = 22
End Sub

' ===================== the look of the output ===============================

' The framework workbook in the desk's colours. Black, emerald and white go
' in as the THEME, so every built-in pivot, slicer and chart style Excel
' offers inside it comes out in the bank's palette - not only the one this
' tool applies - and the document says what it is when someone finds it in a
' folder a month later.
Private Sub Brand(ByVal wb As Workbook, ByVal fw As String)
    On Error Resume Next
    modPD_Theme.DarkNormal wb
    With wb.Theme.ThemeColorScheme
        .Colors(1).RGB = modPD_Theme.HX("000000")      ' dark 1
        .Colors(2).RGB = modPD_Theme.HX("FFFFFF")      ' light 1
        .Colors(3).RGB = modPD_Theme.HX("0C1512")      ' dark 2
        .Colors(4).RGB = modPD_Theme.HX("F4F7F5")      ' light 2
        .Colors(5).RGB = modPD_Theme.HX("009060")      ' accent 1 - the brand emerald
        .Colors(6).RGB = modPD_Theme.HX("16B07F")
        .Colors(7).RGB = modPD_Theme.HX("4FC79C")
        .Colors(8).RGB = modPD_Theme.HX("006141")
        .Colors(9).RGB = modPD_Theme.HX("8FDBBE")
        .Colors(10).RGB = modPD_Theme.HX("5F7068")
        .Colors(11).RGB = modPD_Theme.HX("00794F")     ' hyperlink
        .Colors(12).RGB = modPD_Theme.HX("004A32")     ' followed hyperlink
    End With
    With wb.Theme.ThemeFontScheme
        .MajorFont(1).Name = modPD_Theme.UI_FONT
        .MinorFont(1).Name = modPD_Theme.UI_FONT
    End With
    With wb.Styles("Normal").Font
        .Name = modPD_Theme.UI_FONT
        .Size = 10
    End With
    wb.BuiltinDocumentProperties("Title").value = FwLabel(fw) & " pivots  -  " & BANK_NAME
    wb.BuiltinDocumentProperties("Subject").value = FwLabel(fw) & " output, staged into live PivotTables"
    wb.BuiltinDocumentProperties("Keywords").value = "ALM; " & FwLabel(fw) & "; Avati; " & BANK_NAME
    wb.BuiltinDocumentProperties("Comments").value = "Built by " & TOOL_NAME & " " & TOOL_VERSION & " on " & _
        Format$(Now, "dd mmm yyyy hh:nn") & "."
    Err.Clear
End Sub

' ===================== tidying up ===========================================

Private Sub Tidy(ByVal wb As Workbook)
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Application.DisplayAlerts = False
    ' The blank sheet Workbooks.Add leaves behind.
    For i = wb.Worksheets.count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If ws.visible = xlSheetVisible Then
            If ws.UsedRange.Address = "$A$1" And IsEmpty(ws.Range("A1").Value2) _
               And ws.PivotTables.count = 0 Then ws.Delete
        End If
    Next i
    Application.DisplayAlerts = True
    wb.Worksheets(1).Activate
    Err.Clear
End Sub
