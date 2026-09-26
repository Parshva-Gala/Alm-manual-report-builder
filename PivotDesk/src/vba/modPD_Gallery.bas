Option Explicit

' ============================================================================
'  Gallery: reports worth having, one click from being in the next build.
'
'  Each card is a finished Pivot config or Chart config row - a top-N league,
'  a heatmap, a folded summary, a haircut table, a label-filtered list - that
'  shows what the configuration can do. Add writes the row, switched on, and
'  opens it; a card already on the sheet says so and offers to open it
'  instead. Nothing here is a report of its own: every card becomes an
'  ordinary row, changed like any other.
' ============================================================================

Private Const CARD_W As Double = 300
Private Const CARD_H As Double = 164
Private Const GAP As Double = 14
Private Const PER_ROW As Long = 3
Private Const LEFT0 As Double = 14
Private Const FIRST_ROW As Long = 7

' key, kind ("pivot" / "chart"), title, what it does, what it uses, the row
Private Function Templates() As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Array("topcp", "pivot", "Top counterparties", _
        "The 25 largest counterparties by gross exposure, each one's share, and the running share down the list.", _
        "Top 25 filter  " & ChrW(183) & "  % of total  " & ChrW(183) & "  % running  " & ChrW(183) & "  data bars", _
        Array("On=Yes", "Pivot={fw} top counterparties", "Frameworks=All", "Rows=Counterparty", _
              "Values=Gross pre-factor sum as Exposure; Gross pre-factor %total as Share; " & _
              "Gross pre-factor %running in Counterparty as Cumulative", _
              "Show only / hide=Counterparty <> (no counterparty)", "Top / value filter=Top 25 by Exposure", _
              "Slicers=LCY / FCY", "Layout=Tabular", "Grand totals=Bottom row", "Sort=Exposure desc", _
              "Units=Millions", "Highlight=Data bars on Exposure", "Tiles=Yes", _
              "Tab=Emerald", "Description=The 25 largest counterparties by gross exposure - each one's share of " & _
              "the 25 and the running share down the list. Slice by LCY / FCY."))
    c.Add Array("heatmap", "pivot", "Counterparty heatmap", _
        "Where the 30 largest counterparties fall across the maturity buckets, coloured by size.", _
        "Top 30  " & ChrW(183) & "  buckets across  " & ChrW(183) & "  heatmap", _
        Array("On=Yes", "Pivot={fw} counterparty heatmap", "Frameworks=All", "Rows=Counterparty", "Columns=Bucket", _
              "Values=Gross pre-factor sum as Exposure", _
              "Show only / hide=Bucket <> (no bucket); Counterparty <> (no counterparty)", _
              "Top / value filter=Top 30 by Exposure", "Layout=Tabular", "Grand totals=Both", "Sort=Exposure desc", _
              "Units=Millions", "Highlight=Heatmap", "Tab=Emerald", _
              "Description=Where the 30 largest counterparties fall across the maturity buckets."))
    c.Add Array("prodconc", "pivot", "Product concentration", _
        "Each product's ten largest counterparties, and the share of the product each one holds.", _
        "Top 10 within each product  " & ChrW(183) & "  % of parent", _
        Array("On=Yes", "Pivot={fw} product concentration", "Frameworks=All", "Rows=Product, Counterparty", _
              "Values=Gross pre-factor sum as Exposure; Gross pre-factor %parent as Share of product", _
              "Show only / hide=Counterparty <> (no counterparty)", "Top / value filter=Top 10 Counterparty by Exposure", "Layout=Tabular", "Subtotals=Product", _
              "Subtotals at=Bottom", "Grand totals=Bottom row", "Blank line=Yes", "Sort=Exposure desc", _
              "Units=Millions", "Highlight=Data bars on Exposure", "Tab=Deep", _
              "Description=Each product's ten largest counterparties and their share of the product."))
    c.Add Array("matyear", "pivot", "Maturity by year", _
        "Balances by the year they mature, local and foreign, with the running total.", _
        "Grouped by year  " & ChrW(183) & "  running total", _
        Array("On=Yes", "Pivot={fw} maturity by year", "Frameworks=All", "Rows=Maturity date", _
              "Columns=LCY / FCY", "Group=Maturity date by year", _
              "Values=Pre factor amount sum as Pre-factor; Pre factor amount running in Maturity date as Cumulative", _
              "Layout=Tabular", "Grand totals=Bottom row", "Units=Millions", "Tab=Deep", _
              "Description=Balances by the year they mature, with the running total."))
    c.Add Array("rulerank", "pivot", "Rule ranking", _
        "The fifteen rules that carry the most, ranked, gross and net.", _
        "Top 15  " & ChrW(183) & "  rank  " & ChrW(183) & "  data bars", _
        Array("On=Yes", "Pivot={fw} rule ranking", "Frameworks=All", "Rows=Rule name", _
              "Values=Gross pre-factor sum as Exposure; Gross pre-factor rank as Rank; Pre factor amount sum as Net", _
              "Top / value filter=Top 15 by Exposure", "Layout=Tabular", "Grand totals=Bottom row", _
              "Sort=Exposure desc", "Units=Millions", "Highlight=Data bars on Exposure", "Widths=Rule name=60", _
              "Tab=Emerald", "Description=The fifteen rules that carry the most, ranked, gross and net."))
    c.Add Array("sector", "pivot", "Sector league table", _
        "Every sector by gross exposure, with its share of the book and its rank.", _
        "% of total  " & ChrW(183) & "  rank  " & ChrW(183) & "  data bars", _
        Array("On=Yes", "Pivot={fw} sector league", "Frameworks=All", "Rows=Sector", _
              "Values=Gross pre-factor sum as Exposure; Gross pre-factor %total as Share; Gross pre-factor rank as Rank", _
              "Layout=Tabular", "Grand totals=Bottom row", "Sort=Exposure desc", "Units=Millions", _
              "Highlight=Data bars on Exposure", "Tab=Emerald", _
              "Description=Every sector by gross exposure, its share of the book and its rank."))
    c.Add Array("ccybucket", "pivot", "Currency by bucket", _
        "Each currency's balances across the maturity buckets, coloured by size.", _
        "Buckets across  " & ChrW(183) & "  heatmap  " & ChrW(183) & "  millions", _
        Array("On=Yes", "Pivot={fw} currency by bucket", "Frameworks=All", "Rows=Currency", "Columns=Bucket", _
              "Values=Gross pre-factor sum as Gross", "Show only / hide=Bucket <> (no bucket)", "Layout=Tabular", _
              "Grand totals=Both", "Sort=Gross desc", "Units=Millions", "Highlight=Heatmap", "Tab=Deep", _
              "Description=Each currency's balances across the maturity buckets."))
    c.Add Array("haircut", "pivot", "Haircut by rule", _
        "The twenty rules whose factors take the most off, and the factor each one really applied.", _
        "Calculated fields  " & ChrW(183) & "  top 20  " & ChrW(183) & "  data bars", _
        Array("On=Yes", "Pivot={fw} haircut by rule", "Frameworks=All", "Rows=Rule name", _
              "Values=Pre factor amount sum as Pre-factor; Post factor amount sum as Post-factor; " & _
              "Haircut sum as Taken off; Effective factor sum as Applied factor", _
              "Top / value filter=Top 20 by Taken off", "Layout=Tabular", "Grand totals=Bottom row", _
              "Sort=Taken off desc", "Units=Millions", "Highlight=Data bars on Taken off", "Widths=Rule name=60", _
              "Tab=Emerald", "Description=Where the factors take the most off, and the factor each rule applied."))
    c.Add Array("lcyfcy", "pivot", "Local and foreign mix", _
        "How much of each line is local currency and how much foreign, as shares of the line.", _
        "% of row  " & ChrW(183) & "  outline  " & ChrW(183) & "  subtotals", _
        Array("On=Yes", "Pivot={fw} local and foreign mix", "Frameworks=All", "Rows=Type, Line", _
              "Columns=LCY / FCY", "Values=Gross pre-factor %row as Share", "Layout=Outline", "Subtotals=Type", _
              "Subtotals at=Top", "Grand totals=Both", "Tab=Deep", _
              "Description=How much of each line is local currency and how much foreign."))
    c.Add Array("banks", "pivot", "Bank counterparties", _
        "Every counterparty with BANK in its name - change the word to follow any other group.", _
        "Label filter: contains  " & ChrW(183) & "  sorted", _
        Array("On=Yes", "Pivot={fw} bank counterparties", "Frameworks=All", "Rows=Counterparty", _
              "Columns=LCY / FCY", "Values=Gross pre-factor sum as Exposure", _
              "Show only / hide=Counterparty contains BANK", "Layout=Tabular", "Grand totals=Both", _
              "Sort=Exposure desc", "Units=Millions", "Tab=Slate", _
              "Description=Every counterparty with BANK in its name."))
    c.Add Array("large", "pivot", "Large exposures", _
        "Counterparties above 100 million gross, each folded to open onto its products.", _
        "Value filter  " & ChrW(183) & "  expand to  " & ChrW(183) & "  outline", _
        Array("On=Yes", "Pivot={fw} large exposures", "Frameworks=All", "Rows=Counterparty, Product", _
              "Values=Gross pre-factor sum as Exposure", "Show only / hide=Counterparty <> (no counterparty)", _
              "Top / value filter=Counterparty: Exposure > 100m", _
              "Layout=Outline", "Subtotals=Counterparty", "Grand totals=Bottom row", "Expand to=Counterparty", _
              "Sort=Exposure desc", "Units=Millions", "Tab=Slate", _
              "Description=Counterparties above 100 million gross - open one to see its products."))
    c.Add Array("bssummary", "pivot", "Balance sheet summary", _
        "The balance sheet folded to its lines, subtotals on top - open any line to its sublines and COAs.", _
        "Outline  " & ChrW(183) & "  expand to  " & ChrW(183) & "  blank lines", _
        Array("On=Yes", "Pivot={fw} balance summary", "Frameworks=All", "Rows=Type, Line, Subline, COA name", _
              "Columns=LCY / FCY", "Values=Pre factor amount sum as Pre-factor", _
              "Show only / hide=Bucket <> (no bucket)", "Layout=Outline", "Subtotals=All", "Subtotals at=Top", _
              "Grand totals=Both", "Blank line=Yes", "Expand to=Line", "Units=Millions", "Tab=Emerald", _
              "Description=The balance sheet folded to its lines - open any line to its sublines and COAs."))
    c.Add Array("ch_ccy", "chart", "Currency mix", "The eight largest currencies by gross amount, as a pie.", _
        "Pie  " & ChrW(183) & "  top 8  " & ChrW(183) & "  percent labels", _
        Array("On=Yes", "Chart=Currency mix", "Frameworks=All", "Type=Pie", "Categories=Currency", _
              "Values=Gross pre-factor sum as Gross", "Top / value filter=Top 8 by Gross", "Sort=Gross desc", _
              "Where=Start here", "Size=Third", "Labels=Percent", "Legend=Right", "Colours=Avati", _
              "Description=The eight largest currencies by gross amount."))
    c.Add Array("ch_buckets", "chart", "Across the buckets", _
        "Balances in each maturity bucket, local and foreign stacked, in millions.", _
        "Stacked column  " & ChrW(183) & "  series by LCY / FCY", _
        Array("On=Yes", "Chart=Across the buckets", "Frameworks=All", "Type=Stacked column", "Categories=Bucket", _
              "Series=LCY / FCY", "Values=Pre factor amount sum as Pre-factor", _
              "Show only / hide=Bucket <> (no bucket)", "Where=Start here", "Size=Full", "Legend=Top", _
              "Units=Auto", "Colours=Two tone", "Description=Balances in each maturity bucket, local and foreign."))
    c.Add Array("ch_matyear", "chart", "Maturing by year", _
        "How much matures each year as columns, and the running share of the book as a line.", _
        "Column + line  " & ChrW(183) & "  grouped by year  " & ChrW(183) & "  own sheet", _
        Array("On=Yes", "Chart=Maturing by year", "Frameworks=All", "Type=Column + line", _
              "Categories=Maturity date", "Group=Maturity date by year", _
              "Values=Gross pre-factor sum as Maturing; Gross pre-factor %running in Maturity date as Cumulative", _
              "Where=Own sheet", "Legend=Top", "Units=Auto", "Colours=Two tone", _
              "Description=How much matures each year, and the running share of the book."))
    c.Add Array("ch_products", "chart", "Products across the buckets", _
        "Each bucket's mix of products, every column to 100%.", _
        "100% column  " & ChrW(183) & "  series by product  " & ChrW(183) & "  own sheet", _
        Array("On=Yes", "Chart=Products across the buckets", "Frameworks=All", "Type=100% column", _
              "Categories=Bucket", "Series=Product", "Values=Gross pre-factor sum as Gross", _
              "Show only / hide=Bucket <> (no bucket)", "Where=Own sheet", "Legend=Right", "Colours=Avati", _
              "Description=Each bucket's mix of products."))
    c.Add Array("ch_sectors", "chart", "Sectors", "The eight largest sectors by gross amount, as a doughnut.", _
        "Doughnut  " & ChrW(183) & "  top 8  " & ChrW(183) & "  emerald", _
        Array("On=Yes", "Chart=Sectors", "Frameworks=All", "Type=Doughnut", "Categories=Sector", _
              "Values=Gross pre-factor sum as Gross", "Top / value filter=Top 8 by Gross", "Sort=Gross desc", _
              "Where=Start here", "Size=Third", "Labels=Percent", "Legend=Right", "Colours=Emerald", _
              "Description=The eight largest sectors by gross amount."))
    c.Add Array("ch_haircut", "chart", "Haircut by rule category", _
        "What the factors take off in each category of rule, biggest first.", _
        "Bar  " & ChrW(183) & "  calculated field  " & ChrW(183) & "  value labels", _
        Array("On=Yes", "Chart=Haircut by rule category", "Frameworks=All", "Type=Bar", _
              "Categories=Rule category", "Values=Haircut sum as Taken off", "Sort=Taken off desc", _
              "Where=Start here", "Size=Half", "Labels=Values", "Legend=None", "Units=Auto", "Colours=Avati", _
              "Description=What the factors take off in each category of rule."))
    Set Templates = c
End Function

' ===================== the sheet ============================================

Public Sub BuildGallerySheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_GALLERY)
    On Error Resume Next
    ws.Cells.Clear
    modPD_Theme.Dress ws, "Gallery", _
        "Reports worth having, ready to build. Add puts the report on Pivot config or Chart config, switched " & _
        "on; from there it is a row like any other - change what you like.", _
        "REPORTS  " & ChrW(183) & "  GALLERY"
    ws.Rows(6).RowHeight = 8
    Err.Clear
    PaintCards ws
End Sub

' Every card, drawn again with what is on the sheets now.
Public Sub PaintCards(ByVal ws As Worksheet)
    Dim t As Variant, i As Long, nPivot As Long, nChart As Long, r As Long, kind As String, lastKind As String
    Dim x As Double, y As Double, col As Long, nOn As Long, st As String
    On Error Resume Next
    modPD_Theme.ClearButtons ws, "pdg_"
    r = FIRST_ROW
    For Each t In Templates()
        kind = CStr(t(1))
        If kind <> lastKind Then
            If Len(lastKind) > 0 Then r = r + 1
            ' a band naming the kind, then card rows
            ws.Rows(r).RowHeight = 30
            modPD_Theme.BarText ws, "pdg_band_" & kind, IIf(kind = "pivot", "PIVOT REPORTS", "CHARTS"), _
                LEFT0 + 2, ws.Rows(r).Top + 12, 8, modPD_Theme.C_LINK, modPD_Theme.UI_SEMI, 1.2
            r = r + 1
            col = 0
            lastKind = kind
        End If
        If col = PER_ROW Then
            r = r + 1
            col = 0
        End If
        ws.Rows(r).RowHeight = CARD_H + GAP
        x = LEFT0 + col * (CARD_W + GAP)
        y = ws.Rows(r).Top
        st = StateOf(t)
        If st = "on" Then nOn = nOn + 1
        Card ws, t, x, y, st
        col = col + 1
        If kind = "pivot" Then nPivot = nPivot + 1 Else nChart = nChart + 1
    Next t
    modPD_Theme.SetStatus ws, nPivot & " pivot reports and " & nChart & " charts. " & _
        IIf(nOn > 0, nOn & " already in the next build.", "None in the next build yet."), "Idle"
    Err.Clear
End Sub

' One card: what kind, the name, what it shows, what it is made of, and the
' button - Add, Switch on, or Open when it is already on.
Private Sub Card(ByVal ws As Worksheet, ByVal t As Variant, ByVal x As Double, ByVal y As Double, _
                 ByVal st As String)
    Dim sh As Shape, key As String, cap As String, kind As Long
    On Error Resume Next
    key = CStr(t(0))
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, CARD_W, CARD_H)
    sh.Name = "pdg_card_" & key
    sh.Adjustments.Item(1) = 0.06
    sh.Fill.ForeColor.RGB = modPD_Theme.C_SURFACE
    sh.Line.ForeColor.RGB = IIf(st = "on", modPD_Theme.C_BRAND, modPD_Theme.C_HAIRLINE_2)
    sh.Line.Weight = 0.75
    sh.Shadow.visible = msoFalse
    sh.Placement = xlFreeFloating
    sh.AlternativeText = CStr(t(2))
    Glyph ws, key, CStr(t(1)), x + CARD_W - 64, y + 16
    modPD_Theme.BarText ws, "pdg_kind_" & key, IIf(CStr(t(1)) = "pivot", "PIVOT", "CHART"), x + 16, y + 16, 7, _
        modPD_Theme.C_LINK, modPD_Theme.UI_SEMI, 1.2
    modPD_Theme.BarText ws, "pdg_title_" & key, CStr(t(2)), x + 16, y + 30, 12.5, modPD_Theme.C_TEXT, _
        modPD_Theme.UI_SEMI
    Wrapped ws, "pdg_desc_" & key, CStr(t(3)), x + 16, y + 54, CARD_W - 32, 44, 9, modPD_Theme.C_TEXT_2
    Wrapped ws, "pdg_uses_" & key, CStr(t(4)), x + 16, y + 102, CARD_W - 32, 16, 7.5, modPD_Theme.C_TEXT_3
    Select Case st
        Case "on": cap = "Open":  kind = 2
        Case "off": cap = "Switch on": kind = 1
        Case Else: cap = "Add": kind = 1
    End Select
    modPD_Theme.Pill ws, "pdg_btn_" & key, cap, "PD_GalleryAdd", x + 16, y + CARD_H - 36, 86, 24, kind
    If st = "on" Then
        modPD_Theme.BarText ws, "pdg_state_" & key, ChrW(9679) & "  IN THE NEXT BUILD", x + 112, y + CARD_H - 29, _
            7, modPD_Theme.C_BRAND_BRIGHT, modPD_Theme.UI_SEMI, 1
    ElseIf st = "off" Then
        modPD_Theme.BarText ws, "pdg_state_" & key, "ON THE SHEET, SWITCHED OFF", x + 112, y + CARD_H - 29, _
            7, modPD_Theme.C_TEXT_3, modPD_Theme.UI_SEMI, 1
    End If
    Err.Clear
End Sub

' A small picture of the kind of report: rows of a table, or a chart's bars.
Private Sub Glyph(ByVal ws As Worksheet, ByVal key As String, ByVal kind As String, ByVal x As Double, _
                  ByVal y As Double)
    Dim i As Long, sh As Shape, h As Variant, w As Variant
    On Error Resume Next
    If kind = "chart" Then
        h = Array(14, 24, 19, 30, 11)
        For i = 0 To 4
            Set sh = ws.Shapes.AddShape(msoShapeRectangle, x + i * 10, y + 32 - h(i), 7, h(i))
            Paint sh, "pdg_glyph_" & key & "_" & i, IIf(i = 3, modPD_Theme.C_BRAND_BRIGHT, modPD_Theme.C_BRAND_DEEP)
        Next i
    Else
        w = Array(48, 40, 44, 30)
        For i = 0 To 3
            Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y + 2 + i * 8, 14, 4)
            Paint sh, "pdg_glyph_" & key & "_l" & i, modPD_Theme.C_HAIRLINE_2
            Set sh = ws.Shapes.AddShape(msoShapeRectangle, x + 18, y + 2 + i * 8, CDbl(w(i)) - 18, 4)
            Paint sh, "pdg_glyph_" & key & "_v" & i, IIf(i = 0, modPD_Theme.C_BRAND_BRIGHT, modPD_Theme.C_BRAND_DEEP)
        Next i
    End If
    Err.Clear
End Sub

Private Sub Paint(ByVal sh As Shape, ByVal nm As String, ByVal clr As Long)
    If sh Is Nothing Then Exit Sub
    sh.Name = nm
    sh.Fill.ForeColor.RGB = clr
    sh.Line.visible = msoFalse
    sh.Shadow.visible = msoFalse
    sh.Placement = xlFreeFloating
End Sub

Private Sub Wrapped(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal x As Double, _
                    ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal size As Single, ByVal clr As Long)
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    If sh Is Nothing Then Exit Sub
    sh.Name = nm
    modPD_Theme.Plain sh
    With sh.TextFrame2
        .WordWrap = msoTrue
        .TextRange.Text = txt
        .TextRange.Font.Name = modPD_Theme.UI_FONT
        .TextRange.Font.Size = size
        .TextRange.Font.Fill.ForeColor.RGB = clr
        .TextRange.ParagraphFormat.SpaceWithin = 1.1
    End With
    Err.Clear
End Sub

' ===================== adding ===============================================

' "on", "off" - the row is on its sheet - or "" when it is not.
Private Function StateOf(ByVal t As Variant) As String
    Dim ws As Worksheet, r As Long
    r = FindRow(t, ws)
    If r = 0 Then Exit Function
    If StrComp(SafeText(ws.Cells(r, 1).Value2), "Yes", vbTextCompare) = 0 Then StateOf = "on" Else StateOf = "off"
End Function

' The row already holding this template on its sheet - found by name, which
' is column 2 on both sheets - or 0.
Private Function FindRow(ByVal t As Variant, ByRef ws As Worksheet) As Long
    Dim r As Long, lastR As Long, nm As String
    Set ws = GetSheet(IIf(CStr(t(1)) = "pivot", SH_CONFIG, SH_CHARTS))
    If ws Is Nothing Then Exit Function
    nm = NameIn(t)
    lastR = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    For r = modPD_Theme.R_FIRST To lastR
        If StrComp(SafeText(ws.Cells(r, 2).Value2), nm, vbTextCompare) = 0 Then FindRow = r: Exit Function
    Next r
End Function

' The name the template's row carries: its Pivot= or Chart= value.
Private Function NameIn(ByVal t As Variant) As String
    Dim kv As Variant, i As Long
    kv = t(5)
    For i = LBound(kv) To UBound(kv)
        If Left$(CStr(kv(i)), 6) = "Pivot=" Or Left$(CStr(kv(i)), 6) = "Chart=" Then NameIn = Mid$(CStr(kv(i)), 7)
    Next i
End Function

' A card's button: Add writes the row, switched on; Switch on turns the row
' on; Open goes to it. Every path ends on the row, checked.
Public Sub PD_GalleryAdd()
    Dim key As String, t As Variant, found As Variant, ws As Worksheet, r As Long, ev As Boolean, verb As String
    If PD_Busy Then Exit Sub
    On Error Resume Next
    key = Mid$(CStr(Application.Caller), Len("pdg_btn_") + 1)
    On Error GoTo 0
    For Each t In Templates()
        If CStr(t(0)) = key Then found = t: Exit For
    Next t
    If IsEmpty(found) Then Exit Sub
    modPD_Desk.PressFx
    modPD_Theme.PD_GoConfig                      ' makes the config sheets if they are missing
    ev = Application.EnableEvents
    Application.EnableEvents = False
    r = FindRow(found, ws)
    If r > 0 Then
        If StrComp(SafeText(ws.Cells(r, 1).Value2), "Yes", vbTextCompare) <> 0 Then
            ws.Cells(r, 1).Value2 = "Yes"
            verb = "switched on"
        Else
            verb = "already on"
        End If
    Else
        r = FreeRow(ws)
        If CStr(found(1)) = "pivot" Then
            modPD_Config.WriteByHeads ws, r, modPD_Config.RecipeHeads(), found(5)
        Else
            modPD_Config.WriteByHeads ws, r, modPD_Charts.ChartHeads(), found(5)
        End If
        verb = "added"
    End If
    If CStr(found(1)) = "pivot" Then
        modPD_Config.CheckAll True
    Else
        modPD_Charts.CheckCharts True
    End If
    Application.EnableEvents = ev
    PaintCards GetSheet(SH_GALLERY)
    modPD_Desk.RefreshDesk
    modPD_Theme.GoTo_ ws.Name
    ws.Cells(r, 2).Select
    Notify CStr(found(2)) & " " & verb & " - row " & r & " of " & ws.Name & ". The next build " & _
           IIf(CStr(found(1)) = "pivot", "makes it.", "draws it."), V_OK
End Sub

Private Function FreeRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = modPD_Theme.R_FIRST
    Do While Len(SafeText(ws.Cells(r, 2).Value2)) > 0 Or Len(SafeText(ws.Cells(r, 5).Value2)) > 0 Or _
             Len(SafeText(ws.Cells(r, 7).Value2)) > 0
        r = r + 1
    Loop
    FreeRow = r
End Function
