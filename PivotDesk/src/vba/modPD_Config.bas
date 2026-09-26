Option Explicit

' ============================================================================
'  Pivot config: what every framework workbook is built from.
'
'  One row is one pivot - or a family of them, one sheet per value of a
'  "one sheet per" field. The row says which frameworks it is for, what goes
'  in the rows, the columns and the values, which items are shown, and how
'  it looks: layout, subtotals, grand totals, repeated labels, sort, widths,
'  number format, tab colour. Nothing about a built workbook is decided in
'  code any more; it is decided here, where the person who reads the pivots
'  can change it.
'
'  Beside it, on its own sheet, the field catalog - Pivot fields: every column
'  of an output the recipes may name, under a plain name, with what to show
'  for a blank and how wide it wants to be. Thirteen of them are built in -
'  the columns 1.0 always staged - and the rest are there to be used.
'
'  The defaults reproduce 1.0's workbooks exactly. And the built-in 1.0 layout
'  is still here, behind "Engine": if a recipe ever gets in the way of a
'  deadline, one click builds what 1.0 built.
'
'  Check reads every row, says OK or says exactly what will not build, next
'  to the row. A build runs the same check first and stops rather than build
'  something the sheet did not ask for.
' ============================================================================

' --- the recipe table -------------------------------------------------------
Public Const K_ON As Long = 1
Public Const K_NAME As Long = 2
Public Const K_FW As Long = 3
Public Const K_SPLIT As Long = 4
Public Const K_ROWS As Long = 5
Public Const K_COLS As Long = 6
Public Const K_VALUES As Long = 7
Public Const K_FILTERS As Long = 8
Public Const K_VFILTER As Long = 9
Public Const K_GROUP As Long = 10
Public Const K_SLICERS As Long = 11
Public Const K_LAYOUT As Long = 12
Public Const K_SUBTOT As Long = 13
Public Const K_SUBAT As Long = 14
Public Const K_GRAND As Long = 15
Public Const K_TOTAL As Long = 16
Public Const K_REPEAT As Long = 17
Public Const K_BLANKLN As Long = 18
Public Const K_VALIN As Long = 19
Public Const K_EXPAND As Long = 20
Public Const K_SORT As Long = 21
Public Const K_UNITS As Long = 22
Public Const K_FORMAT As Long = 23
Public Const K_HILITE As Long = 24
Public Const K_WIDTHS As Long = 25
Public Const K_TILES As Long = 26
Public Const K_TAB As Long = 27
Public Const K_MAX As Long = 28
Public Const K_DESC As Long = 29
Public Const K_CHECK As Long = 30
Public Const K_WHY As Long = 31
Public Const K_LAST As Long = 31

' Rows dressed for recipes. More can be added below; they are read to the
' last row in use.
Private Const RECIPE_ROWS As Long = 30

' Pivot config or Pivot fields edited since the Check column was last written.
Private mDirty As Boolean
Private Const R_GROUPS As Long = 6

' --- the field catalog --------------------------------------------------------
Private Const G_NAME As Long = 1
Private Const G_SOURCE As Long = 2
Private Const G_KIND As Long = 3
Private Const G_LABELS As Long = 4
Private Const G_BLANK As Long = 5
Private Const G_WIDTH As Long = 6
Private Const G_FORMAT As Long = 7
Private Const G_NOTE As Long = 8
Private Const G_LAST As Long = 8
' "Words kept in capitals", beside the table: its heading and its value.
Private Const G_KEEP_COL As Long = 10

' Computed sources: not columns of the file, but worked out while staging.
Public Const SRC_PRE As String = "=PRE"
Public Const SRC_POST As String = "=POST"
Public Const SRC_CCYCLASS As String = "=CCYCLASS"
Public Const SRC_FACTOR As String = "=FACTOR"
' The amounts without their sign: what is at stake either way.
Public Const SRC_ABS_PRE As String = "=|PRE|"
Public Const SRC_ABS_POST As String = "=|POST|"

Public Const MAX_SHEETS_DEFAULT As Long = 120

' ===================== which engine =========================================

' "recipes" builds from this sheet; "classic" builds what 1.0 built.
Public Function Engine() As String
    If SettingGet("engine", "recipes") = "classic" Then Engine = "classic" Else Engine = "recipes"
End Function

Public Sub PD_ConfigEngine()
    If Engine() = "recipes" Then
        If MsgBox("Build with the 1.0 layout instead of this sheet?" & vbCrLf & vbCrLf & _
                  "Your recipes are kept. Switch back any time.", vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
        SettingSet "engine", "classic"
    Else
        SettingSet "engine", "recipes"
    End If
    PD_ConfigCheck
    modPD_Theme.RailEverywhere
    modPD_Desk.RefreshDesk
End Sub

' ===================== the sheet ============================================

Private Function ConfigSheet() As Worksheet
    Set ConfigSheet = GetSheet(SH_CONFIG)
End Function

' Dresses both sheets and, when they are new or asked to, fills them with the
' defaults. An existing configuration is never overwritten by a restyle.
Public Sub BuildConfigSheet(Optional ByVal withDefaults As Boolean = False)
    Dim ws As Worksheet, fresh As Boolean, ev As Boolean, old As Collection
    ' Writing a few hundred cells here must not run the live check on each.
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done
    Set ws = GetSheet(SH_CONFIG)
    fresh = ws Is Nothing
    Set ws = EnsureSheet(SH_CONFIG)
    If Not fresh And Not withDefaults Then
        If Not HeadersMatch(ws, RecipeHeads()) Then Set old = Remember(ws)
    End If
    If fresh Or withDefaults Or Not old Is Nothing Then
        ws.Cells.Clear
        ws.Cells.Validation.Delete
    End If

    modPD_Theme.Dress ws, "Pivot config", _
        "What every framework workbook is built from. One row is one pivot, or one sheet per value of a field. " & _
        "Edit a row or add one; Check says whether it will build. Select any cell for how to fill it.", _
        "REPORTS  " & ChrW(183) & "  PIVOTS"
    Groups ws
    modPD_Theme.Head ws, RecipeHeads(), RecipeWidths()
    If Not old Is Nothing Then PutBack ws, old, RecipeHeads()

    If fresh Or withDefaults Or Len(SafeText(ws.Cells(modPD_Theme.R_FIRST, K_NAME).Value2)) = 0 Then
        WriteDefaultRecipes ws
    End If
    DressRecipes ws
    Hints ws
    modPD_Theme.PrintReady ws, K_LAST
    BuildFieldsSheet withDefaults
    mDirty = False
Done:
    ' Events are Excel's, not this workbook's: they go back on whatever happened.
    Application.EnableEvents = ev
End Sub

Public Sub BuildFieldsSheet(Optional ByVal withDefaults As Boolean = False)
    Dim ws As Worksheet, fresh As Boolean, old As Collection
    Set ws = GetSheet(SH_FIELDS)
    fresh = ws Is Nothing
    Set ws = EnsureSheet(SH_FIELDS)
    ' A sheet from an older version has other columns: its rows are read by
    ' heading and put back under the same headings after the rebuild.
    If Not fresh And Not withDefaults Then
        If Not HeadersMatch(ws, FieldHeads()) Then Set old = Remember(ws)
    End If
    If fresh Or withDefaults Or Not old Is Nothing Then
        ws.Cells.Clear
        ws.Cells.Validation.Delete
    End If
    modPD_Theme.Dress ws, "Pivot fields", _
        "Every column a pivot or a chart may name. The first thirteen are built in; add any column of an " & _
        "output under a name of your own, and use that name anywhere.", _
        "REPORTS  " & ChrW(183) & "  FIELDS"
    modPD_Theme.Head ws, FieldHeads(), Array(26, 42, 11, 24, 20, 8, 22, 56)
    If Not old Is Nothing Then PutBack ws, old, FieldHeads()
    If fresh Or withDefaults Or Len(SafeText(ws.Cells(modPD_Theme.R_FIRST, G_NAME).Value2)) = 0 Then
        WriteDefaultFields ws
    End If
    KeepBlock ws
    modPD_Theme.SetStatus ws, FieldCount(ws) & " fields. Built-in fields cannot be renamed; any other row can " & _
        "be changed, and new ones added at the bottom.", "Idle"
    DressFields ws
    FieldHints ws
    modPD_Theme.PrintReady ws, G_LAST
End Sub

Private Function FieldCount(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = modPD_Theme.R_FIRST
    Do While Len(SafeText(ws.Cells(r, G_NAME).Value2)) > 0
        r = r + 1
    Loop
    FieldCount = r - modPD_Theme.R_FIRST
End Function

' A band over the header naming what each run of columns is for - a
' twenty-column table is only navigable when it says where you are.
Private Sub Groups(ByVal ws As Worksheet)
    ws.Rows(R_GROUPS).RowHeight = 20
    Band ws, K_ON, K_SPLIT, "WHICH PIVOT"
    Band ws, K_ROWS, K_SLICERS, "WHAT IT SHOWS"
    Band ws, K_LAYOUT, K_TILES, "HOW IT LOOKS"
    Band ws, K_TAB, K_DESC, "THE SHEET"
    Band ws, K_CHECK, K_WHY, "CHECK"
End Sub

Public Function RecipeHeads() As Variant
    RecipeHeads = Array("On", "Pivot", "Frameworks", "One sheet per", "Rows", "Columns", "Values", _
                        "Show only / hide", "Top / value filter", "Group", "Slicers", "Layout", "Subtotals", _
                        "Subtotals at", "Grand totals", "Total label", "Repeat labels", "Blank line", "Values in", _
                        "Expand to", "Sort", "Units", "Number format", "Highlight", "Widths", "Tiles", "Tab", _
                        "Max sheets", "Description", "Check", "What to fix")
End Function

Private Function RecipeWidths() As Variant
    RecipeWidths = Array(7, 22, 16, 20, 34, 14, 46, 30, 26, 24, 24, 10, 12, 11, 13, 11, 9, 9, 10, 12, 16, _
                         10, 16, 12, 24, 7, 9, 9, 40, 10, 60)
End Function

Private Sub Band(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, ByVal label As String)
    On Error Resume Next
    With ws.Range(ws.Cells(R_GROUPS, c1), ws.Cells(R_GROUPS, c2))
        .Interior.Color = modPD_Theme.C_BRAND_950
        .Font.Name = modPD_Theme.UI_SEMI
        .Font.Size = 7.5
        .Font.Color = modPD_Theme.C_LINK
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = modPD_Theme.C_BRAND
        .Borders(xlEdgeLeft).Weight = xlMedium
    End With
    With ws.Cells(R_GROUPS, c1)
        .Value2 = label
        .IndentLevel = 1
    End With
    Err.Clear
End Sub

Private Sub DressRecipes(ByVal ws As Worksheet)
    Dim lastR As Long, rng As Range
    On Error Resume Next
    lastR = RecipeLastRow(ws)
    Set rng = ws.Range(ws.Cells(modPD_Theme.R_FIRST, 1), ws.Cells(lastR, K_LAST))
    modPD_Theme.DressTable ws, K_LAST, lastR
    rng.NumberFormat = "@"
    rng.WrapText = False
    With ws.Range(ws.Cells(modPD_Theme.R_FIRST, K_NAME), ws.Cells(lastR, K_NAME)).Font
        .Name = modPD_Theme.UI_SEMI
        .Color = modPD_Theme.C_TEXT
    End With
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, K_ON), ws.Cells(lastR, K_ON)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, K_DESC), ws.Cells(lastR, K_DESC)).Font.Color = modPD_Theme.C_TEXT_3
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, K_WHY), ws.Cells(lastR, K_WHY)).Font.Color = modPD_Theme.C_TEXT_2
    ' The syntax columns in the mono face: they are read character by character.
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, K_SPLIT), ws.Cells(lastR, K_SLICERS)).Font.Name = modPD_Theme.UI_FONT
    Err.Clear
End Sub

Private Sub DressFields(ByVal ws As Worksheet)
    Dim lastR As Long
    On Error Resume Next
    lastR = modPD_Theme.R_FIRST + FieldCount(ws) + 9       ' room to add fields
    modPD_Theme.DressTable ws, G_LAST, lastR
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, 1), ws.Cells(lastR, G_LAST)).NumberFormat = "@"
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, G_NAME), ws.Cells(lastR, G_NAME)).Font.Name = modPD_Theme.UI_SEMI
    With ws.Range(ws.Cells(modPD_Theme.R_FIRST, G_SOURCE), ws.Cells(lastR, G_SOURCE)).Font
        .Name = modPD_Theme.UI_MONO
        .Size = 9
    End With
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, G_NOTE), ws.Cells(lastR, G_NOTE)).Font.Color = modPD_Theme.C_TEXT_3
    Err.Clear
End Sub

' The last recipe row in use, and never fewer than the rows dressed for them.
Public Function RecipeLastRow(ByVal ws As Worksheet) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.count, K_NAME).End(xlUp).Row
    b = ws.Cells(ws.Rows.count, K_ROWS).End(xlUp).Row
    If b > a Then a = b
    If a < modPD_Theme.R_FIRST + RECIPE_ROWS - 1 Then a = modPD_Theme.R_FIRST + RECIPE_ROWS - 1
    RecipeLastRow = a
End Function

' The cell-by-cell help: select a cell, and Excel shows what goes in it.
' Each is kept under 255 characters - Excel's own limit on an input message.
Private Sub Hints(ByVal ws As Worksheet)
    Dim r1 As Long, r2 As Long
    r1 = modPD_Theme.R_FIRST
    r2 = RecipeLastRow(ws) + 20
    ListRule ws, r1, r2, K_ON, "Yes,No", "On", "Yes builds this pivot; No keeps the row but skips it."
    Hint ws, r1, r2, K_NAME, "Pivot", "The sheet name. {fw} becomes the framework (" & Chr$(34) & "{fw} Output" & _
         Chr$(34) & " -> LCR Output). With One sheet per, leave it as {split} and each sheet takes its value."
    Hint ws, r1, r2, K_FW, "Frameworks", "All, or any of: LCR, NSFR, Maturity ladder - separated by commas."
    Hint ws, r1, r2, K_SPLIT, "One sheet per", "Optional. A field (or two, comma-separated) to make one sheet " & _
         "per value of - e.g. Rule name, LCY / FCY. Biggest first; Max sheets caps how many."
    Hint ws, r1, r2, K_ROWS, "Rows", "Fields down the side, in order, separated by commas - e.g. Type, Line, " & _
         "Subline, COA name. Names come from the Pivot fields sheet."
    Hint ws, r1, r2, K_COLS, "Columns", "Fields across the top, separated by commas - e.g. LCY / FCY, or Bucket."
    Hint ws, r1, r2, K_VALUES, "Values", "Separated by ; - field, how, as caption. How: sum count average max " & _
         "min, then %row %col %total %parent running %running rank diff %diff index. " & _
         "e.g. Gross pre-factor %running in Counterparty as Cumulative"
    Hint ws, r1, r2, K_FILTERS, "Show only / hide", "Separated by ; - Field = a | b shows only a and b; " & _
         "Field <> a hides a; Field contains, does not contain, begins with or ends with some text. " & _
         "A field not in rows or columns becomes a report filter."
    Hint ws, r1, r2, K_VFILTER, "Top / value filter", "Top 25 by Exposure, Bottom 10 by Net, Top 5% by " & _
         "Exposure, Exposure > 1m, Exposure between 1m and 5m. By names a value caption. It filters the first " & _
         "row field; name another in front: Top 10 Counterparty by Exposure."
    Hint ws, r1, r2, K_GROUP, "Group", "A date by year, quarter, month or day; a number by a step. " & _
         "e.g. Maturity date by year; Interest rate by 0.5. The field must be in Rows or Columns."
    Hint ws, r1, r2, K_SLICERS, "Slicers", "Fields to put slicers on, separated by commas. Not added on " & _
         "One-sheet-per pivots: one slicer would filter every one of those sheets at once."
    ListRule ws, r1, r2, K_LAYOUT, "Tabular,Outline,Compact", "Layout", _
         "Tabular: one column per row field. Outline: nested with headers. Compact: all row fields in one column."
    Hint ws, r1, r2, K_SUBTOT, "Subtotals", "None, All, or the row fields to subtotal, separated by commas."
    ListRule ws, r1, r2, K_SUBAT, "Top,Bottom", "Subtotals at", _
         "Above or below each group. Outline and Compact only - Tabular always totals below. Blank: Excel's choice."
    ListRule ws, r1, r2, K_GRAND, "Both,Bottom row,Right column,None", "Grand totals", _
         "Bottom row totals each column; Right column totals each row."
    Hint ws, r1, r2, K_TOTAL, "Total label", "What the grand totals are called. Blank: Total."
    ListRule ws, r1, r2, K_REPEAT, "Yes,No", "Repeat labels", _
         "Yes repeats a row label on every line it covers - easier to filter and copy out."
    ListRule ws, r1, r2, K_BLANKLN, "Yes,No", "Blank line", _
         "Yes leaves a blank line after each group of the outer row fields."
    ListRule ws, r1, r2, K_VALIN, "Columns,Rows", "Values in", _
         "With two or more values: side by side in columns (the default), or stacked under the row labels."
    Hint ws, r1, r2, K_EXPAND, "Expand to", "Optional. A row field to fold the table to: the fields below it " & _
         "start closed, with +/- to open them. e.g. Line."
    Hint ws, r1, r2, K_SORT, "Sort", "Blank keeps the data's order. label asc / label desc sorts by the names; " & _
         "a value caption then asc / desc sorts by it - e.g. Pre-factor desc."
    ListRule ws, r1, r2, K_UNITS, "As is,Thousands,Millions,Billions", "Units", _
         "Shows the figures in thousands, millions or billions, and says so over the title. Filters and totals " & _
         "still use the full amounts."
    Hint ws, r1, r2, K_FORMAT, "Number format", "Blank: the desk's own, in the Units chosen. Or any Excel " & _
         "number format, e.g. #,##0.00 - it wins over Units."
    ListRule ws, r1, r2, K_HILITE, "None,Data bars,Heatmap,Negatives,Top 10", "Highlight", _
         "Marks the plain values. Add on and a caption for one value only - Data bars on Exposure. Top takes any " & _
         "number - Top 5.", False
    Hint ws, r1, r2, K_WIDTHS, "Widths", "Column widths, separated by ; - field=width, and values=width for " & _
         "the figures. e.g. COA name=44; values=16. Unset fields fit their longest label."
    ListRule ws, r1, r2, K_TILES, "Yes,No", "Tiles", _
         "Yes puts each plain value's live total in a tile above the pivot."
    ListRule ws, r1, r2, K_TAB, "Auto,Emerald,Deep,Slate,Black", "Tab", _
         "Tab colour. Auto: overviews emerald, LCY sheets deep green, FCY sheets slate."
    Hint ws, r1, r2, K_MAX, "Max sheets", "For One sheet per: the most sheets to make. Blank means " & _
         MAX_SHEETS_DEFAULT & "."
    Hint ws, r1, r2, K_DESC, "Description", "Shown on the sheet under its title, and on Start here."
    Hint ws, r1, r2, K_CHECK, "Check", "Written by Check. OK, or Break with the reason beside it."
End Sub

Private Sub FieldHints(ByVal ws As Worksheet)
    Dim r1 As Long, r2 As Long
    r1 = modPD_Theme.R_FIRST
    r2 = r1 + FieldCount(ws) + 60
    Hint ws, r1, r2, G_NAME, "Field", "The name recipes use. The first thirteen are built in and " & _
         "cannot be renamed."
    Hint ws, r1, r2, G_SOURCE, "Source column", "The column header in the output, e.g. " & _
         "COUNTERPARTY_NAME. Case and spacing do not matter. For a Calculated field, a formula over other " & _
         "fields' names in single quotes."
    ListRule ws, r1, r2, G_KIND, "Text,Number,Date,Calculated", "Kind", _
         "Text is grouped by; Number can be summed; Date is grouped as a date. Calculated: the source is a " & _
         "formula over other fields, worked out by the pivot - e.g. ='Pre factor amount' - 'Post factor amount'."
    ListRule ws, r1, r2, G_LABELS, "As is,Drop codes,Drop codes + title case,Title case", "Labels", _
         "How the values read in a pivot. Drop codes turns 1.07.00.MBGL.1360.LOANS TO CUSTOMERS into LOANS TO " & _
         "CUSTOMERS; title case makes it Loans to Customers. Blank: this field's default."
    Hint ws, r1, r2, G_BLANK, "Blank shows as", "What a blank cell reads as in a pivot, " & _
         "e.g. (no sector). A named item can be hidden; a true blank cannot, reliably."
    Hint ws, r1, r2, G_WIDTH, "Width", "Column width when this field is on the rows."
    Hint ws, r1, r2, G_FORMAT, "Number format", "For Number and Date fields."
End Sub

Private Sub Hint(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c As Long, _
                 ByVal title As String, ByVal msg As String)
    On Error Resume Next
    With ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputTitle = Left$(title, 32)
        .InputMessage = Left$(msg, 255)
        .ShowInput = True
        .ShowError = False
    End With
    Err.Clear
End Sub

' A dropdown. Strict refuses anything else; not strict offers the list and
' takes what is typed - "Data bars on Share".
Private Sub ListRule(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c As Long, _
                     ByVal items As String, ByVal title As String, ByVal msg As String, _
                     Optional ByVal strict As Boolean = True)
    On Error Resume Next
    With ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=items
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = Left$(title, 32)
        .InputMessage = Left$(msg, 255)
        .ErrorTitle = TOOL_NAME
        .ErrorMessage = "Choose one of: " & Replace(items, ",", ", ")
        .ShowInput = True
        .ShowError = strict
    End With
    Err.Clear
End Sub

' ===================== defaults =============================================

' The four recipes that make exactly what 1.0 made.
Private Sub WriteDefaultRecipes(ByVal ws As Worksheet)
    Dim r As Long
    r = modPD_Theme.R_FIRST
    ws.Range(ws.Cells(r, 1), ws.Cells(RecipeLastRow(ws), K_LAST)).ClearContents
    Rec ws, r, "On=Yes", "Pivot={fw} Output", "Frameworks=All", _
        "Rows=Rule order, Rule category, Rule name, Factor", "Columns=LCY / FCY", _
        "Values=Pre factor amount sum as Pre-factor; Post factor amount sum as Post-factor", _
        "Show only / hide=Bucket <> (no bucket)", "Slicers=Rule category, LCY / FCY, Bucket", "Layout=Tabular", _
        "Subtotals=None", "Grand totals=Both", "Tiles=Yes", "Tab=Emerald", _
        "Description=Every rule, in the order the engine evaluates them, against what it read and what it kept."
    Rec ws, r + 1, "On=Yes", "Pivot=Balance sheet", "Frameworks=All", "Rows=Type, Line, Subline, COA name", _
        "Columns=LCY / FCY", "Values=Pre factor amount sum as Pre-factor", "Show only / hide=Bucket <> (no bucket)", _
        "Slicers=Type, LCY / FCY, Rule category", "Layout=Tabular", "Subtotals=Type", "Grand totals=Both", _
        "Tiles=Yes", "Tab=Emerald", _
        "Description=The same balances as the balance sheet reads them, down to the COA. Pre-factor only."
    Rec ws, r + 2, "On=Yes", "Pivot={split}", "Frameworks=LCR, NSFR", "One sheet per=Rule name, LCY / FCY", _
        "Rows=Type, Line, Subline, COA name", "Columns=Bucket", "Values=Pre factor amount sum as Pre-factor", _
        "Show only / hide=Bucket <> (no bucket)", "Layout=Tabular", "Subtotals=None", "Grand totals=Both", _
        "Tiles=Yes", "Tab=Auto", "Max sheets=" & MAX_SHEETS_DEFAULT, _
        "Description=One rule, one currency side - balances across the maturity buckets."
    ' The ladder is one workbook per currency (Workbooks), so inside each the
    ' sheets are one per rule, as for LCR and NSFR.
    Rec ws, r + 3, "On=Yes", "Pivot={split}", "Frameworks=Maturity ladder", "One sheet per=Rule name", _
        "Rows=Type, Line, Subline, COA name", "Columns=Bucket", "Values=Pre factor amount sum as Pre-factor", _
        "Show only / hide=Bucket <> (no bucket)", "Layout=Tabular", "Subtotals=None", "Grand totals=Both", _
        "Tiles=Yes", "Tab=Auto", "Max sheets=" & MAX_SHEETS_DEFAULT, _
        "Description=One rule - its balances across the maturity buckets."
    Rec ws, r + 4, "On=Yes", "Pivot={fw} top counterparties", "Frameworks=All", "Rows=Counterparty", _
        "Columns=LCY / FCY", _
        "Values=Gross pre-factor sum as Exposure; Gross pre-factor %total as Share; " & _
               "Gross pre-factor %running in Counterparty as Cumulative", _
        "Top / value filter=Top 25 by Exposure", "Layout=Tabular", "Grand totals=Both", "Values in=Columns", _
        "Sort=Exposure desc", "Units=Millions", "Highlight=Data bars", "Widths=Counterparty=40", "Tiles=Yes", _
        "Tab=Emerald", "Description=The 25 largest counterparties by exposure, local and foreign - each one's " & _
        "share of the 25, and the running share down the list."
    ' Ready to switch on.
    Rec ws, r + 5, "On=No", "Pivot={fw} counterparty heatmap", "Frameworks=All", "Rows=Counterparty", _
        "Columns=Bucket", "Values=Gross pre-factor sum as Exposure", "Show only / hide=Bucket <> (no bucket)", _
        "Top / value filter=Top 30 by Exposure", "Layout=Tabular", "Grand totals=Both", "Sort=Exposure desc", _
        "Units=Millions", "Highlight=Heatmap", "Widths=Counterparty=40", "Tab=Emerald", _
        "Description=Where the 30 largest counterparties fall across the maturity buckets."
    Rec ws, r + 6, "On=No", "Pivot={fw} product concentration", "Frameworks=All", "Rows=Product, Counterparty", _
        "Values=Gross pre-factor sum as Exposure; Gross pre-factor %parent as Share of product", _
        "Top / value filter=Top 10 Counterparty by Exposure", "Layout=Tabular", "Subtotals=Product", _
        "Subtotals at=Bottom", "Grand totals=Bottom row", "Blank line=Yes", "Sort=Exposure desc", _
        "Units=Millions", "Highlight=Data bars", "Tab=Deep", _
        "Description=Each product's ten largest counterparties and their share of the product."
    Rec ws, r + 7, "On=No", "Pivot={fw} maturity by year", "Frameworks=All", "Rows=Maturity date", _
        "Columns=LCY / FCY", "Group=Maturity date by year", _
        "Values=Pre factor amount sum as Pre-factor; Pre factor amount running in Maturity date as Cumulative", _
        "Layout=Tabular", "Grand totals=Bottom row", "Units=Millions", "Tab=Deep", _
        "Description=Balances by the year they mature, with the running total."
    Rec ws, r + 8, "On=No", "Pivot={fw} rule ranking", "Frameworks=All", "Rows=Rule name", _
        "Values=Gross pre-factor sum as Exposure; Gross pre-factor rank as Rank; Pre factor amount sum as Net", _
        "Top / value filter=Top 15 by Exposure", "Layout=Tabular", "Grand totals=Bottom row", _
        "Sort=Exposure desc", "Units=Millions", "Highlight=Data bars", "Widths=Rule name=60", "Tab=Emerald", _
        "Description=The fifteen rules that carry the most, ranked, gross and net."
    Rec ws, r + 9, "On=No", "Pivot={split}", "Frameworks=All", "One sheet per=Product", _
        "Rows=Type, Line, COA name", "Columns=Bucket", "Values=Pre factor amount sum as Pre-factor", _
        "Show only / hide=Bucket <> (no bucket)", "Layout=Tabular", "Subtotals=Type", "Grand totals=Both", _
        "Tab=Deep", "Max sheets=20", "Description=One sheet per product: its balances across the maturity buckets."
End Sub

' A recipe row written by heading - Rec ws, r, "On=Yes", "Pivot={fw} Output", ...
' A heading not named stays blank, so a new column never shifts a default,
' and the build checks every heading named here is one the sheet has.
Private Sub Rec(ByVal ws As Worksheet, ByVal r As Long, ParamArray kv() As Variant)
    Dim heads As Variant, i As Long, c As Long, p As Long, k As String
    heads = RecipeHeads()
    ws.Range(ws.Cells(r, 1), ws.Cells(r, K_LAST)).NumberFormat = "@"
    For i = 0 To UBound(kv)
        p = InStr(CStr(kv(i)), "=")
        If p > 1 Then
            k = Left$(CStr(kv(i)), p - 1)
            For c = 0 To UBound(heads)
                If StrComp(CStr(heads(c)), k, vbTextCompare) = 0 Then
                    ws.Cells(r, c + 1).Value2 = Mid$(CStr(kv(i)), p + 1)
                    Exit For
                End If
            Next c
        End If
    Next i
End Sub

Private Sub WriteDefaultFields(ByVal ws As Worksheet)
    Dim r As Long, f As Variant
    r = modPD_Theme.R_FIRST
    ws.Range(ws.Cells(r, 1), ws.Cells(r + 200, G_LAST)).ClearContents
    For Each f In DefaultFields()
        ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST)).NumberFormat = "@"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST)).Value2 = _
            Array(f(0), f(1), f(2), DefaultLabels(CStr(f(0))), f(3), f(4), f(5), f(6))
        r = r + 1
    Next f
End Sub

Public Function FieldHeads() As Variant
    FieldHeads = Array("Field", "Source column", "Kind", "Labels", "Blank shows as", "Width", "Number format", "Note")
End Function

' How each field's values read in a pivot. The ledger's own names carry their
' codes in front and are set in capitals; those are tidied by default.
Private Function DefaultLabels(ByVal nm As String) As String
    Select Case nm
        Case H_LINE, H_SUBLINE, H_COA_NAME: DefaultLabels = "Drop codes + title case"
        Case "Counterparty", "Sector", "Report class", "Cashflow element", "Product": DefaultLabels = "Title case"
        Case Else: DefaultLabels = "As is"
    End Select
End Function

' The Labels setting as TidyLabel's mode. Blank means the field's default.
Public Function LabelMode(ByVal nm As String, ByVal setting As String) As Long
    If Len(Trim$(setting)) = 0 Then setting = DefaultLabels(nm)
    Select Case LCase$(Trim$(setting))
        Case "drop codes": LabelMode = 1
        Case "drop codes + title case", "drop codes and title case": LabelMode = 2
        Case "title case": LabelMode = 3
        Case Else: LabelMode = 0
    End Select
End Function

' Beside the table: acronyms of the bank's own to keep in capitals when a
' field is set to title case.
Private Sub KeepBlock(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Columns(G_KEEP_COL - 1).ColumnWidth = 3
    ws.Columns(G_KEEP_COL).ColumnWidth = 44
    With ws.Cells(modPD_Theme.R_HDR, G_KEEP_COL)
        .Value2 = "Words kept in capitals"
        .Font.Name = modPD_Theme.UI_SEMI
        .Font.Size = 8.5
        .Font.Color = modPD_Theme.C_BRAND_SOFT
        .Interior.Color = modPD_Theme.C_INK
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = modPD_Theme.C_BRAND
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    With ws.Cells(modPD_Theme.R_FIRST, G_KEEP_COL)
        .NumberFormat = "@"
        .Interior.Color = modPD_Theme.C_ROW_ALT
        .Font.Color = modPD_Theme.C_TEXT
        .IndentLevel = 1
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Rows(modPD_Theme.R_FIRST).RowHeight = 34
    With ws.Cells(modPD_Theme.R_FIRST + 1, G_KEEP_COL)
        .Value2 = "Yours, separated by commas. Built in already: ALM, ECL, FVTOCI, IFRS, LCR, LCY, FCY, " & _
                  "NSFR, OCI, NPL, SME, USD, EGP and about fifty more."
        .Font.Size = 8.5
        .Font.Color = modPD_Theme.C_TEXT_3
        .WrapText = True
        .IndentLevel = 1
        .VerticalAlignment = xlTop
    End With
    ws.Rows(modPD_Theme.R_FIRST + 1).RowHeight = 34
    With ws.Cells(modPD_Theme.R_FIRST, G_KEEP_COL).Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .InputTitle = "Words kept in capitals"
        .InputMessage = "Acronyms your outputs use that title case should leave alone, e.g. QNB, CIB, NBE."
        .ShowInput = True
    End With
    Err.Clear
End Sub

Public Function KeepCapsList() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet(SH_FIELDS)
    If Not ws Is Nothing Then KeepCapsList = SafeText(ws.Cells(modPD_Theme.R_FIRST, G_KEEP_COL).Value2)
    Err.Clear
End Function

' ===================== older sheets =========================================
'
' A sheet built by an older version has other columns. Its rows are read by
' heading before the sheet is rebuilt and written back under the same headings
' after, so adding a column never shifts anyone's rows.

Private Function HeadersMatch(ByVal ws As Worksheet, ByVal heads As Variant) As Boolean
    Dim c As Long
    For c = 0 To UBound(heads)
        If StrComp(SafeText(ws.Cells(modPD_Theme.R_HDR, c + 1).Value2), CStr(heads(c)), vbTextCompare) <> 0 Then Exit Function
    Next c
    HeadersMatch = True
End Function

Private Function Remember(ByVal ws As Worksheet) As Collection
    Dim out As Collection, r As Long, c As Long, lastC As Long, lastR As Long, d As Object, h As String, any_ As Boolean
    Set out = New Collection
    Set Remember = out
    lastC = ws.Cells(modPD_Theme.R_HDR, ws.Columns.count).End(xlToLeft).Column
    lastR = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    For r = modPD_Theme.R_FIRST To lastR
        Set d = NewMap()
        any_ = False
        For c = 1 To lastC
            h = SafeText(ws.Cells(modPD_Theme.R_HDR, c).Value2)
            If Len(h) > 0 Then
                d(h) = SafeText(ws.Cells(r, c).Value2)
                If Len(d(h)) > 0 Then any_ = True
            End If
        Next c
        If any_ Then out.Add d
    Next r
End Function

Private Sub PutBack(ByVal ws As Worksheet, ByVal rows As Collection, ByVal heads As Variant)
    Dim d As Object, r As Long, c As Long
    r = modPD_Theme.R_FIRST
    For Each d In rows
        For c = 0 To UBound(heads)
            If d.Exists(CStr(heads(c))) Then
                ws.Cells(r, c + 1).NumberFormat = "@"
                ws.Cells(r, c + 1).Value2 = d(CStr(heads(c)))
            End If
        Next c
        r = r + 1
    Next d
End Sub

' Built in first - the thirteen columns 1.0 always staged, under the names its
' pivots use - then the rest of an output's columns, ready to be named in a
' recipe.
' Written one field per statement: a single Array( ... ) over all of them
' needs more line continuations than VBA allows (24), and does not compile.
Private Function DefaultFields() As Collection
    Dim c As Collection
    Set c = New Collection
    c.Add Array(H_RULE_ORDER, F_RULE_ORDER, "Number", "", "8", "0", "Built in. The order the engine evaluates rules in.")
    c.Add Array(H_RULE_CAT, F_RULE_CAT, "Text", "(no category)", "18", "", "Built in.")
    c.Add Array(H_RULE_NAME, F_RULE_NAME, "Text", "(no rule)", "40", "", "Built in.")
    c.Add Array(H_FACTOR, SRC_FACTOR, "Text", "", "9", "", "Built in. Post-factor over pre-factor, written as a label.")
    c.Add Array(H_TYPE, F_TYPE, "Text", "(no type)", "14", "", "Built in.")
    c.Add Array(H_LINE, F_LINE, "Text", "(no line)", "34", "", "Built in.")
    c.Add Array(H_SUBLINE, F_SUBLINE, "Text", "(no subline)", "34", "", "Built in.")
    c.Add Array(H_COA_NAME, F_COA_NAME, "Text", "(no COA)", "38", "", "Built in.")
    c.Add Array(H_CURRENCY, F_CURRENCY, "Text", "(no currency)", "16", "", "Built in.")
    c.Add Array(H_CCYCLASS, SRC_CCYCLASS, "Text", "", "10", "", "Built in. LCY when the row's currency is the most used one.")
    c.Add Array(H_BUCKET, F_BUCKET, "Text", "(no bucket)", "14", "", "Built in.")
    c.Add Array(H_PRE, SRC_PRE, "Number", "", "14", NUM_FMT, "Built in. Native-currency pre-factor if the file has it, else LCY.")
    c.Add Array(H_POST, SRC_POST, "Number", "", "14", NUM_FMT, "Built in. The post-factor pair of the same.")
    ' Computed while staging: the amounts without their sign - how much is at
    ' stake either way - which is what a ranking or a share should be read on.
    c.Add Array("Gross pre-factor", SRC_ABS_PRE, "Number", "", "14", NUM_FMT, "The pre-factor amount without its sign. Rank and share on this.")
    c.Add Array("Gross post-factor", SRC_ABS_POST, "Number", "", "14", NUM_FMT, "The post-factor amount without its sign.")
    ' Worked out by the pivot from the sums, whatever the rows add up to.
    c.Add Array("Haircut", "='" & H_PRE & "' - '" & H_POST & "'", "Calculated", "", "14", NUM_FMT, "Calculated: pre-factor less post-factor - what the factors took off.")
    c.Add Array("Effective factor", "=IF('" & H_PRE & "'=0,0,'" & H_POST & "'/'" & H_PRE & "')", "Calculated", "", "10", "0.0%", "Calculated: post over pre for whatever the row adds up - the factor actually applied.")
    c.Add Array("COA code", "COA_CODE", "Text", "(no COA code)", "18", "", "")
    c.Add Array("Account", "ACCOUNT_NUMBER", "Text", "(no account)", "24", "", "")
    c.Add Array("Counterparty", "COUNTERPARTY_NAME", "Text", "(no counterparty)", "30", "", "")
    c.Add Array("Counterparty class", "COUNTERPARTY_CLASSIFICATION_CODE", "Text", "(none)", "18", "", "")
    c.Add Array("Sector", "SECTOR_NAME", "Text", "(no sector)", "24", "", "")
    c.Add Array("Product", "PRODUCT_TYPE", "Text", "(no product)", "16", "", "")
    c.Add Array("Data source", "DATA_SOURCE", "Text", "(no source)", "12", "", "")
    c.Add Array("Cashflow element", "ALM_CASHFLOW_ELEMENT_NAME", "Text", "(no element)", "20", "", "")
    c.Add Array("Engine bucket", "ALM_BUCKET_NAME", "Text", "(none)", "12", "", "")
    c.Add Array("Maturity date", "MATURITY_DATE", "Date", "", "12", "d mmm yyyy", "")
    c.Add Array("As of", F_AS_OF, "Date", "", "12", "d mmm yyyy", "")
    c.Add Array("Interest rate", "INTEREST_RATE", "Number", "", "10", "0.00", "The first of the two INTEREST_RATE columns.")
    c.Add Array("Stage", "STAGE_NAME", "Text", "(no stage)", "12", "", "")
    c.Add Array("Customer", "FLG_CUSTOMER", "Text", "(blank)", "10", "", "")
    c.Add Array("Rate sensitive", "FLAG_RATE_SENSITIVE", "Text", "(blank)", "12", "", "")
    c.Add Array("Branch", "BRANCH_CODE", "Text", "(no branch)", "12", "", "")
    c.Add Array("Rule code", "ALM_PORTFOLIO_SEGMENTATION_RULE_CODE", "Text", "(no rule code)", "16", "", "")
    c.Add Array("Report class", "JCB_IND_REPORT_CLASSIFICATION_NAME", "Text", "(none)", "26", "", "")
    c.Add Array("Engine factor", F_FACTOR, "Number", "", "10", "0.00", "The source's own factor column; units not stated.")
    c.Add Array("Deposits outstanding", "CUST_DEPOSIT_OUTSTANDING", "Number", "", "14", NUM_FMT, "")
    c.Add Array("Amount with segment", "CASHFLOW_AMOUNT_LCY_WITH_SEGMENT", "Number", "", "14", NUM_FMT, "")
    c.Add Array("Amount without segment", "CASHFLOW_AMOUNT_LCY_WITHOUT_SEGMENT", "Number", "", "14", NUM_FMT, "")
    c.Add Array("Amount with bucket", "CASHFLOW_AMOUNT_LCY_WITH_BUCKET", "Number", "", "14", NUM_FMT, "")
    c.Add Array("Amount without bucket", "CASHFLOW_AMOUNT_LCY_WITHOUT_BUCKET", "Number", "", "14", NUM_FMT, "")
    Set DefaultFields = c
End Function

Public Sub PD_ConfigDefaults()
    Dim ws As Worksheet
    If PD_Busy Then Exit Sub
    If MsgBox("Put the Pivot config back to its defaults?" & vbCrLf & vbCrLf & _
              "Every recipe, and every field on Pivot fields, is replaced. The defaults build exactly what " & _
              "version 1.0 built.", vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    BuildConfigSheet True
    modPD_Theme.Rail GetSheet(SH_CONFIG)
    modPD_Theme.Rail GetSheet(SH_FIELDS)
    Application.ScreenUpdating = True
    PD_ConfigCheck
End Sub

' ===================== reading ==============================================

' name -> Dictionary(Name, Source, Kind, Blank, Width, Format, Builtin)
Public Function Fields() As Object
    Dim ws As Worksheet, d As Object, r As Long, nm As String, bi As Object
    Dim v As Variant
    Set d = NewMap()
    Set Fields = d
    Set bi = BuiltinNames()
    Set ws = GetSheet(SH_FIELDS)
    If ws Is Nothing Then
        For Each v In DefaultFields()
            AddField d, v(0), v(1), v(2), v(3), v(4), v(5), bi, ""
        Next v
        Exit Function
    End If
    r = modPD_Theme.R_FIRST
    Do While Len(SafeText(ws.Cells(r, G_NAME).Value2)) > 0
        nm = SafeText(ws.Cells(r, G_NAME).Value2)
        If Not d.Exists(nm) Then
            AddField d, nm, SafeText(ws.Cells(r, G_SOURCE).Value2), SafeText(ws.Cells(r, G_KIND).Value2), _
                     SafeText(ws.Cells(r, G_BLANK).Value2), SafeText(ws.Cells(r, G_WIDTH).Value2), _
                     SafeText(ws.Cells(r, G_FORMAT).Value2), bi, SafeText(ws.Cells(r, G_LABELS).Value2)
        End If
        r = r + 1
    Loop
    ' A built-in that was deleted or renamed on the sheet is still built in.
    For Each v In DefaultFields()
        If bi.Exists(CStr(v(0))) And Not d.Exists(CStr(v(0))) Then AddField d, v(0), v(1), v(2), v(3), v(4), v(5), bi, ""
    Next v
End Function

Private Sub AddField(ByVal d As Object, ByVal nm As String, ByVal src As String, ByVal kind As String, _
                     ByVal blank As String, ByVal w As String, ByVal fmt As String, ByVal bi As Object, _
                     ByVal labels As String)
    Dim f As Object
    Set f = NewMap()
    f("Name") = nm
    f("Source") = src
    Select Case LCase$(kind)
        Case "number": f("Kind") = "Number"
        Case "date": f("Kind") = "Date"
        Case "calculated": f("Kind") = "Calculated"
        Case Else: f("Kind") = "Text"
    End Select
    f("Blank") = blank
    f("Width") = SafeNum(w)
    f("Format") = fmt
    f("Builtin") = bi.Exists(nm)
    f("Labels") = LabelMode(nm, labels)
    Set d(nm) = f
End Sub

Private Function BuiltinNames() As Object
    Dim d As Object, h As Variant
    Set d = NewMap()
    For Each h In StageHeadings()
        d(CStr(h)) = True
    Next h
    Set BuiltinNames = d
End Function

Private Function Cell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Cell = SafeText(ws.Cells(r, c).Value2)
End Function

' ===================== Check ================================================

' Reads every row, writes OK / Break and the reason beside it, and says on
' the status line how many will build. Returns the number of rows that are
' on and will not build.
Public Function CheckAll(Optional ByVal quiet As Boolean = False) As Long
    Dim ws As Worksheet, rc As Object, nOn As Long, nBad As Long, r As Long, lastR As Long
    Set ws = ConfigSheet()
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    lastR = RecipeLastRow(ws)
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, K_CHECK), ws.Cells(lastR, K_WHY)).ClearContents
    For Each rc In modPD_Recipe.AllRecipes()
        r = CLng(rc("Row"))
        If Not CBool(rc("On")) Then
            ws.Cells(r, K_CHECK).Value2 = "Off"
            If Len(rc("Problem")) > 0 Then ws.Cells(r, K_WHY).Value2 = "When switched on: " & rc("Problem")
        ElseIf Len(rc("Problem")) > 0 Then
            nOn = nOn + 1
            nBad = nBad + 1
            ws.Cells(r, K_CHECK).Value2 = V_BREAK
            ws.Cells(r, K_WHY).Value2 = rc("Problem")
        Else
            nOn = nOn + 1
            ws.Cells(r, K_CHECK).Value2 = V_OK
            If rc("Split").count > 0 And rc("Slicers").count > 0 Then
                ws.Cells(r, K_CHECK).Value2 = V_CHECK
                ws.Cells(r, K_WHY).Value2 = "Slicers are left off one-sheet-per pivots - one slicer would filter " & _
                                            "every sheet of the family at once."
            End If
        End If
    Next rc
    For r = modPD_Theme.R_FIRST To lastR
        If Len(SafeText(ws.Cells(r, K_CHECK).Value2)) > 0 Then
            modPD_Theme.PaintVerdict ws.Cells(r, K_CHECK)
        Else
            ws.Cells(r, K_CHECK).Interior.Color = IIf((r - modPD_Theme.R_FIRST) Mod 2 = 1, modPD_Theme.C_ROW_ALT, modPD_Theme.C_ROW)
            ws.Cells(r, K_CHECK).NumberFormat = "@"
        End If
    Next r
    If Engine() = "classic" Then
        modPD_Theme.SetStatus ws, "Building with the 1.0 layout - this sheet is kept but not used. " & _
            "Engine switches back.", "Check"
    ElseIf nBad > 0 Then
        modPD_Theme.SetStatus ws, nBad & " of " & nOn & " pivot(s) that are on will not build - the reason is " & _
            "beside each, under What to fix.", "Break"
    ElseIf nOn = 0 Then
        modPD_Theme.SetStatus ws, "No pivot is switched on - a build would make nothing.", "Check"
    Else
        modPD_Theme.SetStatus ws, nOn & " pivot(s) switched on, every one ready to build.", "OK"
    End If
    SettingSet "config_on", CStr(nOn)
    SettingSet "config_bad", CStr(nBad)
    CheckAll = nBad
    Err.Clear
End Function

' A new row, ready to edit: switched off, so a half-finished pivot cannot stop
' a build, and holding a working pivot to change rather than a blank to fill.
Public Sub PD_ConfigAdd()
    Dim ws As Worksheet, r As Long, lastR As Long, ev As Boolean
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    Set ws = ConfigSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    lastR = RecipeLastRow(ws)
    For r = modPD_Theme.R_FIRST To lastR + 1
        If Len(Cell(ws, r, K_NAME)) = 0 And Len(Cell(ws, r, K_ROWS)) = 0 And Len(Cell(ws, r, K_VALUES)) = 0 Then Exit For
    Next r
    ev = Application.EnableEvents
    Application.EnableEvents = False
    Rec ws, r, "On=No", "Pivot={fw} new pivot " & (r - modPD_Theme.R_FIRST + 1), "Frameworks=All", _
        "Rows=Type, Line", "Columns=Bucket", "Values=Pre factor amount sum as Pre-factor", _
        "Show only / hide=Bucket <> (no bucket)", "Layout=Tabular", "Subtotals=None", "Grand totals=Both", _
        "Tiles=Yes", "Tab=Auto", "Description=A new pivot. Change any cell, then set On to Yes."
    If r > lastR Then
        DressRecipes ws
        Hints ws
    End If
    CheckAll True
    Application.EnableEvents = ev
    mDirty = False
    ws.Activate
    ws.Cells(r, K_NAME).Select
    Notify "A new pivot is on row " & r & ", switched off. Change what you need, then set On to Yes.", V_OK
    Err.Clear
End Sub

' ===================== live check ============================================
'
' Typing on Pivot config answers at once, in the status bar: whether the row
' just edited will build and, if not, what to fix. Nothing is written to the
' sheet while you type - a macro that writes cells empties Excel's undo, and
' undo matters most while editing. The Check column catches up when you leave
' the sheet, and a build always checks first.
Public Sub LiveCheck(ByVal sh As Object, ByVal target As Range)
    Dim rc As Object, r As Long, nBad As Long, nOn As Long, msg As String
    On Error Resume Next
    If PD_Busy Then Exit Sub
    If StrComp(sh.Name, SH_CONFIG, vbTextCompare) = 0 Then
        r = target.Row
        If r < modPD_Theme.R_FIRST Or target.Column > K_DESC Then Exit Sub
        mDirty = True
        For Each rc In modPD_Recipe.AllRecipes()
            If CLng(rc("Row")) = r Then
                If Not CBool(rc("On")) Then
                    msg = "Row " & r & " is off"
                    If Len(rc("Problem")) > 0 Then msg = msg & ". When switched on it will not build: " & rc("Problem")
                ElseIf Len(rc("Problem")) > 0 Then
                    msg = "Row " & r & " will not build: " & rc("Problem")
                Else
                    msg = "Row " & r & " is ready to build."
                End If
                Exit For
            End If
        Next rc
        If Len(msg) = 0 Then msg = "Row " & r & " is empty."
    ElseIf StrComp(sh.Name, SH_FIELDS, vbTextCompare) = 0 Then
        If target.Row < modPD_Theme.R_FIRST Then Exit Sub
        mDirty = True
        For Each rc In modPD_Recipe.AllRecipes()
            If CBool(rc("On")) Then
                nOn = nOn + 1
                If Len(rc("Problem")) > 0 Then nBad = nBad + 1
            End If
        Next rc
        If nBad = 0 Then
            msg = "Fields changed. Every pivot that is on still builds."
        Else
            msg = "Fields changed. " & nBad & " of " & nOn & " pivot(s) that are on will not build now."
        End If
    Else
        Exit Sub
    End If
    Application.StatusBar = TOOL_NAME & "   " & ChrW(183) & "   Pivot config   " & ChrW(183) & "   " & msg
    Err.Clear
End Sub

' Leaving Pivot config or Pivot fields after an edit: the Check column and the
' Desk are brought up to date.
Public Sub LeaveCheck(ByVal sh As Object)
    Dim ev As Boolean
    On Error Resume Next
    If Not mDirty Then Exit Sub
    If StrComp(sh.Name, SH_CONFIG, vbTextCompare) <> 0 And StrComp(sh.Name, SH_FIELDS, vbTextCompare) <> 0 Then Exit Sub
    mDirty = False
    ev = Application.EnableEvents
    Application.EnableEvents = False
    CheckAll True
    Application.EnableEvents = ev
    Application.StatusBar = False
    modPD_Desk.RefreshDesk
    Err.Clear
End Sub

Public Sub PD_ConfigCheck()
    Dim n As Long
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    n = CheckAll()
    modPD_Desk.RefreshDesk
    If ActiveSheet.Name <> SH_CONFIG Then modPD_Theme.GoTo_ SH_CONFIG
    If n = 0 Then
        Notify "Every pivot that is on is ready to build.", V_OK
    Else
        Notify n & " pivot(s) will not build - see What to fix beside each.", V_BREAK
    End If
End Sub
