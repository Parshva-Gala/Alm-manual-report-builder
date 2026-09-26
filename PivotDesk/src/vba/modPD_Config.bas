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
Private Const K_ON As Long = 1
Private Const K_NAME As Long = 2
Private Const K_FW As Long = 3
Private Const K_SPLIT As Long = 4
Private Const K_ROWS As Long = 5
Private Const K_COLS As Long = 6
Private Const K_VALUES As Long = 7
Private Const K_FILTERS As Long = 8
Private Const K_SLICERS As Long = 9
Private Const K_LAYOUT As Long = 10
Private Const K_SUBTOT As Long = 11
Private Const K_GRAND As Long = 12
Private Const K_REPEAT As Long = 13
Private Const K_SORT As Long = 14
Private Const K_WIDTHS As Long = 15
Private Const K_FORMAT As Long = 16
Private Const K_TAB As Long = 17
Private Const K_MAX As Long = 18
Private Const K_DESC As Long = 19
Private Const K_CHECK As Long = 20
Private Const K_WHY As Long = 21
Private Const K_LAST As Long = 21

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
Private Const G_BLANK As Long = 4
Private Const G_WIDTH As Long = 5
Private Const G_FORMAT As Long = 6
Private Const G_NOTE As Long = 7
Private Const G_LAST As Long = 7

' Computed sources: not columns of the file, but worked out while staging.
Public Const SRC_PRE As String = "=PRE"
Public Const SRC_POST As String = "=POST"
Public Const SRC_CCYCLASS As String = "=CCYCLASS"
Public Const SRC_FACTOR As String = "=FACTOR"

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
    Dim ws As Worksheet, fresh As Boolean, ev As Boolean
    ' Writing a few hundred cells here must not run the live check on each.
    ev = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done
    Set ws = GetSheet(SH_CONFIG)
    fresh = ws Is Nothing
    Set ws = EnsureSheet(SH_CONFIG)
    If fresh Or withDefaults Then
        ws.Cells.Clear
        ws.Cells.Validation.Delete
    End If

    modPD_Theme.Dress ws, "Pivot config", _
        "What every framework workbook is built from. One row is one pivot, or one sheet per value of a field. " & _
        "Edit a row or add one; Check says whether it will build. Select any cell for how to fill it.", _
        "REPORTS  " & ChrW(183) & "  PIVOTS"
    Groups ws
    modPD_Theme.Head ws, Array("On", "Pivot", "Frameworks", "One sheet per", "Rows", "Columns", "Values", _
                               "Show only / hide", "Slicers", "Layout", "Subtotals", "Grand totals", _
                               "Repeat labels", "Sort", "Widths", "Number format", "Tab", "Max sheets", _
                               "Description", "Check", "What to fix"), _
                        Array(7, 20, 16, 20, 36, 14, 38, 30, 26, 10, 11, 14, 9, 16, 24, 16, 9, 9, 40, 10, 60)

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
    Dim ws As Worksheet, fresh As Boolean
    Set ws = GetSheet(SH_FIELDS)
    fresh = ws Is Nothing
    Set ws = EnsureSheet(SH_FIELDS)
    If fresh Or withDefaults Then
        ws.Cells.Clear
        ws.Cells.Validation.Delete
    End If
    modPD_Theme.Dress ws, "Pivot fields", _
        "Every column a pivot or a chart may name. The first thirteen are built in; add any column of an " & _
        "output under a name of your own, and use that name anywhere.", _
        "REPORTS  " & ChrW(183) & "  FIELDS"
    modPD_Theme.Head ws, Array("Field", "Source column", "Kind", "Blank shows as", "Width", "Number format", "Note"), _
                        Array(26, 42, 10, 20, 8, 22, 64)
    If fresh Or withDefaults Or Len(SafeText(ws.Cells(modPD_Theme.R_FIRST, G_NAME).Value2)) = 0 Then
        WriteDefaultFields ws
    End If
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
    Band ws, 1, 4, "WHICH PIVOT"
    Band ws, 5, 9, "WHAT IT SHOWS"
    Band ws, 10, 16, "HOW IT LOOKS"
    Band ws, 17, 19, "THE SHEET"
    Band ws, 20, 21, "CHECK"
End Sub

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
Private Function RecipeLastRow(ByVal ws As Worksheet) As Long
    Dim a As Long, b As Long
    a = ws.Cells(ws.Rows.count, K_NAME).End(xlUp).Row
    b = ws.Cells(ws.Rows.count, K_ROWS).End(xlUp).Row
    If b > a Then a = b
    If a < modPD_Theme.R_FIRST + RECIPE_ROWS - 1 Then a = modPD_Theme.R_FIRST + RECIPE_ROWS - 1
    RecipeLastRow = a
End Function

' The cell-by-cell help: select a cell, and Excel shows what goes in it.
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
    Hint ws, r1, r2, K_VALUES, "Values", "One or more, separated by ; - field, then sum / count / average / " & _
         "max / min / %row / %col / %total, then as and a caption. e.g. Pre factor amount sum as Pre-factor"
    Hint ws, r1, r2, K_FILTERS, "Show only / hide", "Separated by ; - Field = a | b shows only a and b; " & _
         "Field <> a hides a. e.g. Bucket <> (no bucket). A field not in rows or columns becomes a report filter."
    Hint ws, r1, r2, K_SLICERS, "Slicers", "Fields to put slicers on, separated by commas. Not added on " & _
         "One-sheet-per pivots: one slicer would filter every one of those sheets at once."
    ListRule ws, r1, r2, K_LAYOUT, "Tabular,Outline,Compact", "Layout", _
         "Tabular: one column per row field. Outline: nested with headers. Compact: all row fields in one column."
    Hint ws, r1, r2, K_SUBTOT, "Subtotals", "None, All, or the row fields to subtotal, separated by commas."
    ListRule ws, r1, r2, K_GRAND, "Both,Bottom row,Right column,None", "Grand totals", _
         "Bottom row totals each column; Right column totals each row."
    ListRule ws, r1, r2, K_REPEAT, "Yes,No", "Repeat labels", _
         "Yes repeats a row label on every line it covers - easier to filter and copy out."
    Hint ws, r1, r2, K_SORT, "Sort", "Blank keeps the data's order. label asc / label desc sorts by the names; " & _
         "a value caption then asc / desc sorts by it - e.g. Pre-factor desc."
    Hint ws, r1, r2, K_WIDTHS, "Widths", "Column widths, separated by ; - field=width, and values=width for " & _
         "the figures. e.g. COA name=44; values=16. Unset fields use the FIELDS list."
    Hint ws, r1, r2, K_FORMAT, "Number format", "Blank uses the desk's own: #,##0;[Red](#,##0);-"
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
         "COUNTERPARTY_NAME. Case and spacing do not matter."
    ListRule ws, r1, r2, G_KIND, "Text,Number,Date", "Kind", _
         "Text is grouped by; Number can be summed; Date is grouped as a date."
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

Private Sub ListRule(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c As Long, _
                     ByVal items As String, ByVal title As String, ByVal msg As String)
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
        .ShowError = True
    End With
    Err.Clear
End Sub

' ===================== defaults =============================================

' The four recipes that make exactly what 1.0 made.
Private Sub WriteDefaultRecipes(ByVal ws As Worksheet)
    Dim r As Long
    r = modPD_Theme.R_FIRST
    ws.Range(ws.Cells(r, 1), ws.Cells(RecipeLastRow(ws), K_LAST)).ClearContents
    Recipe ws, r, "Yes", "{fw} Output", "All", "", "Rule order, Rule category, Rule name, Factor", "LCY / FCY", _
        "Pre factor amount sum as Pre-factor; Post factor amount sum as Post-factor", "Bucket <> (no bucket)", _
        "Rule category, LCY / FCY, Bucket", "Tabular", "None", "Both", "No", "", "", "", "Emerald", "", _
        "Every rule, in the order the engine evaluates them, against what it read and what it kept."
    Recipe ws, r + 1, "Yes", "Balance sheet", "All", "", "Type, Line, Subline, COA name", "LCY / FCY", _
        "Pre factor amount sum as Pre-factor", "Bucket <> (no bucket)", "Type, LCY / FCY, Rule category", _
        "Tabular", "None", "Both", "No", "", "", "", "Emerald", "", _
        "The same balances as the balance sheet reads them, down to the COA. Pre-factor only."
    Recipe ws, r + 2, "Yes", "{split}", "LCR, NSFR", "Rule name, LCY / FCY", "Type, Line, Subline, COA name", _
        "Bucket", "Pre factor amount sum as Pre-factor", "Bucket <> (no bucket)", "", "Tabular", "None", "Both", _
        "No", "", "", "", "Auto", CStr(MAX_SHEETS_DEFAULT), _
        "One rule, one currency side - balances across the maturity buckets."
    Recipe ws, r + 3, "Yes", "{split}", "Maturity ladder", "Currency", "Rule name, Type, Line, Subline, COA name", _
        "Bucket", "Pre factor amount sum as Pre-factor", "Bucket <> (no bucket)", "", "Tabular", "None", "Both", _
        "No", "", "", "", "Auto", CStr(MAX_SHEETS_DEFAULT), _
        "Every rule for one currency, across the maturity buckets."
    ' A worked example that is off: what a recipe of your own looks like.
    Recipe ws, r + 4, "No", "{fw} by counterparty", "All", "", "Counterparty, Product", "Bucket", _
        "Pre factor amount sum as Pre-factor; Pre factor amount %col as Share", "Bucket <> (no bucket)", _
        "Product", "Tabular", "Counterparty", "Both", "Yes", "Pre-factor desc", "Counterparty=34; values=15", _
        "", "Emerald", "", "An example of your own: switch it On to build it."
    ' A few more worth having, ready to switch on.
    Recipe ws, r + 5, "No", "{fw} maturity profile", "All", "", "Product, Cashflow element", "Bucket", _
        "Pre factor amount sum as Pre-factor; Post factor amount sum as Post-factor", "Bucket <> (no bucket)", _
        "Product", "Tabular", "Product", "Both", "No", "Pre-factor desc", "Product=22; Cashflow element=24", _
        "", "Emerald", "", "What each product contributes to each bucket, before and after the factors."
    Recipe ws, r + 6, "No", "{fw} top counterparties", "All", "", "Counterparty", "LCY / FCY", _
        "Pre factor amount sum as Pre-factor; Pre factor amount %total as Share", "", "", "Tabular", "None", _
        "Both", "No", "Pre-factor desc", "Counterparty=40; values=15", "", "Emerald", "", _
        "Counterparties by size, with each one's share of the whole book."
    Recipe ws, r + 7, "No", "{split}", "All", "Product", "Type, Line, COA name", "Bucket", _
        "Pre factor amount sum as Pre-factor", "Bucket <> (no bucket)", "", "Tabular", "None", "Both", "No", _
        "", "", "", "Deep", "20", "One sheet per product: its balances across the maturity buckets."
End Sub

Private Sub Recipe(ByVal ws As Worksheet, ByVal r As Long, ParamArray v() As Variant)
    Dim i As Long
    For i = 0 To UBound(v)
        ws.Cells(r, i + 1).NumberFormat = "@"
        ws.Cells(r, i + 1).Value2 = CStr(v(i))
    Next i
End Sub

Private Sub WriteDefaultFields(ByVal ws As Worksheet)
    Dim r As Long, f As Variant
    r = modPD_Theme.R_FIRST
    ws.Range(ws.Cells(r, 1), ws.Cells(r + 200, G_LAST)).ClearContents
    For Each f In DefaultFields()
        ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST)).NumberFormat = "@"
        ws.Range(ws.Cells(r, 1), ws.Cells(r, G_LAST)).Value2 = f
        r = r + 1
    Next f
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
            AddField d, v(0), v(1), v(2), v(3), v(4), v(5), bi
        Next v
        Exit Function
    End If
    r = modPD_Theme.R_FIRST
    Do While Len(SafeText(ws.Cells(r, G_NAME).Value2)) > 0
        nm = SafeText(ws.Cells(r, G_NAME).Value2)
        If Not d.Exists(nm) Then
            AddField d, nm, SafeText(ws.Cells(r, G_SOURCE).Value2), SafeText(ws.Cells(r, G_KIND).Value2), _
                     SafeText(ws.Cells(r, G_BLANK).Value2), SafeText(ws.Cells(r, G_WIDTH).Value2), _
                     SafeText(ws.Cells(r, G_FORMAT).Value2), bi
        End If
        r = r + 1
    Loop
    ' A built-in that was deleted or renamed on the sheet is still built in.
    For Each v In DefaultFields()
        If bi.Exists(CStr(v(0))) And Not d.Exists(CStr(v(0))) Then AddField d, v(0), v(1), v(2), v(3), v(4), v(5), bi
    Next v
End Function

Private Sub AddField(ByVal d As Object, ByVal nm As String, ByVal src As String, ByVal kind As String, _
                     ByVal blank As String, ByVal w As String, ByVal fmt As String, ByVal bi As Object)
    Dim f As Object
    Set f = NewMap()
    f("Name") = nm
    f("Source") = src
    Select Case LCase$(kind)
        Case "number": f("Kind") = "Number"
        Case "date": f("Kind") = "Date"
        Case Else: f("Kind") = "Text"
    End Select
    f("Blank") = blank
    f("Width") = SafeNum(w)
    f("Format") = fmt
    f("Builtin") = bi.Exists(nm)
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

' Fields named by the recipes that 1.0 did not stage - what the stager has to
' add for this build.
Public Function ExtraFields(ByVal recipes As Collection) As Collection
    Dim out As Collection, seen As Object, fl As Object, rc As Object, nm As Variant
    Set out = New Collection
    Set seen = NewMap()
    Set fl = Fields()
    For Each rc In recipes
        For Each nm In FieldNamesOf(rc)
            If fl.Exists(CStr(nm)) Then
                If Not CBool(fl(CStr(nm))("Builtin")) And Not seen.Exists(CStr(nm)) Then
                    seen(CStr(nm)) = True
                    out.Add fl(CStr(nm))
                End If
            End If
        Next nm
    Next rc
    Set ExtraFields = out
End Function

Private Function FieldNamesOf(ByVal rc As Object) As Collection
    Dim c As Collection, x As Variant, v As Object
    Set c = New Collection
    For Each x In rc("Split"): c.Add x: Next x
    For Each x In rc("Rows"): c.Add x: Next x
    For Each x In rc("Cols"): c.Add x: Next x
    For Each x In rc("Slicers"): c.Add x: Next x
    For Each v In rc("Values"): c.Add v("Field"): Next v
    For Each v In rc("Filters"): c.Add v("Field"): Next v
    Set FieldNamesOf = c
End Function

' Every row of the table, parsed and checked. Rows that are off are parsed
' too, so Check can say what is wrong with them before anyone switches them on.
Public Function AllRecipes() As Collection
    Dim ws As Worksheet, out As Collection, r As Long, lastR As Long, fl As Object
    Dim seenNames As Object, rc As Object
    Set out = New Collection
    Set AllRecipes = out
    Set ws = ConfigSheet()
    If ws Is Nothing Then Exit Function
    Set fl = Fields()
    Set seenNames = NewMap()
    lastR = RecipeLastRow(ws)
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

Private Function NormFw(ByVal s As String) As String
    Select Case UCase$(Replace(Replace(Trim$(s), " ", ""), "_", ""))
        Case "ALL", "": NormFw = "ALL"
        Case "LCR": NormFw = FW_LCR
        Case "NSFR": NormFw = FW_NSFR
        Case "MATURITYLADDER", "LADDER", "ML": NormFw = FW_ML
        Case Else: NormFw = "?" & s
    End Select
End Function

Private Function ParseRow(ByVal ws As Worksheet, ByVal r As Long, ByVal fl As Object) As Object
    Dim rc As Object, s As String, prob As String, x As Variant, subs As Object
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
    Set rc("Values") = ParseValues(Cell(ws, r, K_VALUES), fl, prob)
    If Len(prob) = 0 Then Set rc("Filters") = ParseFilters(Cell(ws, r, K_FILTERS), prob) Else Set rc("Filters") = New Collection
    rc("Desc") = Cell(ws, r, K_DESC)
    rc("Format") = Cell(ws, r, K_FORMAT)
    If Len(rc("Format")) = 0 Then rc("Format") = NUM_FMT

    Select Case LCase$(Cell(ws, r, K_LAYOUT))
        Case "", "tabular": rc("Layout") = 1          ' xlTabularRow
        Case "outline": rc("Layout") = 2              ' xlOutlineRow
        Case "compact": rc("Layout") = 0              ' xlCompactRow
        Case Else: If Len(prob) = 0 Then prob = "Layout must be Tabular, Outline or Compact."
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

    Select Case LCase$(Cell(ws, r, K_GRAND))
        Case "", "both": rc("ColGrand") = True: rc("RowGrand") = True
        Case "bottom row": rc("ColGrand") = True: rc("RowGrand") = False
        Case "right column": rc("ColGrand") = False: rc("RowGrand") = True
        Case "none": rc("ColGrand") = False: rc("RowGrand") = False
        Case Else
            rc("ColGrand") = True: rc("RowGrand") = True
            If Len(prob) = 0 Then prob = "Grand totals must be Both, Bottom row, Right column or None."
    End Select
    rc("Repeat") = (StrComp(Cell(ws, r, K_REPEAT), "Yes", vbTextCompare) = 0)

    ParseSort Cell(ws, r, K_SORT), rc
    Set rc("Widths") = ParseWidths(Cell(ws, r, K_WIDTHS), rc, prob)
    rc("Tab") = Cell(ws, r, K_TAB)
    s = Cell(ws, r, K_MAX)
    If Len(s) = 0 Then
        rc("Max") = MAX_SHEETS_DEFAULT
    ElseIf IsNumeric(s) Then
        rc("Max") = CLng(Val(s))
        If CLng(rc("Max")) < 1 Or CLng(rc("Max")) > 250 Then
            If Len(prob) = 0 Then prob = "Max sheets must be between 1 and 250."
        End If
    Else
        rc("Max") = MAX_SHEETS_DEFAULT
        If Len(prob) = 0 Then prob = "Max sheets must be a number."
    End If

    If Len(prob) = 0 Then prob = Validate(rc, fl)
    rc("Problem") = prob
End Function

Private Function Cell(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Cell = SafeText(ws.Cells(r, c).Value2)
End Function

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

' "Pre factor amount sum as Pre-factor; Pre factor amount %col as Share"
Private Function ParseValues(ByVal s As String, ByVal fl As Object, ByRef prob As String) As Collection
    Dim out As Collection, it As Variant, item As String, cap As String, agg As String, p As Long
    Dim v As Object, seenCap As Object, low As String
    Set out = New Collection
    Set ParseValues = out
    Set seenCap = NewMap()
    For Each it In SplitList(s, ";")
        item = CStr(it)
        cap = ""
        agg = "sum"
        low = LCase$(item)
        p = InStr(1, low, " as ")
        If p > 0 Then
            cap = Trim$(Mid$(item, p + 4))
            item = Trim$(Left$(item, p - 1))
        End If
        p = InStrRev(item, " ")
        If p > 0 Then
            If IsAgg(Mid$(item, p + 1)) Then
                agg = LCase$(Mid$(item, p + 1))
                item = Trim$(Left$(item, p - 1))
            End If
        ElseIf IsAgg(item) Then
            prob = "A value is missing its field: " & Chr$(34) & CStr(it) & Chr$(34) & "."
        End If
        Set v = NewMap()
        v("Field") = item
        v("Agg") = agg
        If Len(cap) = 0 Then cap = AggWord(agg) & " of " & item
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

Private Function IsAgg(ByVal w As String) As Boolean
    Select Case LCase$(Trim$(w))
        Case "sum", "count", "average", "avg", "mean", "max", "min", "%row", "%col", "%column", "%total"
            IsAgg = True
    End Select
End Function

Private Function AggWord(ByVal agg As String) As String
    Select Case agg
        Case "count": AggWord = "Count"
        Case "average", "avg", "mean": AggWord = "Average"
        Case "max": AggWord = "Max"
        Case "min": AggWord = "Min"
        Case "%row": AggWord = "% of row"
        Case "%col", "%column": AggWord = "% of column"
        Case "%total": AggWord = "% of total"
        Case Else: AggWord = "Sum"
    End Select
End Function

' "Bucket <> (no bucket); LCY / FCY = LCY"
Private Function ParseFilters(ByVal s As String, ByRef prob As String) As Collection
    Dim out As Collection, it As Variant, item As String, p As Long, f As Object, incl As Boolean, rest As String
    Set out = New Collection
    Set ParseFilters = out
    For Each it In SplitList(s, ";")
        item = CStr(it)
        p = InStr(item, "<>")
        If p > 0 Then
            incl = False
            rest = Mid$(item, p + 2)
        Else
            p = InStr(item, "=")
            If p = 0 Then
                prob = "A show-only / hide rule needs = or <> : " & Chr$(34) & item & Chr$(34) & "."
                Exit Function
            End If
            incl = True
            rest = Mid$(item, p + 1)
        End If
        Set f = NewMap()
        f("Field") = Trim$(Left$(item, p - 1))
        f("Include") = incl
        Set f("Items") = SplitList(rest, "|")
        If f("Items").count = 0 Then
            prob = "A show-only / hide rule names no items: " & Chr$(34) & item & Chr$(34) & "."
            Exit Function
        End If
        out.Add f
    Next it
End Function

Private Sub ParseSort(ByVal s As String, ByVal rc As Object)
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
Private Function Validate(ByVal rc As Object, ByVal fl As Object) As String
    Dim place As Object, x As Variant, v As Object, t As Variant, nm As String, caps As Object
    If Len(rc("Name")) = 0 Then Validate = "Give the pivot a name - it becomes the sheet's name.": Exit Function
    ' measured with the longest framework label, which {fw} can become
    If Len(Replace(Replace(rc("Name"), "{fw}", "Maturity Ladder"), "{split}", "")) > 31 Then
        Validate = "The pivot name is longer than a sheet name can be (31 characters).": Exit Function
    End If
    For Each t In SplitList(CStr(rc("Frameworks")), ",")
        If Left$(NormFw(CStr(t)), 1) = "?" Then
            Validate = "Frameworks: " & Chr$(34) & CStr(t) & Chr$(34) & " is not All, LCR, NSFR or Maturity ladder."
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
            Validate = Chr$(34) & CStr(x) & Chr$(34) & " is not on the Pivot fields sheet - check the spelling, " & _
                       "or add it there."
            Exit Function
        End If
    Next x
    For Each x In rc("Split")
        If fl(CStr(x))("Kind") <> "Text" Then
            Validate = "One sheet per takes text fields - " & Chr$(34) & CStr(x) & Chr$(34) & " is " & _
                       LCase$(fl(CStr(x))("Kind")) & "."
            Exit Function
        End If
    Next x
    For Each x In rc("Widths").keys
        If Not fl.Exists(CStr(x)) Then
            Validate = "Widths: " & Chr$(34) & CStr(x) & Chr$(34) & " is not a field.": Exit Function
        End If
    Next x
    For Each x In rc("SubFields").keys
        If Not InCollection(rc("Rows"), CStr(x)) Then
            Validate = "Subtotals: " & Chr$(34) & CStr(x) & Chr$(34) & " is not one of the Rows.": Exit Function
        End If
    Next x

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
    For Each v In rc("Filters")
        If place.Exists(v("Field")) Then
            If place(v("Field")) = "One sheet per" Then
                Validate = Chr$(34) & v("Field") & Chr$(34) & " is already one sheet per value; it cannot be " & _
                           "filtered as well."
                Exit Function
            End If
        End If
    Next v

    ' values: numbers can be summed, anything can be counted
    Set caps = NewMap()
    For Each v In rc("Values")
        caps(Trim$(v("Caption"))) = True
        If fl(v("Field"))("Kind") = "Text" And v("Agg") <> "count" Then
            Validate = Chr$(34) & v("Field") & Chr$(34) & " is text, so it can only be counted - write " & _
                       v("Field") & " count."
            Exit Function
        End If
    Next v
    If rc("SortBy") <> "" And rc("SortBy") <> "label" Then
        If Not caps.Exists(rc("SortBy")) Then
            Validate = "Sort: " & Chr$(34) & rc("SortBy") & Chr$(34) & " is not one of the value captions, or label."
            Exit Function
        End If
    End If
    Select Case LCase$(CStr(rc("Tab")))
        Case "", "auto", "emerald", "deep", "slate", "black"
        Case Else: Validate = "Tab must be Auto, Emerald, Deep, Slate or Black.": Exit Function
    End Select
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
    For Each rc In AllRecipes()
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
    Recipe ws, r, "No", "{fw} new pivot " & (r - modPD_Theme.R_FIRST + 1), "All", "", "Type, Line", "Bucket", _
           "Pre factor amount sum as Pre-factor", "Bucket <> (no bucket)", "", "Tabular", "None", "Both", "No", _
           "", "", "", "Auto", "", "A new pivot. Change any cell, then set On to Yes."
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
        For Each rc In AllRecipes()
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
        For Each rc In AllRecipes()
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
