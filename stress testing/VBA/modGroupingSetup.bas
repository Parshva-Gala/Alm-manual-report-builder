Attribute VB_Name = "modGroupingSetup"
Option Explicit

' ============================================================================
'  Detail grouping setup on Config_ElementTypeRules.
'
'  JKB_ApplyGroupingSetup is a one-time migration for a workbook that still has
'  the old layout (a catch-all OTHER group, and every row's group typed by hand in
'  column U). It:
'
'    - replaces the group list (Y:AE) with one where ECL and RWA each have their
'      own group in every phase - pre-shock, ratios, impact and post-shock - and
'      there is no OTHER group;
'    - adds the Group assignment rules table (tblGroupRules, AG:AK): one line per
'      label pattern instead of one cell per row, first match wins;
'    - clears column U so it becomes an optional per-row override;
'    - adds the "Rows with no group" policy cell, dropdowns and a Check grouping
'      button.
'
'  It is safe to run again: an existing tblGroupRules is kept unless resetRules
'  is True, so rules someone has edited are never overwritten by accident.
'
'  JKB_CheckGrouping is the everyday tool: it lists how many rows each group
'  holds and every row that has no group, without running a generation.
' ============================================================================

Private Const RULES_FIRST_COL As String = "AG"

Public Sub JKB_ApplyGroupingSetup(Optional ByVal quiet As Boolean = False, Optional ByVal resetRules As Boolean = False)
    Dim ws As Worksheet, lo As ListObject, rulesLo As ListObject, r As Long, a() As Variant, i As Long, n As Long
    Dim sh As Shape, keepRules As Boolean, cell As Range
    On Error GoTo Failed
    If Not quiet Then
        If Not UiAsk("Detail grouping", "Set up rule-based detail grouping?", _
                     "ECL and RWA get their own groups in every phase, and OTHER is removed." & vbCrLf & _
                     "A group assignment rules table is added next to the group list." & vbCrLf & _
                     "Column U (DETAIL_GROUP_KEY) is cleared and becomes an optional override for a single row.") Then Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SHEET_RULES)
    Set lo = ws.ListObjects(TABLE_RULES)
    Application.ScreenUpdating = False

    ' --- group list -------------------------------------------------------------
    ws.Range("Y7:AE60").ClearContents
    ws.Range("Y7:AE7").Copy
    ws.Range("Y8:AE" & (GROUP_DEF_FIRST_ROW + 30)).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Range("Y4").Value2 = "The groups shown on every generated sheet, in DISPLAY_ORDER. Rows are placed by the Group assignment rules on the right " & _
                            "(first match wins). Type a group in column U only to override a single row."
    ws.Range("Y5").Value2 = "Rows with no group:"
    ws.Range("Y5").Font.Bold = True
    ws.Range(GROUP_POLICY_CELL).Value2 = GROUP_POLICY_STOP
    ws.Range(GROUP_POLICY_CELL).Resize(1, 2).Interior.Color = UI_WARN_BG
    ws.Range("AC5").Value2 = "Stop lists them before anything is generated. Show puts them under a flagged 'Needs a group' header."
    ws.Range("AC5").Font.Italic = True: ws.Range("AC5").Font.Color = UI_MUTED
    r = GROUP_DEF_FIRST_ROW
    AddGroupRow ws, r, "CONTEXT", "Context & identifiers", 10, "SECTION_CONTEXT", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "SHOCK_INPUT", "Shock & scenario inputs", 20, "SECTION_SHOCK_INPUT", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "PRE_MARKET_ALM", "Pre-shock market / ALM", 30, "SECTION_SHOCK_INPUT", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "PRE_IFRS9", "Pre-shock ECL / IFRS 9", 40, "SECTION_PRE_IFRS9", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "PRE_CAPITAL", "Pre-shock RWA", 50, "SECTION_CAPITAL", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "RATIOS_IFRS9", "Ratios - ECL / IFRS 9", 60, "SECTION_RATIOS", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "RATIOS_RWA", "Ratios - RWA", 65, "SECTION_RATIOS", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "SHOCK_IFRS9", "Impact - ECL / IFRS 9", 70, "SECTION_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "SHOCK_RWA", "Impact - RWA", 72, "SECTION_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "SHOCK_CAPITAL", "Impact - profit & capital", 74, "SECTION_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "SHOCK_LIQUIDITY", "Impact - liquidity", 76, "SECTION_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "SHOCK", "Impact - market & FX", 78, "SECTION_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "POST_IFRS9", "Post-shock ECL / IFRS 9", 80, "SECTION_POST_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "POST_RWA", "Post-shock RWA", 82, "SECTION_POST_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "POST_SHOCK", "Post-shock market & FX", 84, "SECTION_POST_SHOCK", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "CAPITAL", "Profit & capital", 90, "SECTION_CAPITAL", "No", "Yes", 1: r = r + 1
    AddGroupRow ws, r, "LIQUIDITY", "Liquidity results", 100, "SECTION_POST_SHOCK", "No", "Yes", 1: r = r + 1

    ' --- column U becomes an override --------------------------------------------
    If Not lo.DataBodyRange Is Nothing Then lo.ListColumns("DETAIL_GROUP_KEY").DataBodyRange.ClearContents
    Set cell = lo.HeaderRowRange.Cells(1, lo.ListColumns("DETAIL_GROUP_KEY").Index)
    cell.ClearComments
    cell.AddComment "Optional. Leave blank to use the Group assignment rules (AG:AK). Pick a group here only to override this one row."

    ' --- rules table -------------------------------------------------------------
    On Error Resume Next
    Set rulesLo = ws.ListObjects(TABLE_GROUP_RULES)
    On Error GoTo Failed
    keepRules = Not rulesLo Is Nothing And Not resetRules
    If Not keepRules Then
        If Not rulesLo Is Nothing Then rulesLo.Delete
        ws.Range(RULES_FIRST_COL & "3:AK400").Clear
        ReDim a(1 To 49, 1 To 5)
        i = 0
        AddRuleRow a, i, 10, "AS_OF_DATE", "", "CONTEXT", "Identifiers"
        AddRuleRow a, i, 11, "ELEMENT_TYPE", "", "CONTEXT", ""
        AddRuleRow a, i, 12, "ENTITY_*", "", "CONTEXT", ""
        AddRuleRow a, i, 13, "RUN_ID", "", "CONTEXT", ""
        AddRuleRow a, i, 14, "SCENARIO_*", "", "CONTEXT", ""
        AddRuleRow a, i, 15, "SEVERITY_*", "", "CONTEXT", ""
        AddRuleRow a, i, 20, "AMOUNT_CHANGE", "", "SHOCK_INPUT", "Shock parameters"
        AddRuleRow a, i, 21, "NUM_CUSTOMERS", "", "SHOCK_INPUT", ""
        AddRuleRow a, i, 22, "PCT_CHANGE", "", "SHOCK_INPUT", ""
        AddRuleRow a, i, 30, "NSFR_IMPACT_*", "", "SHOCK_LIQUIDITY", "Impact rows first, so they never fall into a results group"
        AddRuleRow a, i, 31, "IMPACT_*RWA*", "", "SHOCK_RWA", ""
        AddRuleRow a, i, 32, "IMPACT_*ECL*", "", "SHOCK_IFRS9", ""
        AddRuleRow a, i, 33, "IMPACT_IIS*", "", "SHOCK_IFRS9", ""
        AddRuleRow a, i, 34, "IMPACT_NPL*", "", "SHOCK_IFRS9", ""
        AddRuleRow a, i, 35, "IMPACT_OUTST_LCY_*", "", "SHOCK_IFRS9", "Outstanding by stage"
        AddRuleRow a, i, 36, "MAX_ECL_*", "", "SHOCK_IFRS9", "Max ECL stress"
        AddRuleRow a, i, 37, "IMPACT_*CAPITAL*", "", "SHOCK_CAPITAL", ""
        AddRuleRow a, i, 38, "IMPACT_*CAR", "", "SHOCK_CAPITAL", ""
        AddRuleRow a, i, 39, "IMPACT_PROFITS_*", "", "SHOCK_CAPITAL", ""
        AddRuleRow a, i, 40, "IMPACT_LCR*", "", "SHOCK_LIQUIDITY", ""
        AddRuleRow a, i, 41, "IMPACT_LEG*", "", "SHOCK_LIQUIDITY", ""
        AddRuleRow a, i, 42, "IMPACT_NSFR*", "", "SHOCK_LIQUIDITY", ""
        AddRuleRow a, i, 43, "IMPACT_*", "", "SHOCK", "Any other impact row"
        AddRuleRow a, i, 50, "*RWA*_PRE_SHOCK", "", "PRE_CAPITAL", "Pre-shock"
        AddRuleRow a, i, 51, "*ECL*_PRE_SHOCK", "", "PRE_IFRS9", ""
        AddRuleRow a, i, 52, "IIS_*_PRE_SHOCK", "", "PRE_IFRS9", ""
        AddRuleRow a, i, 53, "OUTST_LCY_*_PRE_SHOCK", "", "PRE_IFRS9", ""
        AddRuleRow a, i, 54, "FX_*_PRE_SHOCK", "", "PRE_MARKET_ALM", ""
        AddRuleRow a, i, 55, "RSA_LCY", "", "PRE_MARKET_ALM", ""
        AddRuleRow a, i, 56, "RSL_LCY", "", "PRE_MARKET_ALM", ""
        AddRuleRow a, i, 57, "NET_RSA_GAP_LCY", "", "PRE_MARKET_ALM", ""
        AddRuleRow a, i, 60, "*RWA*_RATIO*", "", "RATIOS_RWA", "Ratios"
        AddRuleRow a, i, 61, "ECL_RATIO*", "", "RATIOS_IFRS9", ""
        AddRuleRow a, i, 62, "IIS_RATIO*", "", "RATIOS_IFRS9", ""
        AddRuleRow a, i, 70, "PRE_LCR_*", "", "LIQUIDITY", "Liquidity, pre and post"
        AddRuleRow a, i, 71, "PRE_LEG_*", "", "LIQUIDITY", ""
        AddRuleRow a, i, 72, "PRE_NSFR_*", "", "LIQUIDITY", ""
        AddRuleRow a, i, 73, "LCR*", "", "LIQUIDITY", ""
        AddRuleRow a, i, 74, "LEG*", "", "LIQUIDITY", ""
        AddRuleRow a, i, 75, "NSFR*", "", "LIQUIDITY", ""
        AddRuleRow a, i, 80, "*RWA*", "", "POST_RWA", "Post-shock"
        AddRuleRow a, i, 81, "*ECL*", "", "POST_IFRS9", ""
        AddRuleRow a, i, 82, "*CAPITAL_LCY", "", "CAPITAL", ""
        AddRuleRow a, i, 83, "*CAR", "", "CAPITAL", ""
        AddRuleRow a, i, 84, "PROFITS_*", "", "CAPITAL", ""
        AddRuleRow a, i, 85, "TAX_RATE", "", "CAPITAL", ""
        AddRuleRow a, i, 86, "FX_*", "", "POST_SHOCK", ""
        AddRuleRow a, i, 87, "MARKET_VALUE_LCY", "", "POST_SHOCK", ""
        AddRuleRow a, i, 88, "OUTST_LCY", "", "POST_SHOCK", ""
        ws.Range("AG6:AK6").Value = Array("PRIORITY", "LABEL_PATTERN", "ELEMENT_TYPE", "GROUP_KEY", "NOTES")
        ws.Range("AG7").Resize(i, 5).Value = a
        Set rulesLo = ws.ListObjects.Add(xlSrcRange, ws.Range("AG6").Resize(i + 1, 5), , xlYes)
        rulesLo.Name = TABLE_GROUP_RULES
        rulesLo.TableStyle = lo.TableStyle
    End If
    ws.Range("Y3").Copy: ws.Range("AG3:AK3").PasteSpecial xlPasteFormats
    ws.Range("Y4").Copy: ws.Range("AG4:AK4").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    On Error Resume Next
    ws.Range("AG3:AK3").Merge: ws.Range("AG4:AK4").Merge
    On Error GoTo Failed
    ws.Range("AG3").Value2 = "GROUP ASSIGNMENT RULES"
    ws.Range("AG4").Value2 = "Lowest PRIORITY is tried first and the first match wins. In LABEL_PATTERN, * means any text and ? one character. " & _
                             "Leave ELEMENT_TYPE blank for every element type."
    ws.Range("AG4").WrapText = True
    ws.Columns("AF").ColumnWidth = 3
    ws.Columns("AG").ColumnWidth = 10: ws.Columns("AH").ColumnWidth = 30: ws.Columns("AI").ColumnWidth = 26
    ws.Columns("AJ").ColumnWidth = 20: ws.Columns("AK").ColumnWidth = 44

    ' --- dropdowns ---------------------------------------------------------------
    If Not lo.DataBodyRange Is Nothing Then GroupListValidation lo.ListColumns("DETAIL_GROUP_KEY").DataBodyRange, "=$Y$7:$Y$60", "Pick a group from column Y, or leave blank to use the rules."
    GroupListValidation rulesLo.ListColumns("GROUP_KEY").DataBodyRange, "=$Y$7:$Y$60", "Pick a group from column Y."
    GroupListValidation ws.Range("AB7:AB60"), "=" & SHEET_FORMATTING & "!$Z$48:$Z$60", "Pick a section style from " & SHEET_FORMATTING & "."
    GroupListValidation ws.Range("AC7:AD60"), "Yes,No", "Yes or No."
    GroupListValidation ws.Range("AE7:AE60"), "1,2,3", "Outline level 1, 2 or 3."
    GroupListValidation ws.Range(GROUP_POLICY_CELL), GROUP_POLICY_STOP & "," & GROUP_POLICY_SHOW, "Choose what happens to rows with no group."

    ' --- Check grouping button -----------------------------------------------------
    On Error Resume Next
    ws.Shapes("JKB_CheckGrouping").Delete
    On Error GoTo Failed
    ' Design system v2 secondary button (white, hairline border, navy text).
    If ws.Rows(5).RowHeight < 26 Then ws.Rows(5).RowHeight = 26
    UiButton ws, "JKB_CheckGrouping", "Check grouping", "JKB_CheckGrouping", ws.Range("AK5"), UI_BTN_SECONDARY, 9.5
    Set sh = ws.Shapes("JKB_CheckGrouping")
    sh.Left = ws.Range("AK5").Left + 4: sh.Width = 116: sh.Height = 22
    sh.Top = ws.Range("AK5").Top + (ws.Range("AK5").Height - sh.Height) / 2

    ' --- the neutral section sample no longer reads "Other calculations" -----------
    With ThisWorkbook.Worksheets(SHEET_FORMATTING)
        If .Range("C55").Value2 = "Other calculations" Then .Range("C55").Value2 = "Neutral section"
    End With

    Application.ScreenUpdating = True
    If Not quiet Then JKB_CheckGrouping
    Exit Sub
Failed:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    If quiet Then Err.Raise Err.Number, "JKB_ApplyGroupingSetup", Err.Description
    UiProblem "Detail grouping", "Detail grouping could not be set up.", Err.Description
End Sub

Public Sub JKB_CheckGrouping()
    Dim s As String, problems As Long
    On Error GoTo Failed
    s = GroupingHealthReport(ThisWorkbook, problems)
    UiNotice "Check grouping", IIf(problems = 0, "Every row has a group.", "Some rows need attention."), s, , (problems > 0)
    Exit Sub
Failed:
    UiProblem "Check grouping", "Grouping could not be checked.", Err.Description
End Sub

Private Sub AddGroupRow(ByVal ws As Worksheet, ByVal r As Long, ByVal key As String, ByVal label As String, ByVal order As Long, _
                        ByVal styleKey As String, ByVal collapsed As String, ByVal showHeader As String, ByVal outlineLevel As Long)
    ws.Cells(r, GROUP_DEF_KEY_COL).Resize(1, 7).Value = Array(key, label, order, styleKey, collapsed, showHeader, outlineLevel)
End Sub

Private Sub AddRuleRow(ByRef a() As Variant, ByRef i As Long, ByVal priority As Long, ByVal pattern As String, _
                       ByVal elementType As String, ByVal groupKey As String, ByVal note As String)
    i = i + 1
    a(i, 1) = priority: a(i, 2) = pattern: a(i, 3) = elementType: a(i, 4) = groupKey: a(i, 5) = note
End Sub

Private Sub GroupListValidation(ByVal target As Range, ByVal source As String, ByVal hint As String)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=source
        .IgnoreBlank = True: .InCellDropdown = True
        .ErrorTitle = "Detail grouping": .ErrorMessage = hint
        .ShowError = True: .ShowInput = False
    End With
End Sub
