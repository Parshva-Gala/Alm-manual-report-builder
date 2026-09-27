Attribute VB_Name = "modScenarioBuilder_Multi"
Option Explicit

' ============================================================================
'  Five extracts, one entity rule, and the filter conditions as the bank
'  maintains them.
'
'  Base and pre-shock used to be rebuilt from two extracts - ECL and CAPRWA -
'  which between them answer credit questions and nothing else. The liquidity
'  and capital test cases had no independent figure at all. Three more extracts
'  now load through the same registry:
'
'      LL    legal liquidity ALM output    assets / liabilities by bucket
'      LCR   LCR ALM output                HQLA / outflow / inflow by bucket
'      CAP   capital component             base capital by element, NOT deal level
'
'  Three things in this module make that work.
'
'  1  THE ENTITY RULE.  Jordan and Cyprus are one legal entity in every one of
'     these files, separated only by branch code 800. Jordan is everything
'     EXCEPT 800; Cyprus is ONLY 800. Getting that backwards, or forgetting it,
'     silently reports the group as Jordan - a figure that looks entirely
'     reasonable and is wrong by the whole of Cyprus. It is applied centrally, to
'     every source, rather than remembered in each test case. CAP also needs
'     its explicit bank and component-code selectors to exclude consolidation
'     totals and the other banks present in the supplied capital extract.
'
'  2  THE ALM AMOUNT MAPPING.  An ALM row carries two amounts, pre-factor and
'     post-factor, and which one a figure wants depends on which side of the
'     ratio it is:
'
'         legal liquidity   assets = PRE-factor      liabilities = POST-factor
'         LCR               HQLA   = PRE-factor      outflow     = POST-factor
'
'     That is a policy, not a fact about the data, so it is seeded as ordinary
'     metric rows the user can edit rather than compiled in. Each row names the
'     column to sum AND the condition that picks its rows, because the two
'     frameworks do not even select on the same column: LCR splits on the
'     segmentation rule category, legal liquidity on the balance sheet category.
'
'  3  THE FILTER CONDITIONS.  The bank maintains these in a scenario rules
'     export - one row per test case and element, with the condition in
'     SCENARIO_ELEMENT_FORMULA_NAME. Importing it is how the conditions get in,
'     rather than being retyped into the case sheet by hand.
' ============================================================================

' Where the entity rule lives, on the Pre_Shock sheet, as editable text.
Public Const ENTITY_SCOPE_LABEL As String = "Entity scope"
Public Const ENTITY_FILTER_LABEL As String = "Entity branch filter"

Public Const SCOPE_JORDAN As String = "Jordan (excludes branch 800)"
Public Const SCOPE_CYPRUS As String = "Cyprus (branch 800 only)"
Public Const SCOPE_BOTH As String = "Both (whole entity)"

' ===================== the entity rule ======================================

' The branch condition for a scope, in the same language as every other filter.
'
' Returned as an expression rather than applied as code so that it can be shown
' to the user, edited, and - most importantly - carried in the aggregate cache
' key, which is keyed on the filter text. A scope applied outside the expression
' would be invisible to that cache and the second test case would silently get
' the first one's scope.
Public Function EntityFilterFor(ByVal scope As String) As String
    Select Case SafeUpperText(scope)
        Case SafeUpperText(SCOPE_JORDAN), "JORDAN", "JO", "JOR"
            EntityFilterFor = "BRANCH_CODE NOT IN ('800')"
        Case SafeUpperText(SCOPE_CYPRUS), "CYPRUS", "CY", "CYP"
            EntityFilterFor = "BRANCH_CODE IN ('800')"
        Case Else
            EntityFilterFor = ""
    End Select
End Function

' The scope the user has chosen, from the Pre_Shock sheet.
Public Function CurrentEntityScope() As String
    Dim s As String
    s = SafeText(ReadSetting(ENTITY_SCOPE_LABEL))
    If Len(s) = 0 Then s = SCOPE_JORDAN
    CurrentEntityScope = s
End Function

' The branch condition actually in force. An override typed on the sheet wins,
' so a bank that reorganises its branches can fix this without a new build.
Public Function CurrentEntityFilter() As String
    Dim override As String
    override = SafeText(ReadSetting(ENTITY_FILTER_LABEL))
    If Len(override) > 0 Then CurrentEntityFilter = override: Exit Function
    CurrentEntityFilter = EntityFilterFor(CurrentEntityScope())
End Function

' A test case's filter with the entity rule ANDed onto it.
'
' Empty case filter means the whole portfolio, so the entity rule alone is the
' answer - not "no filter", which would quietly restore the group.
Public Function ApplyEntityScope(ByVal caseFilter As String) As String
    Dim ent As String, f As String
    ent = Trim$(CurrentEntityFilter())
    f = Trim$(caseFilter)
    If Len(ent) = 0 Then ApplyEntityScope = f: Exit Function
    If Len(f) = 0 Then ApplyEntityScope = ent: Exit Function
    ApplyEntityScope = "( " & f & " ) AND ( " & ent & " )"
End Function

 ' The supplied group extracts contain subsidiaries as well as JKB. Branch
' scope alone cannot identify the bank. Prefer the physical bank identifier;
' ALM group extracts have BANK_CODE instead. Missing identity is a failed check.
Public Function SourceBankFilter(ByVal headers As Object, ByVal sourceKey As String) As String
    If headers Is Nothing Then Err.Raise vbObjectError + 750, "SourceBankFilter", sourceKey & ": source headers are unavailable for bank scoping."
    If headers.Exists("BANK_ID") Then
        SourceBankFilter = "BANK_ID IN ('101')"
    ElseIf headers.Exists("BANK_CODE") Then
        SourceBankFilter = "BANK_CODE IN ('JKB')"
    Else
        Err.Raise vbObjectError + 750, "SourceBankFilter", sourceKey & ": BANK_ID or BANK_CODE is required to independently identify JKB; branch scope alone is insufficient."
    End If
End Function

' ===================== reading and writing the settings =====================

' Held as a defined NAME in the workbook, not as a cell on a sheet.
'
' A cell needs a row somebody has reserved for it, and reserving one on the
' Pre_Shock sheet means every future layout change has to remember it. A name
' persists in the file exactly the same way, survives every rebuild of every
' sheet, and cannot be overwritten by a block of configuration growing past it.
' The first attempt wrote to a labelled row that did not exist, so the setting
' silently never changed and all three scopes returned the same figure.
Private Function SettingName(ByVal label As String) As String
    SettingName = "JKB_" & Replace(Replace(label, " ", "_"), "-", "_")
End Function

Private Function ReadSetting(ByVal label As String) As String
    Dim nm As name, s As String
    On Error Resume Next
    Set nm = ThisWorkbook.names(SettingName(label))
    If nm Is Nothing Then Err.Clear: Exit Function
    s = nm.RefersTo
    Err.Clear
    ' A name holding text refers to ="the text"; strip the = and the quotes.
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    If Left$(s, 1) = """" And Right$(s, 1) = """" And Len(s) >= 2 Then s = Mid$(s, 2, Len(s) - 2)
    ReadSetting = s
End Function

Public Sub SetEntityScope(ByVal scope As String)
    WriteSetting ENTITY_SCOPE_LABEL, scope
    ' The override is cleared: a scope chosen from the list means the list's
    ' condition, and leaving a stale override behind would silently ignore it.
    WriteSetting ENTITY_FILTER_LABEL, ""
End Sub

Private Sub WriteSetting(ByVal label As String, ByVal value As String)
    On Error Resume Next
    ThisWorkbook.names(SettingName(label)).Delete
    Err.Clear
    ThisWorkbook.names.Add name:=SettingName(label), RefersTo:="=""" & Replace(value, """", "'") & """"
    Err.Clear
End Sub

' Cycles Jordan -> Cyprus -> Both, for the button on the Pre_Shock sheet.
Public Sub CycleEntityScope()
    Dim s As String
    s = CurrentEntityScope()
    If s = SCOPE_JORDAN Then
        SetEntityScope SCOPE_CYPRUS
    ElseIf s = SCOPE_CYPRUS Then
        SetEntityScope SCOPE_BOTH
    Else
        SetEntityScope SCOPE_JORDAN
    End If
    UiNotice "Entity scope", "Entity scope is now " & CurrentEntityScope() & _
             IIf(Len(CurrentEntityFilter()) = 0, ", with no branch restriction.", ", " & CurrentEntityFilter() & "."), _
             "Every figure derived from the input files is limited to this scope.", "Run the tests again for it to take effect."
End Sub

' ===================== the metric seeds for the new sources =================

' Liquidity and capital metrics, written onto the metric sheet the same way the
' ECL and CAPRWA ones are: one row each, editable, with the row selector stated.
'
' Columns are  name | source | (enabled) | formula | stage | purpose | where.
'
' The row selector is part of the FORMULA, after WHERE, because that is where the
' engine parses it from - and because it puts the whole policy on one editable
' line. Swapping which side takes the pre-factor amount is a two-word edit here.
Public Function SeedMetricsLiquidity() As Variant
    SeedMetricsLiquidity = Array( _
      Array("LL_ASSETS_PRE_SHOCK", "LL", _
            "SUM(CASHFLOW_AMOUNT_LCY_PRE_FACTOR) WHERE COA_BALANCESHEET_CATEGORY IN ('Assets')", "ALL", _
            "Legal liquidity - assets", "Assets take the PRE-factor amount"), _
      Array("LL_LIABILITIES_PRE_SHOCK", "LL", _
            "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE COA_BALANCESHEET_CATEGORY IN ('Liabilities')", "ALL", _
            "Legal liquidity - liabilities", "Liabilities take the POST-factor amount"), _
      Array("LCR_HQLA_PRE_SHOCK", "LCR", _
            "SUM(CASHFLOW_AMOUNT_LCY_PRE_FACTOR) WHERE ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY IN ('HQLA_LEVEL1', 'HQLA_LEVEL2A', 'HQLA_LEVEL2B')", "ALL", _
            "LCR - high quality liquid assets", "HQLA takes the PRE-factor amount"), _
      Array("LCR_OUTFLOW_PRE_SHOCK", "LCR", _
            "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY IN ('OUTFLOW')", "ALL", _
            "LCR - cash outflow", "Outflow takes the POST-factor amount"), _
      Array("LCR_INFLOW_PRE_SHOCK", "LCR", _
            "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY IN ('INFLOW')", "ALL", _
            "LCR - cash inflow", ""))
End Function

' Capital, from the component file. Base only: this extract is not deal level,
' so nothing here can be filtered to a test case's population - which is exactly
' why RWA_CR is NOT in this list. It stays on CAPRWA, where it is deal level and
' can be.
Public Function SeedMetricsCapital() As Variant
    ' CAPITAL_ELEMENT is a display group and contains RWA, capital and ratio
    ' rows. Select exact component codes, and keep the bank scope explicit.
    ' BANK_ID 101 is JKB in the supplied extract. The formula remains editable
    ' on Pre_Shock_Metrics; branch scope is applied by the ordinary engine.
    ' Blank-bank consolidation rows 901/902 must not be added to branch data.
    SeedMetricsCapital = Array( _
      Array("CET1_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_CET1')", "ALL", _
            "Regulatory common equity tier 1", "JKB bank 101; exact net component; branch scope applies. Edit the bank filter for another bank."), _
      Array("NET_CET1_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_CET1')", "ALL", _
            "Net common equity tier 1", "Exact net regulatory component; excludes consolidation rows"), _
      Array("AT1_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_AT1')", "ALL", _
            "Regulatory additional tier 1", "Exact net regulatory component"), _
      Array("NET_AT1_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_AT1')", "ALL", _
            "Net additional tier 1", "Exact net regulatory component"), _
      Array("T2_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_T2')", "ALL", _
            "Regulatory tier 2", "Exact net regulatory component"), _
      Array("NET_T2_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_T2')", "ALL", _
            "Net tier 2", "Exact net regulatory component"), _
      Array("T2_ECL_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('CALCULATED_STAGE1_ECL')", "ALL", _
            "Tier 2 ECL add-back", "Calculated Stage 1 ECL before the source capital adjustment"), _
      Array("TOTAL_CAPITAL_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('TOTAL_REG_CAPITAL')", "ALL", _
            "Total regulatory capital", "Use TOTAL_REG_CAPITAL only; the CAPITAL group also contains RWA and ratios"))
End Function

' Upgrade known shipped defaults without replacing a user's own definitions.
Public Sub UpgradeCapitalMetricDefaults(ByVal ws As Worksheet)
    Dim seeds As Variant, seed As Variant, r As Long, lastRow As Long
    Dim metricName As String, oldFormula As String, oldElement As String
    If ws Is Nothing Then Exit Sub
    seeds = SeedMetricsCapital()
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = 1 To lastRow
        metricName = SafeUpperText(ws.Cells(r, 1).Value2)
        If SafeUpperText(ws.Cells(r, 2).Value2) = "CAP" Then
            For Each seed In seeds
                If metricName = CStr(seed(0)) Then
                    oldElement = Replace(metricName, "_BASE", "")
                    If metricName = "TOTAL_CAPITAL_BASE" Then oldElement = "CAPITAL"
                    oldFormula = "SUM(REPORT_BALANCE) WHERE CAPITAL_ELEMENT IN ('" & oldElement & "')"
                    If Replace(SafeUpperText(ws.Cells(r, 4).Value2), " ", "") = Replace(oldFormula, " ", "") Then
                        ws.Cells(r, 4).Value2 = seed(2)
                        ws.Cells(r, 6).Value2 = seed(4)
                        ws.Cells(r, 9).Value2 = seed(5)
                    End If
                    Exit For
                End If
            Next seed
        End If
    Next r
End Sub

' The max-ECL columns the new ECL extract carries beyond an ordinary one.
Public Function SeedMetricsMaxEcl() As Variant
    SeedMetricsMaxEcl = Array( _
      Array("MAX_ECL_INPUT", "ECL", "SUM(ORI_MOODY_STRESS_ECL)", "ALL", _
            "Max ECL input, as supplied in the extract", ""), _
      Array("MAX_ECL", "ECL", "SUM(MAX_PRE_FINAL_ECL_MOODY_ECL)", "ALL", _
            "Max ECL", ""), _
      Array("IMPACT_MAX_ECL", "ECL", "SUM(IMPACT_ECL_LCY_MOODY)", "ALL", _
            "Impact of max ECL", ""))
End Function

' ===================== importing the bank's filter conditions ===============

' Reads a scenario rules export and puts its conditions on the case sheet.
'
' The bank maintains these as one row per test case and element:
'
'     SCENARIO_TEST_CASE_CODE | SCENARIO_ELEMENT_CODE | SCENARIO_ELEMENT_FORMULA_NAME
'
' and the last of those IS the filter condition, in the same language the case
' sheet already uses. Matching is on the PAIR, not on the test case alone: one
' test case can have several elements with different conditions, and collapsing
' them onto the test case would give every element the last one's filter.
Public Function ImportScenarioRules(ByVal path As String, ByRef report As String) As Long
    Dim wb As Workbook, ws As Worksheet, h As Object, a As Variant
    Dim r As Long, hdr As Long, n As Long, matched As Long, unmatched As String, nUnmatched As Long
    Dim tc As String, el As String, cond As String
    Dim elementCounts As Object, matchMode As String

    On Error GoTo Failed
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    On Error GoTo CloseAndFail

    For Each ws In wb.Worksheets
        hdr = FindRuleHeaderRow(ws, h)
        If hdr > 0 Then Exit For
    Next ws
    If hdr = 0 Then
        report = "No sheet in that file has the scenario rule columns. Expected a header row with " & _
                 "SCENARIO_TEST_CASE_CODE, SCENARIO_ELEMENT_CODE and SCENARIO_ELEMENT_FORMULA_NAME."
        GoTo CloseAndDone
    End If

    PersistScenarioRules ws, hdr, path
    Set elementCounts = NewMap()
    For r = hdr + 1 To LastUsedRow(ws)
        el = SafeUpperText(ws.Cells(r, h("SCENARIO_ELEMENT_CODE")).Value2)
        If Len(el) > 0 Then
            If Not elementCounts.Exists(el) Then elementCounts(el) = 0
            elementCounts(el) = CLng(elementCounts(el)) + 1
        End If
    Next r
    For r = hdr + 1 To LastUsedRow(ws)
        tc = SafeText(ws.Cells(r, h("SCENARIO_TEST_CASE_CODE")).Value2)
        el = SafeText(ws.Cells(r, h("SCENARIO_ELEMENT_CODE")).Value2)
        cond = SafeText(ws.Cells(r, h("SCENARIO_ELEMENT_FORMULA_NAME")).Value2)
        If Len(tc) = 0 Then GoTo NextRow
        ' A literal "NULL" is how this export writes "no condition". Carrying it
        ' through as text would produce a filter that matches nothing at all.
        If SafeUpperText(cond) = "NULL" Or Len(Trim$(cond)) = 0 Then cond = "ALL"
        n = n + 1
        matchMode = "No matching output element"
        If WriteCaseFilter(tc, el, cond, (CLng(elementCounts(SafeUpperText(el))) = 1), (h("SCENARIO_ELEMENT_CODE") = h("SCENARIO_TEST_CASE_CODE")), matchMode) Then
            matched = matched + 1
        Else
            nUnmatched = nUnmatched + 1
            If nUnmatched <= 12 Then unmatched = unmatched & vbCrLf & "    " & tc & " / " & el
        End If
        ThisWorkbook.Worksheets("Rule_Register").Cells(r - hdr + 5, 9).Value2 = matchMode
NextRow:
    Next r

    report = n & " rule row(s) read; " & matched & " matched a test case on the case sheet."
    If nUnmatched > 0 Then
        report = report & vbCrLf & vbCrLf & nUnmatched & " did not match any test case here:" & unmatched
        If nUnmatched > 12 Then report = report & vbCrLf & "    ... and " & (nUnmatched - 12) & " more."
        report = report & vbCrLf & vbCrLf & "Those are rules for test cases this scenario output does not contain, " & _
                 "which is normal when the rules file covers every framework."
    End If
    ImportScenarioRules = matched

CloseAndDone:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Err.Clear
    Exit Function
CloseAndFail:
    report = "The rules file could not be read: " & Err.description
    On Error Resume Next
    wb.Close SaveChanges:=False
    Err.Clear
    Exit Function
Failed:
    report = "That file could not be opened: " & Err.description
End Function

' The header row of a scenario rules export, wherever it sits.
Private Function FindRuleHeaderRow(ByVal ws As Worksheet, ByRef hOut As Object) As Long
    Dim r As Long, c As Long, h As Object, v As String
    For r = 1 To 12
        Set h = NewMap()
        For c = 1 To 40
            v = NormalHeader(ws.Cells(r, c).Value2)
            If Len(v) > 0 Then If Not h.Exists(v) Then h(v) = c
        Next c
        If h.Exists("SCENARIO_TEST_CASE_CODE") And h.Exists("SCENARIO_ELEMENT_FORMULA_NAME") Then
            ' The element code is optional: a rules file that has only test cases
            ' is still usable, every element of the case then gets its condition.
            If Not h.Exists("SCENARIO_ELEMENT_CODE") Then h("SCENARIO_ELEMENT_CODE") = h("SCENARIO_TEST_CASE_CODE")
            Set hOut = h
            FindRuleHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error Resume Next
    LastUsedRow = ws.UsedRange.row + ws.UsedRange.rows.count - 1
    If LastUsedRow > 100000 Then LastUsedRow = 100000
    Err.Clear
End Function

' Puts one condition onto the case sheet, against the matching test case and
' element. Returns False when this scenario output has no such case.
Private Function WriteCaseFilter(ByVal tc As String, ByVal el As String, ByVal cond As String, _
                                 ByVal allowElementFallback As Boolean, ByVal wholeCase As Boolean, ByRef matchMode As String) As Boolean
    Dim ws As Worksheet, r As Long, rowTc As String, rowEl As String
    Dim exactRows As Collection, fallbackRow As Long, fallbackCount As Long, item As Variant
    Set ws = modScenarioBuilder_PreShock.PsCasesSheet()
    If ws Is Nothing Then Exit Function
    Set exactRows = New Collection
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        rowTc = SafeText(ws.Cells(r, 2).Value2)
        If Len(rowTc) = 0 Then Exit For
        rowEl = SafeText(ws.Cells(r, 3).Value2)
        If StrComp(rowTc, tc, vbTextCompare) = 0 Then
            If StrComp(rowEl, el, vbTextCompare) = 0 Or wholeCase Then exactRows.Add r
        End If
        If StrComp(rowEl, el, vbTextCompare) = 0 Then
            fallbackRow = r: fallbackCount = fallbackCount + 1
        End If
    Next r
    If exactRows.count > 0 Then
        matchMode = "Test case and element"
    ElseIf allowElementFallback And fallbackCount = 1 Then
        exactRows.Add fallbackRow
        matchMode = "Unique element code; rule test case " & tc
    ElseIf fallbackCount > 1 Then
        matchMode = "Ambiguous element code - mapping required"
    End If
    For Each item In exactRows
        r = CLng(item)
        ws.Cells(r, 7).Value2 = cond
        ws.Cells(r, 8).Value2 = cond
        ws.Cells(r, 12).Value2 = "Imported ST rule. " & matchMode & ". Original definition: Rule_Register."
        WriteCaseFilter = True
    Next item
End Function

Public Function RuleImportMatchMode(ByVal testCase As String, ByVal element As String) As String
    Dim ws As Worksheet, h As Object, c As Long, r As Long, tc As String, el As String, fallback As String
    Set ws = GetWorksheetSafe(ThisWorkbook, "Rule_Register")
    If ws Is Nothing Then RuleImportMatchMode = "Rules not imported": Exit Function
    Set h = NewMap()
    For c = 1 To ws.Cells(5, ws.columns.count).End(xlToLeft).Column
        h(NormalHeader(ws.Cells(5, c).Value2)) = c
    Next c
    If Not h.Exists("SCENARIO_TEST_CASE_CODE") Or Not h.Exists("SCENARIO_ELEMENT_CODE") Then
        RuleImportMatchMode = "Rule identity unavailable": Exit Function
    End If
    For r = 6 To LastUsedRow(ws)
        tc = SafeText(ws.Cells(r, h("SCENARIO_TEST_CASE_CODE")).Value2)
        el = SafeText(ws.Cells(r, h("SCENARIO_ELEMENT_CODE")).Value2)
        If StrComp(el, element, vbTextCompare) = 0 Then
            If StrComp(tc, testCase, vbTextCompare) = 0 Then RuleImportMatchMode = SafeText(ws.Cells(r, 9).Value2): Exit Function
            fallback = SafeText(ws.Cells(r, 9).Value2)
        End If
    Next r
    If Len(fallback) > 0 Then RuleImportMatchMode = fallback Else RuleImportMatchMode = "No supplied rule for this element"
End Function

' ===================== proving the mapping against the real files ===========

' Loads the real extracts, aggregates each new metric three ways - Jordan,
' Cyprus and the whole entity - and writes what it found.
'
' The mapping in this module is a claim about somebody else's files: that a
' legal liquidity asset is a pre-factor amount, that LCR splits HQLA from
' outflow on the segmentation rule category, that branch 800 is Cyprus. Every
' one of those is wrong in a way that still produces a plausible number, so the
' build checks them against the files rather than against my reading of them.
'
' spec is the folder holding LL_group, LCR_group, CAP_Component and the ECL
' extract, then ";;" and the rules file.
Public Sub Diag_MultiSource(ByVal outPath As String, ByVal spec As String)
    Dim ff As Integer, parts As Variant, folder As String, rulesFile As String
    Dim t0 As Single, report As String, n As Long

    ff = FreeFile
    Open outPath For Output As #ff
    Print #ff, "=== five sources, one entity rule ==="
    On Error GoTo Fatal

    parts = Split(spec, ";;")
    folder = Trim$(CStr(parts(0)))
    If UBound(parts) >= 1 Then rulesFile = Trim$(CStr(parts(1)))

    modScenarioBuilder_PreShock.EnsurePreShockWorkspace

    t0 = Timer
    LoadOne ff, folder & "\LCR_group.xlsx", "LCR"
    LoadOne ff, folder & "\LL_group.xlsx", "LL"
    LoadOne ff, folder & "\CAP_Component.xlsx", "CAP"
    Print #ff, "  loaded in " & format$(Timer - t0, "0.0") & "s"

    Print #ff, ""
    Print #ff, "-- what the metric sheet actually holds for the new sources"
    Print #ff, modScenarioBuilder_PreShock.DescribeMetricsForDiag("LL|LCR|CAP")
    Print #ff, ""
    Print #ff, "-- one metric, traced end to end"
    Print #ff, modScenarioBuilder_PreShock.TraceMetricForDiag("LCR_OUTFLOW_PRE_SHOCK")

    Print #ff, ""
    Print #ff, "-- metric values by entity scope (base, no test-case filter)"
    Print #ff, String(96, "-")
    Print #ff, "  " & PadR("metric", 28) & PadR("Jordan (not 800)", 22) & PadR("Cyprus (800)", 22) & "Both"
    Print #ff, String(96, "-")
    ReportMetric ff, "LL_ASSETS_PRE_SHOCK"
    ReportMetric ff, "LL_LIABILITIES_PRE_SHOCK"
    ReportMetric ff, "LCR_HQLA_PRE_SHOCK"
    ReportMetric ff, "LCR_OUTFLOW_PRE_SHOCK"
    ReportMetric ff, "LCR_INFLOW_PRE_SHOCK"
    ReportMetric ff, "CET1_BASE"
    ReportMetric ff, "TOTAL_CAPITAL_BASE"

    Print #ff, ""
    Print #ff, "-- the entity rule adds up:  Jordan + Cyprus should equal Both"
    CheckAdditive ff, "LL_ASSETS_PRE_SHOCK"
    CheckAdditive ff, "LCR_OUTFLOW_PRE_SHOCK"
    CheckAdditive ff, "CET1_BASE"

    Print #ff, ""
    Print #ff, "-- the join: an ECL-only filter measured against the ALM extracts"
    Print #ff, "   (REPORT CLASSIFICATION exists only in ECL, so LL and LCR must borrow it)"
    LoadOne ff, folder & "\ECL_Report_JKB_Pre_shock_test_new.xlsx", "ECL"
    JoinCase ff, "REPORT CLASSIFICATION IN ('Corporate Facilities')"
    JoinCase ff, "DATA SOURCE IN ('CORPORATE') AND IFRS STAGE IN ('Stage2 - SICR')"

    Print #ff, ""
    Print #ff, "-- building the joined input file"
    Dim jp As String, t1 As Single
    t1 = Timer
    jp = modScenarioBuilder_Joined.BuildJoinedInputTo(folder & "\joined_out")
    Print #ff, "   took " & format$(Timer - t1, "0.0") & "s"
    If Len(jp) = 0 Then
        Print #ff, "   !!! nothing was written"
    Else
        Print #ff, "   " & jp
        Print #ff, DescribeJoined(jp)
    End If

    If Len(rulesFile) > 0 Then
        Print #ff, ""
        Print #ff, "-- importing the scenario rules"
        n = ImportScenarioRules(rulesFile, report)
        Print #ff, "  " & Replace(report, vbCrLf, vbCrLf & "  ")
    End If

    Print #ff, ""
    Print #ff, "=== complete ==="
    Close #ff
    Exit Sub
Fatal:
    Print #ff, "!!! FATAL " & Err.Number & ": " & Err.description
    On Error Resume Next
    Close #ff
End Sub

Private Sub LoadOne(ByVal ff As Integer, ByVal path As String, ByVal key As String)
    On Error GoTo Failed
    If Len(Dir$(path)) = 0 Then Print #ff, "  " & key & ": NO FILE at " & path: Exit Sub
    modScenarioBuilder_PreShock.LoadEclOutput path, False, key
    Print #ff, "  " & key & ": loaded  " & format$(modScenarioBuilder_PreShock.SourceRowCount(key), "#,##0") & " row(s), family " & _
               modScenarioBuilder_PreShock.SourceFamilyFor(key)
    Exit Sub
Failed:
    Print #ff, "  " & key & ": FAILED - " & Err.description
End Sub

' Just the joined input file, from the real extracts, with nothing else in the
' way. The full diag above re-scans every source twenty-one times to build its
' entity matrix, which is minutes of work that says nothing about this.
Public Sub Diag_JoinedOnly(ByVal outPath As String, ByVal spec As String)
    Dim ff As Integer, parts As Variant, folder As String, rulesFile As String
    Dim t0 As Single, jp As String, report As String

    ff = FreeFile
    Open outPath For Output As #ff
    Print #ff, "=== the joined input file ==="
    On Error GoTo Fatal

    parts = Split(spec, ";;")
    folder = Trim$(CStr(parts(0)))
    If UBound(parts) >= 1 Then rulesFile = Trim$(CStr(parts(1)))

    modScenarioBuilder_PreShock.EnsurePreShockWorkspace
    t0 = Timer
    LoadOne ff, folder & "\ECL_Report_JKB_Pre_shock_test_new.xlsx", "ECL"
    LoadOne ff, folder & "\LCR_group.xlsx", "LCR"
    LoadOne ff, folder & "\LL_group.xlsx", "LL"
    LoadOne ff, folder & "\CAP_Component.xlsx", "CAP"
    Print #ff, "  loaded in " & format$(Timer - t0, "0.0") & "s"

    ' Conditions first, so the test-case columns have something to mark.
    If Len(rulesFile) > 0 Then
        Print #ff, ""
        ImportScenarioRules rulesFile, report
        Print #ff, "  rules: " & Split(report, vbCrLf)(0)
    End If

    Print #ff, ""
    Print #ff, "-- building"
    t0 = Timer
    jp = modScenarioBuilder_Joined.BuildJoinedInputTo(folder & "\joined_out")
    Print #ff, "   took " & format$(Timer - t0, "0.0") & "s"
    If Len(jp) = 0 Then
        Print #ff, "   !!! nothing was written"
    Else
        Print #ff, "   " & jp
        Print #ff, DescribeJoined(jp)
    End If

    Print #ff, "=== complete ==="
    Close #ff
    Exit Sub
Fatal:
    Print #ff, "!!! FATAL " & Err.Number & ": " & Err.description
    On Error Resume Next
    Close #ff
End Sub

' What actually landed in the joined workbook: its sheets, its shape, and the
' first columns of the joined table so the field names can be seen to be intact.
Private Function DescribeJoined(ByVal path As String) As String
    Dim wb As Workbook, ws As Worksheet, s As String, c As Long, n As Long, lastC As Long
    On Error GoTo Done
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error GoTo CloseAndDone
    For Each ws In wb.Worksheets
        n = n + 1
        If n <= 14 Then
            s = s & "      " & PadR(ws.name, 34) & ws.UsedRange.rows.count & " row(s), " & _
                    ws.UsedRange.columns.count & " col(s), " & ws.PivotTables.count & " pivot" & vbCrLf
        End If
    Next ws
    If n > 14 Then s = s & "      ... and " & (n - 14) & " more sheet(s)" & vbCrLf
    Set ws = wb.Worksheets("Joined_Data")
    lastC = ws.UsedRange.columns.count
    s = s & "      field names, first 16 of " & lastC & ":" & vbCrLf & "         "
    For c = 1 To 16
        If c > lastC Then Exit For
        s = s & SafeText(ws.Cells(1, c).Value2) & "  "
    Next c
    s = s & vbCrLf & "      last 6 (the test-case columns):" & vbCrLf & "         "
    For c = lastC - 5 To lastC
        If c >= 1 Then s = s & SafeText(ws.Cells(1, c).Value2) & "  "
    Next c
CloseAndDone:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Err.Clear
Done:
    DescribeJoined = s
End Function

' One ECL-only condition, measured against every source that has to borrow ECL
' to understand it.
Private Sub JoinCase(ByVal ff As Integer, ByVal expr As String)
    Print #ff, ""
    Print #ff, "   filter: " & expr
    Print #ff, "      " & modScenarioBuilder_PreShock.JoinProbeForDiag("ECL", "OUTST_LCY", expr)
    Print #ff, "      " & modScenarioBuilder_PreShock.JoinProbeForDiag("LL", "LL_ASSETS_PRE_SHOCK", expr)
    Print #ff, "      " & modScenarioBuilder_PreShock.JoinProbeForDiag("LL", "LL_LIABILITIES_PRE_SHOCK", expr)
    Print #ff, "      " & modScenarioBuilder_PreShock.JoinProbeForDiag("LCR", "LCR_OUTFLOW_PRE_SHOCK", expr)
End Sub

Private Sub ReportMetric(ByVal ff As Integer, ByVal metric As String)
    Print #ff, "  " & PadR(metric, 28) & _
               PadR(Amt(ValueUnder(SCOPE_JORDAN, metric)), 22) & _
               PadR(Amt(ValueUnder(SCOPE_CYPRUS, metric)), 22) & _
               Amt(ValueUnder(SCOPE_BOTH, metric))
End Sub

Private Sub CheckAdditive(ByVal ff As Integer, ByVal metric As String)
    Dim j As Double, c As Double, b As Double, d As Double
    j = ValueUnder(SCOPE_JORDAN, metric)
    c = ValueUnder(SCOPE_CYPRUS, metric)
    b = ValueUnder(SCOPE_BOTH, metric)
    d = (j + c) - b
    Print #ff, "  " & PadR(metric, 28) & " difference " & Amt(d) & "   " & _
               IIf(Abs(d) < 0.51, "OK", "!!! Jordan and Cyprus do not partition the entity")
End Sub

' One metric, aggregated over the whole population under one entity scope.
Private Function ValueUnder(ByVal scope As String, ByVal metric As String) As Double
    Dim keep As String
    keep = CurrentEntityScope()
    SetEntityScope scope
    ValueUnder = modScenarioBuilder_PreShock.MetricValueForDiag(metric)
    SetEntityScope keep
End Function

Private Function Amt(ByVal v As Double) As String
    Amt = format$(v, "#,##0")
End Function

Private Function PadR(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then PadR = Left$(s, n - 1) & " " Else PadR = s & Space$(n - Len(s))
End Function

' The button on the case sheet.
Public Sub UploadScenarioRules()
    Dim path As String, report As String, n As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Choose the scenario rules file (jkb st rules, or the same export under any name)"
        .filters.Clear
        .filters.Add "Excel files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = False
        If Not .Show = -1 Then Exit Sub
        path = .SelectedItems(1)
    End With
    n = ImportScenarioRules(path, report)
    If n > 0 Then
        UiNotice "Scenario rules", "The filter conditions were imported.", report
    Else
        UiNotice "Scenario rules", "Nothing was imported.", report, , True
    End If
End Sub


Private Sub PersistScenarioRules(ByVal source As Worksheet, ByVal headerRow As Long, ByVal sourcePath As String)
    Dim target As Worksheet, lastRow As Long, lastCol As Long, values As Variant
    Set target = GetWorksheetSafe(ThisWorkbook, "Rule_Register")
    If target Is Nothing Then
        Set target = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        target.name = "Rule_Register"
    End If
    target.Cells.Clear
    target.Range("A1").Value2 = "Imported ST rules"
    target.Range("A2").Value2 = sourcePath
    target.Range("A3").Value2 = "Imported " & format$(Now, "yyyy-mm-dd hh:mm:ss") & ". Original supplied values."
    lastRow = LastUsedRow(source)
    lastCol = source.Cells(headerRow, source.columns.count).End(xlToLeft).Column
    values = source.Range(source.Cells(headerRow, 1), source.Cells(lastRow, lastCol)).Value2
    WriteLiteralValues target.Range("A5").Resize(UBound(values, 1), UBound(values, 2)), values
    target.Cells(5, 9).Value2 = "Import match"
    target.rows(5).Font.Bold = True
    target.rows(5).Interior.Color = RGB(14, 34, 64)
    target.rows(5).Font.Color = RGB(255, 255, 255)
    target.columns.ColumnWidth = 25
    target.columns(lastCol).ColumnWidth = 65
    target.Range(target.Cells(5, 1), target.Cells(4 + UBound(values, 1), lastCol)).AutoFilter
End Sub


Public Function ReconRulesImportRegressionTests(ByVal rulesPath As String) As String
    Dim ws As Worksheet, savedTypes As Object, savedFilters As Object
    Dim r As Long, key As String, matched As Long, report As String, tested As Long
    Set ws = modScenarioBuilder_PreShock.PsCasesSheet()
    Set savedTypes = NewMap(): Set savedFilters = NewMap()
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
        key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
        savedTypes(key) = SafeText(ws.Cells(r, 4).Value2)
    Next r
    matched = ImportScenarioRules(rulesPath, report)
    If matched = 0 Then Err.Raise vbObjectError + 750, , "Rule import regression: " & report
    If GetWorksheetSafe(ThisWorkbook, "Rule_Register") Is Nothing Then Err.Raise vbObjectError + 751, , "Rule register was not retained."
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
        key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
        If SafeText(ws.Cells(r, 4).Value2) <> CStr(savedTypes(key)) Then Err.Raise vbObjectError + 752, , "Rules import changed element type: " & key
        savedFilters(key) = SafeText(ws.Cells(r, 7).Value2)
    Next r
    modScenarioBuilder_PreShock.SyncPreShockTestCasesFromCurrentSource
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
        key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
        If savedFilters.Exists(key) Then
            If SafeText(ws.Cells(r, 7).Value2) <> CStr(savedFilters(key)) Then Err.Raise vbObjectError + 753, , "Synchronization lost the imported rule: " & key
            If Len(CStr(savedFilters(key))) > 0 Then
                If SafeText(ws.Cells(r, 8).Value2) <> CStr(savedFilters(key)) Then Err.Raise vbObjectError + 754, , "Effective filter differs from imported rule: " & key
                tested = tested + 1
            End If
        End If
    Next r
    ReconRulesImportRegressionTests = "PASS: " & tested & " actual imported filters retained with original element types; " & report
End Function
