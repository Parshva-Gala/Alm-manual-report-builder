Attribute VB_Name = "modDesignBasePreShock"
Option Explicit

' ============================================================================
'  Base and pre-shock as the ST Design defines them.
'
'  Every value on an element sheet is one of three things:
'    BASE        the whole portfolio from the input files (there is one base);
'    PRE-SHOCK   the same figure filtered to the test case;
'    calculated  shock and post-shock, from base and pre-shock only.
'
'  This module corrects the base and pre-shock definitions that did not match
'  the design or the system's own base row. It only replaces a cell that still
'  holds the value an earlier release seeded, so a row someone has edited is
'  never touched and running it twice changes nothing.
'
'  What it changes:
'    1. Liquidity pre-shock (deposit withdrawal, limit drawdown, liquid asset
'       change, specific customers, deposits without HQLA). The design takes
'       the SEC accounts' own ALM cashflows, at the pre- or post-factor amount
'       depending on the value, with NO category restriction. The old metrics
'       kept the base's category filter (HQLA levels, OUTFLOW, Assets...), so a
'       deposit test case summed to almost nothing. New metrics *_SEC_PRE_FACTOR
'       and *_SEC_POST_FACTOR, and each element type points at the one the
'       design names for it.
'    2. RWA_CR_RATIO_STAGE1: RWA over RWA exposure, no ECL deducted (the
'       the system's base row is RWA / exposure).
'    3. FX_TOTAL_LONG / SHORT_POSITION_LCY base: the long and the short
'       position, not long minus short (the system's base row).
'    4. RWA_MR_FOREX_LCY: the larger of the long and short position, also
'       inside TOTAL_RWA and the CARs (the system's base row equals the long).
'    5. RWA_CR_LCY base: capital component RWA_CREDIT_RISK (before the
'       excess-ECL reduction), which is the system's base to the cent.
'    6. RWA_OR_BASE: component code RWA_OPERATIONAL_RISK. The capital component
'       file has no NET_RWA_OPERATIONAL_RISK row.
'    7. Test-case fields the liquidity and market-risk filters use: BALANCE
'       SHEET TYPE, RULE IDENTIFIER, BALANCESHEET LINE had no column, so every
'       LR / MR case failed; COUNTERPARTY CODE and CURRENCY gain the ALM
'       columns as a fallback only; COUNTERPARTY CLASSIFICATION (ALM) reads
'       the ALM code RETAIL_PRODUCTS as the filters' 'Retail - Products'.
'    8. Liquidity and capital BASE, pinned with the per-bank split of the
'       sample files: LL assets = rule category ASSETS, LL liabilities =
'       LIABILITIES post-factor, LCR HQLA = HQLA levels post-factor (all to
'       the cent); LCR outflow = OUTFLOW 'Upto 1 month' post-factor (0.32%
'       short, open). New *_BASE metrics; the base links point at them.
' ============================================================================

Private Const DBP_VS_FIRST_ROW As Long = 10
Private Const DBP_VS_LABEL_COL As Long = 1
Private Const DBP_VS_BASE_COL As Long = 3
Private Const DBP_VS_PRE_COL As Long = 5
Private Const DBP_VS_NOTES_COL As Long = 10

Private Const DBP_LCR_PRE As String = "derived.LCR_SEC_PRE_FACTOR"
Private Const DBP_LCR_POST As String = "derived.LCR_SEC_POST_FACTOR"
Private Const DBP_LL_PRE As String = "derived.LL_SEC_PRE_FACTOR"
Private Const DBP_LL_POST As String = "derived.LL_SEC_POST_FACTOR"
Private Const DBP_NSFR_PRE As String = "derived.NSFR_SEC_PRE_FACTOR"
Private Const DBP_NSFR_POST As String = "derived.NSFR_SEC_POST_FACTOR"

' The SEC accounts' own cashflows: no category condition, because the test
' case's filter is what selects the rows. Appended to Pre_Shock_Metrics by
' AppendMissingMetrics like every other seed.
'   name | source | formula | stage | purpose | note
Public Function DBP_SeedMetrics() As Variant
    Dim src As Variant, a(0 To 5) As Variant, i As Long, nm As String
    src = Array("LCR", "LL", "NSFR")
    For i = 0 To 2
        nm = CStr(src(i))
        a(2 * i) = Array(nm & "_SEC_PRE_FACTOR", nm, "SUM(CASHFLOW_AMOUNT_LCY_PRE_FACTOR)", "ALL", _
                         nm & " - the test case's accounts, pre-factor amount", _
                         "ST Design liquidity pre-shock: the SEC accounts' cashflows, no category restriction. Base values keep their own category metric.")
        a(2 * i + 1) = Array(nm & "_SEC_POST_FACTOR", nm, "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR)", "ALL", _
                             nm & " - the test case's accounts, post-factor amount", _
                             "ST Design liquidity pre-shock: the SEC accounts' cashflows, no category restriction. Base values keep their own category metric.")
    Next i
    DBP_SeedMetrics = DbpJoin(a, DbpBaseMetrics())
End Function

' The base figures, pinned to the system's base row with the per-bank split of
' the sample files (JKB, Jordan scope; the bank and branch scope are applied by
' the engine as for every other metric). Each one matched to the cent except
' LCR outflow, which is 0.32% short and still open.
Private Function DbpBaseMetrics() As Variant
    Const RC As String = "ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY"
    DbpBaseMetrics = Array( _
      Array("LL_ASSETS_BASE", "LL", "SUM(CASHFLOW_AMOUNT_LCY_PRE_FACTOR) WHERE " & RC & " IN ('ASSETS')", "ALL", _
            "Legal liquidity - total liquid assets (base)", "Rule category ASSETS. Matches the system's base to the cent."), _
      Array("LL_LIABILITIES_BASE", "LL", "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE " & RC & " IN ('LIABILITIES')", "ALL", _
            "Legal liquidity - total liabilities (base)", "Rule category LIABILITIES, post-factor. Matches the system's base to the cent."), _
      Array("LCR_HQLA_BASE", "LCR", "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE " & RC & " IN ('HQLA_LEVEL1', 'HQLA_LEVEL2A', 'HQLA_LEVEL2B')", "ALL", _
            "LCR - HQLA after haircuts (base)", "HQLA levels at the post-factor (haircut) amount. Matches the system's base to the cent."), _
      Array("LCR_OUTFLOW_BASE", "LCR", "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE " & RC & " IN ('OUTFLOW') AND BUCKET_DISPLAY_NAME IN ('Upto 1 month')", "ALL", _
            "LCR - 30-day outflow (base)", "Outflow in the 'Upto 1 month' bucket, post-factor. 0.32% below the system's base: open."), _
      Array("RWA_CR_GROSS_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('RWA_CREDIT_RISK')", "ALL", _
            "Credit risk RWA before the excess-ECL reduction (base)", "Matches the system's RWA_CR_LCY base to the cent."))
End Function

Private Function DbpToArray(ByVal c As Collection) As Variant
    Dim o() As Variant, i As Long
    ReDim o(0 To c.count - 1)
    For i = 1 To c.count: o(i - 1) = c(i): Next i
    DbpToArray = o
End Function

Private Function DbpJoin(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim o() As Variant, i As Long, n As Long
    ReDim o(0 To UBound(a) + UBound(b) + 1)
    For i = LBound(a) To UBound(a): o(n) = a(i): n = n + 1: Next i
    For i = LBound(b) To UBound(b): o(n) = b(i): n = n + 1: Next i
    DbpJoin = o
End Function

' Run by JKB_ApplyValueSourceSetup after the link table is seeded. Returns how
' many cells it changed, and logs the count.
Public Function DBP_ApplyDesignSetup() As Long
    Dim n As Long
    On Error GoTo Failed
    n = n + UpgradeMetrics()
    n = n + UpgradeFields()
    n = n + UpgradeLinks()
    n = n + UpgradeRules()
    LogIssue LOG_LEVEL_INFO, "Design base and pre-shock", n & " seeded definition(s) brought in line with the ST Design.", "Config_ValueSources"
    DBP_ApplyDesignSetup = n
    Exit Function
Failed:
    LogIssue LOG_LEVEL_WARN, "Design base and pre-shock", "Stopped after " & n & " change(s): " & Err.description, "Config_ValueSources"
    DBP_ApplyDesignSetup = n
End Function

' ------------------------------------------------------------ metrics ----

Private Function UpgradeMetrics() As Long
    Dim ws As Worksheet, r As Long
    Set ws = PsMetricsSheet()
    If ws Is Nothing Then Exit Function
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        If SafeUpperText(ws.Cells(r, 1).Value2) = "RWA_OR_BASE" Then
            If DbpSquash(ws.Cells(r, 4).Value2) = DbpSquash("SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_RWA_OPERATIONAL_RISK')") Then
                ws.Cells(r, 4).Value2 = "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('RWA_OPERATIONAL_RISK')"
                ws.Cells(r, 9).Value2 = "The capital component file has RWA_OPERATIONAL_RISK (there is no NET_ row)."
                UpgradeMetrics = UpgradeMetrics + 1
            End If
        End If
    Next r
End Function

' ------------------------------------------------------------- fields ----

' A field row is replaced only while it still holds the candidates an earlier
' release seeded. "a|b||c" means: a or b, and c only where neither exists, so
' a source that already resolves the field keeps exactly the columns it had.
Private Function UpgradeFields() As Long
    Dim ws As Worksheet, r As Long, specs As Variant, s As Variant, nm As String
    Set ws = PsFieldsSheet()
    If ws Is Nothing Then Exit Function
    specs = DbpFieldSpecs()
    For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        nm = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(nm) = 0 Then Exit For
        For Each s In specs
            If nm = CStr(s(0)) Then
                If DbpSquash(ws.Cells(r, 2).Value2) = DbpSquash(s(1)) And _
                   (DbpSquash(ws.Cells(r, 2).Value2) <> DbpSquash(s(2)) Or SafeUpperText(ws.Cells(r, 3).Value2) <> CStr(s(3))) Then
                    ws.Cells(r, 2).Value2 = s(2)
                    ws.Cells(r, 3).Value2 = s(3)
                    ws.Cells(r, 6).Value2 = s(4)
                    UpgradeFields = UpgradeFields + 1
                End If
            End If
        Next s
    Next r
End Function

'   field | old candidates | new candidates | mode | note
Public Function DbpFieldSpecs() As Variant
    DbpFieldSpecs = Array( _
      Array("BALANCE SHEET TYPE", "", "COA_BALANCESHEET_CATEGORY", "TEXT", _
            "ALM balance-sheet category (Assets, Liabilities, Off Balance Sheet...). Only the ALM files carry it."), _
      Array("RULE IDENTIFIER", "", "RULE_IDENTIFIER_CODE", "RULEID", _
            "ALM rule identifier. RULEID reads STABLE_DEPOSITS, STABLE DEPOSITS and STABLE as one value; the LCR file writes STABLE."), _
      Array("BALANCESHEET LINE", "", "BALANCESHEET_LINE_NAME|BALANCESHEET_SUB_LINE_NAME", "BSLINE", _
            "ALM balance-sheet line or sub-line. BSLINE drops the code prefix (2.01.00_) and reads Customer Deposits as deposits and Cash on Hands and in Central Banks as cash and central bank balances (default until the bank confirms)."), _
      Array("BALANCESHEET LINE", "BALANCESHEET_LINE_NAME|BALANCESHEET_SUB_LINE_NAME", "BALANCESHEET_LINE_NAME|BALANCESHEET_SUB_LINE_NAME", "BSLINE", _
            "ALM balance-sheet line or sub-line. BSLINE drops the code prefix (2.01.00_) and reads Customer Deposits as deposits and Cash on Hands and in Central Banks as cash and central bank balances (default until the bank confirms)."), _
      Array("COUNTERPARTY CODE", "CUSTOMER_CODE|COUNTERPARTY_ID|CUSTOMER_ID", "CUSTOMER_CODE|COUNTERPARTY_ID|CUSTOMER_ID||COUNTERPARTY_CODE", "TEXT", _
            "Customer / counterparty identifier. COUNTERPARTY_CODE (ALM files) is used only where none of the others exists."), _
      Array("CURRENCY", "ISO_CODE|CURRENCY_CODE", "ISO_CODE|CURRENCY_CODE||CURRENCY_NAME", "CURRENCY", _
            "Currency. CURRENCY_NAME (ALM files: Jordan Dinar, US-Dollar...) is used only where no code column exists, and reads as the ISO code."), _
      Array("COUNTERPARTY CLASSIFICATION (ALM)", "COUNTERPARTY_CLASSIFICATION_NAME|COUNTERPARTY_CLASSIFICATION_CODE", _
            "COUNTERPARTY_CLASSIFICATION_NAME|COUNTERPARTY_CLASSIFICATION_CODE", "LOOSE_TEXT", _
            "The ALM files carry only the code (RETAIL_PRODUCTS); LOOSE_TEXT reads it as the filters' 'Retail - Products'."))
End Function

' -------------------------------------------------------------- links ----

Private Function UpgradeLinks() As Long
    Dim ws As Worksheet, r As Long, lastRow As Long, specs As Variant, s As Variant, lbl As String, col As Long
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_SHEET)
    If ws Is Nothing Then Exit Function
    specs = DbpLinkSpecs()
    lastRow = ws.Cells(ws.rows.count, DBP_VS_LABEL_COL).End(xlUp).row
    For r = DBP_VS_FIRST_ROW To lastRow
        lbl = SafeUpperText(ws.Cells(r, DBP_VS_LABEL_COL).Value2)
        If Len(lbl) > 0 Then
            For Each s In specs
                If lbl = CStr(s(0)) Then
                    col = IIf(CStr(s(1)) = "BASE", DBP_VS_BASE_COL, DBP_VS_PRE_COL)
                    If DbpSquash(ws.Cells(r, col).Value2) = DbpSquash(s(2)) Then
                        ws.Cells(r, col).Value2 = s(3)
                        If Len(CStr(s(4))) > 0 Then ws.Cells(r, DBP_VS_NOTES_COL).Value2 = s(4)
                        UpgradeLinks = UpgradeLinks + 1
                    End If
                End If
            Next s
        End If
    Next r
End Function

'   label | BASE or PRE | old link | new link | note
Public Function DbpLinkSpecs() As Variant
    Dim fxOld As String, fxNew As String, rwaSum As String, liq As String, crNote As String, lcrOld As String, lcrNew As String
    crNote = "Capital component RWA_CREDIT_RISK, bank 101: the system's base to the cent. TOTAL_RWA keeps NET_RWA_CREDIT_RISK, which is what makes it match."
    lcrOld = "derivedbase.LCR_HQLA_PRE_SHOCK/(derivedbase.LCR_OUTFLOW_PRE_SHOCK-MIN(derivedbase.LCR_INFLOW_PRE_SHOCK,0.75*derivedbase.LCR_OUTFLOW_PRE_SHOCK))"
    lcrNew = "derivedbase.LCR_HQLA_BASE/(derivedbase.LCR_OUTFLOW_BASE-MIN(derivedbase.LCR_INFLOW_PRE_SHOCK,0.75*derivedbase.LCR_OUTFLOW_BASE))"
    fxOld = "figure.RWA_MR_FOREX"
    fxNew = "MAX(figure.FX_LONG_POSITION,figure.FX_SHORT_POSITION)"
    rwaSum = "derivedbase.RWA_CR_BASE+figure.RWA_MR_EQUITY+" & fxOld & "+derivedbase.RWA_OR_BASE"
    liq = "ST Design: the test case's own ALM cashflows (the SEC accounts), no category restriction. " & _
          "Each element type picks pre- or post-factor on Config_ElementTypeRules; this link is the usual one."
    Dim c As New Collection
    ' One statement per row, so no statement comes near VBA's limit of 24 line continuations.
    c.Add Array("RWA_CR_RATIO_STAGE1", "BASE", _
            "derivedbase.RWA_CR_STAGE1_LCY_PRE_SHOCK/(derivedbase.OUTST_RWA_STAGE1_LCY_PRE_SHOCK-derivedbase.ECL_STAGE1_LCY_PRE_SHOCK)", _
            "derivedbase.RWA_CR_STAGE1_LCY_PRE_SHOCK/derivedbase.OUTST_RWA_STAGE1_LCY_PRE_SHOCK", _
            "Stage 1 RWA over stage 1 RWA exposure, no ECL deducted: the system's base row (RWA / exposure). Stage 2 and NPA do deduct ECL.")
    c.Add Array("RWA_CR_RATIO_STAGE1", "PRE", _
            "derived.RWA_CR_STAGE1_LCY_PRE_SHOCK/(derived.OUTST_RWA_STAGE1_LCY_PRE_SHOCK-derived.ECL_STAGE1_LCY_PRE_SHOCK)", _
            "derived.RWA_CR_STAGE1_LCY_PRE_SHOCK/derived.OUTST_RWA_STAGE1_LCY_PRE_SHOCK", "")
    c.Add Array("FX_TOTAL_LONG_POSITION_LCY", "BASE", "figure.FX_LONG_POSITION-figure.FX_SHORT_POSITION", "figure.FX_LONG_POSITION", _
            "The long position (cap_market_risk_forex_total.total_long_position): the system's base row. Long minus short is IMPACT_NET_OPEN_POSITION's base.")
    c.Add Array("FX_TOTAL_SHORT_POSITION_LCY", "BASE", "figure.FX_LONG_POSITION-figure.FX_SHORT_POSITION", "figure.FX_SHORT_POSITION", _
            "The short position (cap_market_risk_forex_total.total_short_position): the system's base row.")
    c.Add Array("RWA_MR_FOREX_LCY", "BASE", fxOld, fxNew, _
            "Forex market risk RWA = the larger of the long and short position: the system's base row equals the long position.")
    c.Add Array("RWA_CR_LCY", "BASE", "derivedbase.RWA_CR_BASE", "derivedbase.RWA_CR_GROSS_BASE", crNote)
    c.Add Array("RWA_CR_LCY", "BASE", "derivedbase.RWA_CR_LCY", "derivedbase.RWA_CR_GROSS_BASE", crNote)
    c.Add Array("TOTAL_RWA", "BASE", rwaSum, Replace(rwaSum, fxOld, fxNew), "")
    c.Add Array("REGULATORY_CAR", "BASE", "derivedbase.TOTAL_CAPITAL_BASE/(" & rwaSum & ")", "derivedbase.TOTAL_CAPITAL_BASE/(" & Replace(rwaSum, fxOld, fxNew) & ")", "")
    c.Add Array("CET1_CAR", "BASE", "derivedbase.CET1_BASE/(" & rwaSum & ")", "derivedbase.CET1_BASE/(" & Replace(rwaSum, fxOld, fxNew) & ")", "")
    c.Add Array("PRE_LEG_LIQ_TOTAL_ASSETS", "BASE", "derivedbase.LL_ASSETS_PRE_SHOCK", "derivedbase.LL_ASSETS_BASE", "")
    c.Add Array("LEG_LIQ_TOTAL_ASSETS", "BASE", "derivedbase.LL_ASSETS_PRE_SHOCK", "derivedbase.LL_ASSETS_BASE", "Rule category ASSETS: the system's base to the cent.")
    c.Add Array("PRE_LEG_LIQ_TOTAL_LIAB", "BASE", "derivedbase.LL_LIABILITIES_PRE_SHOCK", "derivedbase.LL_LIABILITIES_BASE", "")
    c.Add Array("LEG_LIQ_TOTAL_LIAB", "BASE", "derivedbase.LL_LIABILITIES_PRE_SHOCK", "derivedbase.LL_LIABILITIES_BASE", "Rule category LIABILITIES, post-factor: the system's base to the cent.")
    c.Add Array("LEGAL_LIQUIDITY_RATIO", "BASE", "derivedbase.LL_ASSETS_PRE_SHOCK/derivedbase.LL_LIABILITIES_PRE_SHOCK", "derivedbase.LL_ASSETS_BASE/derivedbase.LL_LIABILITIES_BASE", "")
    c.Add Array("PRE_LCR_HQLA", "BASE", "derivedbase.LCR_HQLA_PRE_SHOCK", "derivedbase.LCR_HQLA_BASE", "")
    c.Add Array("LCR_HQLA", "BASE", "derivedbase.LCR_HQLA_PRE_SHOCK", "derivedbase.LCR_HQLA_BASE", "HQLA levels, post-factor (after haircuts): the system's base to the cent.")
    c.Add Array("PRE_LCR_OUTFLOW", "BASE", "derivedbase.LCR_OUTFLOW_PRE_SHOCK", "derivedbase.LCR_OUTFLOW_BASE", "")
    c.Add Array("LCR_OUTFLOW", "BASE", "derivedbase.LCR_OUTFLOW_PRE_SHOCK", "derivedbase.LCR_OUTFLOW_BASE", "Outflow, 'Upto 1 month' bucket, post-factor: 0.32% below the system's base (open).")
    c.Add Array("LCR", "BASE", lcrOld, lcrNew, "HQLA / (outflow - min(inflow, 75% of outflow)). Inflow's base is still open.")
    c.Add Array("PRE_LCR_HQLA", "PRE", "derived.LCR_HQLA_PRE_SHOCK", DBP_LCR_PRE, liq)
    c.Add Array("PRE_LCR_OUTFLOW", "PRE", "derived.LCR_OUTFLOW_PRE_SHOCK", DBP_LCR_POST, liq)
    c.Add Array("PRE_LCR_INFLOW", "PRE", "derived.LCR_INFLOW_PRE_SHOCK", DBP_LCR_POST, liq)
    c.Add Array("PRE_LEG_LIQ_TOTAL_ASSETS", "PRE", "derived.LL_ASSETS_PRE_SHOCK", DBP_LL_PRE, liq)
    c.Add Array("PRE_LEG_LIQ_TOTAL_LIAB", "PRE", "derived.LL_LIABILITIES_PRE_SHOCK", DBP_LL_POST, liq)
    c.Add Array("PRE_NSFR_ASF", "PRE", "derived.NSFR_ASF_PRE_SHOCK", DBP_NSFR_POST, liq)
    c.Add Array("PRE_NSFR_RSF", "PRE", "derived.NSFR_RSF_PRE_SHOCK", DBP_NSFR_PRE, liq)
    DbpLinkSpecs = DbpToArray(c)
End Function

' -------------------------------------------------------- config rules ----

' The pre-shock formula per liquidity element type, from the ST Design's
' liquidity columns (Q, R, S, T, Z). Replaced only while the cell still holds
' the old seeded metric, is blank, or only copies the system value.
Private Function UpgradeRules() As Long
    Dim lo As ListObject, ws As Worksheet, cType As Long, cLabel As Long, cFormula As Long
    Dim r As Long, et As String, lbl As String, want As String, cur As String, i As Long
    Set ws = GetWorksheetSafe(ThisWorkbook, SHEET_RULES)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_RULES)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    cType = DbpListCol(lo, "ELEMENT_TYPE")
    cLabel = DbpListCol(lo, "OUTPUT_ROW_LABEL")
    cFormula = DbpListCol(lo, "MANUAL_FORMULA_DEFAULT")
    If cType = 0 Or cLabel = 0 Or cFormula = 0 Then Exit Function
    For r = 1 To lo.DataBodyRange.rows.count
        et = SafeUpperText(lo.DataBodyRange.Cells(r, cType).Value2)
        lbl = SafeUpperText(lo.DataBodyRange.Cells(r, cLabel).Value2)
        want = DbpRuleFormula(et, lbl)
        If Len(want) > 0 Then
            cur = DbpSquash(lo.DataBodyRange.Cells(r, cFormula).Value2)
            If cur <> DbpSquash(want) Then
                If Len(cur) = 0 Or cur = DbpSquash(DbpOldRuleFormula(lbl)) Or cur = "SYSTEMOUTPUT." & lbl Or cur = "SYSTEM." & lbl Then
                    lo.DataBodyRange.Cells(r, cFormula).Value2 = want
                    UpgradeRules = UpgradeRules + 1
                End If
            End If
        End If
    Next r
End Function

Private Function DbpListCol(ByVal lo As ListObject, ByVal header As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.count
        If SafeUpperText(lo.ListColumns(i).name) = header Then DbpListCol = i: Exit Function
    Next i
End Function

' What the earlier release seeded for each liquidity pre-shock row.
Private Function DbpOldRuleFormula(ByVal lbl As String) As String
    Select Case lbl
        Case "PRE_LCR_HQLA": DbpOldRuleFormula = "derived.LCR_HQLA_PRE_SHOCK"
        Case "PRE_LCR_OUTFLOW": DbpOldRuleFormula = "derived.LCR_OUTFLOW_PRE_SHOCK"
        Case "PRE_LCR_INFLOW": DbpOldRuleFormula = "derived.LCR_INFLOW_PRE_SHOCK"
        Case "PRE_LEG_LIQ_TOTAL_ASSETS": DbpOldRuleFormula = "derived.LL_ASSETS_PRE_SHOCK"
        Case "PRE_LEG_LIQ_TOTAL_LIAB": DbpOldRuleFormula = "derived.LL_LIABILITIES_PRE_SHOCK"
        Case "PRE_NSFR_ASF": DbpOldRuleFormula = "derived.NSFR_ASF_PRE_SHOCK"
        Case "PRE_NSFR_RSF": DbpOldRuleFormula = "derived.NSFR_RSF_PRE_SHOCK"
    End Select
End Function

' The ST Design, liquidity columns. "" leaves the row alone: the design sets
' its impact to zero, so its pre-shock does not enter any figure.
Public Function DbpRuleFormula(ByVal elementType As String, ByVal lbl As String) As String
    Select Case SafeUpperText(elementType)
        Case "SPECIFIED PERCENTAGE OF DEPOSITS ARE WITHDRAWN"          ' Q
            Select Case lbl
                Case "PRE_LCR_HQLA": DbpRuleFormula = DBP_LCR_PRE
                Case "PRE_LCR_OUTFLOW": DbpRuleFormula = DBP_LCR_POST
                Case "PRE_LEG_LIQ_TOTAL_ASSETS": DbpRuleFormula = DBP_LL_PRE
                Case "PRE_LEG_LIQ_TOTAL_LIAB": DbpRuleFormula = DBP_LL_POST
                Case "PRE_NSFR_ASF": DbpRuleFormula = DBP_NSFR_POST
            End Select
        Case "LIMIT DRAWDOWN"                                          ' R
            Select Case lbl
                Case "PRE_LCR_HQLA": DbpRuleFormula = DBP_LCR_PRE
                Case "PRE_LCR_INFLOW": DbpRuleFormula = DBP_LCR_POST
                Case "PRE_LEG_LIQ_TOTAL_ASSETS": DbpRuleFormula = DBP_LL_PRE
                Case "PRE_LEG_LIQ_TOTAL_LIAB": DbpRuleFormula = DBP_LL_POST
            End Select
        Case "CHANGE IN VALUE OF LIQUID ASSETS"                        ' S
            Select Case lbl
                Case "PRE_LCR_HQLA": DbpRuleFormula = DBP_LCR_POST
                Case "PRE_LEG_LIQ_TOTAL_ASSETS": DbpRuleFormula = DBP_LL_POST
            End Select
        Case "SPECIFIC CUSTOMERS WITHDRAW DEPOSITS"                    ' T
            Select Case lbl
                Case "PRE_LCR_HQLA": DbpRuleFormula = DBP_LCR_PRE
                Case "PRE_LCR_OUTFLOW": DbpRuleFormula = DBP_LCR_POST
                Case "PRE_LEG_LIQ_TOTAL_ASSETS": DbpRuleFormula = DBP_LL_PRE
                Case "PRE_LEG_LIQ_TOTAL_LIAB": DbpRuleFormula = DBP_LL_POST
                Case "PRE_NSFR_ASF": DbpRuleFormula = DBP_NSFR_POST
            End Select
        Case "DEPOSIT INCREASE OR WITHDRAWN WITHOUT CHANGE IN LIQUID ASSETS"   ' Z
            Select Case lbl
                Case "PRE_LCR_OUTFLOW": DbpRuleFormula = DBP_LCR_POST
                Case "PRE_LEG_LIQ_TOTAL_LIAB": DbpRuleFormula = DBP_LL_POST
                Case "PRE_NSFR_ASF": DbpRuleFormula = DBP_NSFR_POST
                Case "PRE_NSFR_RSF": DbpRuleFormula = DBP_NSFR_PRE
            End Select
    End Select
End Function

' ------------------------------------------------------------- helpers ----

Private Function DbpSquash(ByVal v As Variant) As String
    Dim s As String
    s = SafeUpperText(NormalizeConfigFormulaText(v))
    s = Replace(Replace(Replace(s, "@", ""), " ", ""), vbTab, "")
    DbpSquash = s
End Function
