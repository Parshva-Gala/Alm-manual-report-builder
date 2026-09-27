Attribute VB_Name = "modReconSetup"
Option Explicit

Public Sub PrepareReconRelease(ByVal rulePath As String)
    Dim report As String, n As Long, stage As String, errorText As String
    On Error GoTo Failed
    stage = "confirmed corrections"
    ApplyConfirmedCorrections
    stage = "shock catalog"
    InstallShockCatalog
    stage = "workbook interface"
    JKB_UpgradeWorkbook
    stage = "rules import"
    n = ImportScenarioRules(rulePath, report)
    HomeSetActivity "Ready. " & n & " rule conditions imported. Next: load the system output and its input files."
    OpenReconWorkbench
    Exit Sub
Failed:
    errorText = stage & ": " & Err.description
    Err.Raise vbObjectError + 1731, "PrepareReconRelease", errorText
End Sub

' The catalog rows live in shock_catalog.tsv beside the workbook: one row per line,
' seven tab-separated columns, with tab, line feed, carriage return and backslash
' written as \t \n \r \\. They used to be 8,400 lines of cell assignments in this
' module. The shipped workbook's Shock_Catalog sheet already holds them, so when the
' file is not there the sheet is left exactly as it is.
Public Sub InstallShockCatalog()
    Dim ws As Worksheet, path As String, rowsIn As Collection, data() As Variant
    Dim f As Integer, ln As String, parts As Variant, r As Long, c As Long, lastRow As Long
    Dim er As String, errNo As Long
    On Error GoTo Failed
    path = ShockCatalogPath()
    If Len(path) = 0 Then
        If Not GetWorksheetSafe(ThisWorkbook, "Shock_Catalog") Is Nothing Then Exit Sub
        Err.Raise vbObjectError + 1732, "InstallShockCatalog", "shock_catalog.tsv is not beside the workbook, and there is no Shock_Catalog sheet to keep."
    End If
    Set rowsIn = New Collection
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        If Len(ln) > 0 Then rowsIn.Add ln
    Loop
    Close #f: f = 0
    If rowsIn.count = 0 Then Err.Raise vbObjectError + 1733, "InstallShockCatalog", path & " has no rows."
    ReDim data(1 To rowsIn.count, 1 To 7)
    For r = 1 To rowsIn.count
        parts = Split(rowsIn(r), vbTab)
        If UBound(parts) <> 6 Then Err.Raise vbObjectError + 1733, "InstallShockCatalog", "Line " & r & " of " & path & " does not have 7 columns."
        For c = 1 To 7
            data(r, c) = UnescapeCatalogField(CStr(parts(c - 1)))
        Next c
    Next r
    lastRow = 7 + rowsIn.count

    Set ws = GetWorksheetSafe(ThisWorkbook, "Shock_Catalog")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)): ws.name = "Shock_Catalog"
    ws.Cells.Clear
    ws.columns("A:G").NumberFormat = "@"
    ws.Range("A1").Value2 = "JKB"
    ws.Range("A3").Value2 = "Shock field guide"
    ws.Range("A4").Value2 = "BRD and design definitions beside configured formulas. ALIGNED/CORRECTED applies only to the exact formula shown; data and dependencies are checked at run time."
    ws.Range("A7:G7").value = Array("Shock type", "Field", "Design definition", "Source reference", "Review status", "Reviewed formula", "Review note")
    ws.Range("A8").Resize(rowsIn.count, 7).Value2 = data
    StyleReconTable ws, lastRow, 7
    ws.columns("A:B").ColumnWidth = 40
    ws.columns("C:D").ColumnWidth = 65
    ws.columns("E").ColumnWidth = 22
    ws.columns("F:G").ColumnWidth = 65
    ws.Range("A8:G" & lastRow).WrapText = True
    ws.rows("8:" & lastRow).RowHeight = 55
    Exit Sub
Failed:
    er = Err.Description: errNo = Err.Number
    If f <> 0 Then Close #f
    Err.Raise errNo, "InstallShockCatalog", er
End Sub

Private Function ShockCatalogPath() As String
    Dim p As String
    If Len(ThisWorkbook.path) = 0 Then Exit Function
    p = JoinPath(ThisWorkbook.path, "shock_catalog.tsv")
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then ShockCatalogPath = p
    Err.Clear
    On Error GoTo 0
End Function

Private Function UnescapeCatalogField(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    If InStr(1, s, "\", vbBinaryCompare) = 0 Then UnescapeCatalogField = s: Exit Function
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            Select Case Mid$(s, i + 1, 1)
                Case "t": out = out & vbTab
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case Else: out = out & Mid$(s, i + 1, 1)
            End Select
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    UnescapeCatalogField = out
End Function

Public Sub ApplyConfirmedCorrections()
    Dim ws As Worksheet, log As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_RULES)
    Set log = GetWorksheetSafe(ThisWorkbook, "Formula_Changes")
    If log Is Nothing Then Set log = ThisWorkbook.Worksheets.Add: log.name = "Formula_Changes"
    log.Cells.Clear
    log.Range("A1").Value2 = "JKB"
    log.Range("A3").Value2 = "Confirmed formula corrections"
    log.Range("A4").Value2 = "Each change is tied to a specific source definition. Unresolved formulas remain marked for review in Shock_Catalog."
    log.Range("A7:F7").value = Array("Shock type", "Field", "Previous formula", "Corrected formula", "Source reference", "Reason")
    If Not SameFormula(ws.Range("J138").Value2, "0") Then
        If Not SameFormula(ws.Range("J138").Value2, "manual.PRE_LEG_LIQ_TOTAL_LIAB*systemoutput.PCT_CHANGE*(-1)*(0.01)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J138"
        ws.Range("J138").Value2 = "0"
    End If
    log.Cells(8, 1).Value2 = "Change in Value of Liquid Assets"
    log.Cells(8, 2).Value2 = "IMPACT_LEG_LIQ_TOTAL_LIAB"
    log.Cells(8, 3).Value2 = "manual.PRE_LEG_LIQ_TOTAL_LIAB*systemoutput.PCT_CHANGE*(-1)*(0.01)"
    log.Cells(8, 4).Value2 = "0"
    log.Cells(8, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!S152"
    log.Cells(8, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J140").Value2, "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO") Then
        If Not SameFormula(ws.Range("J140").Value2, "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J140"
        ws.Range("J140").Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    End If
    log.Cells(9, 1).Value2 = "Change in Value of Liquid Assets"
    log.Cells(9, 2).Value2 = "IMPACT_LEGAL_LIQUIDITY_RATIO"
    log.Cells(9, 3).Value2 = "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO"
    log.Cells(9, 4).Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    log.Cells(9, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!S155"
    log.Cells(9, 6).Value2 = "The source impact sign is base legal liquidity ratio minus stressed ratio."
    If Not SameFormula(ws.Range("J146").Value2, "0") Then
        If Not SameFormula(ws.Range("J146").Value2, "manual.PRE_LCR_OUTFLOW*systemoutput.PCT_CHANGE*(0.01)*(-1)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J146"
        ws.Range("J146").Value2 = "0"
    End If
    log.Cells(10, 1).Value2 = "Change in Value of Liquid Assets"
    log.Cells(10, 2).Value2 = "IMPACT_LCR_OUTFLOW"
    log.Cells(10, 3).Value2 = "manual.PRE_LCR_OUTFLOW*systemoutput.PCT_CHANGE*(0.01)*(-1)"
    log.Cells(10, 4).Value2 = "0"
    log.Cells(10, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!S141"
    log.Cells(10, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J149").Value2, "0") Then
        If Not SameFormula(ws.Range("J149").Value2, "manual.PRE_LCR_INFLOW*systemoutput.PCT_CHANGE*(0.01)*(-1)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J149"
        ws.Range("J149").Value2 = "0"
    End If
    log.Cells(11, 1).Value2 = "Change in Value of Liquid Assets"
    log.Cells(11, 2).Value2 = "IMPACT_LCR_INFLOW"
    log.Cells(11, 3).Value2 = "manual.PRE_LCR_INFLOW*systemoutput.PCT_CHANGE*(0.01)*(-1)"
    log.Cells(11, 4).Value2 = "0"
    log.Cells(11, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!S143"
    log.Cells(11, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J151").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW") Then
        If Not SameFormula(ws.Range("J151").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J151"
        ws.Range("J151").Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    End If
    log.Cells(12, 1).Value2 = "Change in Value of Liquid Assets"
    log.Cells(12, 2).Value2 = "IMPACT_LCR_NET_OUTFLOW"
    log.Cells(12, 3).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75"
    log.Cells(12, 4).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    log.Cells(12, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!S145"
    log.Cells(12, 6).Value2 = "Net outflow impact is outflow impact minus inflow impact. The current formula incorrectly assumes inflow always equals 75% of outflow."
    If Not SameFormula(ws.Range("J152").Value2, "base.LCR-manual.LCR") Then
        If Not SameFormula(ws.Range("J152").Value2, "manual.LCR-base.LCR") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J152"
        ws.Range("J152").Value2 = "base.LCR-manual.LCR"
    End If
    log.Cells(13, 1).Value2 = "Change in Value of Liquid Assets"
    log.Cells(13, 2).Value2 = "IMPACT_LCR"
    log.Cells(13, 3).Value2 = "manual.LCR-base.LCR"
    log.Cells(13, 4).Value2 = "base.LCR-manual.LCR"
    log.Cells(13, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!S148"
    log.Cells(13, 6).Value2 = "The source impact sign is base LCR minus stressed LCR."
    If Not SameFormula(ws.Range("J224").Value2, "0") Then
        If Not SameFormula(ws.Range("J224").Value2, "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J224"
        ws.Range("J224").Value2 = "0"
    End If
    log.Cells(14, 1).Value2 = "Change in profits"
    log.Cells(14, 2).Value2 = "IMPACT_OUTST_ECL_STAGE1_LCY"
    log.Cells(14, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1"
    log.Cells(14, 4).Value2 = "0"
    log.Cells(14, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P66"
    log.Cells(14, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J225").Value2, "0") Then
        If Not SameFormula(ws.Range("J225").Value2, "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J225"
        ws.Range("J225").Value2 = "0"
    End If
    log.Cells(15, 1).Value2 = "Change in profits"
    log.Cells(15, 2).Value2 = "IMPACT_OUTST_ECL_STAGE2_LCY"
    log.Cells(15, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2"
    log.Cells(15, 4).Value2 = "0"
    log.Cells(15, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P67"
    log.Cells(15, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J226").Value2, "0") Then
        If Not SameFormula(ws.Range("J226").Value2, "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J226"
        ws.Range("J226").Value2 = "0"
    End If
    log.Cells(16, 1).Value2 = "Change in profits"
    log.Cells(16, 2).Value2 = "IMPACT_OUTST_ECL_NPA_LCY"
    log.Cells(16, 3).Value2 = "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA"
    log.Cells(16, 4).Value2 = "0"
    log.Cells(16, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P68"
    log.Cells(16, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J227").Value2, "0") Then
        If Not SameFormula(ws.Range("J227").Value2, "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J227"
        ws.Range("J227").Value2 = "0"
    End If
    log.Cells(17, 1).Value2 = "Change in profits"
    log.Cells(17, 2).Value2 = "IMPACT_OUTST_RWA_STAGE1_LCY"
    log.Cells(17, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1"
    log.Cells(17, 4).Value2 = "0"
    log.Cells(17, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P72"
    log.Cells(17, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J228").Value2, "0") Then
        If Not SameFormula(ws.Range("J228").Value2, "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J228"
        ws.Range("J228").Value2 = "0"
    End If
    log.Cells(18, 1).Value2 = "Change in profits"
    log.Cells(18, 2).Value2 = "IMPACT_OUTST_RWA_STAGE2_LCY"
    log.Cells(18, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2"
    log.Cells(18, 4).Value2 = "0"
    log.Cells(18, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P73"
    log.Cells(18, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J229").Value2, "0") Then
        If Not SameFormula(ws.Range("J229").Value2, "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J229"
        ws.Range("J229").Value2 = "0"
    End If
    log.Cells(19, 1).Value2 = "Change in profits"
    log.Cells(19, 2).Value2 = "IMPACT_OUTST_RWA_NPA_LCY"
    log.Cells(19, 3).Value2 = "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA"
    log.Cells(19, 4).Value2 = "0"
    log.Cells(19, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P74"
    log.Cells(19, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J233").Value2, "base.OUTST_LCY") Then
        If Not SameFormula(ws.Range("J233").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J233"
        ws.Range("J233").Value2 = "base.OUTST_LCY"
    End If
    log.Cells(20, 1).Value2 = "Change in profits"
    log.Cells(20, 2).Value2 = "OUTST_LCY"
    log.Cells(20, 3).Value2 = ""
    log.Cells(20, 4).Value2 = "base.OUTST_LCY"
    log.Cells(20, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P81"
    log.Cells(20, 6).Value2 = "Profit/loss shock has zero exposure impact; gross outstanding is unchanged."
    If Not SameFormula(ws.Range("J258").Value2, "0") Then
        If Not SameFormula(ws.Range("J258").Value2, "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J258"
        ws.Range("J258").Value2 = "0"
    End If
    log.Cells(21, 1).Value2 = "Change in profits"
    log.Cells(21, 2).Value2 = "IMPACT_RWA_MR_EQUITY_LCY"
    log.Cells(21, 3).Value2 = "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)"
    log.Cells(21, 4).Value2 = "0"
    log.Cells(21, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P108"
    log.Cells(21, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J260").Value2, "0") Then
        If Not SameFormula(ws.Range("J260").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J260"
        ws.Range("J260").Value2 = "0"
    End If
    log.Cells(22, 1).Value2 = "Change in profits"
    log.Cells(22, 2).Value2 = "IMPACT_RWA_MR_FOREX_LCY"
    log.Cells(22, 3).Value2 = ""
    log.Cells(22, 4).Value2 = "0"
    log.Cells(22, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P110"
    log.Cells(22, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J262").Value2, "0") Then
        If Not SameFormula(ws.Range("J262").Value2, "((manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE*(0.01)*(0.01)*(0.01))/3)*0.15*12.5") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J262"
        ws.Range("J262").Value2 = "0"
    End If
    log.Cells(23, 1).Value2 = "Change in profits"
    log.Cells(23, 2).Value2 = "IMPACT_RWA_OR_LCY"
    log.Cells(23, 3).Value2 = "((manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE*(0.01)*(0.01)*(0.01))/3)*0.15*12.5"
    log.Cells(23, 4).Value2 = "0"
    log.Cells(23, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P112"
    log.Cells(23, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J269").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J269").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J269"
        ws.Range("J269").Value2 = "base.TAX_RATE"
    End If
    log.Cells(24, 1).Value2 = "Change in profits"
    log.Cells(24, 2).Value2 = "TAX_RATE"
    log.Cells(24, 3).Value2 = ""
    log.Cells(24, 4).Value2 = "base.TAX_RATE"
    log.Cells(24, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!P119"
    log.Cells(24, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J386").Value2, "(manual.FX_TOTAL_LONG_POSITION_LCY-manual.FX_TOTAL_SHORT_POSITION_LCY)-(base.FX_TOTAL_LONG_POSITION_LCY-base.FX_TOTAL_SHORT_POSITION_LCY)") Then
        If Not SameFormula(ws.Range("J386").Value2, "(manual.FX_TOTAL_LONG_POSITION_LCY-manual.FX_TOTAL_SHORT_POSITION_LCY)-base.IMPACT_NET_OPEN_POSITION") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J386"
        ws.Range("J386").Value2 = "(manual.FX_TOTAL_LONG_POSITION_LCY-manual.FX_TOTAL_SHORT_POSITION_LCY)-(base.FX_TOTAL_LONG_POSITION_LCY-base.FX_TOTAL_SHORT_POSITION_LCY)"
    End If
    log.Cells(25, 1).Value2 = "Exchange rates changing"
    log.Cells(25, 2).Value2 = "IMPACT_NET_OPEN_POSITION"
    log.Cells(25, 3).Value2 = "(manual.FX_TOTAL_LONG_POSITION_LCY-manual.FX_TOTAL_SHORT_POSITION_LCY)-base.IMPACT_NET_OPEN_POSITION"
    log.Cells(25, 4).Value2 = "(manual.FX_TOTAL_LONG_POSITION_LCY-manual.FX_TOTAL_SHORT_POSITION_LCY)-(base.FX_TOTAL_LONG_POSITION_LCY-base.FX_TOTAL_SHORT_POSITION_LCY)"
    log.Cells(25, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K80"
    log.Cells(25, 6).Value2 = "Subtract base net long-minus-short position, not the base impact field."
    If Not SameFormula(ws.Range("J387").Value2, "base.OUTST_LCY+manual.IMPACT_OUTST_RWA_STAGE1_LCY+manual.IMPACT_OUTST_RWA_STAGE2_LCY+manual.IMPACT_OUTST_RWA_NPA_LCY") Then
        If Not SameFormula(ws.Range("J387").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J387"
        ws.Range("J387").Value2 = "base.OUTST_LCY+manual.IMPACT_OUTST_RWA_STAGE1_LCY+manual.IMPACT_OUTST_RWA_STAGE2_LCY+manual.IMPACT_OUTST_RWA_NPA_LCY"
    End If
    log.Cells(26, 1).Value2 = "Exchange rates changing"
    log.Cells(26, 2).Value2 = "OUTST_LCY"
    log.Cells(26, 3).Value2 = ""
    log.Cells(26, 4).Value2 = "base.OUTST_LCY+manual.IMPACT_OUTST_RWA_STAGE1_LCY+manual.IMPACT_OUTST_RWA_STAGE2_LCY+manual.IMPACT_OUTST_RWA_NPA_LCY"
    log.Cells(26, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K81"
    log.Cells(26, 6).Value2 = "Gross outstanding follows bank-wide base plus total RWA-basis exposure impacts per matrix."
    If Not SameFormula(ws.Range("J412").Value2, "0") Then
        If Not SameFormula(ws.Range("J412").Value2, "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J412"
        ws.Range("J412").Value2 = "0"
    End If
    log.Cells(27, 1).Value2 = "Exchange rates changing"
    log.Cells(27, 2).Value2 = "IMPACT_RWA_MR_EQUITY_LCY"
    log.Cells(27, 3).Value2 = "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)"
    log.Cells(27, 4).Value2 = "0"
    log.Cells(27, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K108"
    log.Cells(27, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J414").Value2, "base.RWA_MR_FOREX_LCY*ABS(systemoutput.PCT_CHANGE)/100") Then
        If Not SameFormula(ws.Range("J414").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J414"
        ws.Range("J414").Value2 = "base.RWA_MR_FOREX_LCY*ABS(systemoutput.PCT_CHANGE)/100"
    End If
    log.Cells(28, 1).Value2 = "Exchange rates changing"
    log.Cells(28, 2).Value2 = "IMPACT_RWA_MR_FOREX_LCY"
    log.Cells(28, 3).Value2 = ""
    log.Cells(28, 4).Value2 = "base.RWA_MR_FOREX_LCY*ABS(systemoutput.PCT_CHANGE)/100"
    log.Cells(28, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K110"
    log.Cells(28, 6).Value2 = "FX market RWA impact uses absolute magnitude percent once."
    If Not SameFormula(ws.Range("J416").Value2, "0") Then
        If Not SameFormula(ws.Range("J416").Value2, "((manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE*(0.01)*(0.01)*(0.01))/3)*0.15*12.5") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J416"
        ws.Range("J416").Value2 = "0"
    End If
    log.Cells(29, 1).Value2 = "Exchange rates changing"
    log.Cells(29, 2).Value2 = "IMPACT_RWA_OR_LCY"
    log.Cells(29, 3).Value2 = "((manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE*(0.01)*(0.01)*(0.01))/3)*0.15*12.5"
    log.Cells(29, 4).Value2 = "0"
    log.Cells(29, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K112"
    log.Cells(29, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J422").Value2, "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY+manual.IMPACT_NET_OPEN_POSITION") Then
        If Not SameFormula(ws.Range("J422").Value2, "systemoutput.PROFITS_BEFORE_TAX_LCY") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J422"
        ws.Range("J422").Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY+manual.IMPACT_NET_OPEN_POSITION"
    End If
    log.Cells(30, 1).Value2 = "Exchange rates changing"
    log.Cells(30, 2).Value2 = "PROFITS_BEFORE_TAX_LCY"
    log.Cells(30, 3).Value2 = "systemoutput.PROFITS_BEFORE_TAX_LCY"
    log.Cells(30, 4).Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY+manual.IMPACT_NET_OPEN_POSITION"
    log.Cells(30, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K118"
    log.Cells(30, 6).Value2 = "FX PBT includes credit ECL impact and net open position change."
    If Not SameFormula(ws.Range("J423").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J423").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J423"
        ws.Range("J423").Value2 = "base.TAX_RATE"
    End If
    log.Cells(31, 1).Value2 = "Exchange rates changing"
    log.Cells(31, 2).Value2 = "TAX_RATE"
    log.Cells(31, 3).Value2 = ""
    log.Cells(31, 4).Value2 = "base.TAX_RATE"
    log.Cells(31, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!K119"
    log.Cells(31, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J535").Value2, "base.OUTST_LCY") Then
        If Not SameFormula(ws.Range("J535").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J535"
        ws.Range("J535").Value2 = "base.OUTST_LCY"
    End If
    log.Cells(32, 1).Value2 = "Increase in NPA percentage of segment"
    log.Cells(32, 2).Value2 = "OUTST_LCY"
    log.Cells(32, 3).Value2 = ""
    log.Cells(32, 4).Value2 = "base.OUTST_LCY"
    log.Cells(32, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!H81"
    log.Cells(32, 6).Value2 = "NPA migration preserves total gross outstanding."
    If Not SameFormula(ws.Range("J570").Value2, "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY") Then
        If Not SameFormula(ws.Range("J570").Value2, "systemoutput.PROFITS_BEFORE_TAX_LCY") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J570"
        ws.Range("J570").Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY"
    End If
    log.Cells(33, 1).Value2 = "Increase in NPA percentage of segment"
    log.Cells(33, 2).Value2 = "PROFITS_BEFORE_TAX_LCY"
    log.Cells(33, 3).Value2 = "systemoutput.PROFITS_BEFORE_TAX_LCY"
    log.Cells(33, 4).Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY"
    log.Cells(33, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!H118"
    log.Cells(33, 6).Value2 = "Recalculate stressed PBT from base and independently calculated ECL impact instead of copying the system result."
    If Not SameFormula(ws.Range("J571").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J571").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J571"
        ws.Range("J571").Value2 = "base.TAX_RATE"
    End If
    log.Cells(34, 1).Value2 = "Increase in NPA percentage of segment"
    log.Cells(34, 2).Value2 = "TAX_RATE"
    log.Cells(34, 3).Value2 = ""
    log.Cells(34, 4).Value2 = "base.TAX_RATE"
    log.Cells(34, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!H119"
    log.Cells(34, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J708").Value2, "0") Then
        If Not SameFormula(ws.Range("J708").Value2, "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J708"
        ws.Range("J708").Value2 = "0"
    End If
    log.Cells(35, 1).Value2 = "Interest rate change"
    log.Cells(35, 2).Value2 = "IMPACT_RWA_MR_EQUITY_LCY"
    log.Cells(35, 3).Value2 = "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)"
    log.Cells(35, 4).Value2 = "0"
    log.Cells(35, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!M108"
    log.Cells(35, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J710").Value2, "0") Then
        If Not SameFormula(ws.Range("J710").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J710"
        ws.Range("J710").Value2 = "0"
    End If
    log.Cells(36, 1).Value2 = "Interest rate change"
    log.Cells(36, 2).Value2 = "IMPACT_RWA_MR_FOREX_LCY"
    log.Cells(36, 3).Value2 = ""
    log.Cells(36, 4).Value2 = "0"
    log.Cells(36, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!M110"
    log.Cells(36, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J718").Value2, "base.PROFITS_BEFORE_TAX_LCY+manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE/10000") Then
        If Not SameFormula(ws.Range("J718").Value2, "systemoutput.PROFITS_BEFORE_TAX_LCY") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J718"
        ws.Range("J718").Value2 = "base.PROFITS_BEFORE_TAX_LCY+manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE/10000"
    End If
    log.Cells(37, 1).Value2 = "Interest rate change"
    log.Cells(37, 2).Value2 = "PROFITS_BEFORE_TAX_LCY"
    log.Cells(37, 3).Value2 = "systemoutput.PROFITS_BEFORE_TAX_LCY"
    log.Cells(37, 4).Value2 = "base.PROFITS_BEFORE_TAX_LCY+manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE/10000"
    log.Cells(37, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!M118"
    log.Cells(37, 6).Value2 = "Interest shock magnitude is basis points; PBT is base plus repricing gap times bps/10000."
    If Not SameFormula(ws.Range("J719").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J719").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J719"
        ws.Range("J719").Value2 = "base.TAX_RATE"
    End If
    log.Cells(38, 1).Value2 = "Interest rate change"
    log.Cells(38, 2).Value2 = "TAX_RATE"
    log.Cells(38, 3).Value2 = ""
    log.Cells(38, 4).Value2 = "base.TAX_RATE"
    log.Cells(38, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!M119"
    log.Cells(38, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J893").Value2, "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO") Then
        If Not SameFormula(ws.Range("J893").Value2, "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J893"
        ws.Range("J893").Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    End If
    log.Cells(39, 1).Value2 = "Limit Drawdown"
    log.Cells(39, 2).Value2 = "IMPACT_LEGAL_LIQUIDITY_RATIO"
    log.Cells(39, 3).Value2 = "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO"
    log.Cells(39, 4).Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    log.Cells(39, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!R155"
    log.Cells(39, 6).Value2 = "The source impact sign is base legal liquidity ratio minus stressed ratio."
    If Not SameFormula(ws.Range("J904").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW") Then
        If Not SameFormula(ws.Range("J904").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J904"
        ws.Range("J904").Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    End If
    log.Cells(40, 1).Value2 = "Limit Drawdown"
    log.Cells(40, 2).Value2 = "IMPACT_LCR_NET_OUTFLOW"
    log.Cells(40, 3).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75"
    log.Cells(40, 4).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    log.Cells(40, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!R145"
    log.Cells(40, 6).Value2 = "Net outflow impact is outflow impact minus inflow impact. The current formula incorrectly assumes inflow always equals 75% of outflow."
    If Not SameFormula(ws.Range("J905").Value2, "base.LCR-manual.LCR") Then
        If Not SameFormula(ws.Range("J905").Value2, "manual.LCR-base.LCR") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J905"
        ws.Range("J905").Value2 = "base.LCR-manual.LCR"
    End If
    log.Cells(41, 1).Value2 = "Limit Drawdown"
    log.Cells(41, 2).Value2 = "IMPACT_LCR"
    log.Cells(41, 3).Value2 = "manual.LCR-base.LCR"
    log.Cells(41, 4).Value2 = "base.LCR-manual.LCR"
    log.Cells(41, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!R148"
    log.Cells(41, 6).Value2 = "The source impact sign is base LCR minus stressed LCR."
    If Not SameFormula(ws.Range("J986").Value2, "base.OUTST_LCY+manual.IMPACT_OUTST_RWA_STAGE1_LCY+manual.IMPACT_OUTST_RWA_STAGE2_LCY+manual.IMPACT_OUTST_RWA_NPA_LCY") Then
        If Not SameFormula(ws.Range("J986").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J986"
        ws.Range("J986").Value2 = "base.OUTST_LCY+manual.IMPACT_OUTST_RWA_STAGE1_LCY+manual.IMPACT_OUTST_RWA_STAGE2_LCY+manual.IMPACT_OUTST_RWA_NPA_LCY"
    End If
    log.Cells(42, 1).Value2 = "Market Price Change of Bonds and Stocks"
    log.Cells(42, 2).Value2 = "OUTST_LCY"
    log.Cells(42, 3).Value2 = ""
    log.Cells(42, 4).Value2 = "base.OUTST_LCY+manual.IMPACT_OUTST_RWA_STAGE1_LCY+manual.IMPACT_OUTST_RWA_STAGE2_LCY+manual.IMPACT_OUTST_RWA_NPA_LCY"
    log.Cells(42, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!O81"
    log.Cells(42, 6).Value2 = "Gross outstanding follows bank-wide base plus total RWA-basis exposure impacts per matrix."
    If Not SameFormula(ws.Range("J1022").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J1022").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1022"
        ws.Range("J1022").Value2 = "base.TAX_RATE"
    End If
    log.Cells(43, 1).Value2 = "Market Price Change of Bonds and Stocks"
    log.Cells(43, 2).Value2 = "TAX_RATE"
    log.Cells(43, 3).Value2 = ""
    log.Cells(43, 4).Value2 = "base.TAX_RATE"
    log.Cells(43, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!O119"
    log.Cells(43, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J1134").Value2, "base.OUTST_LCY") Then
        If Not SameFormula(ws.Range("J1134").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1134"
        ws.Range("J1134").Value2 = "base.OUTST_LCY"
    End If
    log.Cells(44, 1).Value2 = "Specific customers becomes NPA"
    log.Cells(44, 2).Value2 = "OUTST_LCY"
    log.Cells(44, 3).Value2 = ""
    log.Cells(44, 4).Value2 = "base.OUTST_LCY"
    log.Cells(44, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!D81"
    log.Cells(44, 6).Value2 = "NPA migration preserves total gross outstanding."
    If Not SameFormula(ws.Range("J1143").Value2, "-manual.ECL_STAGE1_LCY_PRE_SHOCK") Then
        If Not SameFormula(ws.Range("J1143").Value2, "manual.IMPACT_OUTST_ECL_STAGE1_LCY* manual.ECL_RATIO_STAGE1") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1143"
        ws.Range("J1143").Value2 = "-manual.ECL_STAGE1_LCY_PRE_SHOCK"
    End If
    log.Cells(45, 1).Value2 = "Specific customers becomes NPA"
    log.Cells(45, 2).Value2 = "IMPACT_ECL_STAGE1_LCY"
    log.Cells(45, 3).Value2 = "manual.IMPACT_OUTST_ECL_STAGE1_LCY* manual.ECL_RATIO_STAGE1"
    log.Cells(45, 4).Value2 = "-manual.ECL_STAGE1_LCY_PRE_SHOCK"
    log.Cells(45, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!D90"
    log.Cells(45, 6).Value2 = "Customer default explicitly reverses the full selected original stage ECL. Multiplying zero exposure by a fallback ratio can fail to reverse nonzero ECL."
    If Not SameFormula(ws.Range("J1144").Value2, "-manual.ECL_STAGE2_LCY_PRE_SHOCK") Then
        If Not SameFormula(ws.Range("J1144").Value2, "manual.IMPACT_OUTST_ECL_STAGE2_LCY* manual.ECL_RATIO_STAGE2") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1144"
        ws.Range("J1144").Value2 = "-manual.ECL_STAGE2_LCY_PRE_SHOCK"
    End If
    log.Cells(46, 1).Value2 = "Specific customers becomes NPA"
    log.Cells(46, 2).Value2 = "IMPACT_ECL_STAGE2_LCY"
    log.Cells(46, 3).Value2 = "manual.IMPACT_OUTST_ECL_STAGE2_LCY* manual.ECL_RATIO_STAGE2"
    log.Cells(46, 4).Value2 = "-manual.ECL_STAGE2_LCY_PRE_SHOCK"
    log.Cells(46, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!D91"
    log.Cells(46, 6).Value2 = "Customer default explicitly reverses the full selected original stage ECL. Multiplying zero exposure by a fallback ratio can fail to reverse nonzero ECL."
    If Not SameFormula(ws.Range("J1169").Value2, "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY") Then
        If Not SameFormula(ws.Range("J1169").Value2, "systemoutput.PROFITS_BEFORE_TAX_LCY") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1169"
        ws.Range("J1169").Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY"
    End If
    log.Cells(47, 1).Value2 = "Specific customers becomes NPA"
    log.Cells(47, 2).Value2 = "PROFITS_BEFORE_TAX_LCY"
    log.Cells(47, 3).Value2 = "systemoutput.PROFITS_BEFORE_TAX_LCY"
    log.Cells(47, 4).Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY"
    log.Cells(47, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!D118"
    log.Cells(47, 6).Value2 = "Recalculate stressed PBT from base and independently calculated ECL impact instead of copying the system result."
    If Not SameFormula(ws.Range("J1170").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J1170").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1170"
        ws.Range("J1170").Value2 = "base.TAX_RATE"
    End If
    log.Cells(48, 1).Value2 = "Specific customers becomes NPA"
    log.Cells(48, 2).Value2 = "TAX_RATE"
    log.Cells(48, 3).Value2 = ""
    log.Cells(48, 4).Value2 = "base.TAX_RATE"
    log.Cells(48, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!D119"
    log.Cells(48, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J1344").Value2, "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO") Then
        If Not SameFormula(ws.Range("J1344").Value2, "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1344"
        ws.Range("J1344").Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    End If
    log.Cells(49, 1).Value2 = "Specific customers withdraw deposits"
    log.Cells(49, 2).Value2 = "IMPACT_LEGAL_LIQUIDITY_RATIO"
    log.Cells(49, 3).Value2 = "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO"
    log.Cells(49, 4).Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    log.Cells(49, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!T155"
    log.Cells(49, 6).Value2 = "The source impact sign is base legal liquidity ratio minus stressed ratio."
    If Not SameFormula(ws.Range("J1353").Value2, "0") Then
        If Not SameFormula(ws.Range("J1353").Value2, "manual.PRE_LCR_INFLOW*(-1)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1353"
        ws.Range("J1353").Value2 = "0"
    End If
    log.Cells(50, 1).Value2 = "Specific customers withdraw deposits"
    log.Cells(50, 2).Value2 = "IMPACT_LCR_INFLOW"
    log.Cells(50, 3).Value2 = "manual.PRE_LCR_INFLOW*(-1)"
    log.Cells(50, 4).Value2 = "0"
    log.Cells(50, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!T143"
    log.Cells(50, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1355").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW") Then
        If Not SameFormula(ws.Range("J1355").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1355"
        ws.Range("J1355").Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    End If
    log.Cells(51, 1).Value2 = "Specific customers withdraw deposits"
    log.Cells(51, 2).Value2 = "IMPACT_LCR_NET_OUTFLOW"
    log.Cells(51, 3).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75"
    log.Cells(51, 4).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    log.Cells(51, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!T145"
    log.Cells(51, 6).Value2 = "Net outflow impact is outflow impact minus inflow impact. The current formula incorrectly assumes inflow always equals 75% of outflow."
    If Not SameFormula(ws.Range("J1356").Value2, "base.LCR-manual.LCR") Then
        If Not SameFormula(ws.Range("J1356").Value2, "manual.LCR-base.LCR") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1356"
        ws.Range("J1356").Value2 = "base.LCR-manual.LCR"
    End If
    log.Cells(52, 1).Value2 = "Specific customers withdraw deposits"
    log.Cells(52, 2).Value2 = "IMPACT_LCR"
    log.Cells(52, 3).Value2 = "manual.LCR-base.LCR"
    log.Cells(52, 4).Value2 = "base.LCR-manual.LCR"
    log.Cells(52, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!T148"
    log.Cells(52, 6).Value2 = "The source impact sign is base LCR minus stressed LCR."
    If Not SameFormula(ws.Range("J1428").Value2, "0") Then
        If Not SameFormula(ws.Range("J1428").Value2, "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1428"
        ws.Range("J1428").Value2 = "0"
    End If
    log.Cells(53, 1).Value2 = "Specific loss"
    log.Cells(53, 2).Value2 = "IMPACT_OUTST_ECL_STAGE1_LCY"
    log.Cells(53, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1"
    log.Cells(53, 4).Value2 = "0"
    log.Cells(53, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W66"
    log.Cells(53, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1429").Value2, "0") Then
        If Not SameFormula(ws.Range("J1429").Value2, "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1429"
        ws.Range("J1429").Value2 = "0"
    End If
    log.Cells(54, 1).Value2 = "Specific loss"
    log.Cells(54, 2).Value2 = "IMPACT_OUTST_ECL_STAGE2_LCY"
    log.Cells(54, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2"
    log.Cells(54, 4).Value2 = "0"
    log.Cells(54, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W67"
    log.Cells(54, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1430").Value2, "0") Then
        If Not SameFormula(ws.Range("J1430").Value2, "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1430"
        ws.Range("J1430").Value2 = "0"
    End If
    log.Cells(55, 1).Value2 = "Specific loss"
    log.Cells(55, 2).Value2 = "IMPACT_OUTST_ECL_NPA_LCY"
    log.Cells(55, 3).Value2 = "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA"
    log.Cells(55, 4).Value2 = "0"
    log.Cells(55, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W68"
    log.Cells(55, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1431").Value2, "0") Then
        If Not SameFormula(ws.Range("J1431").Value2, "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1431"
        ws.Range("J1431").Value2 = "0"
    End If
    log.Cells(56, 1).Value2 = "Specific loss"
    log.Cells(56, 2).Value2 = "IMPACT_OUTST_RWA_STAGE1_LCY"
    log.Cells(56, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE1-manual.IMPACT_IIS_STAGE1"
    log.Cells(56, 4).Value2 = "0"
    log.Cells(56, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W72"
    log.Cells(56, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1432").Value2, "0") Then
        If Not SameFormula(ws.Range("J1432").Value2, "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1432"
        ws.Range("J1432").Value2 = "0"
    End If
    log.Cells(57, 1).Value2 = "Specific loss"
    log.Cells(57, 2).Value2 = "IMPACT_OUTST_RWA_STAGE2_LCY"
    log.Cells(57, 3).Value2 = "manual.IMPACT_OUTST_LCY_STAGE2-manual.IMPACT_IIS_STAGE2"
    log.Cells(57, 4).Value2 = "0"
    log.Cells(57, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W73"
    log.Cells(57, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1433").Value2, "0") Then
        If Not SameFormula(ws.Range("J1433").Value2, "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1433"
        ws.Range("J1433").Value2 = "0"
    End If
    log.Cells(58, 1).Value2 = "Specific loss"
    log.Cells(58, 2).Value2 = "IMPACT_OUTST_RWA_NPA_LCY"
    log.Cells(58, 3).Value2 = "manual.IMPACT_OUTST_LCY_NPA- manual.IMPACT_IIS_NPA"
    log.Cells(58, 4).Value2 = "0"
    log.Cells(58, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W74"
    log.Cells(58, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1437").Value2, "base.OUTST_LCY") Then
        If Not SameFormula(ws.Range("J1437").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1437"
        ws.Range("J1437").Value2 = "base.OUTST_LCY"
    End If
    log.Cells(59, 1).Value2 = "Specific loss"
    log.Cells(59, 2).Value2 = "OUTST_LCY"
    log.Cells(59, 3).Value2 = ""
    log.Cells(59, 4).Value2 = "base.OUTST_LCY"
    log.Cells(59, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W81"
    log.Cells(59, 6).Value2 = "Profit/loss shock has zero exposure impact; gross outstanding is unchanged."
    If Not SameFormula(ws.Range("J1462").Value2, "0") Then
        If Not SameFormula(ws.Range("J1462").Value2, "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1462"
        ws.Range("J1462").Value2 = "0"
    End If
    log.Cells(60, 1).Value2 = "Specific loss"
    log.Cells(60, 2).Value2 = "IMPACT_RWA_MR_EQUITY_LCY"
    log.Cells(60, 3).Value2 = "base.RWA_MR_EQUITY_LCY* systemoutput.PCT_CHANGE*(0.01)*(0.01)"
    log.Cells(60, 4).Value2 = "0"
    log.Cells(60, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W108"
    log.Cells(60, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1464").Value2, "0") Then
        If Not SameFormula(ws.Range("J1464").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1464"
        ws.Range("J1464").Value2 = "0"
    End If
    log.Cells(61, 1).Value2 = "Specific loss"
    log.Cells(61, 2).Value2 = "IMPACT_RWA_MR_FOREX_LCY"
    log.Cells(61, 3).Value2 = ""
    log.Cells(61, 4).Value2 = "0"
    log.Cells(61, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W110"
    log.Cells(61, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1466").Value2, "0") Then
        If Not SameFormula(ws.Range("J1466").Value2, "((manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE*(0.01)*(0.01)*(0.01))/3)*0.15*12.5") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1466"
        ws.Range("J1466").Value2 = "0"
    End If
    log.Cells(62, 1).Value2 = "Specific loss"
    log.Cells(62, 2).Value2 = "IMPACT_RWA_OR_LCY"
    log.Cells(62, 3).Value2 = "((manual.NET_RSA_GAP_LCY*systemoutput.PCT_CHANGE*(0.01)*(0.01)*(0.01))/3)*0.15*12.5"
    log.Cells(62, 4).Value2 = "0"
    log.Cells(62, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W112"
    log.Cells(62, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1472").Value2, "base.PROFITS_BEFORE_TAX_LCY+manual.IMPACT_PROFITS_BEFORE_TAX_LCY") Then
        If Not SameFormula(ws.Range("J1472").Value2, "systemoutput.PROFITS_BEFORE_TAX_LCY") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1472"
        ws.Range("J1472").Value2 = "base.PROFITS_BEFORE_TAX_LCY+manual.IMPACT_PROFITS_BEFORE_TAX_LCY"
    End If
    log.Cells(63, 1).Value2 = "Specific loss"
    log.Cells(63, 2).Value2 = "PROFITS_BEFORE_TAX_LCY"
    log.Cells(63, 3).Value2 = "systemoutput.PROFITS_BEFORE_TAX_LCY"
    log.Cells(63, 4).Value2 = "base.PROFITS_BEFORE_TAX_LCY+manual.IMPACT_PROFITS_BEFORE_TAX_LCY"
    log.Cells(63, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W118"
    log.Cells(63, 6).Value2 = "Recalculate stressed PBT from base plus direct profit impact."
    If Not SameFormula(ws.Range("J1473").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J1473").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1473"
        ws.Range("J1473").Value2 = "base.TAX_RATE"
    End If
    log.Cells(64, 1).Value2 = "Specific loss"
    log.Cells(64, 2).Value2 = "TAX_RATE"
    log.Cells(64, 3).Value2 = ""
    log.Cells(64, 4).Value2 = "base.TAX_RATE"
    log.Cells(64, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!W119"
    log.Cells(64, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J1647").Value2, "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO") Then
        If Not SameFormula(ws.Range("J1647").Value2, "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1647"
        ws.Range("J1647").Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    End If
    log.Cells(65, 1).Value2 = "Specified Percentage of Deposits are Withdrawn"
    log.Cells(65, 2).Value2 = "IMPACT_LEGAL_LIQUIDITY_RATIO"
    log.Cells(65, 3).Value2 = "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO"
    log.Cells(65, 4).Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    log.Cells(65, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Q155"
    log.Cells(65, 6).Value2 = "The source impact sign is base legal liquidity ratio minus stressed ratio."
    If Not SameFormula(ws.Range("J1656").Value2, "0") Then
        If Not SameFormula(ws.Range("J1656").Value2, "manual.PRE_LCR_INFLOW*systemoutput.PCT_CHANGE*(0.01)*(-1)") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1656"
        ws.Range("J1656").Value2 = "0"
    End If
    log.Cells(66, 1).Value2 = "Specified Percentage of Deposits are Withdrawn"
    log.Cells(66, 2).Value2 = "IMPACT_LCR_INFLOW"
    log.Cells(66, 3).Value2 = "manual.PRE_LCR_INFLOW*systemoutput.PCT_CHANGE*(0.01)*(-1)"
    log.Cells(66, 4).Value2 = "0"
    log.Cells(66, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Q143"
    log.Cells(66, 6).Value2 = "The source explicitly specifies no impact for this field and shock type."
    If Not SameFormula(ws.Range("J1658").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW") Then
        If Not SameFormula(ws.Range("J1658").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1658"
        ws.Range("J1658").Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    End If
    log.Cells(67, 1).Value2 = "Specified Percentage of Deposits are Withdrawn"
    log.Cells(67, 2).Value2 = "IMPACT_LCR_NET_OUTFLOW"
    log.Cells(67, 3).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75"
    log.Cells(67, 4).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    log.Cells(67, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Q145"
    log.Cells(67, 6).Value2 = "Net outflow impact is outflow impact minus inflow impact. The current formula incorrectly assumes inflow always equals 75% of outflow."
    If Not SameFormula(ws.Range("J1659").Value2, "base.LCR-manual.LCR") Then
        If Not SameFormula(ws.Range("J1659").Value2, "manual.LCR-base.LCR") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1659"
        ws.Range("J1659").Value2 = "base.LCR-manual.LCR"
    End If
    log.Cells(68, 1).Value2 = "Specified Percentage of Deposits are Withdrawn"
    log.Cells(68, 2).Value2 = "IMPACT_LCR"
    log.Cells(68, 3).Value2 = "manual.LCR-base.LCR"
    log.Cells(68, 4).Value2 = "base.LCR-manual.LCR"
    log.Cells(68, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Q148"
    log.Cells(68, 6).Value2 = "The source impact sign is base LCR minus stressed LCR."
    If Not SameFormula(ws.Range("J1733").Value2, "base.OUTST_LCY") Then
        If Not SameFormula(ws.Range("J1733").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1733"
        ws.Range("J1733").Value2 = "base.OUTST_LCY"
    End If
    log.Cells(69, 1).Value2 = "Specified portfolio segment moves from performing to NPA"
    log.Cells(69, 2).Value2 = "OUTST_LCY"
    log.Cells(69, 3).Value2 = ""
    log.Cells(69, 4).Value2 = "base.OUTST_LCY"
    log.Cells(69, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!F81"
    log.Cells(69, 6).Value2 = "NPA migration preserves total gross outstanding."
    If Not SameFormula(ws.Range("J1768").Value2, "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY") Then
        If Not SameFormula(ws.Range("J1768").Value2, "systemoutput.PROFITS_BEFORE_TAX_LCY") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1768"
        ws.Range("J1768").Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY"
    End If
    log.Cells(70, 1).Value2 = "Specified portfolio segment moves from performing to NPA"
    log.Cells(70, 2).Value2 = "PROFITS_BEFORE_TAX_LCY"
    log.Cells(70, 3).Value2 = "systemoutput.PROFITS_BEFORE_TAX_LCY"
    log.Cells(70, 4).Value2 = "base.PROFITS_BEFORE_TAX_LCY-manual.IMPACT_ECL_LCY"
    log.Cells(70, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!F118"
    log.Cells(70, 6).Value2 = "Recalculate stressed PBT from base and independently calculated ECL impact instead of copying the system result."
    If Not SameFormula(ws.Range("J1769").Value2, "base.TAX_RATE") Then
        If Not SameFormula(ws.Range("J1769").Value2, "") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1769"
        ws.Range("J1769").Value2 = "base.TAX_RATE"
    End If
    log.Cells(71, 1).Value2 = "Specified portfolio segment moves from performing to NPA"
    log.Cells(71, 2).Value2 = "TAX_RATE"
    log.Cells(71, 3).Value2 = ""
    log.Cells(71, 4).Value2 = "base.TAX_RATE"
    log.Cells(71, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!F119"
    log.Cells(71, 6).Value2 = "Tax rate is unchanged from independently reconciled base."
    If Not SameFormula(ws.Range("J1943").Value2, "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO") Then
        If Not SameFormula(ws.Range("J1943").Value2, "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1943"
        ws.Range("J1943").Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    End If
    log.Cells(72, 1).Value2 = "Deposit increase or withdrawn without change in liquid assets"
    log.Cells(72, 2).Value2 = "IMPACT_LEGAL_LIQUIDITY_RATIO"
    log.Cells(72, 3).Value2 = "manual.LEGAL_LIQUIDITY_RATIO-base.LEGAL_LIQUIDITY_RATIO"
    log.Cells(72, 4).Value2 = "base.LEGAL_LIQUIDITY_RATIO-manual.LEGAL_LIQUIDITY_RATIO"
    log.Cells(72, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Z155"
    log.Cells(72, 6).Value2 = "The source impact sign is base legal liquidity ratio minus stressed ratio."
    If Not SameFormula(ws.Range("J1954").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW") Then
        If Not SameFormula(ws.Range("J1954").Value2, "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1954"
        ws.Range("J1954").Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    End If
    log.Cells(73, 1).Value2 = "Deposit increase or withdrawn without change in liquid assets"
    log.Cells(73, 2).Value2 = "IMPACT_LCR_NET_OUTFLOW"
    log.Cells(73, 3).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_OUTFLOW*0.75"
    log.Cells(73, 4).Value2 = "manual.IMPACT_LCR_OUTFLOW-manual.IMPACT_LCR_INFLOW"
    log.Cells(73, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Z145"
    log.Cells(73, 6).Value2 = "Net outflow impact is outflow impact minus inflow impact. The current formula incorrectly assumes inflow always equals 75% of outflow."
    If Not SameFormula(ws.Range("J1955").Value2, "base.LCR-manual.LCR") Then
        If Not SameFormula(ws.Range("J1955").Value2, "manual.LCR-base.LCR") Then Err.Raise vbObjectError + 1730, , "Correction target changed: J1955"
        ws.Range("J1955").Value2 = "base.LCR-manual.LCR"
    End If
    log.Cells(74, 1).Value2 = "Deposit increase or withdrawn without change in liquid assets"
    log.Cells(74, 2).Value2 = "IMPACT_LCR"
    log.Cells(74, 3).Value2 = "manual.LCR-base.LCR"
    log.Cells(74, 4).Value2 = "base.LCR-manual.LCR"
    log.Cells(74, 5).Value2 = "ST Design - 10Jul2026 - shock type info.xlsx / Sheet2!Z148"
    log.Cells(74, 6).Value2 = "The source impact sign is base LCR minus stressed LCR."
    StyleReconTable log, 74, 6
    log.columns("A:B").ColumnWidth = 42
    log.columns("C:F").ColumnWidth = 65
    log.UsedRange.WrapText = True
    log.rows("8:74").RowHeight = 55
End Sub

Private Function SameFormula(ByVal a As Variant, ByVal b As Variant) As Boolean
    SameFormula = (StrComp(Replace(NormalizeConfigFormulaText(a), " ", ""), Replace(NormalizeConfigFormulaText(b), " ", ""), vbTextCompare) = 0)
End Function
