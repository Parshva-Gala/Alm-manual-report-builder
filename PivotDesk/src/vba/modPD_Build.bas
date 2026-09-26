Option Explicit

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
    Dim Path As String, t0 As Single

    t0 = Timer
    Step_ FwLabel(fw) & " - opening a new workbook"
    Set wb = Workbooks.Add(xlWBATWorksheet)

    modPD_Pivot.ResetPivots
    Step_ FwLabel(fw) & " - staging the output"
    Set lo = modPD_Stage.StageFramework(fw, wb, errOut)
    If lo Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        Err.Clear
        Exit Function
    End If

    modPD_Pivot.UseCache wb, lo

    Step_ FwLabel(fw) & " - the Output pivot"
    modPD_Pivot.BuildOutputSheet wb, fw
    Step_ FwLabel(fw) & " - the Balance sheet pivot"
    modPD_Pivot.BuildBalanceSheet wb, fw

    If SplitsByCurrency(fw) Then
        nSheets = CurrencySheets(wb, capped)
    Else
        nSheets = RuleSheets(wb, capped)
    End If

    Guide wb, fw, nSheets, capped
    Tidy wb

    Path = PathJoin(outFolder, SafeFileName(FwLabel(fw)) & ".xlsx")
    Step_ FwLabel(fw) & " - saving"
    On Error GoTo SaveFailed
    Application.DisplayAlerts = False
    wb.SaveAs Path, 51            ' xlOpenXMLWorkbook - no macros, opens anywhere
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    On Error GoTo 0

    LogIt V_OK, "Pivots", Fmt(modPD_Stage.StagedRows()) & " row(s) staged, " & _
          (nSheets + 2) & " pivot sheet(s), " & Format$(Timer - t0, "0.0") & "s.", FwLabel(fw)
    BuildFramework = Path
    Exit Function

SaveFailed:
    errOut = "could not save: " & Err.Description
    On Error Resume Next
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    Err.Clear
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

Private Sub Guide(ByVal wb As Workbook, ByVal fw As String, ByVal nSheets As Long, ByVal capped As Boolean)
    Dim ws As Worksheet, made As Collection, i As Long, r As Long, e As Variant

    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    On Error Resume Next
    ws.Name = SH_GUIDE
    Err.Clear
    On Error GoTo 0

    modPD_Theme.Dress ws, FwLabel(fw) & "  -  " & BANK_NAME, _
        "Built by " & TOOL_NAME & " " & TOOL_VERSION & " on " & Format$(Now, "dd mmm yyyy hh:nn") & _
        IIf(Len(modPD_Stage.StagedAsOf()) > 0, "   -   data as of " & modPD_Stage.StagedAsOf(), "")

    modPD_Theme.SetStatus ws, Fmt(modPD_Stage.StagedRows()) & " rows staged into one pivot cache.  " & _
        "Every sheet below is a live PivotTable over it - drag a field, drop a slicer, drill a total.", "OK"

    modPD_Theme.Head ws, Array("Sheet", "What is on it", "", "", ""), Array(34, 62, 2, 2, 2)

    r = modPD_Theme.R_FIRST
    Set made = modPD_Pivot.MadeSheets()
    For i = 1 To made.count
        e = made(i)
        ws.Cells(r, 1).Value2 = CStr(e(0))
        ws.Cells(r, 2).Value2 = CStr(e(1))
        On Error Resume Next
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="", _
                          SubAddress:="'" & CStr(e(0)) & "'!A1", ScreenTip:="Open " & CStr(e(0))
        Err.Clear
        On Error GoTo 0
        r = r + 1
    Next i
    modPD_Theme.DressTable ws, 2, r - 1

    ' What the tool had to decide for itself goes here, where it is read, not
    ' into a log nobody opens.
    r = r + 2
    Note ws, r, "How to read this", ""
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
    Note ws, r, "Blanks", "Rows with no bucket are excluded from the bucket filter by default. " & _
        "They are still in the data - clear the filter to see them."
    If capped Then
        r = r + 1
        Note ws, r, "Sheet cap", "Stopped at " & MAX_RULE_SHEETS & " sheets. The rest are not here; " & _
            "everything is still in the Output and Balance sheet pivots."
    End If

    ws.Columns(1).ColumnWidth = 34
    ws.Columns(2).ColumnWidth = 96
End Sub

Private Sub Note(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal body As String)
    ws.Cells(r, 1).Value2 = label
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = modPD_Theme.C_INK
    ws.Cells(r, 2).Value2 = body
    ws.Cells(r, 2).Font.Color = modPD_Theme.C_MUTED
    ws.Cells(r, 2).WrapText = False
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
