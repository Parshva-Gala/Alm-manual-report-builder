Attribute VB_Name = "modPD_LayoutAcceptance"
Option Explicit

' Optional native acceptance harness. Never included in the release workbook.
' Inject this module and chart_acceptance.bas into a disposable *QA*.xlsm.
' Uses generated synthetic balances only and writes beside that QA workbook.
Private mLayoutLog As String
Private mLayoutChecks As Long

Public Sub PD_LayoutAcceptance()
    Dim wb As Workbook, stage As Worksheet, lo As ListObject, balance As Worksheet
    Dim output As Worksheet, guide As Worksheet, dates As Worksheet
    Dim oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean, oldQuiet As Boolean
    Dim logPath As String, outputPath As String, why As String
    If InStr(1, ThisWorkbook.Name, "QA", vbTextCompare) = 0 Then Exit Sub
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub
    logPath = ThisWorkbook.Path & Application.PathSeparator & "layout-acceptance.log"
    outputPath = LayoutOutputPath(ThisWorkbook.Path)
    mLayoutLog = "NATIVE EXCEL LAYOUT ACCEPTANCE" & vbCrLf & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    mLayoutChecks = 0
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldQuiet = PD_Quiet
    On Error GoTo Failed
    LayoutWriteLog logPath
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    PD_Quiet = True
    LayoutProgress "Creating synthetic staging table", logPath
    Set wb = Workbooks.Add(xlWBATWorksheet)
    modPD_Theme.DarkNormal wb
    Set stage = wb.Worksheets(1)
    stage.Name = SH_STAGE
    Set lo = LayoutStage(stage)
    modPD_Pivot.ResetPivots
    modPD_Pivot.SetBook "SYNTHETIC LAYOUT QA"
    modPD_Pivot.UseCache wb, lo
    LayoutProgress "Generating actual Balance sheet", logPath
    Set balance = modPD_Pivot.BuildBalanceSheet(wb, FW_LCR)
    LayoutAssert Not balance Is Nothing, "Balance sheet generator returned a sheet"
    LayoutProgress "Generating actual Output sheet", logPath
    Set output = modPD_Pivot.BuildOutputSheet(wb, FW_LCR)
    LayoutAssert Not output Is Nothing, "Output generator returned a sheet"
    LayoutAssert balance.PivotTables.count = 1, "Balance sheet contains a native pivot"
    LayoutAssert output.PivotTables.count = 1, "Output sheet contains a native pivot"
    LayoutAssert LayoutHashCount(balance.PivotTables(1).TableRange2) = 0, "Freshly generated balance values fit"
    LayoutAssert LayoutHashCount(output.PivotTables(1).TableRange2) = 0, "Freshly generated output values fit"
    LayoutProgress "Refitting large contra balances", logPath
    LayoutCheckReport balance, "#,##0.00;[Red](#,##0.00);-", "Large contra balances"
    LayoutProgress "Refitting twelve decimal places", logPath
    LayoutCheckReport output, "0.000000000000;[Red](0.000000000000);-", "Twelve decimal places"
    LayoutProgress "Refitting long currency format beyond former width cap", logPath
    LayoutCheckReport balance, """SYNTHETIC CURRENCY ""#,##0.000000;[Red](""SYNTHETIC CURRENCY ""#,##0.000000);-", _
                      "Long currency format beyond the old width cap"
    LayoutAssert balance.PivotTables(1).DataBodyRange.Columns(1).ColumnWidth > 30, _
                 "Formatted value columns can exceed the former width cap"
    LayoutProgress "Generating and refitting long date labels", logPath
    Set dates = LayoutDatePivot(wb, lo)
    LayoutCheckReport dates, "#,##0.000000;[Red](#,##0.000000);-", "Long date and time labels"
    ' Keep one native filter active while refitting, then compare every result.
    LayoutProgress "Refitting while native report filter is selected", logPath
    balance.PivotTables(1).PivotFields(H_BUCKET).CurrentPage = "UPTO 1 MONTH"
    LayoutCheckReport balance, "#,##0.00;[Red](#,##0.00);-", "Active report filter"
    LayoutAssert CStr(balance.PivotTables(1).PivotFields(H_BUCKET).CurrentPage) = "UPTO 1 MONTH", _
                 "Fit preserves the selected report filter"
    Set guide = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    guide.Name = SH_GUIDE
    modPD_Theme.Dress guide, "Synthetic layout acceptance", _
        "Generated fixtures only. Includes large contra amounts, precise decimals, dates and charts.", "NATIVE EXCEL QA"
    modPD_Theme.BookBar guide, "SYNTHETIC LAYOUT QA", False
    guide.Columns(1).ColumnWidth = 38
    guide.Columns(2).ColumnWidth = 110
    LayoutProgress "Generating and checking actual charts", logPath
    mLayoutLog = mLayoutLog & ChartAcceptance(wb, guide) & vbCrLf
    modPD_Pivot.LinkSiblings wb
    stage.visible = xlSheetVeryHidden
    modPD_Theme.PrintReady guide, 2
    guide.Activate
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayGridlines = False
    LayoutProgress "Saving synthetic acceptance output", logPath
    wb.SaveAs outputPath, 51
    mLayoutLog = mLayoutLog & "OUTPUT " & outputPath & vbCrLf
    mLayoutLog = mLayoutLog & "PASS " & mLayoutChecks & " pivot/layout assertions; chart assertions above." & vbCrLf
    GoTo Done
Failed:
    why = CStr(Err.Number) & " " & Err.Description
    Err.Clear
    mLayoutLog = mLayoutLog & "FAIL " & why & vbCrLf
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.SaveAs outputPath, 51
        mLayoutLog = mLayoutLog & "PARTIAL OUTPUT " & outputPath & vbCrLf
    End If
Done:
    On Error Resume Next
    LayoutWriteLog logPath
    PD_Quiet = oldQuiet
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    Err.Clear
End Sub

Private Function LayoutStage(ByVal ws As Worksheet) As ListObject
    Dim headers As Variant, values As Variant, data(1 To 8, 1 To 14) As Variant
    Dim r As Long, c As Long, lo As ListObject
    headers = StageHeadings()
    For c = LBound(headers) To UBound(headers)
        ws.Cells(1, c + 1).Value2 = headers(c)
    Next c
    ws.Cells(1, 14).Value2 = "Fixture date"
    ws.Columns(C_FACTOR).NumberFormat = "@"
    values = Array(9876543210123.45, -9876543210123.45, 0.000000123456, -0.000000123456, _
                   123456789.125, -123456789.125, 43210.5, -43210.5)
    For r = 1 To 8
        data(r, C_RULE_ORDER) = r
        data(r, C_RULE_CAT) = IIf(r <= 4, "Synthetic liquidity", "Synthetic funding")
        data(r, C_RULE_NAME) = "Synthetic rule " & r
        data(r, C_FACTOR) = "50%"
        data(r, C_TYPE) = IIf(r Mod 2 = 1, "Assets", "Liabilities")
        data(r, C_LINE) = "Synthetic line " & IIf(r <= 4, "A", "B")
        data(r, C_SUBLINE) = "Synthetic subline"
        data(r, C_COA_NAME) = "Synthetic account " & Format$(r, "00")
        data(r, C_CURRENCY) = IIf(r Mod 2 = 1, "EGP", "USD")
        data(r, C_CCYCLASS) = IIf(r Mod 2 = 1, "LCY", "FCY")
        data(r, C_BUCKET) = IIf(r <= 4, "UPTO 1 MONTH", "1 - 3 MONTHS")
        data(r, C_PRE) = values(r - 1)
        data(r, C_POST) = CDbl(values(r - 1)) / 2
        data(r, 14) = CDbl(DateSerial(2026, 9, r) + TimeSerial(15, 44, 59))
    Next r
    ws.Range(ws.Cells(2, 1), ws.Cells(9, 14)).Value2 = data
    ws.Range(ws.Cells(2, C_PRE), ws.Cells(9, C_POST)).NumberFormat = "#,##0.000000"
    ws.Range(ws.Cells(2, 14), ws.Cells(9, 14)).NumberFormat = "dddd, mmmm d, yyyy hh:mm:ss"
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(9, 14)), , xlYes)
    lo.Name = "SyntheticLayoutData"
    Set LayoutStage = lo
End Function

Private Function LayoutDatePivot(ByVal wb As Workbook, ByVal lo As ListObject) As Worksheet
    Dim ws As Worksheet, pt As PivotTable, df As PivotField
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = "Date formats"
    modPD_Theme.Dress ws, "Date and time formatting", "Synthetic date labels and amounts.", "NATIVE EXCEL QA"
    modPD_Theme.BookBar ws, "SYNTHETIC LAYOUT QA", True
    Set pt = wb.PivotCaches(1).CreatePivotTable(TableDestination:=ws.Range("A10"), TableName:="qa_dates")
    pt.PivotFields("Fixture date").Orientation = xlRowField
    pt.PivotFields("Fixture date").Position = 1
    pt.PivotFields("Fixture date").NumberFormat = "dddd, mmmm d, yyyy hh:mm:ss"
    Set df = pt.AddDataField(pt.PivotFields(H_PRE), "Fixture balance", xlSum)
    pt.RowAxisLayout xlTabularRow
    pt.DisplayFieldCaptions = True
    pt.ManualUpdate = False
    Set LayoutDatePivot = ws
End Function

Private Sub LayoutCheckReport(ByVal ws As Worksheet, ByVal numberFormat As String, ByVal label As String)
    Dim pt As PivotTable, df As PivotField, before As Variant, after As Variant
    Dim beforeHashes As Long, afterHashes As Long, selected As String, z As Long, topRow As Long, leftCol As Long
    Set pt = ws.PivotTables(1)
    For Each df In pt.DataFields
        df.NumberFormat = numberFormat
    Next df
    ws.Calculate
    before = pt.TableRange2.Value2
    pt.TableRange2.Columns.ColumnWidth = 8
    beforeHashes = LayoutHashCount(pt.TableRange2)
    LayoutAssert beforeHashes > 0, label & ": deliberately narrow fixture reproduces hashes"
    ws.Activate
    pt.TableRange2.Cells(1, 1).Select
    ActiveWindow.Zoom = 85
    selected = ActiveCell.Address
    z = ActiveWindow.Zoom
    topRow = ActiveWindow.ScrollRow
    leftCol = ActiveWindow.ScrollColumn
    modPD_Pivot.FitReportValues ws
    after = pt.TableRange2.Value2
    afterHashes = LayoutHashCount(pt.TableRange2)
    LayoutAssert afterHashes = 0, label & ": no numeric display contains hashes after fitting"
    LayoutAssert LayoutSameValues(before, after), label & ": every pivot value and total is unchanged"
    LayoutAssert pt.DisplayFieldCaptions, label & ": native field captions remain available"
    LayoutAssert ActiveWindow.DisplayHeadings, label & ": native column headings remain available"
    LayoutAssert ActiveCell.Address = selected And ActiveWindow.Zoom = z And _
                 ActiveWindow.ScrollRow = topRow And ActiveWindow.ScrollColumn = leftCol, _
                 label & ": selected cell, zoom and scroll remain unchanged"
    mLayoutLog = mLayoutLog & "  hashes before=" & beforeHashes & "; after=" & afterHashes & vbCrLf
End Sub

Private Function LayoutHashCount(ByVal rng As Range) As Long
    Dim cell As Range, v As Variant
    For Each cell In rng.Cells
        v = cell.Value2
        If Not IsError(v) And Not IsEmpty(v) Then
            If IsNumeric(v) Then
                If InStr(1, cell.Text, "###", vbBinaryCompare) > 0 Then LayoutHashCount = LayoutHashCount + 1
            End If
        End If
    Next cell
End Function

Private Function LayoutSameValues(ByVal before As Variant, ByVal after As Variant) As Boolean
    Dim r As Long, c As Long
    If Not IsArray(before) Or Not IsArray(after) Then Exit Function
    If UBound(before, 1) <> UBound(after, 1) Or UBound(before, 2) <> UBound(after, 2) Then Exit Function
    For r = LBound(before, 1) To UBound(before, 1)
        For c = LBound(before, 2) To UBound(before, 2)
            If VarType(before(r, c)) <> VarType(after(r, c)) Then Exit Function
            If CStr(before(r, c)) <> CStr(after(r, c)) Then Exit Function
        Next c
    Next r
    LayoutSameValues = True
End Function

Private Sub LayoutAssert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 930, "Layout acceptance", message
    mLayoutChecks = mLayoutChecks + 1
    mLayoutLog = mLayoutLog & "OK " & message & vbCrLf
End Sub

Private Function LayoutOutputPath(ByVal folder As String) As String
    Dim path As String, suffix As Long
    path = folder & Application.PathSeparator & "layout-acceptance-" & Format$(Now, "yyyymmdd-hhnnss")
    LayoutOutputPath = path & ".xlsx"
    Do While Len(Dir$(LayoutOutputPath)) > 0
        suffix = suffix + 1
        LayoutOutputPath = path & "-" & suffix & ".xlsx"
    Loop
End Function

Private Sub LayoutProgress(ByVal message As String, ByVal logPath As String)
    mLayoutLog = mLayoutLog & "START " & Format$(Now, "hh:nn:ss") & " " & message & vbCrLf
    LayoutWriteLog logPath
End Sub

Private Sub LayoutWriteLog(ByVal path As String)
    Dim fh As Integer
    fh = FreeFile
    Open path For Output As #fh
    Print #fh, mLayoutLog
    Close #fh
End Sub
