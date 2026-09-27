Attribute VB_Name = "modJKBDiag"
Option Explicit

' ============================================================================
'  Headless diagnostics / test harness.
'
'  Nothing in this module shows a dialog, so the whole pipeline can be driven
'  from a build script and asserted on. It is a development aid: it reads the
'  same code paths the ribbon uses, but reports to a text file instead of a
'  message box.
' ============================================================================

Private mFF As Integer
Private mTracePath As String

Private Sub w(ByVal s As String)
    Print #mFF, s
End Sub

Private Sub OpenLog(ByVal outPath As String)
    mFF = FreeFile
    mTracePath = outPath & ".trace"
    On Error Resume Next: Kill mTracePath: Err.Clear: On Error GoTo 0
    Open outPath For Output As #mFF
End Sub

' Writes one line and closes the handle immediately, so the marker survives even if the
' next statement takes the whole process down.
Private Sub Trace(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open mTracePath For Append As #f
    Print #f, format$(Now, "hh:nn:ss") & "  " & s
    Close #f
    Err.Clear
    On Error GoTo 0
End Sub

' What the breakdown actually produced, so "why is the pre-shock that much" can
' be checked against the extract rather than taken on trust.
Private Sub DumpBreakdown(Optional ByVal maxRows As Long = 22)
    Dim ws As Worksheet, lastRow As Long, r As Long, n As Long, line As String, c As Long
    w ""
    w "-- breakdown: why each figure is what it is"
    w "   switched on: " & BreakdownsEnabled()
    On Error Resume Next
    Set ws = BreakdownSheet()
    On Error GoTo 0
    If ws Is Nothing Then w "   (no breakdown sheet)": Exit Sub
    lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    w "   rows: " & BreakdownCount()
    If lastRow < PS_FIRST_ROW Then
        w "   " & SafeText(ws.Cells(PS_FIRST_ROW, 1).Value2)
        Exit Sub
    End If
    line = ""
    For c = 3 To 12
        line = line & SafeText(ws.Cells(PS_HDR_ROW, c).Value2) & " | "
    Next c
    w "   HDR: " & line
    For r = PS_FIRST_ROW To lastRow
        n = n + 1
        If n > maxRows Then w "   ... (" & (lastRow - PS_FIRST_ROW + 1) & " rows total)": Exit For
        line = ""
        For c = 3 To 12
            line = line & Left$(SafeText(ws.Cells(r, c).text), 26) & " | "
        Next c
        w "   " & line
    Next r
End Sub

' A handful of breakdown rows, so the smoke test exercises the Why sheet and its
' calculated share rather than the empty-table path.
Private Sub SeedBreakdownForTest()
    Dim ws As Worksheet, i As Long, dims As Variant, vals As Variant, d As Long, v As Long, r As Long
    On Error Resume Next
    Set ws = BreakdownSheet(True)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB"
    ws.Cells(PS_HDR_ROW, 1).Resize(1, 17).Value2 = Array( _
        "As-of date", "Entity", "Test case", "Test element", "Source", "Dimension", "Value", _
        "Rows selected", "Rows in portfolio", "Selected", "In portfolio", "Share selected", _
        "Selected (2nd)", "In portfolio (2nd)", "Measure", "Measure (2nd)", "Filter applied")
    dims = Array("IFRS STAGE", "SECTOR")
    vals = Array(Array("1", "2", "3"), Array("TRADE", "INDUSTRY", "RETAIL"))
    r = PS_FIRST_ROW
    For i = 0 To 2                                   ' three test cases
        For d = 0 To 1
            For v = 0 To 2
                ws.Cells(r, 1).Value2 = "2025-12-31"
                ws.Cells(r, 2).Value2 = "JKB_JORDAN"
                ws.Cells(r, 3).Value2 = "CR_T00" & (i + 1)
                ws.Cells(r, 4).Value2 = "EL_1"
                ws.Cells(r, 5).Value2 = "ECL"
                ws.Cells(r, 6).Value2 = dims(d)
                ws.Cells(r, 7).Value2 = vals(d)(v)
                ws.Cells(r, 8).Value2 = 100 + i * 10 + v * 7
                ws.Cells(r, 9).Value2 = 900 + i * 30 + v * 40
                ws.Cells(r, 10).Value2 = 250000 + i * 31000 + v * 17000
                ws.Cells(r, 11).Value2 = 1800000 + i * 90000 + v * 220000
                ws.Cells(r, 17).Value2 = "SECTOR IN ('X')"
                r = r + 1
            Next v
        Next d
    Next i
    Err.Clear
End Sub

' ---------------------------------------------- manual reports smoke test ----
'
' The pivot and slicer machinery is otherwise exercised only at the end of a very
' long pass, which is a terrible place to find out that a slicer silently failed
' to appear. This seeds a dozen derived figures, builds the real report workbook
' from them, and reports exactly what arrived - in about a second.
Public Sub JKB_DiagManual(ByVal outPath As String, ByVal folder As String)
    Dim ws As Worksheet, i As Long, path As String, wb As Workbook
    Dim nSheets As Long, nPivots As Long, nSlicers As Long, nShapes As Long, w2 As Worksheet

    OpenLog outPath
    w "=== manual reports smoke test ==="
    On Error GoTo Fatal
    Application.DisplayAlerts = False

    w "template: " & IIf(Len(ReportTemplatePath()) > 0, ReportTemplatePath(), "(not found)")
    w "   from: " & modShared_Template.TemplateOrigin()

    EnsurePreShockSheet
    Set ws = DerivedSheet(True)
    ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB"
    ws.Cells(DV_HEADER_ROW, 1).Resize(1, DV_COLS).Value2 = Array( _
        "As-of date", "Entity", "Entity ID", "Test case", "Element", "Metric", "Source", _
        "Base (derived)", "Pre-shock (derived)", "Rows matched", "Filter applied", "Status", _
        "Base (system)", "Pre-shock (system)", "Value to use", "Which one was used")
    ws.Range(ws.Cells(DV_FIRST_ROW, 1), ws.Cells(DV_FIRST_ROW + 40, 1)).NumberFormat = "@"
    For i = 0 To 11
        ws.Cells(DV_FIRST_ROW + i, 1).Value2 = "2025-12-31"
        ws.Cells(DV_FIRST_ROW + i, 2).Value2 = "JKB_JORDAN"
        ws.Cells(DV_FIRST_ROW + i, 3).Value2 = "1"
        ws.Cells(DV_FIRST_ROW + i, 4).Value2 = "CR_T00" & ((i Mod 3) + 1)
        ws.Cells(DV_FIRST_ROW + i, 5).Value2 = "EL_" & ((i Mod 2) + 1)
        ws.Cells(DV_FIRST_ROW + i, 6).Value2 = "OUTST_LCY_STAGE" & ((i Mod 3) + 1) & "_PRE_SHOCK"
        ws.Cells(DV_FIRST_ROW + i, 7).Value2 = IIf(i Mod 4 = 0, "CAPRWA", "ECL")
        ws.Cells(DV_FIRST_ROW + i, 8).Value2 = 1000000 + i * 137000
        ws.Cells(DV_FIRST_ROW + i, 9).Value2 = 250000 + i * 31000
        ws.Cells(DV_FIRST_ROW + i, 10).Value2 = 400 + i * 13
        ws.Cells(DV_FIRST_ROW + i, 11).Value2 = "SECTOR IN ('X')"
        ws.Cells(DV_FIRST_ROW + i, 12).Value2 = IIf(i Mod 5 = 0, "Break", "Matched")
        ws.Cells(DV_FIRST_ROW + i, 13).Value2 = 1000000 + i * 137000
        ws.Cells(DV_FIRST_ROW + i, 14).Value2 = 250000 + i * 31000 + IIf(i Mod 5 = 0, 900, 0)
        ws.Cells(DV_FIRST_ROW + i, 15).Value2 = USE_DERIVED
        ws.Cells(DV_FIRST_ROW + i, 16).Value2 = USE_DERIVED
    Next i
    w "seeded: " & PersistedDerivedCount() & " derived figure(s)"
    SeedBreakdownForTest
    w "seeded: " & BreakdownCount() & " breakdown row(s)"

    path = Ps_BuildManualReportsTo(folder)
    w "written: " & path

    Set wb = Workbooks.Open(path)
    nSheets = wb.Worksheets.count
    For Each w2 In wb.Worksheets
        nPivots = nPivots + w2.PivotTables.count
        nShapes = nShapes + w2.Shapes.count
    Next w2
    nSlicers = wb.SlicerCaches.count
    w "sheets: " & nSheets & "   pivots: " & nPivots & "   slicer caches: " & nSlicers & "   shapes: " & nShapes
    For Each w2 In wb.Worksheets
        w "   " & w2.name & "   pivots=" & w2.PivotTables.count & " shapes=" & w2.Shapes.count
    Next w2
    wb.Close SaveChanges:=False
    w "=== complete ==="
    Close #mFF
    Exit Sub
Fatal:
    w "!!! FATAL " & Err.Number & ": " & Err.description
    On Error Resume Next
    Close #mFF
End Sub

' -------------------------------------------------- configuration checks ----

Public Sub JKB_Diag(ByVal outPath As String)
    Dim a As Variant, headers As Object, hr As Long, idx As Object, stats As Object, rules As Object
    Dim tk As Variant, ek As Variant, tc As Object, el As Object, t As String, seen As Object
    Dim rl As Collection, ri As Object, rm As Object, sev As Long, f As String, conv As String
    Dim nRules As Long, nFail As Long, nOk As Long, ctx As String
    Dim grouping As Object, displayRules As Collection, totalRows As Long

    OpenLog outPath
    On Error GoTo Fatal
    w "=== JKB headless diagnostic ==="
    w "Workbook: " & ThisWorkbook.FullName

    w "-- ReadSource"
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, headers, hr
    w "   header row: " & hr & "  rows: " & UBound(a, 1) & "  cols: " & UBound(a, 2)

    w "-- IndexSource"
    Set idx = IndexSource(a, headers, hr, ThisWorkbook.date1904, stats)
    w "   testcase/date/entity combinations: " & idx.count
    w "   dates=" & stats("Dates").count & " entities=" & stats("Entities").count & _
      " cases=" & stats("Cases").count & " categories=" & stats("Categories").count

    w "-- LoadRulesByElementType"
    Set rules = LoadRulesByElementType(ThisWorkbook)
    w "   element types with enabled rules: " & rules.count
    Set grouping = LoadDetailGrouping(ThisWorkbook)

    w "-- element types present in the uploaded source"
    Set seen = NewMap()
    For Each tk In idx.keys
        Set tc = idx(tk)
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            t = SafeUpperText(el("ElementType"))
            If Not seen.Exists(t) Then
                seen(t) = True
                w "   [" & IIf(rules.Exists(t), "config OK ", "NO CONFIG") & "] " & el("ElementType")
            End If
        Next ek
    Next tk

    w ""
    w "-- per-rule formula conversion"
    ResetFormulaCache
    For Each tk In seen.keys
        t = CStr(tk)
        If rules.Exists(t) Then
            Set rl = rules(t)
            On Error Resume Next
            Set rm = Nothing
            Set rm = BuildOutputRowMap(rl, grouping, t, displayRules, totalRows)
            If Err.Number <> 0 Then
                w "  !! BuildOutputRowMap failed for " & t & " -> Error " & Err.Number & ": " & Err.description
                Err.Clear
                On Error GoTo Fatal
                GoTo NextType
            End If
            On Error GoTo Fatal
            For Each ri In rl
                If CBool(ri("ShowManual")) Then
                    For sev = 0 To 2
                        nRules = nRules + 1
                        ctx = t & " / " & ri("OutputRowLabel") & " / sev " & sev
                        f = "": conv = ""
                        On Error Resume Next
                        f = ManualFormulaText(ri, sev)
                        If Err.Number <> 0 Then
                            nFail = nFail + 1
                            w "  !! ManualFormulaText  " & ctx & "  Error " & Err.Number & ": " & Err.description
                            Err.Clear
                        ElseIf Len(f) > 0 Then
                            conv = ConvertManualFormula(f, rm, sev, ctx)
                            If Err.Number <> 0 Then
                                nFail = nFail + 1
                                w "  !! ConvertManualFormula  " & ctx
                                w "       Error " & Err.Number & ": " & Err.description
                                w "       source formula: " & f
                                Err.Clear
                            Else
                                nOk = nOk + 1
                            End If
                        End If
                        On Error GoTo Fatal
                    Next sev
                End If
            Next ri
        End If
NextType:
    Next tk

    w ""
    w "formula conversions attempted: " & nRules & "   ok: " & nOk & "   FAILED: " & nFail
    w "=== diagnostic complete ==="
    Close #mFF
    Exit Sub
Fatal:
    w ""
    w "!!! FATAL Error " & Err.Number & ": " & Err.description
    Close #mFF
End Sub

' ------------------------------------------------- end-to-end generation ----

' Generates output for the first `maxCases` test cases into `folder`.
' Pass maxCases = 0 to generate everything the upload contains.
Public Sub JKB_DiagGenerate(ByVal outPath As String, ByVal spec As String)
    Dim a As Variant, headers As Object, hr As Long, idx As Object, stats As Object
    Dim selections As Object, cats As Object, dates As Object, ents As Object, cases As Object
    Dim k As Variant, n As Long, maxCases As Long, folder As String, parts As Variant
    Dim report As Object, i As Long, res As Object, t0 As Single

    parts = Split(spec, "|")
    folder = CStr(parts(0))
    If UBound(parts) >= 1 Then maxCases = CLng(parts(1))

    OpenLog outPath
    On Error GoTo Fatal
    w "=== JKB headless generation ==="
    w "output folder: " & folder & "   max test cases: " & IIf(maxCases = 0, "all", CStr(maxCases))

    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, headers, hr
    Set idx = IndexSource(a, headers, hr, ThisWorkbook.date1904, stats)

    Set selections = NewMap()
    Set cats = NewMap(): Set dates = NewMap(): Set ents = NewMap(): Set cases = NewMap()
    For Each k In stats("Categories").keys: cats(CStr(k)) = True: Next k
    For Each k In stats("Dates").keys: dates(CStr(k)) = True: Next k
    For Each k In stats("Entities").keys: ents(CStr(k)) = True: Next k
    n = 0
    For Each k In stats("Cases").keys
        If maxCases > 0 And n >= maxCases Then Exit For
        cases(CStr(k)) = True: n = n + 1
    Next k
    selections.Add "Categories", cats
    selections.Add "Dates", dates
    selections.Add "Entities", ents
    selections.Add "Cases", cases
    w "selected test cases: " & cases.count

    t0 = Timer
    Set report = GenerateOutputs(selections, folder, False)
    w ""
    w "elapsed seconds: " & format$(Timer - t0, "0.0")
    w "Ok        : " & report("Ok")
    w "Stage     : " & report("Stage")
    w "ErrorNo   : " & report("ErrorNumber")
    w "ErrorText : " & report("ErrorText")
    w "Selected  : " & report("Selected")
    w "Generated : " & report("Generated")
    w "Failed    : " & report("Failed")
    w "Cancelled : " & report("Cancelled")
    w "Root      : " & report("Root")
    w ""
    w "problems (" & report("Problems").count & "):"
    For i = 1 To report("Problems").count
        w "   " & CStr(report("Problems")(i))
        If i >= 40 Then w "   ...": Exit For
    Next i
    w ""
    w "results (" & report("Results").count & "):"
    For i = 1 To report("Results").count
        Set res = report("Results")(i)
        w "   " & IIf(CBool(res("Succeeded")), "OK   ", "FAIL ") & res("TestCase") & " | " & res("Date") & " | elements=" & res("Elements") & " | " & res("Message")
        If i >= 60 Then w "   ...": Exit For
    Next i
    w "=== generation complete ==="
    Close #mFF
    Exit Sub
Fatal:
    w ""
    w "!!! FATAL Error " & Err.Number & ": " & Err.description
    On Error Resume Next
    JKB_Busy = False
    Close #mFF
End Sub

' ---------------------------------------------------- ECL pre-shock path ----

' Drives the whole independent-check path without any dialog: load an ECL extract,
' rebuild the test-case list, run the checks, and report what the sheet ends up holding.
Public Sub JKB_DiagEcl(ByVal outPath As String, ByVal eclPath As String)
    Dim a As Variant, h As Object, hr As Long, idx As Object, stats As Object
    Dim cache As Object, notes As Collection, ws As Worksheet, r As Long, i As Long, n As Long
    Dim parts As Variant

    parts = Split(eclPath, "|")
    OpenLog outPath
    On Error GoTo Fatal
    w "=== JKB ECL pre-shock diagnostic ==="
    w "ECL file: " & CStr(parts(0))

    EnsurePreShockSheet
    Set ws = PreShockSheet()
    w "pre-shock sheet: " & ws.name

    ' parts: eclPath [| Ignore] [| caprwaPath]
    Trace "before LoadEclOutput"
    LoadEclOutput CStr(parts(0)), False, "ECL"
    Trace "after ECL load"
    If UBound(parts) >= 2 Then
        If Len(Trim$(CStr(parts(2)))) > 0 Then
            LoadEclOutput CStr(parts(2)), False, "CAPRWA"
            Trace "after CAPRWA load"
            w "CAPRWA loaded from " & CStr(parts(2))
        End If
    End If
    ' Append "|Ignore" to the argument to exercise the path where the extract's reporting
    ' date differs from the upload's - the usual case for a trimmed sample file.
    If UBound(parts) >= 1 Then
        If SafeUpperText(parts(1)) = "IGNORE" Then
            SetAsOfMatching "Ignore"
            w "as-of date matching: Ignore (requested by the caller)"
        End If
    End If
    w ""
    w "-- block 1: source"
    For r = 6 To 10
        w "   A" & r & "=" & SafeText(ws.Cells(r, 1).Value2) & " | B=" & Left$(SafeText(ws.Cells(r, 2).Value2), 120) & _
          " | E=" & SafeText(ws.Cells(r, 5).Value2) & " | F=" & SafeText(ws.Cells(r, 6).Value2) & _
          " | H=" & SafeText(ws.Cells(r, 8).Value2) & " | I=" & Left$(SafeText(ws.Cells(r, 9).Value2), 80)
    Next r

    Trace "ReadSource"
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Trace "IndexSource"
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    Trace "SyncPreShockTestCases"
    SyncPreShockTestCases idx
    Trace "PrepareBaseTestCache"
    Set cache = PrepareBaseTestCache(idx, Nothing, a, h, notes)
    Trace "PrepareBaseTestCache done: " & cache.count
    w ""
    w "independent metric results cached: " & cache.count
    DumpBreakdown
    ' Same step the generator and the Run tests button take: make the configuration read the
    ' derived figures instead of copying the system output.
    w "config rows repointed at derived values: " & MigratePreShockConfigTokens(False)
    w "notes: " & IIf(notes Is Nothing, 0, notes.count)
    If Not notes Is Nothing Then
        For i = 1 To notes.count
            w "   " & CStr(notes(i))
            If i >= 25 Then w "   ...": Exit For
        Next i
    End If

    w ""
    w "-- block 2: formulas (metric | source | purpose | availability)"
    Set ws = PsMetricsSheet()
    n = 0
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then Exit For
        n = n + 1
        w "   " & SafeText(ws.Cells(r, 1).Value2) & " | " & SafeText(ws.Cells(r, 2).Value2) & _
          " | " & SafeText(ws.Cells(r, 6).Value2) & " | " & Left$(SafeText(ws.Cells(r, 7).Value2), 90)
    Next r
    w "   (" & n & " metric rows)"

    w ""
    w "-- block 3: dimensions (name | availability | distinct values)"
    Set ws = PsFieldsSheet()
    For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then Exit For
        w "   " & SafeText(ws.Cells(r, 1).Value2) & " | " & SafeText(ws.Cells(r, 4).Value2) & " | " & SafeText(ws.Cells(r, 5).Value2)
    Next r

    w ""
    w "-- block 4: test cases (first 20: case | element | source | mapping status | rows | last result)"
    Set ws = PsCasesSheet()
    n = 0
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
        n = n + 1
        If n <= 20 Then
            w "   " & SafeText(ws.Cells(r, 2).Value2) & " | " & SafeText(ws.Cells(r, 3).Value2) & " | " & _
              SafeText(ws.Cells(r, 5).Value2) & " | " & SafeText(ws.Cells(r, 9).Value2) & " | " & _
              SafeText(ws.Cells(r, 10).Value2) & " | " & SafeText(ws.Cells(r, 11).Value2)
        End If
    Next r
    w "   (" & n & " test-case rows)"

    w ""
    w "-- block 5: results from " & DERIVED_SHEET & " (first 25: element | metric | BASE sys/derived | PRE sys/derived | reconciles)"
    Set ws = DerivedSheet()
    n = 0
    If Not ws Is Nothing Then
        For r = DV_FIRST_ROW To DV_FIRST_ROW + 4000
            If Len(SafeText(ws.Cells(r, 6).Value2)) = 0 Then Exit For
            n = n + 1
            If n <= 25 Then
                w "   " & SafeText(ws.Cells(r, 5).Value2) & " | " & SafeText(ws.Cells(r, 6).Value2) & _
                  " | base sys=" & SafeText(ws.Cells(r, 13).Value2) & " drv=" & SafeText(ws.Cells(r, 8).Value2) & _
                  " | pre sys=" & SafeText(ws.Cells(r, 14).Value2) & " drv=" & SafeText(ws.Cells(r, 9).Value2) & _
                  " | " & SafeText(ws.Cells(r, 19).Value2)
            End If
        Next r
    End If
    w "   (" & n & " result rows)"

    DumpDerivedValues
    DumpConfigPreShockRows
    w "=== ECL diagnostic complete ==="
    Close #mFF
    Exit Sub
Fatal:
    w ""
    w "!!! FATAL Error " & Err.Number & ": " & Err.description
    On Error Resume Next
    JKB_Busy = False
    Close #mFF
End Sub

' The derived-values store: what base and pre-shock actually came out as, per source.
Private Sub DumpDerivedValues()
    Dim ws As Worksheet, r As Long, n As Long, lastRow As Long
    On Error Resume Next
    w ""
    w "-- block 6: " & DERIVED_SHEET & " (first 30)"
    Set ws = DerivedSheet()
    If ws Is Nothing Then w "   (sheet not created)": Exit Sub
    lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    For r = DV_FIRST_ROW To lastRow
        If Len(SafeText(ws.Cells(r, 6).Value2)) = 0 Then Exit For
        n = n + 1
        If n <= 30 Then
            w "   " & SafeText(ws.Cells(r, 5).Value2) & " | " & SafeText(ws.Cells(r, 6).Value2) & _
              " | src=" & SafeText(ws.Cells(r, 7).Value2) & _
              " | base=" & SafeText(ws.Cells(r, 8).Value2) & _
              " | pre=" & SafeText(ws.Cells(r, 9).Value2) & _
              " | rows=" & SafeText(ws.Cells(r, 10).Value2)
        End If
    Next r
    w "   (" & n & " derived rows)"
End Sub

' Proves whether the configuration now reads the derived value or still copies the system.
Private Sub DumpConfigPreShockRows()
    Dim lo As ListObject, h As Object, hdr As Variant, c As Long, a As Variant, r As Long
    Dim label As String, f As String, n As Long, drv As Long, sysCopy As Long
    On Error Resume Next
    w ""
    w "-- block 7: config MANUAL_FORMULA_DEFAULT for pre-shock rows (first 20)"
    Set lo = ThisWorkbook.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    If lo Is Nothing Then w "   (rules table not found)": Exit Sub
    Set h = NewMap()
    hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    a = lo.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        label = SafeUpperText(a(r, h("OUTPUT_ROW_LABEL")))
        If InStr(1, label, "_PRE_SHOCK", vbTextCompare) > 0 Then
            f = SafeText(a(r, h("MANUAL_FORMULA_DEFAULT")))
            n = n + 1
            If InStr(1, f, "derived.", vbTextCompare) > 0 Then drv = drv + 1
            If InStr(1, f, "systemoutput.", vbTextCompare) > 0 Then sysCopy = sysCopy + 1
            If n <= 20 Then w "   " & SafeText(a(r, h("ELEMENT_TYPE"))) & " | " & label & " -> " & f
        End If
    Next r
    w "   (" & n & " pre-shock config rows: " & drv & " read a derived value, " & sysCopy & " still copy the system output)"
End Sub

' Dumps the Build_Log so a harness can assert on what the run recorded.
Public Sub JKB_DumpBuildLog(ByVal outPath As String)
    Dim ws As Worksheet, lastRow As Long, r As Long
    OpenLog outPath
    On Error GoTo Fatal
    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    w "=== Build_Log (" & (lastRow - 3) & " entries) ==="
    For r = 4 To lastRow
        w SafeText(ws.Cells(r, 1).Value2) & " | " & SafeText(ws.Cells(r, 2).Value2) & " | " & _
          SafeText(ws.Cells(r, 3).Value2) & " | " & SafeText(ws.Cells(r, 4).Value2)
    Next r
    Close #mFF
    Exit Sub
Fatal:
    w "!!! FATAL Error " & Err.Number & ": " & Err.description
    Close #mFF
End Sub

' Builds the console payload without opening a window, so the JSON can be checked.
Public Sub JKB_DiagConsole(ByVal outPath As String)
    Dim paths As String, jsPath As String, fso As Object, src As Object, txt As String
    OpenLog outPath
    On Error GoTo Fatal
    w "=== console payload ==="
    paths = OpenStressConsole_Paths()
    w "paths: " & paths
    jsPath = Mid$(paths, InStr(paths, "|") + 1)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set src = fso.OpenTextFile(jsPath, 1, False, -1)
    txt = src.ReadAll
    src.Close
    w "payload bytes: " & Len(txt)
    w txt
    w "=== end ==="
    Close #mFF
    Exit Sub
Fatal:
    w "!!! FATAL " & Err.Number & ": " & Err.description
    On Error Resume Next
    Close #mFF
End Sub
' Drives the folder upload headlessly and reports what it loaded and how fast.
Public Sub JKB_DiagFolder(ByVal outPath As String, ByVal folder As String)
    Dim t0 As Single, el As Single, ws As Worksheet, r As Long, c As Long, line As String
    OpenLog outPath
    On Error GoTo Fatal
    w "=== folder upload ==="
    w "folder: " & folder
    t0 = Timer
    UploadFolderHeadless folder
    el = Timer - t0
    If el < 0 Then el = el + 86400   ' the run crossed midnight
    w "elapsed: " & format$(el, "0.0") & "s"
    w "loader:  " & Replace(LastFolderLoadSummary(), vbCrLf, vbCrLf & "         ")
    w ""
    w "-- sources"
    Set ws = PsSourcesSheet()
    For r = PS_FIRST_ROW To PS_FIRST_ROW + 8
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then Exit For
        line = ""
        For c = 1 To 8
            line = line & Left$(SafeText(ws.Cells(r, c).text), 48) & " | "
        Next c
        w "   " & line
    Next r
    w "=== end ==="
    Close #mFF
    Exit Sub
Fatal:
    w "!!! FATAL " & Err.Number & ": " & Err.description
    On Error Resume Next
    JKB_Busy = False
    Close #mFF
End Sub
