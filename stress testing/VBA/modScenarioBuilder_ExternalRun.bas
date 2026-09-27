Attribute VB_Name = "modScenarioBuilder_ExternalRun"
Option Explicit

' ============================================================================
'  Generation engine.
'
'  GenerateOutputs is headless: it never shows a dialog, never leaves Excel in a
'  modified state, and reports everything through the returned report object.
'  RunExternalScenarioOutput is the thin interactive wrapper the ribbon calls.
'  Keeping the two apart is what makes the generator testable from a harness and
'  keeps a single, auditable cleanup path.
' ============================================================================

Private Const MAX_REPORTED_PROBLEMS As Long = 15

' Detail grouping - see the grouping section below.
Public Const GROUP_UNASSIGNED As String = "UNASSIGNED"
Public Const TABLE_GROUP_RULES As String = "tblGroupRules"
Public Const GROUP_DEF_FIRST_ROW As Long = 7
Public Const GROUP_DEF_KEY_COL As Long = 25          ' Y: GROUP_KEY .. AE: OUTLINE_LEVEL
Public Const GROUP_POLICY_CELL As String = "AA5"
Public Const GROUP_POLICY_STOP As String = "Stop and list them"
Public Const GROUP_POLICY_SHOW As String = "Show under Needs a group"

' ---------------------------------------------------------------- report ----

Private Function NewReport() As Object
    Dim r As Object
    Set r = NewMap()
    r("Ok") = False
    r("Stage") = ""
    r("ErrorNumber") = 0
    r("ErrorText") = ""
    r("Selected") = 0
    r("Generated") = 0
    r("Failed") = 0
    r("Root") = ""
    r("Cancelled") = False
    r("Repointed") = 0
    ' Problems stop a run: they mean the configuration cannot produce correct numbers.
    ' EclNotes never stop a run: they mean the independent second opinion was unavailable.
    Set r("Problems") = New Collection
    Set r("EclNotes") = New Collection
    Set r("Results") = New Collection
    Set NewReport = r
End Function

Private Sub AddProblem(ByVal report As Object, ByVal text As String)
    report("Problems").Add text
End Sub

' Renders at most MAX_REPORTED_PROBLEMS problems for a dialog; the Build_Log holds them all.
Public Function ProblemSummary(ByVal report As Object) As String
    Dim i As Long, s As String, c As Collection
    Set c = report("Problems")
    For i = 1 To c.count
        If i > MAX_REPORTED_PROBLEMS Then
            s = s & vbCrLf & "... and " & (c.count - MAX_REPORTED_PROBLEMS) & " more - see Build_Log."
            Exit For
        End If
        s = s & vbCrLf & "  - " & CStr(c(i))
    Next i
    ProblemSummary = s
End Function

' ---------------------------------------------------------------- engine ----

Public Function GenerateOutputs(ByVal selections As Object, _
                                Optional ByVal outputFolder As String = "", _
                                Optional ByVal allowPrompts As Boolean = True) As Object
    Dim report As Object
    Dim a As Variant, headers As Object, hr As Long, idx As Object, stats As Object
    Dim rules As Object, grouping As Object, baseCache As Object, state As Object
    Dim tk As Variant, tc As Object, root As String, runId As String, folder As String, file As String, baseFolder As String
    Dim count As Long, Done As Long, Failed As Long, result As Object
    Dim start As Date, suffix As Long, busyTaken As Boolean, stateTaken As Boolean
    Dim eclNotes As Collection, repointError As String

    Set report = NewReport()
    If JKB_Busy Then
        report("Stage") = "starting"
        report("ErrorText") = "Another JKB operation is still running. Wait for it to finish and try again."
        Set GenerateOutputs = report
        Exit Function
    End If

    On Error GoTo Failed

    report("Stage") = "reading the testcase source"
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, headers, hr

    report("Stage") = "indexing the testcase source"
    Set idx = IndexSource(a, headers, hr, ThisWorkbook.date1904, stats)

    report("Stage") = "loading calculation rules"
    Set rules = LoadRulesByElementType(ThisWorkbook)

    report("Stage") = "loading detail grouping"
    Set grouping = LoadDetailGrouping(ThisWorkbook)

    report("Stage") = "validating workbook view settings"
    Sheet1.ValidateViewSettings

    report("Stage") = "matching the selection"
    For Each tk In idx.keys
        If IsSelected(idx(tk), selections) Then count = count + 1
    Next tk
    report("Selected") = count
    If count = 0 Then Err.Raise vbObjectError + 600, , "No available records match all four selections. Widen the selection and try again."

    ' The log is the audit trail for this run, so reset it only once we know there is work.
    ResetBuildLog ThisWorkbook
    WarnOnce "__RESET__", "", ""
    ResetFormulaCache

    report("Stage") = "validating selected rules"
    ValidateSelectedRules idx, selections, rules, headers, grouping, report
    If report("Problems").count > 0 Then
        report("ErrorText") = report("Problems").count & " configuration problem(s) must be fixed before generation can run."
        GoTo Cleanup
    End If

    ' The independent ECL check is a second opinion, not a precondition. It has its own
    ' problem channel: if the ECL side cannot run, the affected rows fall back to their
    ' configured manual formula and the run continues.
    report("Stage") = "preparing the independent pre-shock cache"
    ' Quiet from here: the cache writes to several sheets, and repainting them was
    ' most of what this step cost. Calculation stays as it is until the session is
    ' prepared below, because the value-source columns are formulas.
    Set state = CaptureState(): stateTaken = True
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Set baseCache = PrepareBaseTestCache(idx, selections, a, headers, eclNotes)
    Set report("EclNotes") = eclNotes

    ' Any pre-shock row still configured to copy the system output is repointed at the derived
    ' value, so the change is visible in the configuration afterwards rather than being a
    ' hidden generator behaviour. This run already honours it: BuildElement promotes such a
    ' formula in memory whether or not the sheet write succeeds.
    report("Stage") = "pointing pre-shock rows at the derived values"
    On Error Resume Next
    report("Repointed") = MigratePreShockConfigTokens(False)
    If Err.Number <> 0 Then
        ' Said, not swallowed. The run itself is unaffected: BuildElement
        ' promotes those formulas in memory either way.
        repointError = Err.description: Err.Clear
        LogIssue LOG_LEVEL_WARN, "ECL pre-shock", "The pre-shock configuration rows could not be switched to the derived values: " & _
                 repointError & ". This run still uses the derived values; the configuration sheet was left as it was.", PRE_SHOCK_SHEET
    End If
    Err.Clear
    On Error GoTo Failed

    report("Stage") = "choosing the output folder"
    baseFolder = outputFolder
    If Len(baseFolder) = 0 Then baseFolder = ThisWorkbook.path
    If Len(baseFolder) = 0 Or InStr(1, baseFolder, "://", vbBinaryCompare) > 0 Or Len(baseFolder) > 100 Then
        If Not allowPrompts Then
            Err.Raise vbObjectError + 605, , "No usable output folder. Save the tool to a local folder with a short path, or pass one in."
        End If
        With Application.FileDialog(4)
            .title = IIf(Len(baseFolder) > 100, "Choose a shorter output folder to fit Excel's path limit", "Choose the output folder")
            If .Show <> -1 Then report("Cancelled") = True: GoTo Cleanup
            baseFolder = .SelectedItems(1)
        End With
    End If

    report("Stage") = "creating the output folder"
    root = JoinPath(baseFolder, "Output"): EnsureFolder root
    start = Now: runId = "RUN_" & format$(start, "yyyymmdd_hhnnss")
    folder = JoinPath(root, runId)
    Do While Len(Dir$(folder, vbDirectory)) > 0
        suffix = suffix + 1: folder = JoinPath(root, runId & "_" & suffix)
    Loop
    root = folder: EnsureFolder root
    report("Root") = root

    report("Stage") = "preparing the Excel session"
    If Not stateTaken Then Set state = CaptureState(): stateTaken = True
    JKB_Busy = True: busyTaken = True
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    ' Convenience-only properties that differ between Excel builds must never abort a run.
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    Application.DefaultSheetDirection = xlLTR
    Err.Clear
    On Error GoTo Failed

    report("Stage") = "preparing report styles"
    Sheet1.PrepareStyles

    For Each tk In idx.keys
        Set tc = idx(tk)
        If IsSelected(tc, selections) Then
            report("Stage") = "generating " & tc("TestCaseCode") & " / " & tc("Date") & " / " & tc("EntityCode")
            Application.StatusBar = "JKB: generating " & (Done + Failed + 1) & " of " & count & " - " & tc("TestCaseCode")
            folder = JoinPath(root, FileComponent(tc("Category"))): EnsureFolder folder
            folder = JoinPath(folder, tc("Date")): EnsureFolder folder
            folder = JoinPath(folder, Left$(FileComponent(tc("EntityCode")), 24) & "_ID_" & Left$(FileComponent(tc("EntityID")), 12)): EnsureFolder folder
            file = JoinPath(folder, format$(Done + Failed + 1, "0000") & "_" & Left$(FileComponent(tc("TestCaseCode")), 42) & ".xlsm")
            Set result = GenerateOne(a, headers, rules, grouping, baseCache, tc, file, runId)
            report("Results").Add result
            If CBool(result("Succeeded")) Then
                Done = Done + 1
            Else
                Failed = Failed + 1
                AddProblem report, tc("TestCaseCode") & " / " & tc("Date") & " / " & tc("EntityCode") & ": " & CStr(result("Message"))
            End If
            DoEvents
        End If
    Next tk
    report("Generated") = Done: report("Failed") = Failed

    report("Stage") = "writing the run summary"
    WriteRunSummary report("Results"), root, runId, start

    With ThisWorkbook.Worksheets(SHEET_DASHBOARD)
        .Range("F27").Value2 = root
        .Range("F28").Value2 = Done & " generated; " & Failed & " failed. Review Build_Log for configuration warnings."
    End With
    LogIssue LOG_LEVEL_INFO, "Generation", Done & " files generated; " & Failed & " failed.", root
    report("Ok") = (Failed = 0)
    report("Stage") = ""

Cleanup:
    On Error Resume Next
    Application.StatusBar = False
    If stateTaken Then RestoreState state
    If busyTaken Then JKB_Busy = False
    On Error GoTo 0
    Set GenerateOutputs = report
    Exit Function

Failed:
    report("ErrorNumber") = Err.Number
    report("ErrorText") = Err.description
    On Error Resume Next
    LogIssue LOG_LEVEL_ERROR, "Generation", "Step: " & report("Stage") & " | Error " & report("ErrorNumber") & " | " & report("ErrorText"), report("Root")
    On Error GoTo 0
    Resume Cleanup
End Function

' ------------------------------------------------------------------- UI -----

Public Sub RunExternalScenarioOutput(ByVal selections As Object)
    Dim report As Object, msg As String
    Set report = GenerateOutputs(selections)

    If CBool(report("Cancelled")) Then Exit Sub

    If report("ErrorNumber") <> 0 Then
        UiProblem "Generate outputs", "Generation stopped while " & IIf(Len(report("Stage")) > 0, report("Stage"), "generating") & ".", _
                  report("ErrorText") & " (error " & report("ErrorNumber") & ")", _
                  IIf(Len(report("Root")) > 0, "Files finished before the stop are in " & report("Root") & ". ", "") & "Build_Log has the full detail."
        Exit Sub
    End If

    If report("Problems").count > 0 And report("Generated") = 0 Then
        UiProblem "Generate outputs", "Generation did not start.", report("ErrorText") & vbCrLf & ProblemSummary(report), _
                  "Build_Log lists every problem with its element type and output row."
        ThisWorkbook.Worksheets(SHEET_LOG).Activate
        Exit Sub
    End If

    OpenReconWorkbench
    msg = "Generated " & report("Generated") & " of " & report("Selected") & " output workbooks." & vbCrLf & _
          "Failed: " & report("Failed") & vbCrLf & vbCrLf & "Folder:" & vbCrLf & report("Root") & vbCrLf & vbCrLf & _
          "Run_Summary.xlsm lists every file and its status."
    If report("Failed") > 0 Then msg = msg & vbCrLf & vbCrLf & "Failures:" & ProblemSummary(report)
    ' Reported apart from failures on purpose: the outputs are complete and usable, only
    ' the independent ECL cross-check was unavailable for some rows.
    If report("EclNotes").count > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Independent ECL check (separate from the manual calculations):" & vbCrLf & _
              "  " & report("EclNotes").count & " item(s) fell back to the configured manual formula." & vbCrLf & _
              "  See Base & pre-shock (" & DERIVED_SHEET & ") and the 'ECL pre-shock' entries in Build_Log."
    End If
    MsgBox msg, IIf(report("Failed") > 0, vbExclamation, vbInformation), "JKB Stress Testing - Generate outputs"
End Sub

' --------------------------------------------------------------- validate ---

' Checks every element type the selection actually uses and records EVERY problem
' rather than stopping at the first. A run that cannot produce correct numbers is
' refused outright - silently blanking an untranslatable formula would hide a
' configuration fault inside a reconciliation deliverable.
Private Sub ValidateSelectedRules(ByVal index As Object, ByVal selections As Object, _
                                  ByVal rulesByType As Object, ByVal headers As Object, _
                                  ByVal grouping As Object, ByVal report As Object)
    Dim tk As Variant, ek As Variant, tc As Object, el As Object
    Dim rules As Collection, ri As Object, checked As Object, rm As Object
    Dim displayRules As Collection, totalRows As Long
    Dim t As String, sev As Long, formula As String, label As String, ctx As String

    Set checked = NewMap()
    For Each tk In index.keys
        Set tc = index(tk)
        If IsSelected(tc, selections) Then
            For Each ek In tc("Elements").keys
                Set el = tc("Elements")(ek)
                t = SafeUpperText(el("ElementType"))
                If Not checked.Exists(t) Then
                    checked(t) = True
                    If Not rulesByType.Exists(t) Then
                        Record report, "No enabled configuration rules for element type '" & el("ElementType") & "'.", t
                    Else
                        Set rules = rulesByType(t)
                        Set rm = Nothing
                        On Error Resume Next
                        ' Exactly the layout BuildElement will use, so what is validated is
                        ' what gets generated.
                        Set rm = BuildOutputRowMap(rules, grouping, CStr(el("ElementType")), displayRules, totalRows)
                        If Err.Number <> 0 Then
                            Record report, "Output rows for '" & el("ElementType") & "' are not usable: " & Err.description, t
                            Err.Clear
                        End If
                        On Error GoTo 0
                        If Not rm Is Nothing Then
                            For Each ri In rules
                                label = SafeText(ri("OutputRowLabel"))
                                ctx = t & " / " & label
                                If Len(SafeText(ri("SourceFieldName"))) > 0 Then
                                    If Len(ResolveSourceFieldKey(headers, ri("SourceFieldName"))) = 0 Then
                                        Record report, "Source field '" & SafeText(ri("SourceFieldName")) & "' is required by " & ctx & " but is missing from the upload.", ctx
                                    End If
                                End If
                                If CBool(ri("ShowManual")) Then
                                    For sev = 0 To 2
                                        formula = ManualFormulaText(ri, sev)
                                        If Len(formula) > 0 Then
                                            On Error Resume Next
                                            formula = ConvertManualFormula(formula, rm, sev, ctx)
                                            If Err.Number <> 0 Then
                                                Record report, "Formula for " & ctx & " (" & SeverityName(sev) & ") cannot be translated - error " & Err.Number & ": " & Err.description, ctx
                                                Err.Clear
                                            End If
                                            On Error GoTo 0
                                        End If
                                    Next sev
                                End If
                            Next ri
                        End If
                    End If
                End If
            Next ek
        End If
    Next tk
    ValidateGrouping grouping, rulesByType, checked, report
End Sub

Private Sub Record(ByVal report As Object, ByVal text As String, ByVal context As String)
    AddProblem report, text
    On Error Resume Next
    LogIssue LOG_LEVEL_ERROR, "Configuration", text, context
    On Error GoTo 0
End Sub

Private Function SeverityName(ByVal sev As Long) As String
    Dim s As Variant: s = SeverityList()
    If sev >= 0 And sev <= 2 Then SeverityName = CStr(s(sev)) Else SeverityName = "severity " & sev
End Function

' ---------------------------------------------------------- output books ----

' The generated workbook carries only what a reviewer needs: a dashboard, the element
' sheets and the build log. Master configuration never travels with a deliverable, so
' the config sheets are not copied at all - only the named styles they define.
Private Function NewOutputBook() As Workbook
    Dim wb As Workbook, dash As Worksheet, sh As Shape, er As String, dashState As Long
    On Error GoTo Failed
    ' The master keeps Dashboard hidden; the copy a reviewer opens must be visible.
    dashState = ThisWorkbook.Worksheets(SHEET_DASHBOARD).Visible
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Copy
    Set wb = ActiveWorkbook
    ThisWorkbook.Worksheets(SHEET_DASHBOARD).Visible = dashState
    wb.date1904 = False
    On Error Resume Next
    wb.styles.Merge ThisWorkbook
    Err.Clear
    On Error GoTo Failed
    Set dash = wb.Worksheets(SHEET_DASHBOARD)
    dash.Range("N1").Value2 = "OUTPUT"
    For Each sh In dash.Shapes
        Select Case sh.name
            Case "JKB_Upload", "JKB_Generate", "JKB_SideBase", "JKB_SideConfig", "JKB_SideFormat", "JKB_SidePreShock", "JKB_Config"
                sh.Delete
        End Select
    Next sh
    ' The master dashboard is a designed layout, and its run-details block is a
    ' column of MERGED label and value pairs - D:E for the label, F:K for the
    ' value. Excel refuses point blank to clear or write half of a merged cell:
    ' "We can't do that to a merged cell". Every operation below therefore covers
    ' whole merge areas, or writes to a single anchor cell.
    '
    ' This is what stopped every generated workbook when the layout gained its
    ' merges, and it failed at the first file of every run rather than at the
    ' tenth, which is the one mercy in it.
    Dim labels As Variant, i As Long
    dash.Range("D6:K14").UnMerge
    dash.Range("D6:K14").Clear
    dash.Range("D6").Value2 = "Testcase reconciliation"
    dash.Range("D8").Value2 = "Open any element from the list below, or from the sheet tabs."
    dash.Range("D9").Value2 = "Use the slim ribbon to hide formula rows and collapse or expand detail groups."

    ClearMergeSafe dash.Range("D18:K28")
    dash.Range("D16").Value2 = "Run details"
    labels = Array("Run", "Test case", "As-of date", "Entity", "Source upload", _
                   "Generated at", "Elements", "Review status")
    For i = LBound(labels) To UBound(labels)
        ' One anchor cell at a time. An eight-row array written into eight merged
        ' two-column areas is the same refusal in a different costume.
        dash.Cells(18 + i, 4).Value2 = labels(i)
    Next i

    Set NewOutputBook = wb
    Exit Function
Failed:
    er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 604, , "A new output workbook could not be prepared: " & er
End Function

' Clearing HALF of a merged cell is refused outright: "We can't do that to a
' merged cell", error -2147220900, and it takes the whole generation down with
' it. The dashboard is a designed layout whose label and value pairs ARE merged,
' so any clear aimed at part of one has to be widened to the whole area first.
'
' Doing that here rather than at each call site means the next change to the
' layout cannot quietly reintroduce the failure - which is how it arrived.
Private Sub ClearMergeSafe(ByVal rng As Range)
    Dim c As Range, target As Range
    On Error Resume Next
    Set target = rng
    For Each c In rng.Cells
        If c.MergeCells Then Set target = Union(target, c.MergeArea)
    Next c
    target.ClearContents
    Err.Clear
End Sub

' The output dashboard, in the same clothes as the tool it came from.
'
' The per-sheet design pass deliberately skips a dashboard marked OUTPUT: its
' layout writes "Prepare and generate" and a two-step workflow, which is the
' master's job description and not a deliverable's. But clearing the master's
' content also clears its merges and its type, and a stripped dashboard in an
' otherwise designed workbook looks like a mistake. So the blocks this one
' actually uses are put back, and only those.
Private Sub DressOutputDashboard(ByVal dash As Worksheet)
    Dim r As Long
    On Error Resume Next
    dash.Range("D6:K6").Merge
    With dash.Range("D6:K6")
        .Font.name = UI_FONT: .Font.Size = 13: .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
    End With
    With dash.Range("D8:K10")
        .Font.name = UI_FONT: .Font.Size = 10
        .Font.Color = RGB(102, 112, 133)
        .VerticalAlignment = xlCenter
    End With
    dash.Range("D16:K16").Merge
    With dash.Range("D16:K16")
        .Font.name = UI_FONT: .Font.Size = 11: .Font.Bold = True
        .Font.Color = RGB(25, 63, 137)
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ' The run-details pairs: label in D:E, value in F:K, one row each.
    For r = 18 To 25
        dash.Range(dash.Cells(r, 4), dash.Cells(r, 5)).Merge
        dash.Range(dash.Cells(r, 6), dash.Cells(r, 11)).Merge
        With dash.Range(dash.Cells(r, 4), dash.Cells(r, 5))
            .Font.name = UI_FONT: .Font.Size = 10: .Font.Bold = False
            .Font.Color = RGB(102, 112, 133)
            .VerticalAlignment = xlCenter
        End With
        With dash.Range(dash.Cells(r, 6), dash.Cells(r, 11))
            .Font.name = UI_FONT: .Font.Size = 10: .Font.Bold = True
            .Font.Color = RGB(14, 34, 64)
            .VerticalAlignment = xlCenter
        End With
        If r Mod 2 = 0 Then dash.Range(dash.Cells(r, 4), dash.Cells(r, 11)).Interior.Color = RGB(245, 247, 250)
    Next r
    Err.Clear
End Sub

Private Sub FinalizeGeneratedBook(ByVal wb As Workbook)
    Dim ws As Worksheet, sh As Shape
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DASHBOARD)
    For Each sh In ws.Shapes
        Select Case sh.name
            Case "JKB_Config", "JKB_SideConfig", "JKB_SideFormat", "JKB_SideBase", "JKB_SidePreShock"
                sh.Delete
        End Select
    Next sh
    ws.Range("D10").Value2 = "Self-contained review workbook - master configuration stays in the tool."
    ' Last, deliberately. StyleWorkbookUI runs before this and re-styles the whole
    ' dashboard from the master's own routine, which knows nothing about the
    ' design pass - so dressing the output any earlier is dressing something that
    ' is about to be overwritten.
    DressOutputDashboard ws
    On Error GoTo 0
End Sub

Private Function GenerateOne(ByRef a As Variant, ByVal h As Object, ByVal rulesByType As Object, _
                             ByVal grouping As Object, ByVal baseCache As Object, ByVal tc As Object, _
                             ByVal path As String, ByVal runId As String) As Object
    Dim wb As Workbook, dash As Worksheet, ws As Worksheet, result As Object
    Dim ek As Variant, el As Object, row As Long, errors As Long, n As Long
    Dim er As String, errNo As Long, pathReserved As Boolean
    Dim dvSheet As Worksheet, dvRows As Object, derivedCount As Long

    Set result = NewMap()
    result("Succeeded") = False: result("Date") = tc("Date")
    result("Entity") = tc("EntityCode") & " (ID " & tc("EntityID") & ")"
    result("Category") = tc("Category"): result("TestCase") = tc("TestCaseCode")
    result("Path") = path: result("Message") = "": result("Elements") = 0

    On Error GoTo Failed
    If Len(path) > 218 Then Err.Raise vbObjectError + 601, , "Output path exceeds Excel's supported length. Choose a shorter output folder."
    If Len(Dir$(path)) > 0 Then Err.Raise vbObjectError + 602, , "Output file already exists: " & path
    ' Past this point nothing but this run can have written to that path, so a leftover
    ' file there on failure is ours to clean up.
    pathReserved = True

    Set wb = NewOutputBook(): Set dash = wb.Worksheets(SHEET_DASHBOARD)

    ' The derived figures travel with the workbook on their own sheet, so every `derived.X`
    ' reference resolves to a visible, traceable cell rather than a pasted-in number.
    Set dvRows = NewMap()
    Set dvSheet = WriteBookDerivedSheet(wb, baseCache, tc, dvRows, derivedCount)

    row = 32
    dash.Range("D31:G31").value = Array("Element", "Description", "Formula errors", "Source run IDs")
    For Each ek In tc("Elements").keys
        Set el = tc("Elements")(ek)
        Set ws = BuildElement(wb, a, h, rulesByType(SafeUpperText(el("ElementType"))), grouping, baseCache, tc, el, _
                              DerivedRowsFor(dvRows, CStr(el("ScenarioCode"))))
        n = CountFormulaErrors(ws): errors = errors + n
        dash.Cells(row, 4).Value2 = el("ScenarioCode")
        dash.Hyperlinks.Add anchor:=dash.Cells(row, 4), address:="", SubAddress:="'" & Replace(ws.name, "'", "''") & "'!A1"
        dash.Cells(row, 5).Value2 = el("ScenarioName")
        dash.Cells(row, 6).Value2 = n
        dash.Cells(row, 7).Value2 = SourceRuns(a, h, el)
        row = row + 1
    Next ek

    With dash
        .Range("F18").Value2 = runId
        .Range("F19").Value2 = tc("TestCaseCode") & IIf(Len(SafeText(tc("TestCaseName"))) > 0 And SafeText(tc("TestCaseName")) <> SafeText(tc("TestCaseCode")), " - " & SafeText(tc("TestCaseName")), "")
        .Range("F20").value = DateSerial(val(Left$(tc("Date"), 4)), val(Mid$(tc("Date"), 6, 2)), val(Right$(tc("Date"), 2)))
        .Range("F20").NumberFormat = "dd-mmm-yyyy"
        .Range("F21").Value2 = result("Entity")
        .Range("F22").Value2 = ThisWorkbook.Worksheets(SHEET_DASHBOARD).Range("F18").Value2
        .Range("F23").value = Now: .Range("F23").NumberFormat = "dd-mmm-yyyy hh:mm:ss"
        .Range("F24").Value2 = row - 32
        .Range("F25").Value2 = IIf(errors > 0, "Review " & errors & " formula error(s).", "Calculated - review differences.")
    End With

    ' Keep the derived sheet at the end, after the element sheets it feeds.
    If Not dvSheet Is Nothing Then
        On Error Resume Next
        dvSheet.Move After:=wb.Worksheets(wb.Worksheets.count)
        If derivedCount = 0 Then dvSheet.Visible = xlSheetHidden
        Err.Clear
        On Error GoTo Failed
    End If

    ThisWorkbook.Worksheets(SHEET_LOG).Copy After:=wb.Worksheets(wb.Worksheets.count)
    ' A self-contained appearance sheet so the workbook can be restyled after delivery
    ' without the master's rule table travelling with it.
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & dash.CodeName & ".EnsureOutputFormatSheet", wb
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & dash.CodeName & ".StyleWorkbookUI"
    ApplyOutputWindowSettings wb
    FinalizeGeneratedBook wb
    UiFinishReportBook wb, "Stress test output - " & Replace(Mid$(path, InStrRev(path, Application.PathSeparator) + 1), ".xlsm", ""), "Generated scenario element output", False
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & dash.CodeName & ".RebindAll"
    wb.SaveAs fileName:=path, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & dash.CodeName & ".RebindAll"
    wb.Save
    wb.Close SaveChanges:=False: Set wb = Nothing

    result("Succeeded") = True: result("Elements") = row - 32
    result("Message") = IIf(errors > 0, "Review: " & errors & " formula errors", "Generated")
    LogIssue IIf(errors > 0, LOG_LEVEL_WARN, LOG_LEVEL_INFO), "Output", result("Message"), tc("Date") & " / " & result("Entity") & " / " & tc("TestCaseCode")
    Set GenerateOne = result
    Exit Function

Failed:
    errNo = Err.Number: er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ' Only a half-written file from THIS run is removed; an existing user file is never touched.
    If pathReserved Then
        If Len(Dir$(path)) > 0 Then Kill path
    End If
    result("Message") = "Error " & errNo & ": " & er
    LogIssue LOG_LEVEL_ERROR, "Output", result("Message"), path
    On Error GoTo 0
    Set GenerateOne = result
    ' Esc / Ctrl+Break must stop the whole run, not just this file.
    If errNo = 18 Then Err.Raise 18, , "Generation cancelled."
End Function

' Window state is cosmetic and depends on Excel actually having a visible window
' (freeze panes fails outright when it does not). A generated deliverable must never
' be lost over pane splits or zoom, so every step here is advisory.
Private Sub ApplyOutputWindowSettings(ByVal wb As Workbook)
    Dim ws As Worksheet, gridlines As Boolean, zoom As Long
    On Error Resume Next
    gridlines = ParseBool(SettingValue("Show gridlines", "No"))
    zoom = CLng(SettingValue("Zoom (%)", 90))
    If zoom < 50 Or zoom > 200 Then zoom = 90
    If wb.Windows.count = 0 Then Exit Sub
    For Each ws In wb.Worksheets
        If Left$(SafeText(ws.Range("A4").Value2), 14) = "Element type: " Then
            ws.Activate
            ActiveWindow.DisplayGridlines = gridlines
            ActiveWindow.DisplayOutline = True
            ' Freezing is delegated so it always scrolls home first; setting SplitRow while
            ' the window is scrolled freezes the wrong rows and the sheet then jumps about.
            Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & wb.Worksheets(SHEET_DASHBOARD).CodeName & ".EnsureFrozenLayout", ws, zoom
        End If
    Next ws
    wb.Worksheets(SHEET_DASHBOARD).Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.zoom = 85
    Err.Clear
End Sub

' ------------------------------------------------- derived values per book ----
'
' Base and pre-shock are rebuilt from the uploaded extracts, not read from the system output.
' Writing them onto their own sheet inside each output workbook is what lets a configured
' formula say `derived.OUTST_LCY_STAGE1_PRE_SHOCK` and have that resolve to a cell the
' reviewer can click through to - complete with the source file, the filter that produced it
' and the number of rows it matched.
'
' Returns the sheet; fills `rowsOut` with "ELEMENT|METRIC" -> row.
Private Function WriteBookDerivedSheet(ByVal wb As Workbook, ByVal baseCache As Object, ByVal tc As Object, _
                                       ByVal rowsOut As Object, ByRef countOut As Long) As Worksheet
    Dim ws As Worksheet, ek As Variant, el As Object, mk As Variant, rec As Object
    Dim lines As Collection, line As Variant, out As Variant, i As Long, j As Long, hdr As Variant
    Dim metrics As Variant, metric As String, lastRow As Long

    On Error GoTo Failed
    VS_ResetCache
    Set lines = New Collection
    If Not baseCache Is Nothing Then
        metrics = DerivedMetricNames(baseCache)
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            For Each mk In metrics
                metric = CStr(mk)
                Set rec = DerivedLookup(baseCache, tc, el, metric)
                If Not rec Is Nothing Then
                    ' Base and Pre are what derived.<metric> resolves to, which is
                    ' not always the derived figure - a metric can be configured to
                    ' take the system's own number, and the workbook has to say so
                    ' rather than presenting a copied figure as an independent one.
                    Dim baseEvidence As Variant, preEvidence As Variant
                    baseEvidence = rec("Base"): preEvidence = rec("Pre")
                    If ReconStrictMode Then
                        baseEvidence = CVErr(xlErrNA): preEvidence = CVErr(xlErrNA)
                        If CBool(rec("HasBase")) Then baseEvidence = rec("BaseDerived")
                        If CBool(rec("HasPre")) Then preEvidence = rec("PreDerived")
                    End If
                    lines.Add Array(el("ScenarioCode"), metric, rec("Source"), _
                                    baseEvidence, preEvidence, _
                                    rec("Rows"), rec("Filter"), rec("Status"), _
                                    DerivedFieldOf(rec, "Use"), DerivedFieldOf(rec, "UsedPre"), _
                                    VS_Evidence(rec, "HasBase", "BaseDerived"), VS_Evidence(rec, "", "BaseSystem"), _
                                    VS_Evidence(rec, "HasPre", "PreDerived"), VS_Evidence(rec, "", "PreSystem"))
                    rowsOut(SafeUpperText(el("ScenarioCode")) & "|" & SafeUpperText(metric)) = DV_FIRST_ROW + lines.count - 1
                End If
            Next mk
        Next ek
    End If
    countOut = lines.count

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.name = DERIVED_SHEET
    ws.Range("A1").Value2 = "JKB"
    ws.Cells(3, 1).Value2 = "Derived base and pre-shock values"
    ws.Cells(4, 1).Value2 = "Rebuilt independently from the uploaded source extracts. Base is the unfiltered total; " & _
                            "pre-shock is that total restricted to this test case's filter condition."
    ws.Cells(6, 1).Value2 = tc("TestCaseCode") & "  |  " & tc("Date") & "  |  " & tc("EntityCode") & _
                            "  |  " & countOut & " derived figure(s)"
    hdr = Array("Element", "Metric", "Source", "Base", "Pre-shock", "Rows matched", "Filter applied", "Status", _
                "Value to use", "Which one was used", _
                "Base - input files", "Base - system", "Pre-shock - input files", "Pre-shock - system")
    For j = 0 To UBound(hdr): ws.Cells(DV_HEADER_ROW, j + 1).Value2 = hdr(j): Next j

    If lines.count > 0 Then
        ReDim out(1 To lines.count, 1 To 14)
        i = 0
        For Each line In lines
            i = i + 1
            For j = 0 To 13: out(i, j + 1) = line(j): Next j
        Next line
        WriteLiteralValues ws.Cells(DV_FIRST_ROW, 1).Resize(lines.count, 14), out
        lastRow = DV_FIRST_ROW + lines.count - 1
    Else
        ws.Cells(DV_FIRST_ROW, 1).Value2 = "No extract was loaded, so no figure could be rebuilt independently for this test case."
        lastRow = DV_FIRST_ROW
    End If
    If Not ReconStrictMode Then StyleBookDerivedSheet ws, lastRow, lines.count
    If Not ReconStrictMode Then VS_StyleEvidence ws, lastRow, lines.count
    If Not ReconStrictMode Then VS_WriteDifferences ws, lastRow, lines.count
    Set WriteBookDerivedSheet = ws
    Exit Function
Failed:
    On Error Resume Next
    LogIssue LOG_LEVEL_WARN, "Derived values", "The derived-values sheet could not be added to the output: " & Err.description, tc("TestCaseCode")
End Function

' A cache record read back from an older store does not carry every key, and
' asking a Dictionary for a missing key silently ADDS it - so optional reads go
' through here rather than growing the record they were only meant to inspect.
Private Function DerivedFieldOf(ByVal rec As Object, ByVal key As String) As String
    If rec.Exists(key) Then DerivedFieldOf = SafeText(rec(key))
End Function

Private Sub StyleBookDerivedSheet(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal n As Long)
    On Error Resume Next
    ws.Cells.Font.name = "Aptos"
    ws.Cells.Font.Size = 9.5
    ws.Cells.Interior.Color = RGB(249, 250, 251)
    ws.Range("A1").Font.Color = RGB(249, 250, 251)
    ws.rows(1).RowHeight = 6
    ws.Cells(3, 1).Font.Size = 15
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Color = RGB(18, 48, 107)
    ws.Cells(4, 1).Font.Color = RGB(102, 112, 133)
    ws.Cells(6, 1).Font.Bold = True
    ws.Cells(6, 1).Font.Color = RGB(18, 48, 107)
    With ws.Range(ws.Cells(DV_HEADER_ROW, 1), ws.Cells(DV_HEADER_ROW, 8))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(52, 64, 84)
        .HorizontalAlignment = xlCenter
    End With
    If n > 0 Then
        ws.Range(ws.Cells(DV_FIRST_ROW, 4), ws.Cells(lastRow, 5)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(DV_FIRST_ROW, 6), ws.Cells(lastRow, 6)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(DV_HEADER_ROW, 1), ws.Cells(lastRow, 8)).Borders.Color = RGB(208, 213, 221)
    End If
    ws.columns("A:H").AutoFit
    ws.columns("G").ColumnWidth = 55
    ws.rows(DV_HEADER_ROW).RowHeight = 26
    Err.Clear
End Sub

' Every metric name seen in the cache, so the derived sheet lists them in a stable order.
Private Function DerivedMetricNames(ByVal baseCache As Object) As Variant
    Dim seen As Object, k As Variant, parts As Variant, out() As String, i As Long
    Set seen = NewMap()
    For Each k In baseCache.keys
        parts = Split(CStr(k), "|")
        If UBound(parts) >= 5 Then
            If Not seen.Exists(CStr(parts(5))) Then seen(CStr(parts(5))) = True
        End If
    Next k
    If seen.count = 0 Then DerivedMetricNames = Array(): Exit Function
    ReDim out(0 To seen.count - 1)
    For Each k In seen.keys
        out(i) = CStr(k): i = i + 1
    Next k
    DerivedMetricNames = out
End Function

' A safety net for configuration that has not been migrated to `derived.X`.
'
' Two cases are upgraded, and only when an independently derived figure actually exists for
' the row:
'   - no manual formula at all
'   - a formula that is nothing but this row's own system-output cell, i.e. checking the
'     system against itself, which can never disagree and so tests nothing
' Anything else the user has written is left exactly as configured.
Private Function PromoteDerivedFormula(ByVal f As String, ByVal label As String, ByVal derivedRows As Object) As String
    Dim t As String, key As String
    PromoteDerivedFormula = f
    If derivedRows Is Nothing Then Exit Function
    key = SafeUpperText(label)
    If Not derivedRows.Exists(key) Then Exit Function
    t = SafeUpperText(f)
    t = Replace(Replace(Replace(t, "@", ""), " ", ""), vbTab, "")
    If Len(t) = 0 Then PromoteDerivedFormula = "derived." & label: Exit Function
    If t = "SYSTEMOUTPUT." & key Or t = "SYSTEM." & key Then PromoteDerivedFormula = "derived." & label
End Function

' The slice of the row map that belongs to one element, keyed by metric alone.
Private Function DerivedRowsFor(ByVal rowsOut As Object, ByVal elementCode As String) As Object
    Dim d As Object, k As Variant, prefix As String
    Set d = NewMap()
    prefix = SafeUpperText(elementCode) & "|"
    For Each k In rowsOut.keys
        If Left$(CStr(k), Len(prefix)) = prefix Then d(Mid$(CStr(k), Len(prefix) + 1)) = rowsOut(k)
    Next k
    Set DerivedRowsFor = d
End Function

' -------------------------------------------------------------- elements ----

Private Function BuildElement(ByVal wb As Workbook, ByRef a As Variant, ByVal h As Object, ByVal rules As Collection, _
                              ByVal grouping As Object, ByVal baseCache As Object, ByVal tc As Object, ByVal el As Object, _
                              ByVal derivedRows As Object) As Worksheet
    Dim ws As Worksheet, rm As Object, ri As Object, displayRules As Collection
    Dim values As Variant, manuals As Variant, differences As Variant, gd As Object
    Dim totalRows As Long, pos As Long, i As Long, j As Long, sourceCol As Long, field As String
    Dim sev As Variant, r As Long, f As String
    Dim currentGroup As String, gkey As String, showHeader As Boolean
    Dim sysRef As String, manRef As String, reconRec As Object
    Dim vsBase As Variant, vsMarks As Collection, vsF As String

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.name = UniqueSheetName(wb, CStr(el("ScenarioCode")))
    ws.Range("A1").Value2 = "JKB": ws.Range("N1").Value2 = "OUTPUT"
    ws.Range("A3").Value2 = el("ScenarioCode") & " | " & el("ScenarioName")
    ws.Range("A4").Value2 = "Element type: " & el("ElementType") & " | " & tc("Date") & " | " & tc("EntityCode") & " (ID " & tc("EntityID") & ")"
    ws.Range("A3:K3").Merge: ws.Range("A4:K4").Merge
    ws.Range("C5:E5").Merge: ws.Range("F5:H5").Merge: ws.Range("I5:K5").Merge
    ws.Range("A5").Value2 = "Output row": ws.Range("B5").Value2 = "Base"
    ws.Range("C5").Value2 = "System output": ws.Range("F5").Value2 = "Independent / manual check": ws.Range("I5").Value2 = "Difference"
    ws.Range("A6").Value2 = "Metric": ws.Range("B6").Value2 = "Base"
    ws.Range("C6:K6").value = Array("Moderate", "Medium", "Severe", "Moderate", "Medium", "Severe", "Moderate", "Medium", "Severe")

    Set rm = BuildOutputRowMap(rules, grouping, CStr(el("ElementType")), displayRules, totalRows)

    ReDim values(1 To totalRows, 1 To 13)
    ReDim manuals(1 To totalRows, 1 To 3)
    ReDim differences(1 To totalRows, 1 To 3)
    ReDim vsBase(1 To totalRows): Set vsMarks = New Collection
    sev = SeverityList(): currentGroup = "": pos = 0

    For Each ri In displayRules
        gkey = DetailGroupKey(grouping, CStr(el("ElementType")), CStr(ri("OutputRowLabel")))
        If gkey <> currentGroup Then
            Set gd = DetailGroupDef(grouping, gkey): showHeader = CBool(gd("ShowHeader"))
            If showHeader Then
                pos = pos + 1
                values(pos, 1) = gd("Label"): values(pos, 12) = "SECTION_HEADER": values(pos, 13) = gkey
            End If
            currentGroup = gkey
        End If
        pos = pos + 1: r = OUTPUT_FIRST_DATA_ROW + pos - 1
        values(pos, 1) = ri("OutputRowLabel"): values(pos, 13) = gkey
        If CBool(ri("HideRowOnOpen")) Then values(pos, 12) = MARKER_TOGGLE_ROW

        field = ResolveSourceFieldKey(h, ri("SourceFieldName"))
        If Len(NormalHeader(ri("SourceFieldName"))) > 0 And Len(field) = 0 And Not ReconStrictMode Then
            Err.Raise vbObjectError + 572, , "Missing source field '" & ri("SourceFieldName") & "' while generating " & el("ScenarioCode")
        End If
        If Len(field) > 0 Then
            sourceCol = h(field)
            values(pos, 2) = ReportSourceValue(a, tc("BaseRow"), sourceCol, field)
            For j = 0 To 2
                If el("SeverityRows").Exists(sev(j)) Then values(pos, 3 + j) = ReportSourceValue(a, el("SeverityRows")(sev(j)), sourceCol, field)
            Next j
        End If

        If ReconStrictMode Then
            values(pos, 2) = CVErr(xlErrNA)
            Set reconRec = ReconBaseRecord(baseCache, tc, el, CStr(ri("OutputRowLabel")))
            If Not reconRec Is Nothing Then values(pos, 2) = reconRec("BaseDerived")
        End If
        ' Base from the input files, where Config_ValueSources links this row (modValueSources).
        If Not ReconStrictMode Then vsBase(pos) = VS_BaseCell(CStr(ri("OutputRowLabel")), values(pos, 2), rm, baseCache, tc, el, _
                                                              derivedRows, r, vsMarks, el("ElementType") & " / " & ri("OutputRowLabel"))

        For j = 0 To 2
            If el("SeverityRows").Exists(sev(j)) Then
                If CBool(ri("ShowManual")) Or ReconStrictMode Then
                    ' Everything goes through the formula compiler now, including the
                    ' independently derived figures. A derived value used to be pasted straight
                    ' into the cell as a bare number: correct, but untraceable, and invisible to
                    ' the configuration. It is now a reference into the workbook's own
                    ' Derived_Values sheet, which is what `derived.X` in the config resolves to.
                    f = PromoteDerivedFormula(ManualFormulaText(ri, j), CStr(ri("OutputRowLabel")), derivedRows)
                    If ReconStrictMode Then
                        Set reconRec = DerivedLookup(baseCache, tc, el, CStr(ri("OutputRowLabel")))
                        If Not reconRec Is Nothing Then
                            If ReconUseSource(CStr(ri("OutputRowLabel")), f) Then f = "derived." & CStr(ri("OutputRowLabel"))
                        End If
                    End If
                    If Len(f) > 0 Then manuals(pos, j + 1) = ConvertManualFormula(f, rm, j, el("ElementType") & " / " & ri("OutputRowLabel"), derivedRows)
                    ' Pre-shock from the input files, unless the config holds a formula of its own.
                    If Not ReconStrictMode Then
                        vsF = VS_PreShockFormula(CStr(ri("OutputRowLabel")), ManualFormulaText(ri, j), rm, j, _
                                                 el("ElementType") & " / " & ri("OutputRowLabel"), baseCache, tc, el, derivedRows, r, vsMarks)
                        If Len(vsF) > 0 Then manuals(pos, j + 1) = vsF
                    End If
                End If
                If CBool(ri("ShowDifference")) And Len(CStr(manuals(pos, j + 1))) > 0 Then
                    ' A difference against a cell that could not be calculated is meaningless,
                    ' so it stays blank. The system and manual cells keep the real error, which
                    ' is what the reviewer needs to see and what the error count reports.
                    sysRef = colLetter(3 + j) & r
                    manRef = colLetter(6 + j) & r
                    differences(pos, j + 1) = "=IFERROR(IF(OR(" & sysRef & "=""""," & manRef & "=""""),""""," & sysRef & "-" & manRef & "),"""")"
                End If
            End If
        Next j
    Next ri

    WriteLiteralValues ws.Range("A7").Resize(totalRows, 13), values
    ws.Range("F7").Resize(totalRows, 3).formula = manuals
    If Not ReconStrictMode Then VS_WriteBaseFormulas ws, vsBase
    ws.Range("I7").Resize(totalRows, 3).formula = differences

    For j = 0 To 2
        If Not el("SeverityRows").Exists(sev(j)) Then
            ws.columns(3 + j).hidden = True: ws.columns(6 + j).hidden = True: ws.columns(9 + j).hidden = True
        End If
    Next j
    ' One setting read and one hide for the whole sheet, not one per row.
    Dim toggleRows As Range, hideToggles As Boolean, toggleRead As Boolean
    For i = 1 To totalRows
        If values(i, 12) = MARKER_TOGGLE_ROW Then
            If Not toggleRead Then hideToggles = (SafeUpperText(SettingValue("Formula rows on generation", "Hide")) = "HIDE"): toggleRead = True
            If toggleRows Is Nothing Then Set toggleRows = ws.rows(i + 6) Else Set toggleRows = Union(toggleRows, ws.rows(i + 6))
        End If
    Next i
    If Not toggleRows Is Nothing Then toggleRows.EntireRow.hidden = hideToggles

    ' Formatting is driven from the master's configuration, so the output workbook
    ' never needs a copy of the config sheets.
    If Not ReconStrictMode Then Sheet1.FormatReport ws
    If Not ReconStrictMode Then VS_MarkElementSheet ws, vsMarks
    ws.Calculate
    Set BuildElement = ws
End Function

' Reuse the production formula compiler in an unsaved calculation workbook.
' This workbook contains no copied macros or external links.
Public Function BuildReconCalculationBook(ByRef a As Variant, ByVal h As Object, ByVal rulesByType As Object, _
                ByVal grouping As Object, ByVal cache As Object, ByVal tc As Object, ByVal sheets As Object) As Workbook
    Dim wb As Workbook, ws As Worksheet, dv As Worksheet, rows As Object, ek As Variant, el As Object, n As Long
    Dim er As String
    On Error GoTo Failed
    RequireCompleteGrouping grouping, rulesByType, tc
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set rows = NewMap()
    Set dv = WriteBookDerivedSheet(wb, cache, tc, rows, n)
    If dv Is Nothing Then Err.Raise vbObjectError + 1720, , "Could not prepare independent source evidence."
    For Each ek In tc("Elements").keys
        Set el = tc("Elements")(ek)
        If rulesByType.Exists(SafeUpperText(el("ElementType"))) Then
            Set ws = BuildElement(wb, a, h, rulesByType(SafeUpperText(el("ElementType"))), grouping, cache, tc, el, DerivedRowsFor(rows, CStr(el("ScenarioCode"))))
            Set sheets(CStr(ek)) = ws
        End If
    Next ek
    Set BuildReconCalculationBook = wb
    Exit Function
Failed:
    er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 1721, "BuildReconCalculationBook", er
End Function

Private Function ReportSourceValue(ByRef a As Variant, ByVal row As Long, ByVal col As Long, ByVal field As String) As Variant
    Dim d As String
    If field = "AS_OF_DATE" Then
        d = DateKey(a(row, col), ThisWorkbook.date1904)
        ReportSourceValue = CDbl(DateSerial(val(Left$(d, 4)), val(Mid$(d, 6, 2)), val(Right$(d, 2))))
    Else
        ReportSourceValue = a(row, col)
    End If
End Function

Private Function CountFormulaErrors(ByVal ws As Worksheet) As Long
    Dim a As Variant, r As Long, c As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 7 Then Exit Function
    a = ws.Range("F7:K" & lastRow).Value2
    For r = 1 To UBound(a, 1)
        For c = 1 To UBound(a, 2)
            If IsError(a(r, c)) Then CountFormulaErrors = CountFormulaErrors + 1
        Next c
    Next r
End Function

Private Function SourceRuns(ByRef a As Variant, ByVal h As Object, ByVal el As Object) As String
    Dim sev As Variant, k As Variant
    sev = SeverityList()
    For Each k In sev
        If el("SeverityRows").Exists(k) Then
            If Len(SourceRuns) > 0 Then SourceRuns = SourceRuns & "; "
            SourceRuns = SourceRuns & k & ": " & CStr(a(el("SeverityRows")(k), h("RUN_ID")))
        End If
    Next k
End Function

' --------------------------------------------------------------- summary ----

Private Sub WriteRunSummary(ByVal results As Collection, ByVal root As String, ByVal runId As String, ByVal started As Date)
    Dim wb As Workbook, ws As Worksheet, r As Long, ri As Object, path As String, er As String
    On Error GoTo Failed
    Set wb = NewOutputBook(): Set ws = wb.Worksheets(SHEET_DASHBOARD)
    ws.Range("D6").Value2 = "Generation summary"
    ws.Range("D8").Value2 = "One row per generated workbook. Click a link to open it."
    ClearMergeSafe ws.Range("D9")
    ClearMergeSafe ws.Range("D18:D25")
    ws.Range("D18").Value2 = "Run": ws.Range("F18").Value2 = runId
    ws.Range("D19").Value2 = "Started": ws.Range("F19").value = started: ws.Range("F19").NumberFormat = "dd-mmm-yyyy hh:mm:ss"
    ws.Range("D20").Value2 = "Files": ws.Range("F20").Value2 = results.count
    ws.Range("D31:G31").ClearContents
    ws.Range("A31:H31").value = Array("Date", "Entity", "Category", "Test case", "Result", "Elements", "File", "Message")
    r = 32
    For Each ri In results
        ws.Cells(r, 1).Resize(1, 8).value = Array(ri("Date"), ri("Entity"), ri("Category"), ri("TestCase"), _
            IIf(CBool(ri("Succeeded")), "Generated", "Failed"), ri("Elements"), "Open workbook", ri("Message"))
        If CBool(ri("Succeeded")) Then ws.Hyperlinks.Add anchor:=ws.Cells(r, 7), address:=Mid$(CStr(ri("Path")), Len(root) + 2), TextToDisplay:="Open workbook"
        r = r + 1
    Next ri
    ThisWorkbook.Worksheets(SHEET_LOG).Copy After:=wb.Worksheets(wb.Worksheets.count)
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & ws.CodeName & ".StyleWorkbookUI"
    On Error Resume Next
    ws.Activate: ActiveWindow.DisplayGridlines = False: ActiveWindow.zoom = 85
    Err.Clear
    On Error GoTo Failed
    ' The summary is a TABLE with a short header block above it, not a dashboard,
    ' so its columns are sized for the table. Left at the dashboard's widths the
    ' first column is a 2.5-character gutter and every date in it reads "####".
    On Error Resume Next
    ws.columns("A").ColumnWidth = 12: ws.columns("B").ColumnWidth = 18
    ws.columns("C").ColumnWidth = 14: ws.columns("D").ColumnWidth = 24
    ws.columns("E").ColumnWidth = 11: ws.columns("F").ColumnWidth = 10
    ws.columns("G").ColumnWidth = 16: ws.columns("H").ColumnWidth = 58
    With ws.Range("A31:H31")
        .Font.name = UI_FONT: .Font.Size = 9.5: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(25, 63, 137)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.rows(31).RowHeight = 26
    If r > 32 Then
        With ws.Range(ws.Cells(32, 1), ws.Cells(r - 1, 8))
            .Font.name = UI_FONT: .Font.Size = 10
            .Font.Color = RGB(14, 34, 64)
            .VerticalAlignment = xlCenter
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
            .Borders(xlInsideHorizontal).Weight = xlHairline
        End With
        ws.Range(ws.Cells(32, 1), ws.Cells(r - 1, 1)).NumberFormat = "dd-mmm-yyyy"
    End If
    ws.ScrollArea = ""
    Err.Clear
    On Error GoTo Failed

    path = JoinPath(root, "Run_Summary.xlsm")
    FinalizeGeneratedBook wb
    UiFinishReportBook wb, "Generation run summary", "Every generated output workbook and its status", False
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & ws.CodeName & ".RebindAll"
    wb.SaveAs fileName:=path, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    Application.Run "'" & Replace(wb.name, "'", "''") & "'!" & ws.CodeName & ".RebindAll"
    wb.Save: wb.Close SaveChanges:=False
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 603, , "Output files were written, but the run summary could not be saved: " & er
End Sub

' -------------------------------------------------------------- grouping ----
'
' Every output row belongs to exactly one named group. The group is decided, in order, by:
'
'   1. DETAIL_GROUP_KEY on the rule row (column U) - an override for one row;
'   2. the first matching line of tblGroupRules (label pattern, optional element type),
'      taken in PRIORITY order;
'   3. nothing - the row is UNASSIGNED.
'
' There is no catch-all "Other" group any more. An unassigned row, or a row pointing at a
' group that is not defined, is a configuration problem: by default the run stops and lists
' them, so nothing ever lands in a bucket nobody chose. Setting the policy cell to "Show
' under Needs a group" lets a run continue and puts those rows under one clearly flagged
' header at the end of the sheet instead.


Public Function LoadDetailGrouping(ByVal wb As Workbook) As Object
    Dim result As Object, assignments As Object, defs As Object, def As Object, ws As Worksheet, lo As ListObject
    Dim hdr As Variant, a As Variant, h As Object, r As Long, c As Long, lastRow As Long
    Dim et As String, label As String, gkey As String, v As Variant
    Dim rules As Collection, rule As Object, problems As Collection, i As Long, placed As Boolean
    Set result = NewMap(): Set assignments = NewMap(): Set defs = NewMap(): Set h = NewMap()
    Set rules = New Collection: Set problems = New Collection
    Set ws = wb.Worksheets(SHEET_RULES): Set lo = ws.ListObjects(TABLE_RULES)

    ' 1. Group definitions.
    lastRow = ws.Cells(ws.Rows.Count, GROUP_DEF_KEY_COL).End(xlUp).Row
    For r = GROUP_DEF_FIRST_ROW To lastRow
        gkey = SafeUpperText(ws.Cells(r, GROUP_DEF_KEY_COL).Value2)
        If Len(gkey) > 0 Then
            Set def = NewMap()
            def("Label") = SafeText(ws.Cells(r, GROUP_DEF_KEY_COL + 1).Value2): If Len(def("Label")) = 0 Then def("Label") = gkey
            v = ws.Cells(r, GROUP_DEF_KEY_COL + 2).Value2: If IsNumeric(v) And Len(SafeText(v)) > 0 Then def("Order") = CLng(v) Else def("Order") = 900
            def("StyleKey") = SafeText(ws.Cells(r, GROUP_DEF_KEY_COL + 3).Value2): If Len(def("StyleKey")) = 0 Then def("StyleKey") = "SECTION_NEUTRAL"
            def("Collapsed") = ParseBool(ws.Cells(r, GROUP_DEF_KEY_COL + 4).Value2)
            def("ShowHeader") = Not (SafeUpperText(ws.Cells(r, GROUP_DEF_KEY_COL + 5).Value2) = "NO")
            v = ws.Cells(r, GROUP_DEF_KEY_COL + 6).Value2: If IsNumeric(v) And Len(SafeText(v)) > 0 Then def("OutlineLevel") = CLng(v) Else def("OutlineLevel") = 1
            If gkey = GROUP_UNASSIGNED Then
                problems.Add "'" & GROUP_UNASSIGNED & "' is reserved for rows with no group. Rename the group in row " & r & " of " & SHEET_RULES & "."
            ElseIf defs.Exists(gkey) Then
                problems.Add "Group '" & gkey & "' is defined twice in " & SHEET_RULES & " (second time in row " & r & ")."
            Else
                defs.Add gkey, def
            End If
        End If
    Next r
    If defs.Count = 0 Then problems.Add "No detail groups are defined in " & SHEET_RULES & " (column Y from row " & GROUP_DEF_FIRST_ROW & ")."

    ' The flagged header used only when the policy lets unassigned rows through.
    Set def = NewMap(): def("Label") = "Needs a group - set it in Config": def("Order") = 99999
    def("StyleKey") = "SECTION_NEUTRAL": def("Collapsed") = False: def("ShowHeader") = True: def("OutlineLevel") = 1
    result.Add "Unassigned", def

    ' 2. Per-row overrides (column U).
    hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    If h.Exists("DETAIL_GROUP_KEY") And Not lo.DataBodyRange Is Nothing Then
        a = lo.DataBodyRange.Value2
        For r = 1 To UBound(a, 1)
            et = SafeUpperText(a(r, h("ELEMENT_TYPE"))): label = SafeUpperText(a(r, h("OUTPUT_ROW_LABEL")))
            gkey = SafeUpperText(a(r, h("DETAIL_GROUP_KEY")))
            If Len(et) > 0 And Len(label) > 0 And Len(gkey) > 0 Then assignments(et & "|" & label) = gkey
        Next r
    End If

    ' 3. Pattern rules, kept sorted by PRIORITY (ties keep sheet order).
    Set lo = Nothing
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_GROUP_RULES)
    On Error GoTo 0
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            Set h = NewMap(): hdr = lo.HeaderRowRange.Value2
            For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
            If Not (h.Exists("PRIORITY") And h.Exists("LABEL_PATTERN") And h.Exists("ELEMENT_TYPE") And h.Exists("GROUP_KEY")) Then
                problems.Add TABLE_GROUP_RULES & " needs the columns PRIORITY, LABEL_PATTERN, ELEMENT_TYPE and GROUP_KEY."
            Else
                a = lo.DataBodyRange.Value2
                For r = 1 To UBound(a, 1)
                    If Len(SafeText(a(r, h("LABEL_PATTERN")))) > 0 Or Len(SafeText(a(r, h("GROUP_KEY")))) > 0 Then
                        Set rule = NewMap()
                        rule("Pattern") = SafeUpperText(a(r, h("LABEL_PATTERN"))): If Len(rule("Pattern")) = 0 Then rule("Pattern") = "*"
                        rule("ElementType") = SafeUpperText(a(r, h("ELEMENT_TYPE")))
                        rule("Group") = SafeUpperText(a(r, h("GROUP_KEY")))
                        rule("Row") = lo.DataBodyRange.Row + r - 1
                        v = a(r, h("PRIORITY")): If IsNumeric(v) And Len(SafeText(v)) > 0 Then rule("Priority") = CDbl(v) Else rule("Priority") = 1E+15
                        If Len(rule("Group")) = 0 Then
                            problems.Add "Group rule in row " & rule("Row") & " (" & rule("Pattern") & ") has no GROUP_KEY."
                        ElseIf Not defs.Exists(rule("Group")) Then
                            problems.Add "Group rule in row " & rule("Row") & " points at group '" & rule("Group") & "', which is not defined."
                        Else
                            placed = False
                            For i = 1 To rules.Count
                                If rule("Priority") < rules(i)("Priority") Then rules.Add rule, Before:=i: placed = True: Exit For
                            Next i
                            If Not placed Then rules.Add rule
                        End If
                    End If
                Next r
            End If
        End If
    End If

    result("Policy") = GROUP_POLICY_STOP
    If InStr(1, SafeText(ws.Range(GROUP_POLICY_CELL).Value2), "Show", vbTextCompare) > 0 Then result("Policy") = GROUP_POLICY_SHOW

    result.Add "Assignments", assignments: result.Add "Definitions", defs: result.Add "Rules", rules
    result.Add "Problems", problems: result.Add "Resolved", NewMap()
    Set LoadDetailGrouping = result
End Function

' The group key for one output row, or GROUP_UNASSIGNED. Results are cached on the
' grouping object because the layout sort asks for the same row many times.
Public Function ResolveDetailGroup(ByVal grouping As Object, ByVal elementType As String, ByVal rowLabel As String) As String
    Dim key As String, et As String, label As String, rule As Object, cache As Object
    et = SafeUpperText(elementType): label = SafeUpperText(rowLabel): key = et & "|" & label
    Set cache = grouping("Resolved")
    If cache.Exists(key) Then ResolveDetailGroup = cache(key): Exit Function
    ResolveDetailGroup = GROUP_UNASSIGNED
    If grouping("Assignments").Exists(key) Then
        ResolveDetailGroup = CStr(grouping("Assignments")(key))
        If Not grouping("Definitions").Exists(ResolveDetailGroup) Then ResolveDetailGroup = GROUP_UNASSIGNED
    Else
        For Each rule In grouping("Rules")
            If label Like CStr(rule("Pattern")) Then
                If Len(rule("ElementType")) = 0 Then ResolveDetailGroup = CStr(rule("Group")): Exit For
                If et Like CStr(rule("ElementType")) Then ResolveDetailGroup = CStr(rule("Group")): Exit For
            End If
        Next rule
    End If
    cache(key) = ResolveDetailGroup
End Function

' Why a row has no usable group, in words a user can act on; "" when it has one.
Public Function DetailGroupGap(ByVal grouping As Object, ByVal elementType As String, ByVal rowLabel As String) As String
    Dim key As String
    If ResolveDetailGroup(grouping, elementType, rowLabel) <> GROUP_UNASSIGNED Then Exit Function
    key = SafeUpperText(elementType) & "|" & SafeUpperText(rowLabel)
    If grouping("Assignments").Exists(key) Then
        DetailGroupGap = "DETAIL_GROUP_KEY '" & grouping("Assignments")(key) & "' is not a defined group"
    Else
        DetailGroupGap = "no group rule matches it"
    End If
End Function

Private Function DetailGroupKey(ByVal grouping As Object, ByVal elementType As String, ByVal rowLabel As String) As String
    DetailGroupKey = ResolveDetailGroup(grouping, elementType, rowLabel)
End Function

Private Function DetailGroupDef(ByVal grouping As Object, ByVal gkey As String) As Object
    Dim d As Object: Set d = grouping("Definitions")
    If d.Exists(gkey) Then Set DetailGroupDef = d(gkey) Else Set DetailGroupDef = grouping("Unassigned")
End Function

Private Function DetailGroupOrder(ByVal grouping As Object, ByVal gkey As String) As Long
    DetailGroupOrder = CLng(DetailGroupDef(grouping, gkey)("Order"))
End Function

' Adds every grouping problem for the element types in a run to the report. Unassigned
' rows are reported once per label, not once per element type, so a new metric that
' appears on thirteen sheets reads as one line.
Private Sub ValidateGrouping(ByVal grouping As Object, ByVal rulesByType As Object, ByVal elementTypes As Object, ByVal report As Object)
    Dim p As Variant, t As Variant, ri As Object, counts As Object, whys As Object, gap As String, label As String, k As Variant
    For Each p In grouping("Problems")
        Record report, CStr(p), SHEET_RULES
    Next p
    Set counts = NewMap(): Set whys = NewMap()
    For Each t In elementTypes.Keys
        If rulesByType.Exists(t) Then
            For Each ri In rulesByType(t)
                gap = DetailGroupGap(grouping, CStr(t), CStr(ri("OutputRowLabel")))
                If Len(gap) > 0 Then
                    label = SafeUpperText(ri("OutputRowLabel"))
                    counts(label) = CLng(counts(label)) + 1
                    If Not whys.Exists(label) Then whys(label) = gap
                End If
            Next ri
        End If
    Next t
    For Each k In counts.Keys
        If grouping("Policy") = GROUP_POLICY_SHOW Then
            On Error Resume Next
            LogIssue LOG_LEVEL_WARN, "Grouping", "'" & k & "' has no group (" & whys(k) & "); shown under 'Needs a group'.", SHEET_RULES
            On Error GoTo 0
        Else
            Record report, "Output row '" & k & "' has no group (" & whys(k) & "; " & counts(k) & _
                " element type(s)). Add a line to Group assignment rules or type a group in DETAIL_GROUP_KEY.", SHEET_RULES
        End If
    Next k
End Sub

' A label -> group map for one element type, for Formula Studio (called through
' Application.Run so the Dashboard code never binds to this module directly).
Public Function GroupMapForElement(ByVal elementType As String) As Object
    Dim grouping As Object, lo As ListObject, a As Variant, hdr As Variant, h As Object, c As Long, r As Long, m As Object, label As String
    Set m = NewMap(): Set h = NewMap()
    Set grouping = LoadDetailGrouping(ThisWorkbook)
    Set lo = ThisWorkbook.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    If Not lo.DataBodyRange Is Nothing Then
        hdr = lo.HeaderRowRange.Value2
        For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
        a = lo.DataBodyRange.Value2
        For r = 1 To UBound(a, 1)
            If StrComp(SafeText(a(r, h("ELEMENT_TYPE"))), SafeText(elementType), vbTextCompare) = 0 Then
                label = SafeText(a(r, h("OUTPUT_ROW_LABEL")))
                If Len(label) > 0 Then m(label) = ResolveDetailGroup(grouping, elementType, label)
            End If
        Next r
    End If
    Set GroupMapForElement = m
End Function

' Grouping health check for the config sheet's button: every enabled row, all element types.
Public Function GroupingHealthReport(ByVal wb As Workbook, ByRef problemCount As Long) As String
    Dim grouping As Object, rulesByType As Object, t As Variant, ri As Object, counts As Object, gaps As Object
    Dim g As String, s As String, k As Variant, n As Long, p As Variant, defs As Object
    Set grouping = LoadDetailGrouping(wb): Set rulesByType = LoadRulesByElementType(wb)
    Set counts = NewMap(): Set gaps = NewMap(): Set defs = grouping("Definitions")
    For Each t In rulesByType.Keys
        For Each ri In rulesByType(t)
            g = ResolveDetailGroup(grouping, CStr(t), CStr(ri("OutputRowLabel")))
            If g = GROUP_UNASSIGNED Then
                gaps(SafeUpperText(ri("OutputRowLabel"))) = DetailGroupGap(grouping, CStr(t), CStr(ri("OutputRowLabel")))
            Else
                counts(g) = CLng(counts(g)) + 1
            End If
        Next ri
    Next t
    problemCount = grouping("Problems").Count + gaps.Count
    For Each k In defs.Keys
        s = s & vbCrLf & "  " & defs(k)("Label") & " (" & k & "): " & CLng(counts(k)) & " row(s)"
    Next k
    s = "Enabled rows per group:" & s
    If grouping("Problems").Count > 0 Then
        s = s & vbCrLf & vbCrLf & "Setup problems:"
        For Each p In grouping("Problems"): s = s & vbCrLf & "  - " & p: Next p
    End If
    If gaps.Count > 0 Then
        s = s & vbCrLf & vbCrLf & gaps.Count & " output row(s) have no group:"
        For Each k In gaps.Keys
            n = n + 1
            If n > 15 Then s = s & vbCrLf & "  ... and " & (gaps.Count - 15) & " more": Exit For
            s = s & vbCrLf & "  - " & k & " (" & gaps(k) & ")"
        Next k
    ElseIf grouping("Problems").Count = 0 Then
        s = s & vbCrLf & vbCrLf & "Every enabled row has a group."
    End If
    GroupingHealthReport = s
End Function

' The same check for paths that build element sheets without the generation validator
' (the Reconciliation Workbench): raises when the policy is to stop and any row in this
' test case has no group, so no calculation book is built on a layout nobody chose.
Public Sub RequireCompleteGrouping(ByVal grouping As Object, ByVal rulesByType As Object, ByVal tc As Object)
    Dim report As Object, types As Object, ek As Variant
    Set report = NewReport(): Set types = NewMap()
    For Each ek In tc("Elements").Keys
        types(SafeUpperText(tc("Elements")(ek)("ElementType"))) = True
    Next ek
    ValidateGrouping grouping, rulesByType, types, report
    If report("Problems").Count > 0 Then
        Err.Raise vbObjectError + 1722, "Grouping", report("Problems").Count & " grouping problem(s) must be fixed in " & SHEET_RULES & ":" & ProblemSummary(report)
    End If
End Sub

' The single definition of where each output row lands on a generated element sheet.
'
' Validation and generation must agree exactly: a formula resolves to absolute cell
' references, so if the two disagree about the layout the generated sheet points at the
' wrong metrics. Both callers therefore go through this one function, which also hands back
' the display order and row count so the caller need not recompute them.
Public Function BuildOutputRowMap(ByVal rules As Collection, ByVal grouping As Object, ByVal elementType As String, _
                                  ByRef displayRules As Collection, ByRef totalRows As Long) As Object
    Dim rm As Object, ri As Object, gd As Object
    Dim currentGroup As String, gkey As String, labelKey As String
    Set displayRules = SortRulesForDetail(rules, grouping, elementType)
    Set rm = NewMap()
    currentGroup = "": totalRows = 0
    For Each ri In displayRules
        gkey = DetailGroupKey(grouping, elementType, CStr(ri("OutputRowLabel")))
        If gkey <> currentGroup Then
            Set gd = DetailGroupDef(grouping, gkey)
            If CBool(gd("ShowHeader")) Then totalRows = totalRows + 1
            currentGroup = gkey
        End If
        totalRows = totalRows + 1
        labelKey = SafeUpperText(ri("OutputRowLabel"))
        If Len(labelKey) = 0 Then Err.Raise vbObjectError + 554, , "An enabled rule for '" & elementType & "' has no output label."
        If rm.Exists(labelKey) Then Err.Raise vbObjectError + 555, , "Duplicate enabled output label '" & labelKey & "' for '" & elementType & "'."
        rm(labelKey) = OUTPUT_FIRST_DATA_ROW + totalRows - 1
    Next ri
    Set BuildOutputRowMap = rm
End Function

Private Function SortRulesForDetail(ByVal rules As Collection, ByVal grouping As Object, ByVal elementType As String) As Collection
    Dim result As New Collection, ri As Object, current As Object, i As Long, go As Long, cgo As Long, g As String, cg As String
    For Each ri In rules
        g = DetailGroupKey(grouping, elementType, CStr(ri("OutputRowLabel"))): go = DetailGroupOrder(grouping, g)
        For i = 1 To result.count
            Set current = result(i)
            cg = DetailGroupKey(grouping, elementType, CStr(current("OutputRowLabel"))): cgo = DetailGroupOrder(grouping, cg)
            If go < cgo Or (go = cgo And CLng(ri("OutputOrder")) < CLng(current("OutputOrder"))) Then result.Add ri, Before:=i: GoTo added
        Next i
        result.Add ri
added:
    Next ri
    Set SortRulesForDetail = result
End Function
