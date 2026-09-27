Attribute VB_Name = "modValueSources"
Option Explicit

' ============================================================================
'  Where base and pre-shock come from.
'
'  The chain every element sheet follows:
'
'     scenario -> test case / element -> BASE -> PRE-SHOCK -> shock -> POST-SHOCK
'
'  Shock and post-shock were already rebuilt independently by the manual formulas.
'  Base and pre-shock were not: column B held the system's own base value, and most
'  pre-shock rows held systemoutput.<row>, so the whole independent check started
'  from the system's numbers.
'
'  Config_ValueSources fixes that with one row per output row label:
'
'     BASE_LINK        what the base should be, rebuilt from the input files
'                      (the whole portfolio, no test-case filter)
'     PRE_SHOCK_LINK   what the pre-shock should be (the same files, restricted to
'                      the test case's filter on Pre_Shock_Cases)
'     BASE_SOURCE /    Manual file only   - the input files, nothing else. A value
'     PRE_SHOCK_SOURCE                      that cannot be rebuilt shows as missing.
'                      Manual else system - the input files, and the system value
'                                           when a file is missing (marked amber).
'                      System output      - the system value, by choice.
'                      blank              - the default set at the top of the sheet.
'
'  A link is written in the same language as the config formulas:
'     derivedbase.X   metric X from Pre_Shock_Metrics, unfiltered (the base)
'     derived.X       metric X, filtered to the test case (the pre-shock)
'     figure.X        a figure typed into the Manual figures table, for input
'                     files that are not wired in as sources yet
'  plus ordinary arithmetic, so ratios are built from their parts.
'
'  In the generated workbook the base cell becomes a formula into Derived_Values,
'  so post-shock formulas that say base.X now start from the manual base too.
'  Nothing is hidden: every fallback is coloured and carries a note.
' ============================================================================

Public Const VS_SHEET As String = "Config_ValueSources"
Private Const VS_TABLE As String = "tblValueSources"
Private Const VS_FIG_TABLE As String = "tblManualFigures"
Private Const VS_HEADER_ROW As Long = 9
Private Const VS_FIRST_ROW As Long = 10
Private Const VS_FIG_FIRST_COL As Long = 13
Private Const VS_LINK_COLS As Long = 10
Private Const VS_FIG_COLS As Long = 11

Private Const TONE_OK As String = "ok"
Private Const TONE_AMBER As String = "amber"
Private Const TONE_RED As String = "red"

Private mLoaded As Boolean
Private mLinks As Object
Private mFigures As Collection
Private mBaseDefault As String
Private mPreDefault As String

' ------------------------------------------------------------------ reading ----

Public Sub VS_ResetCache()
    mLoaded = False
    Set mLinks = Nothing
    Set mFigures = Nothing
End Sub

Public Function VS_Ready() As Boolean
    VS_Ready = Not (GetWorksheetSafe(ThisWorkbook, VS_SHEET) Is Nothing)
End Function

' The three choices, in the order they appear in the dropdowns.
Public Function VS_Choices() As String
    VS_Choices = USE_DERIVED & "," & USE_DERIVED_ELSE & "," & USE_SYSTEM
End Function

' "BASE" or "PRE". A blank default cell means Manual else system: the manual files
' wherever they exist, and a visible fallback wherever they do not.
Public Function VS_DefaultUse(ByVal phase As String) As String
    EnsureLoaded
    If UCase$(phase) = "BASE" Then VS_DefaultUse = mBaseDefault Else VS_DefaultUse = mPreDefault
End Function

Private Function UseFromText(ByVal s As String, ByVal fallback As String) As String
    Dim t As String
    t = SafeUpperText(s)
    If Len(t) = 0 Then UseFromText = fallback: Exit Function
    If InStr(t, "ELSE") > 0 Then UseFromText = USE_DERIVED_ELSE: Exit Function
    If Left$(t, 3) = "SYS" Then UseFromText = USE_SYSTEM: Exit Function
    If Left$(t, 3) = "MAN" Or Left$(t, 3) = "DER" Then UseFromText = USE_DERIVED: Exit Function
    UseFromText = fallback
End Function

Private Sub EnsureLoaded()
    Dim ws As Worksheet, h As Object, r As Long, label As String, lastRow As Long
    Dim fh As Object
    If mLoaded Then Exit Sub
    mLoaded = True
    Set mLinks = NewMap(): Set mFigures = New Collection
    mBaseDefault = USE_DERIVED_ELSE: mPreDefault = USE_DERIVED_ELSE
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_SHEET)
    If ws Is Nothing Then Exit Sub
    mBaseDefault = UseFromText(SafeText(ws.Range("C6").Value2), USE_DERIVED_ELSE)
    mPreDefault = UseFromText(SafeText(ws.Range("C7").Value2), USE_DERIVED_ELSE)

    Set h = HeaderMap(ws, 1, VS_LINK_COLS)
    If h.Exists("OUTPUT_ROW_LABEL") Then
        lastRow = ws.Cells(ws.rows.count, h("OUTPUT_ROW_LABEL")).End(xlUp).row
        For r = VS_FIRST_ROW To lastRow
            label = SafeUpperText(ws.Cells(r, h("OUTPUT_ROW_LABEL")).Value2)
            If Len(label) > 0 Then
                mLinks(label) = Array(CellAt(ws, r, h, "BASE_LINK"), _
                                      UseFromText(CellAt(ws, r, h, "BASE_SOURCE"), mBaseDefault), _
                                      CellAt(ws, r, h, "PRE_SHOCK_LINK"), _
                                      UseFromText(CellAt(ws, r, h, "PRE_SHOCK_SOURCE"), mPreDefault), r)
            End If
        Next r
    End If

    Set fh = HeaderMap(ws, VS_FIG_FIRST_COL, VS_FIG_COLS)
    If fh.Exists("FIGURE") Then
        lastRow = ws.Cells(ws.rows.count, fh("FIGURE")).End(xlUp).row
        For r = VS_FIRST_ROW To lastRow
            label = SafeUpperText(ws.Cells(r, fh("FIGURE")).Value2)
            If Len(label) > 0 Then
                mFigures.Add Array(label, ws.Cells(r, fh("VALUE")).Value2, _
                                   DateKey(ws.Cells(r, fh("AS_OF_DATE")).Value2), _
                                   SafeUpperText(ws.Cells(r, fh("ENTITY_CODE")).Value2))
            End If
        Next r
    End If
End Sub

Private Function HeaderMap(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal n As Long) As Object
    Dim c As Long, k As String, d As Object
    Set d = NewMap()
    For c = firstCol To firstCol + n - 1
        k = NormalHeader(ws.Cells(VS_HEADER_ROW, c).Value2)
        If Len(k) > 0 Then If Not d.Exists(k) Then d(k) = c
    Next c
    Set HeaderMap = d
End Function

Private Function CellAt(ByVal ws As Worksheet, ByVal r As Long, ByVal h As Object, ByVal key As String) As String
    If h.Exists(key) Then CellAt = NormalizeConfigFormulaText(ws.Cells(r, h(key)).Value2)
End Function

' A date typed as a date, a serial or text all compare as yyyy-mm-dd.
Private Function DateKey(ByVal v As Variant) As String
    On Error Resume Next
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsNumeric(v) And Not VarType(v) = vbString Then
        DateKey = format$(CDate(v), "yyyy-mm-dd")
    ElseIf IsDate(v) Then
        DateKey = format$(CDate(v), "yyyy-mm-dd")
    Else
        DateKey = Trim$(CStr(v))
    End If
    Err.Clear
End Function

' label -> Array(baseLink, baseUse, preLink, preUse, sheetRow)
Private Function LinkFor(ByVal label As String, ByRef ln As Variant) As Boolean
    EnsureLoaded
    If mLinks Is Nothing Then Exit Function
    If Not mLinks.Exists(SafeUpperText(label)) Then Exit Function
    ln = mLinks(SafeUpperText(label))
    LinkFor = True
End Function

' The best-matching figure: an exact date and entity beat a blank (any) one.
Private Function FigureValue(ByVal name As String, ByVal tc As Object, ByRef found As Boolean) As Double
    Dim f As Variant, score As Long, best As Long, want As String, ent As String
    EnsureLoaded
    found = False: best = -1
    If Not tc Is Nothing Then want = DateKey(tc("Date")): ent = SafeUpperText(tc("EntityCode"))
    For Each f In mFigures
        If f(0) = UCase$(name) Then
            score = -1
            ' Without a test case (the Check button) any entered value counts.
            If tc Is Nothing Or ((Len(f(2)) = 0 Or f(2) = want) And (Len(f(3)) = 0 Or f(3) = ent)) Then
                score = IIf(Len(f(2)) > 0, 2, 0) + IIf(Len(f(3)) > 0, 1, 0)
            End If
            If score > best Then
                If Not IsError(f(1)) And Not IsEmpty(f(1)) Then
                    If IsNumeric(f(1)) Then FigureValue = CDbl(f(1)): found = True: best = score
                End If
            End If
        End If
    Next f
End Function

' ---------------------------------------------------------------- compiling ----

' Numbers go into formulas in the invariant form Excel's Formula property expects,
' whatever the machine's decimal separator is.
Private Function NumLiteral(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then NumLiteral = "0": Exit Function
    If Not IsNumeric(v) Then NumLiteral = "0": Exit Function
    NumLiteral = Trim$(Str$(CDbl(v)))
    If Left$(NumLiteral, 1) = "-" Then NumLiteral = "(" & NumLiteral & ")"
End Function

Private Function ResolveFigures(ByVal s As String, ByVal tc As Object) As String
    Dim matches As Object, m As Object, i As Long, v As Double, found As Boolean, rep As String
    Set matches = RegexExecute("\bfigure\.([A-Za-z0-9_]+)\b", s)
    For i = matches.count - 1 To 0 Step -1
        Set m = matches(i)
        v = FigureValue(CStr(m.SubMatches(0)), tc, found)
        If found Then rep = NumLiteral(v) Else rep = "NA()"
        s = Left$(s, m.FirstIndex) & rep & Mid$(s, m.FirstIndex + m.Length + 1)
    Next i
    ResolveFigures = s
End Function

' The link as a formula body (no leading =). Derived tokens are pointed at the
' manual-file-only evidence columns (K base, M pre-shock) rather than at D and E,
' which hold whatever the metric's own setting picked. The row switch decides the
' fallback here, visibly, so a missing file can never pass as a manual value.
Private Function CompileLink(ByVal link As String, ByVal rm As Object, ByVal severity As Long, _
                             ByVal context As String, ByVal derivedRows As Object, ByVal tc As Object) As String
    Dim s As String, q As String
    s = ResolveFigures(link, tc)
    s = ConvertManualFormula(s, rm, severity, context, derivedRows)
    If Left$(s, 1) = "=" Then s = Mid$(s, 2)
    q = "'" & DERIVED_SHEET & "'!"
    s = Replace(s, q & "$D$", q & "$K$")
    s = Replace(s, q & "$E$", q & "$M$")
    CompileLink = s
End Function

' Every input a link needs, and which of them this test case could not rebuild.
Private Function LinkMissing(ByVal link As String, ByVal baseCache As Object, ByVal tc As Object, ByVal el As Object) As String
    Dim matches As Object, m As Object, kind As String, metric As String, rec As Object
    Dim flag As String, why As String, found As Boolean, out As String
    Set matches = RegexExecute("\b(derivedbase|base_derived|derived|figure)\.([A-Za-z0-9_]+)\b", link)
    For Each m In matches
        kind = UCase$(CStr(m.SubMatches(0))): metric = UCase$(CStr(m.SubMatches(1)))
        why = ""
        If kind = "FIGURE" Then
            FigureValue metric, tc, found
            If Not found Then why = "figure " & metric & " (not entered)"
        Else
            Set rec = Nothing
            If Not baseCache Is Nothing Then Set rec = DerivedLookup(baseCache, tc, el, metric)
            If rec Is Nothing Then
                why = metric & " (no input file loaded for it)"
            Else
                If kind = "DERIVED" Then flag = "HasPre" Else flag = "HasBase"
                If Not rec.Exists(flag) Then
                    why = metric
                ElseIf Not CBool(rec(flag)) Then
                    why = metric
                    If rec.Exists("Status") Then If Len(SafeText(rec("Status"))) > 0 Then why = why & " (" & SafeText(rec("Status")) & ")"
                End If
            End If
        End If
        If Len(why) > 0 Then If InStr(1, out, why, vbTextCompare) = 0 Then out = out & IIf(Len(out) > 0, "; ", "") & why
    Next m
    LinkMissing = out
End Function

Private Sub AddMark(ByVal marks As Collection, ByVal sheetRow As Long, ByVal col As Long, ByVal tone As String, ByVal note As String)
    If marks Is Nothing Then Exit Sub
    marks.Add Array(sheetRow, col, tone, note)
End Sub

Private Function SysText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then SysText = "(none)": Exit Function
    If IsNumeric(v) Then SysText = format$(CDbl(v), "#,##0.00##") Else SysText = CStr(v)
End Function

' ------------------------------------------------------- generator hooks ----

' Column B of an element sheet. Returns the formula to write, or "" to keep the
' system's literal base value (no link, or System output chosen).
Public Function VS_BaseCell(ByVal label As String, ByVal systemBase As Variant, ByVal rm As Object, _
                            ByVal baseCache As Object, ByVal tc As Object, ByVal el As Object, _
                            ByVal derivedRows As Object, ByVal sheetRow As Long, ByVal marks As Collection, _
                            ByVal context As String) As String
    Dim ln As Variant, body As String, missing As String
    On Error GoTo Failed
    If Not LinkFor(label, ln) Then Exit Function
    If Len(ln(0)) = 0 Or ln(1) = USE_SYSTEM Then Exit Function
    body = CompileLink(CStr(ln(0)), rm, 0, context & " (base)", derivedRows, tc)
    If Len(body) = 0 Then Exit Function
    missing = LinkMissing(CStr(ln(0)), baseCache, tc, el)
    If ln(1) = USE_DERIVED Then
        VS_BaseCell = "=" & body
        If Len(missing) > 0 Then
            AddMark marks, sheetRow, 2, TONE_RED, "Base: manual file only, and it could not be rebuilt." & vbLf & _
                    "Missing: " & missing & vbLf & "System base: " & SysText(systemBase)
        Else
            AddMark marks, sheetRow, 2, TONE_OK, "Base rebuilt from the input files: " & ln(0) & vbLf & "System base: " & SysText(systemBase)
        End If
    Else
        VS_BaseCell = "=IFERROR(" & body & "," & NumLiteral(systemBase) & ")"
        If Len(missing) > 0 Then
            AddMark marks, sheetRow, 2, TONE_AMBER, "Base: the input file value is not available, so this is the system base." & vbLf & _
                    "Missing: " & missing
        Else
            AddMark marks, sheetRow, 2, TONE_OK, "Base rebuilt from the input files: " & ln(0) & vbLf & "System base: " & SysText(systemBase)
        End If
    End If
    Exit Function
Failed:
    WarnOnce context & "|vs-base", "The base link could not be compiled, so the system base is kept: " & Err.description, context
    VS_BaseCell = ""
End Function

' A pre-shock cell (F:H). The link applies only where the config formula is blank
' or merely points at this row's own system value or at the link itself; a formula
' someone has written by hand always wins.
Public Function VS_PreShockFormula(ByVal label As String, ByVal cfgFormula As String, ByVal rm As Object, _
                                   ByVal severity As Long, ByVal context As String, ByVal baseCache As Object, _
                                   ByVal tc As Object, ByVal el As Object, ByVal derivedRows As Object, _
                                   ByVal sheetRow As Long, ByVal marks As Collection) As String
    Dim ln As Variant, body As String, missing As String, sysRef As String
    On Error GoTo Failed
    If Not LinkFor(label, ln) Then Exit Function
    If Len(ln(2)) = 0 Then Exit Function
    If Not VS_ConfigDefersToLink(cfgFormula, label, CStr(ln(2))) Then Exit Function
    sysRef = colLetter(3 + severity) & sheetRow
    If ln(3) = USE_SYSTEM Then VS_PreShockFormula = "=" & sysRef: Exit Function
    body = CompileLink(CStr(ln(2)), rm, severity, context & " (pre-shock)", derivedRows, tc)
    If Len(body) = 0 Then Exit Function
    missing = LinkMissing(CStr(ln(2)), baseCache, tc, el)
    If ln(3) = USE_DERIVED Then
        VS_PreShockFormula = "=" & body
        If Len(missing) > 0 Then AddMark marks, sheetRow, 6 + severity, TONE_RED, _
            "Pre-shock: manual file only, and it could not be rebuilt." & vbLf & "Missing: " & missing
    Else
        VS_PreShockFormula = "=IFERROR(" & body & "," & sysRef & ")"
        If Len(missing) > 0 Then AddMark marks, sheetRow, 6 + severity, TONE_AMBER, _
            "Pre-shock: the input file value is not available, so this is the system value." & vbLf & "Missing: " & missing
    End If
    Exit Function
Failed:
    WarnOnce context & "|vs-pre", "The pre-shock link could not be compiled, so the config formula is kept: " & Err.description, context
    VS_PreShockFormula = ""
End Function

Public Function VS_ConfigDefersToLink(ByVal cfgFormula As String, ByVal label As String, ByVal link As String) As Boolean
    Dim t As String, k As String
    t = Squash(cfgFormula): k = SafeUpperText(label)
    VS_ConfigDefersToLink = (Len(t) = 0 Or t = "SYSTEMOUTPUT." & k Or t = "SYSTEM." & k Or _
                             t = "DERIVED." & k Or t = Squash(link))
End Function

Private Function Squash(ByVal s As String) As String
    s = SafeUpperText(NormalizeConfigFormulaText(s))
    s = Replace(Replace(Replace(s, "@", ""), " ", ""), vbTab, "")
    Squash = s
End Function

Public Sub VS_WriteBaseFormulas(ByVal ws As Worksheet, ByRef baseFormulas As Variant)
    Dim i As Long
    If Not IsArray(baseFormulas) Then Exit Sub
    For i = LBound(baseFormulas) To UBound(baseFormulas)
        If Len(CStr(baseFormulas(i))) > 0 Then ws.Cells(OUTPUT_FIRST_DATA_ROW + i - 1, 2).formula = CStr(baseFormulas(i))
    Next i
End Sub

' After FormatReport, so the report's own styling cannot paint over the marks.
Public Sub VS_MarkElementSheet(ByVal ws As Worksheet, ByVal marks As Collection)
    Dim m As Variant, c As Range
    If marks Is Nothing Then Exit Sub
    On Error Resume Next
    For Each m In marks
        Set c = ws.Cells(CLng(m(0)), CLng(m(1)))
        Select Case CStr(m(2))
            Case TONE_RED
                c.Interior.Color = UI_BAD_BG: c.Font.Color = UI_BAD
            Case TONE_AMBER
                c.Interior.Color = UI_WARN_BG: c.Font.Color = UI_WARN
        End Select
        If Len(CStr(m(3))) > 0 Then
            c.ClearComments
            c.AddComment CStr(m(3))
            c.Comment.Shape.TextFrame.AutoSize = True
        End If
        Err.Clear
    Next m
End Sub

' The Derived_Values evidence columns K:N - the input-file value and the system
' value side by side, whatever the metric's own setting picked for D and E.
Public Function VS_Evidence(ByVal rec As Object, ByVal hasKey As String, ByVal valueKey As String) As Variant
    VS_Evidence = CVErr(xlErrNA)
    If rec Is Nothing Then Exit Function
    If Len(hasKey) > 0 Then
        If Not rec.Exists(hasKey) Then Exit Function
        If Not CBool(rec(hasKey)) Then Exit Function
    End If
    If Not rec.Exists(valueKey) Then Exit Function
    If IsError(rec(valueKey)) Or IsEmpty(rec(valueKey)) Then Exit Function
    If IsNumeric(rec(valueKey)) Then VS_Evidence = CDbl(rec(valueKey))
End Function

Public Sub VS_StyleEvidence(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal n As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(DV_HEADER_ROW, 11), ws.Cells(DV_HEADER_ROW, 14))
        .Font.Bold = True
        .Font.Color = UI_WHITE
        .Interior.Color = UI_NAVY
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With
    If n > 0 Then
        ws.Range(ws.Cells(DV_FIRST_ROW, 11), ws.Cells(lastRow, 14)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(DV_HEADER_ROW, 11), ws.Cells(lastRow, 14)).Borders.Color = UI_LINE
    End If
    ws.columns("K:N").ColumnWidth = 17
    Err.Clear
End Sub

' --------------------------------------------------- sources on the sheet ----

Public Sub UploadNSFRSource()
    UploadSource "NSFR"
End Sub

' Metrics the links need that Pre_Shock_Metrics did not have. Appended by
' AppendMissingMetrics like every other seed, so an edited row is never replaced.
'   name | source | formula | stage | purpose | where
Public Function VS_SeedMetrics() As Variant
    Dim a(0 To 9) As Variant, i As Long, s As String, stg As Variant, mx As String, ecl As String
    mx = "GREATEST(ORI_MOODY_STRESS_ECL|MOODY_STRESS_ECL|STRESS_ECL, ECL|OVERRIDDEN_ECL|CALCULATED_ECL)"
    ecl = "ECL|OVERRIDDEN_ECL|CALCULATED_ECL"
    For i = 1 To 3
        stg = Array("", "STAGE1", "STAGE2", "STAGE3")
        s = CStr(stg(i))
        a(i - 1) = Array("MAX_ECL_STRESS_ECL_" & s, "ECL", "SUM(" & mx & ") WHERE STAGE=" & i, CStr(i), _
                         "Moody's stress ECL or ECL, whichever is larger per account, stage " & i, _
                         "ST Design: max(stress ECL, ECL) per account. Needs the Moody's stress ECL column in the ECL file.")
        a(i + 2) = Array("IMPACT_ECL_LCY_MOODY_" & s, "ECL", "SUM(" & mx & ")-SUM(" & ecl & ") WHERE STAGE=" & i, CStr(i), _
                         "Impact of Moody's stress ECL over ECL, stage " & i, "Per-account max less ECL.")
    Next i
    a(6) = Array("RWA_CR_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_RWA_CREDIT_RISK')", "ALL", _
                 "Credit risk RWA, capital component", "ST Design base for RWA_CR_LCY.")
    a(7) = Array("RWA_OR_BASE", "CAP", "SUM(REPORT_BALANCE) WHERE BANK_ID IN ('101') AND CAP_COMPONENT_CODE IN ('NET_RWA_OPERATIONAL_RISK')", "ALL", _
                 "Operational risk RWA, capital component", "Confirm the component code against the capital component file.")
    a(8) = Array("NSFR_ASF_PRE_SHOCK", "NSFR", "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY IN ('ASF')", "ALL", _
                 "NSFR - available stable funding", "Category code from the ST Design; confirm against the NSFR file.")
    a(9) = Array("NSFR_RSF_PRE_SHOCK", "NSFR", "SUM(CASHFLOW_AMOUNT_LCY_POST_FACTOR) WHERE ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY IN ('RSF')", "ALL", _
                 "NSFR - required stable funding", "Category code from the ST Design; confirm against the NSFR file.")
    VS_SeedMetrics = a
End Function

' ------------------------------------------------------------------- setup ----

' One-time setup, safe to run again. It:
'   - adds the new metric rows (Moody's per-account max, capital RWA, NSFR) to
'     Pre_Shock_Metrics and registers the NSFR source;
'   - relabels the value-source choices to Manual file only / Manual else system /
'     System output, keeping what each row had chosen;
'   - builds Config_ValueSources with a link for every value row, and the Manual
'     figures table; rows someone has edited are kept unless resetLinks is True;
'   - rewrites pre-shock config formulas that were only systemoutput.<row> to the
'     row's pre-shock link, so the config itself says where the value comes from.
Public Sub JKB_ApplyValueSourceSetup(Optional ByVal quiet As Boolean = False, Optional ByVal resetLinks As Boolean = False)
    Dim ws As Worksheet, st As Object, added As Long, figAdded As Long, cfgChanged As Long, er As String
    On Error GoTo Failed
    Set st = CaptureState()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    EnsurePreShockWorkspace
    RegisterNsfrRow
    Set ws = EnsureValueSourcesSheet()
    added = SeedLinkTable(ws, resetLinks)
    figAdded = SeedFigureTable(ws, resetLinks)
    ConsolidateSwitches ws
    FinishValueSourcesSheet ws
    VS_ResetCache
    cfgChanged = LinkConfigFormulas()
    VS_ResetCache
    RestoreState st
    LogIssue LOG_LEVEL_INFO, "Value sources", added & " link row(s) and " & figAdded & " figure row(s) added; " & _
             cfgChanged & " config formula(s) now point at their pre-shock link.", VS_SHEET
    If Not quiet Then
        UiNotice "Value sources", "Value sources are set up.", _
                 added & " link row(s) added to " & VS_SHEET & vbCrLf & figAdded & " manual figure row(s) added" & vbCrLf & _
                 cfgChanged & " pre-shock config formula(s) that only copied the system value now use the link", _
                 "Click Check sources on " & VS_SHEET & " to see which input files each row still needs."
    End If
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    RestoreState st
    LogIssue LOG_LEVEL_ERROR, "Value sources", "Setup stopped: " & er, VS_SHEET
    If Not quiet Then UiProblem "Value sources", "Value source setup stopped.", er
End Sub

Private Sub RegisterNsfrRow()
    Dim ws As Worksheet, i As Long, k As Variant
    Set ws = PsSourcesSheet()
    If ws Is Nothing Then Exit Sub
    For Each k In SourceKeys()
        If CStr(k) = "NSFR" Then
            If Len(SafeText(ws.Cells(PS_SRC_LIST_ROW + i, 1).Value2)) = 0 Then
                ws.Cells(PS_SRC_LIST_ROW + i, 1).Value2 = "NSFR"
                ws.Cells(PS_SRC_LIST_ROW + i, 2).Value2 = "NSFR ALM output - available and required stable funding"
                ws.Cells(PS_SRC_LIST_ROW + i, 8).Value2 = "Not loaded"
            End If
        End If
        i = i + 1
    Next k
End Sub

Private Sub ApplyList(ByVal rng As Range)
    On Error Resume Next
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=VS_Choices()
        .IgnoreBlank = True
        .InputTitle = "Where the value comes from"
        .InputMessage = "Manual file only: input files, missing shows as missing." & vbLf & _
                        "Manual else system: input files, system value where a file is missing." & vbLf & _
                        "System output: the system value." & vbLf & "Blank: the default at the top."
    End With
    Err.Clear
End Sub

Private Function EnsureValueSourcesSheet() As Worksheet
    Dim ws As Worksheet, anchorWs As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_SHEET)
    If ws Is Nothing Then
        Set anchorWs = GetWorksheetSafe(ThisWorkbook, SHEET_RULES)
        If anchorWs Is Nothing Then Set anchorWs = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
        Set ws = ThisWorkbook.Worksheets.Add(After:=anchorWs)
        ws.name = VS_SHEET
    End If
    ws.Range("A1").Value2 = "JKB"
    ' Rows 1 and 2 belong to the sheet toolbar (Sheet1.EnsureRibbon sets row 2 to
    ' 4pt), so the title starts on row 3.
    ws.Range("A2").ClearContents
    ws.Range("A3").Value2 = "Where base and pre-shock values come from"
    ws.Range("A4").Value2 = "Scenario  >  test case / element  >  BASE (input files, whole portfolio)  >  PRE-SHOCK (input files, test case filter)  >  shock  >  POST-SHOCK"
    ws.Range("A5").Value2 = "Each value row links to the figure the input files should produce. derivedbase.X is metric X from Pre_Shock_Metrics over the whole portfolio, " & _
                            "derived.X is the same metric filtered to the test case, figure.X is a value typed into Manual figures. Post-shock formulas that use base.X start from this base."
    ws.Range("A6").Value2 = "Base values come from"
    ws.Range("A7").Value2 = "Pre-shock values come from"
    If Len(SafeText(ws.Range("C6").Value2)) = 0 Then ws.Range("C6").Value2 = USE_DERIVED_ELSE
    If Len(SafeText(ws.Range("C7").Value2)) = 0 Then ws.Range("C7").Value2 = USE_DERIVED_ELSE
    ws.Range("D6").Value2 = "Default for rows whose BASE_SOURCE is blank"
    ws.Range("D7").Value2 = "Default for rows whose PRE_SHOCK_SOURCE is blank"
    ws.Cells(VS_HEADER_ROW, 1).Resize(1, VS_LINK_COLS).Value2 = Array("OUTPUT_ROW_LABEL", "CHAIN_STEP", "BASE_LINK", "BASE_SOURCE", _
        "PRE_SHOCK_LINK", "PRE_SHOCK_SOURCE", "INPUT_FILES", "ST_DESIGN_BASE", "STATUS", "NOTES")
    ws.Cells(VS_HEADER_ROW - 1, VS_FIG_FIRST_COL).Value2 = "Manual figures - typed, or read from their own file with Refresh figures from files"
    ws.Cells(VS_HEADER_ROW, VS_FIG_FIRST_COL).Resize(1, VS_FIG_COLS).Value2 = Array("FIGURE", "VALUE", "AS_OF_DATE", "ENTITY_CODE", "SOURCE", "NOTES", _
        "FILE", "SHEET", "SUM_COLUMN", "WHERE", "LOADED")
    Set EnsureValueSourcesSheet = ws
End Function

Private Sub SeedLink(ByVal d As Collection, ByVal label As String, ByVal stepName As String, ByVal baseLink As String, _
                     ByVal preLink As String, ByVal files As String, ByVal design As String, ByVal note As String)
    d.Add Array(label, stepName, baseLink, preLink, files, design, note)
End Sub

Private Function LastDataRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastDataRow = ws.Cells(ws.rows.count, col).End(xlUp).row
    If LastDataRow < VS_HEADER_ROW Then LastDataRow = VS_HEADER_ROW
End Function

Private Function SeedLinkTable(ByVal ws As Worksheet, ByVal reset As Boolean) As Long
    Dim have As Object, r As Long, lastRow As Long, s As Variant, n As Long
    lastRow = LastDataRow(ws, 1)
    If reset And lastRow >= VS_FIRST_ROW Then
        ws.Range(ws.Cells(VS_FIRST_ROW, 1), ws.Cells(lastRow, VS_LINK_COLS)).Clear
        lastRow = VS_HEADER_ROW
    End If
    Set have = NewMap()
    For r = VS_FIRST_ROW To lastRow
        If Len(SafeText(ws.Cells(r, 1).Value2)) > 0 Then have(SafeUpperText(ws.Cells(r, 1).Value2)) = True
    Next r
    r = lastRow + 1
    ws.Range(ws.Cells(VS_FIRST_ROW, 3), ws.Cells(VS_FIRST_ROW + 400, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(VS_FIRST_ROW, 5), ws.Cells(VS_FIRST_ROW + 400, 5)).NumberFormat = "@"
    For Each s In SeedLinkRows()
        If Not have.Exists(UCase$(CStr(s(0)))) Then
            ws.Cells(r, 1).Resize(1, VS_LINK_COLS).Value2 = Array(s(0), s(1), s(2), "", s(3), "", s(4), s(5), "", s(6))
            r = r + 1: n = n + 1
        End If
    Next s
    SeedLinkTable = n
End Function

Private Function SeedFigureTable(ByVal ws As Worksheet, ByVal reset As Boolean) As Long
    Dim have As Object, r As Long, lastRow As Long, s As Variant, n As Long, seeds As Variant
    lastRow = LastDataRow(ws, VS_FIG_FIRST_COL)
    If reset And lastRow >= VS_FIRST_ROW Then
        ws.Range(ws.Cells(VS_FIRST_ROW, VS_FIG_FIRST_COL), ws.Cells(lastRow, VS_FIG_FIRST_COL + VS_FIG_COLS - 1)).Clear
        lastRow = VS_HEADER_ROW
    End If
    Set have = NewMap()
    For r = VS_FIRST_ROW To lastRow
        If Len(SafeText(ws.Cells(r, VS_FIG_FIRST_COL).Value2)) > 0 Then have(SafeUpperText(ws.Cells(r, VS_FIG_FIRST_COL).Value2)) = True
    Next r
    seeds = Array( _
        Array("FX_LONG_POSITION", "cap_market_risk_forex_total.total_long_position", "FX pre-shock long position"), _
        Array("FX_SHORT_POSITION", "cap_market_risk_forex_total.total_short_position", "FX pre-shock short position"), _
        Array("MARKET_VALUE_FVTPL", "cap_market_risk_equity: sum(market_value_lcy), FVTPL", "Market value of FVTPL equity"), _
        Array("PBT", "stresstest.aod_pbt_pat_data", "Profit before tax"), _
        Array("PAT", "stresstest.aod_pbt_pat_data", "Profit after tax"), _
        Array("RSA", "GL / balance sheet: current and non-current assets, rate sensitive", "Rate sensitive assets"), _
        Array("RSL", "GL / balance sheet: current and non-current liabilities, rate sensitive", "Rate sensitive liabilities"), _
        Array("RWA_MR_EQUITY", "Market risk RWA - equity", "Part of TOTAL_RWA"), _
        Array("RWA_MR_FOREX", "Market risk RWA - forex", "Part of TOTAL_RWA"))
    r = lastRow + 1
    For Each s In seeds
        If Not have.Exists(CStr(s(0))) Then
            ws.Cells(r, VS_FIG_FIRST_COL).Resize(1, 6).Value2 = Array(s(0), "", "", "", s(1), s(2))
            r = r + 1: n = n + 1
        End If
    Next s
    SeedFigureColumns ws
    SeedFigureTable = n
End Function

Private Sub FinishValueSourcesSheet(ByVal ws As Worksheet)
    Dim lastLink As Long, lastFig As Long
    On Error Resume Next
    lastLink = LastDataRow(ws, 1)
    lastFig = LastDataRow(ws, VS_FIG_FIRST_COL)
    ApplyList ws.Range("C6:C7")
    ApplyList ws.Range(ws.Cells(VS_FIRST_ROW, 4), ws.Cells(VS_FIRST_ROW + 400, 4))
    ApplyList ws.Range(ws.Cells(VS_FIRST_ROW, 6), ws.Cells(VS_FIRST_ROW + 400, 6))

    ws.Cells.Font.name = UI_FONT
    ws.Cells.Font.Size = 9.5
    ws.Range("A1").Font.Color = UI_WHITE
    ws.rows(3).RowHeight = 26
    ws.Range("A3").Font.Size = 15: ws.Range("A3").Font.Bold = True: ws.Range("A3").Font.Color = UI_INK
    ws.Range("A4").Font.Bold = True: ws.Range("A4").Font.Color = UI_TEXT_2
    ws.Range("A5").Font.Color = UI_MUTED
    ws.Range("A6:A7").Font.Bold = True
    With ws.Range("C6:C7")
        .Interior.Color = UI_BRAND_TINT
        .Borders.Color = UI_LINE
        .Font.Bold = True
    End With
    ws.Range("D6:D7").Font.Color = UI_MUTED
    With ws.Range(ws.Cells(VS_HEADER_ROW, 1), ws.Cells(VS_HEADER_ROW, VS_LINK_COLS))
        .Font.Bold = True: .Font.Color = UI_WHITE: .Interior.Color = UI_NAVY
    End With
    With ws.Range(ws.Cells(VS_HEADER_ROW, VS_FIG_FIRST_COL), ws.Cells(VS_HEADER_ROW, VS_FIG_FIRST_COL + VS_FIG_COLS - 1))
        .Font.Bold = True: .Font.Color = UI_WHITE: .Interior.Color = UI_NAVY
    End With
    ws.Cells(VS_HEADER_ROW - 1, VS_FIG_FIRST_COL).Font.Bold = True
    ' The two switch columns and the figure values are where people type.
    ws.Range(ws.Cells(VS_FIRST_ROW, 4), ws.Cells(lastLink, 4)).Interior.Color = UI_BRAND_TINT
    ws.Range(ws.Cells(VS_FIRST_ROW, 6), ws.Cells(lastLink, 6)).Interior.Color = UI_BRAND_TINT
    ws.Range(ws.Cells(VS_FIRST_ROW, VS_FIG_FIRST_COL + 1), ws.Cells(lastFig, VS_FIG_FIRST_COL + 1)).Interior.Color = UI_BRAND_TINT
    ws.Range(ws.Cells(VS_FIRST_ROW, VS_FIG_FIRST_COL + 1), ws.Cells(lastFig, VS_FIG_FIRST_COL + 1)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(VS_HEADER_ROW, 1), ws.Cells(lastLink, VS_LINK_COLS)).Borders.Color = UI_LINE
    ws.Range(ws.Cells(VS_HEADER_ROW, VS_FIG_FIRST_COL), ws.Cells(lastFig, VS_FIG_FIRST_COL + VS_FIG_COLS - 1)).Borders.Color = UI_LINE
    ws.columns("A").ColumnWidth = 40
    ws.columns("B").ColumnWidth = 15
    ws.columns("C").ColumnWidth = 48
    ws.columns("D").ColumnWidth = 19
    ws.columns("E").ColumnWidth = 40
    ws.columns("F").ColumnWidth = 19
    ws.columns("G").ColumnWidth = 30
    ws.columns("H").ColumnWidth = 50
    ws.columns("I").ColumnWidth = 40
    ws.columns("J").ColumnWidth = 50
    ws.columns("K:L").ColumnWidth = 3
    ws.columns("M").ColumnWidth = 22
    ws.columns("N").ColumnWidth = 16
    ws.columns("O:P").ColumnWidth = 13
    ws.columns("Q:R").ColumnWidth = 44
    ws.columns("S").ColumnWidth = 40: ws.columns("T").ColumnWidth = 16: ws.columns("U").ColumnWidth = 24
    ws.columns("V").ColumnWidth = 40: ws.columns("W").ColumnWidth = 36
    ws.Range(ws.Cells(VS_FIRST_ROW, 1), ws.Cells(lastLink, VS_LINK_COLS)).VerticalAlignment = xlTop
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(VS_HEADER_ROW, 1), ws.Cells(lastLink, VS_LINK_COLS)).AutoFilter
    If ws.rows(6).RowHeight < 26 Then ws.rows(6).RowHeight = 26
    If ws.rows(7).RowHeight < 26 Then ws.rows(7).RowHeight = 26
    AddButton ws, "JKB_VS_Check", "Check sources", "JKB_CheckValueSources", "F6", UI_BTN_PRIMARY
    AddButton ws, "JKB_VS_Load", "Load inputs", "UploadAllSources", "G6"
    ws.Shapes("JKB_VS_Reports").Delete   ' Reports: one review pack, from Home.
    ' NSFR is one of the choices under "Load one input" (and the folder loader
    ' recognises it), so the sheet no longer carries a loader of its own.
    AddButton ws, "JKB_VS_One", "Load one input", "UiLoadOneSource", "G7"
    AddButton ws, "JKB_VS_Figures", "Refresh figures from files", "JKB_RefreshFiguresFromFiles", "H7"
    ws.Shapes("JKB_VS_NSFR").Delete
    Err.Clear
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal macro As String, _
                      ByVal anchor As String, Optional ByVal kind As String = "secondary")
    ' Design system v2 button (modUI_Theme), 22pt high, sized to its caption.
    Dim sh As Shape
    On Error Resume Next
    UiButton ws, nm, caption, macro, ws.Range(anchor), kind, 9.5
    Set sh = ws.Shapes(nm)
    If Not sh Is Nothing Then
        sh.Left = ws.Range(anchor).Left + 4
        sh.Width = Len(caption) * 5.6 + 28: sh.Height = 22
        sh.Top = ws.Range(anchor).Top + (ws.Range(anchor).Height - sh.Height) / 2
    End If
    Err.Clear
End Sub

' Config rows whose pre-shock formula only copied the system value get the link
' text instead. A figure link is left out: figure.X is resolved by the generator,
' not by the config formula compiler.
Private Function LinkConfigFormulas() As Long
    Dim lo As ListObject, a As Variant, h As Object, hdr As Variant, c As Long, r As Long
    Dim label As String, ln As Variant, k As Variant, t As String, n As Long
    On Error GoTo Done
    Set lo = ThisWorkbook.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set h = NewMap()
    hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    If Not h.Exists("OUTPUT_ROW_LABEL") Then Exit Function
    a = lo.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        label = SafeUpperText(a(r, h("OUTPUT_ROW_LABEL")))
        If Len(label) > 0 Then
            If LinkFor(label, ln) Then
                If Len(ln(2)) > 0 And InStr(1, CStr(ln(2)), "figure.", vbTextCompare) = 0 Then
                    For Each k In Array("MANUAL_FORMULA_DEFAULT", "MANUAL_FORMULA_MODERATE", "MANUAL_FORMULA_MEDIUM", "MANUAL_FORMULA_SEVERE")
                        If h.Exists(CStr(k)) Then
                            t = Squash(SafeText(a(r, h(CStr(k)))))
                            If t = "SYSTEMOUTPUT." & label Or t = "SYSTEM." & label Then
                                lo.DataBodyRange.Cells(r, h(CStr(k))).Value2 = CStr(ln(2))
                                n = n + 1
                            End If
                        End If
                    Next k
                End If
            End If
        End If
    Next r
Done:
    LinkConfigFormulas = n
End Function

' ------------------------------------------------------------------- check ----

' Writes a STATUS per row: Ready, or exactly which input file, metric or figure
' is still missing. Nothing is generated, so it is quick to run after loading.
Public Sub JKB_CheckValueSources()
    Dim ws As Worksheet, h As Object, r As Long, lastRow As Long, metrics As Object, own As Object
    Dim bs As String, ps As String, status As String, nReady As Long, nNeeds As Long, nProblem As Long, nNone As Long
    Dim label As String, cell As Range
    On Error GoTo Failed
    VS_ResetCache
    EnsureLoaded
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_SHEET)
    If ws Is Nothing Then
        UiNotice "Value sources", "There is no " & VS_SHEET & " sheet yet.", , "Click Value sources under Base and pre-shock on the home screen. It sets the sheet up."
        Exit Sub
    End If
    Set metrics = MetricIndex()
    Set own = OwnFormulaCounts()
    Set h = HeaderMap(ws, 1, VS_LINK_COLS)
    If Not h.Exists("STATUS") Or Not h.Exists("OUTPUT_ROW_LABEL") Then Exit Sub
    lastRow = LastDataRow(ws, h("OUTPUT_ROW_LABEL"))
    For r = VS_FIRST_ROW To lastRow
        label = SafeUpperText(ws.Cells(r, h("OUTPUT_ROW_LABEL")).Value2)
        If Len(label) > 0 Then
            bs = CheckOne(CellAt(ws, r, h, "BASE_LINK"), UseFromText(CellAt(ws, r, h, "BASE_SOURCE"), mBaseDefault), "Base", metrics)
            ps = CheckOne(CellAt(ws, r, h, "PRE_SHOCK_LINK"), UseFromText(CellAt(ws, r, h, "PRE_SHOCK_SOURCE"), mPreDefault), "Pre-shock", metrics)
            If Len(bs) = 0 And Len(ps) = 0 Then
                status = "No link: system value": nNone = nNone + 1
            Else
                status = bs
                If Len(ps) > 0 Then status = status & IIf(Len(status) > 0, " | ", "") & ps
                If InStr(status, "Unknown") > 0 Or InStr(status, "switched off") > 0 Then
                    nProblem = nProblem + 1
                ElseIf InStr(status, "needs") > 0 Then
                    nNeeds = nNeeds + 1
                Else
                    nReady = nReady + 1
                End If
            End If
            If own.Exists(label) And Len(CellAt(ws, r, h, "PRE_SHOCK_LINK")) > 0 Then
                status = status & " | " & own(label) & " config row(s) keep their own pre-shock formula"
            End If
            Set cell = ws.Cells(r, h("STATUS"))
            cell.Value2 = status
            If InStr(status, "Unknown") > 0 Or InStr(status, "switched off") > 0 Then
                cell.Interior.Color = UI_BAD_BG
            ElseIf InStr(status, "needs") > 0 Then
                cell.Interior.Color = UI_WARN_BG
            ElseIf Left$(status, 8) = "No link:" Then
                cell.Interior.Color = UI_FILL
            Else
                cell.Interior.Color = UI_OK_BG
            End If
        End If
    Next r
    DressWorkSheet ws
    UiNotice "Check sources", nReady & " row(s) are ready: every input they need is loaded.", _
             nNeeds & " row(s) still need an input file or a figure (see STATUS)." & vbCrLf & _
             nProblem & " row(s) point at a metric that does not exist or is switched off." & vbCrLf & _
             nNone & " row(s) have no link and keep the system value.", , (nProblem > 0)
    Exit Sub
Failed:
    UiProblem "Check sources", "The check stopped.", Err.description
End Sub

' name -> Array(source, enabled, availability text)
Private Function MetricIndex() As Object
    Dim ws As Worksheet, r As Long, nm As String, d As Object
    Set d = NewMap()
    Set MetricIndex = d
    Set ws = PsMetricsSheet()
    If ws Is Nothing Then Exit Function
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        nm = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(nm) = 0 Then Exit For
        d(nm) = Array(NormalizeSourceKey(SafeText(ws.Cells(r, 2).Value2)), _
                                SafeUpperText(ws.Cells(r, 3).Value2) <> "NO", SafeText(ws.Cells(r, 7).Value2))
    Next r
End Function

' label -> how many config rows carry a hand-written pre-shock formula.
Private Function OwnFormulaCounts() As Object
    Dim lo As ListObject, a As Variant, h As Object, hdr As Variant, c As Long, r As Long, label As String, ln As Variant
    Dim d As Object
    Set d = NewMap()
    Set OwnFormulaCounts = d
    On Error GoTo Done
    Set lo = ThisWorkbook.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set h = NewMap(): hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    a = lo.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        label = SafeUpperText(a(r, h("OUTPUT_ROW_LABEL")))
        If ParseBool(a(r, h("IS_ENABLED"))) And LinkFor(label, ln) Then
            If Len(ln(2)) > 0 Then
                If Not VS_ConfigDefersToLink(SafeText(a(r, h("MANUAL_FORMULA_DEFAULT"))), label, CStr(ln(2))) Then d(label) = d(label) + 1
            End If
        End If
    Next r
Done:
End Function

Private Function CheckOne(ByVal link As String, ByVal use As String, ByVal phase As String, ByVal metrics As Object) As String
    Dim matches As Object, m As Object, kind As String, nm As String, needs As String, bad As String, info As Variant, found As Boolean
    If Len(link) = 0 Then Exit Function
    If use = USE_SYSTEM Then CheckOne = phase & ": system output by choice": Exit Function
    Set matches = RegexExecute("\b(derivedbase|base_derived|derived|figure)\.([A-Za-z0-9_]+)\b", link)
    For Each m In matches
        kind = UCase$(CStr(m.SubMatches(0))): nm = UCase$(CStr(m.SubMatches(1)))
        If kind = "FIGURE" Then
            FigureValue nm, Nothing, found
            If Not found Then AddText needs, "figure " & nm
        ElseIf Not metrics.Exists(nm) Then
            AddText bad, "Unknown metric " & nm
        Else
            info = metrics(nm)
            If Not CBool(info(1)) Then
                AddText bad, nm & " is switched off"
            ElseIf Not SourceLoaded(CStr(info(0))) Then
                AddText needs, CStr(info(0)) & " file"
            ElseIf Left$(SafeUpperText(info(2)), 7) = "MISSING" Then
                AddText needs, nm & " " & LCase$(CStr(info(2)))
            End If
        End If
    Next m
    If Len(bad) > 0 Then
        CheckOne = phase & ": " & bad
    ElseIf Len(needs) > 0 Then
        CheckOne = phase & " needs " & needs & IIf(use = USE_DERIVED, " (manual file only: shows as missing)", " (system value used meanwhile)")
    Else
        CheckOne = phase & ": ready"
    End If
End Function

Private Sub AddText(ByRef s As String, ByVal t As String)
    If InStr(1, s, t, vbTextCompare) > 0 Then Exit Sub
    If Len(s) > 0 Then s = s & ", "
    s = s & t
End Sub

' ----------------------------------------------------------------- reports ----

' The pre-shock pivots, in a new dated folder beside the manual reports:
'   <tool folder>\Output\PreShock_Reports_<yyyymmdd_hhnnss>\
' holding the manual report workbook (base and pre-shock by test case, with the
' value each row used) and the joined input workbook (source records with the
' test case filters applied, and native base and test case pivots).
Public Sub JKB_BuildPreShockPivots()
    Dim root As String, folder As String, st As Object, p1 As String, p2 As String, e1 As String, e2 As String
    If JKB_Busy Then Exit Sub
    If PersistedDerivedCount() = 0 Then
        UiNotice "Pre-shock reports", "There is nothing to report on yet.", , "Load the input files, then click Run on the home screen."
        Exit Sub
    End If
    root = ThisWorkbook.path
    If Len(root) = 0 Then root = Application.DefaultFilePath
    root = JoinPath(root, "Output")
    On Error Resume Next
    EnsureFolder root
    folder = JoinPath(root, "PreShock_Reports_" & format$(Now, "yyyymmdd_hhnnss"))
    EnsureFolder folder
    If Err.Number <> 0 Then
        UiProblem "Pre-shock reports", "The report folder could not be created.", folder & vbCrLf & Err.description
        Exit Sub
    End If
    Err.Clear

    Set st = CaptureState(): JKB_Busy = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ProgressStart "Building the pre-shock reports", "Into " & folder, 100
    ProgressStep "Manual report workbook", 10
    p1 = Ps_BuildManualReportsTo(folder)
    If Err.Number <> 0 Then e1 = Err.description: Err.Clear
    ProgressStep "Joined input workbook", 40
    JKB_Busy = False
    p2 = BuildJoinedInputTo(folder, False)
    If Err.Number <> 0 Then e2 = Err.description: Err.Clear
    ProgressDone "Written."
    RestoreState st: JKB_Busy = False
    On Error GoTo 0

    LogIssue LOG_LEVEL_INFO, "Pre-shock reports", "Written to " & folder, DERIVED_SHEET
    UiNotice "Pre-shock reports", "The pre-shock reports are in " & folder, _
             "Manual reports: " & IIf(Len(e1) > 0, "not built - " & e1, FileLeafOf(p1)) & vbCrLf & _
             "Joined input: " & IIf(Len(e2) > 0, "not built - " & e2, FileLeafOf(p2)), , (Len(e1) + Len(e2) > 0)
    On Error Resume Next
    If Len(e1) = 0 Or Len(e2) = 0 Then Shell "explorer.exe """ & folder & """", vbNormalFocus
End Sub

Private Function FileLeafOf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, Application.PathSeparator)
    If i > 0 Then FileLeafOf = Mid$(p, i + 1) Else FileLeafOf = p
    If Len(FileLeafOf) = 0 Then FileLeafOf = "(written)"
End Function

' ------------------------------------------------------------- link seeds ----
'
' One row per value row in Config_ElementTypeRules. Generated from the ST Design
' (10 Jul 2026) base definitions; see value_links.csv beside this module.

Private Function SeedLinkRows() As Collection
    Dim d As Collection: Set d = New Collection
    SeedLinks1 d
    SeedLinks2 d
    SeedLinks3 d
    SeedLinks4 d
    Set SeedLinkRows = d
End Function

Private Sub SeedLinks1(ByVal d As Collection)
    SeedLink d, "AMOUNT_CHANGE", "Shock input", "", "", "Scenario definition (system)", "", "Shock parameter set in the scenario, not derived from an input file."
    SeedLink d, "PCT_CHANGE", "Shock input", "", "", "Scenario definition (system)", "", "Shock parameter set in the scenario, not derived from an input file."
    SeedLink d, "NUM_CUSTOMERS", "Shock input", "", "", "Scenario definition (system)", "", "Shock parameter set in the scenario, not derived from an input file."
    SeedLink d, "FX_TOTAL_LONG_POSITION_LCY_PRE_SHOCK", "Pre-shock", "figure.FX_LONG_POSITION", "figure.FX_LONG_POSITION", "Market risk forex total (cap_market_risk_forex_total)", "cap_market_risk_forex_total.total_long_position", "Not deal level, so pre-shock equals base."
    SeedLink d, "FX_TOTAL_SHORT_POSITION_LCY_PRE_SHOCK", "Pre-shock", "figure.FX_SHORT_POSITION", "figure.FX_SHORT_POSITION", "Market risk forex total (cap_market_risk_forex_total)", "cap_market_risk_forex_total.total_short_position", "Not deal level, so pre-shock equals base."
    SeedLink d, "OUTST_LCY_STAGE1_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_LCY_STAGE1_PRE_SHOCK", "derived.OUTST_LCY_STAGE1_PRE_SHOCK", "ECL output", "EO.outst_lcy for EO.stage_id=1", ""
    SeedLink d, "OUTST_LCY_STAGE2_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_LCY_STAGE2_PRE_SHOCK", "derived.OUTST_LCY_STAGE2_PRE_SHOCK", "ECL output", "EO.outst_lcy for EO.stage_id=2", ""
    SeedLink d, "OUTST_LCY_NPA_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_LCY_NPA_PRE_SHOCK", "derived.OUTST_LCY_NPA_PRE_SHOCK", "ECL output", "EO.outst_lcy for EO.stage_id=3", ""
    SeedLink d, "IIS_STAGE1_PRE_SHOCK", "Pre-shock", "derivedbase.IIS_STAGE1_PRE_SHOCK", "derived.IIS_STAGE1_PRE_SHOCK", "ECL output", "", ""
    SeedLink d, "IIS_STAGE2_PRE_SHOCK", "Pre-shock", "derivedbase.IIS_STAGE2_PRE_SHOCK", "derived.IIS_STAGE2_PRE_SHOCK", "ECL output", "", ""
    SeedLink d, "IIS_NPA_PRE_SHOCK", "Pre-shock", "derivedbase.IIS_NPA_PRE_SHOCK", "derived.IIS_NPA_PRE_SHOCK", "ECL output", "EO.IIS for EO.stage_id=3", ""
    SeedLink d, "OUTST_ECL_STAGE1_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_ECL_STAGE1_LCY_PRE_SHOCK", "derived.OUTST_ECL_STAGE1_LCY_PRE_SHOCK", "ECL output", "EO.outst_for_ecl for EO.stage_id=1", ""
    SeedLink d, "OUTST_ECL_STAGE2_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_ECL_STAGE2_LCY_PRE_SHOCK", "derived.OUTST_ECL_STAGE2_LCY_PRE_SHOCK", "ECL output", "EO.outst_for_ecl for EO.stage_id=2", ""
    SeedLink d, "OUTST_ECL_NPA_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_ECL_NPA_LCY_PRE_SHOCK", "derived.OUTST_ECL_NPA_LCY_PRE_SHOCK", "ECL output", "EO.outst_for_ecl for EO.stage_id=3", ""
    SeedLink d, "MAX_ECL_STRESS_ECL_STAGE1", "Pre-shock", "derivedbase.MAX_ECL_STRESS_ECL_STAGE1", "derived.MAX_ECL_STRESS_ECL_STAGE1", "ECL output with the Moody's stress ECL column", "", "Per account: the larger of Moody's stress ECL and ECL. Needs ORI_MOODY_STRESS_ECL (or MOODY_STRESS_ECL / STRESS_ECL) in the ECL file."
    SeedLink d, "MAX_ECL_STRESS_ECL_STAGE2", "Pre-shock", "derivedbase.MAX_ECL_STRESS_ECL_STAGE2", "derived.MAX_ECL_STRESS_ECL_STAGE2", "ECL output with the Moody's stress ECL column", "", "Per account: the larger of Moody's stress ECL and ECL. Needs ORI_MOODY_STRESS_ECL (or MOODY_STRESS_ECL / STRESS_ECL) in the ECL file."
    SeedLink d, "MAX_ECL_STRESS_ECL_STAGE3", "Pre-shock", "derivedbase.MAX_ECL_STRESS_ECL_STAGE3", "derived.MAX_ECL_STRESS_ECL_STAGE3", "ECL output with the Moody's stress ECL column", "", "Per account: the larger of Moody's stress ECL and ECL. Needs ORI_MOODY_STRESS_ECL (or MOODY_STRESS_ECL / STRESS_ECL) in the ECL file."
    SeedLink d, "IMPACT_ECL_LCY_MOODY_STAGE1", "Pre-shock", "derivedbase.IMPACT_ECL_LCY_MOODY_STAGE1", "derived.IMPACT_ECL_LCY_MOODY_STAGE1", "ECL output with the Moody's stress ECL column", "", "Per account: the larger of Moody's stress ECL and ECL. Needs ORI_MOODY_STRESS_ECL (or MOODY_STRESS_ECL / STRESS_ECL) in the ECL file."
    SeedLink d, "IMPACT_ECL_LCY_MOODY_STAGE2", "Pre-shock", "derivedbase.IMPACT_ECL_LCY_MOODY_STAGE2", "derived.IMPACT_ECL_LCY_MOODY_STAGE2", "ECL output with the Moody's stress ECL column", "", "Per account: the larger of Moody's stress ECL and ECL. Needs ORI_MOODY_STRESS_ECL (or MOODY_STRESS_ECL / STRESS_ECL) in the ECL file."
    SeedLink d, "IMPACT_ECL_LCY_MOODY_STAGE3", "Pre-shock", "derivedbase.IMPACT_ECL_LCY_MOODY_STAGE3", "derived.IMPACT_ECL_LCY_MOODY_STAGE3", "ECL output with the Moody's stress ECL column", "", "Per account: the larger of Moody's stress ECL and ECL. Needs ORI_MOODY_STRESS_ECL (or MOODY_STRESS_ECL / STRESS_ECL) in the ECL file."
    SeedLink d, "OUTST_RWA_STAGE1_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_RWA_STAGE1_LCY_PRE_SHOCK", "derived.OUTST_RWA_STAGE1_LCY_PRE_SHOCK", "CAPRWA output", "RCO.outst_for_ecl for RCO.stage_id=1", ""
    SeedLink d, "OUTST_RWA_STAGE2_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_RWA_STAGE2_LCY_PRE_SHOCK", "derived.OUTST_RWA_STAGE2_LCY_PRE_SHOCK", "CAPRWA output", "RCO.outst_for_ecl for RCO.stage_id=2", ""
    SeedLink d, "OUTST_RWA_NPA_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.OUTST_RWA_NPA_LCY_PRE_SHOCK", "derived.OUTST_RWA_NPA_LCY_PRE_SHOCK", "CAPRWA output", "RCO.outst_for_ecl for RCO.stage_id=3", ""
    SeedLink d, "ECL_STAGE1_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.ECL_STAGE1_LCY_PRE_SHOCK", "derived.ECL_STAGE1_LCY_PRE_SHOCK", "ECL output", "EO.ecl for EO.stage_id =1", ""
    SeedLink d, "ECL_STAGE2_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.ECL_STAGE2_LCY_PRE_SHOCK", "derived.ECL_STAGE2_LCY_PRE_SHOCK", "ECL output", "EO.ecl for EO.stage_id =2", ""
    SeedLink d, "ECL_NPA_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.ECL_NPA_LCY_PRE_SHOCK", "derived.ECL_NPA_LCY_PRE_SHOCK", "ECL output", "EO.ecl for EO.stage_id =3", ""
    SeedLink d, "RWA_CR_STAGE1_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.RWA_CR_STAGE1_LCY_PRE_SHOCK", "derived.RWA_CR_STAGE1_LCY_PRE_SHOCK", "CAPRWA output", "RWA.RWA_LCY for stage_id=1", ""
    SeedLink d, "RWA_CR_STAGE2_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.RWA_CR_STAGE2_LCY_PRE_SHOCK", "derived.RWA_CR_STAGE2_LCY_PRE_SHOCK", "CAPRWA output", "RWA.RWA_LCY for stage_id=2", ""
    SeedLink d, "RWA_CR_NPA_LCY_PRE_SHOCK", "Pre-shock", "derivedbase.RWA_CR_NPA_LCY_PRE_SHOCK", "derived.RWA_CR_NPA_LCY_PRE_SHOCK", "CAPRWA output", "RWA.RWA_LCY for stage_id=3", ""
    SeedLink d, "ECL_RATIO_STAGE1", "Ratio", "derivedbase.ECL_RATIO_STAGE1", "derived.ECL_RATIO_STAGE1", "ECL output", "ECL_STAGE1_LCY_PRE_SHOCK / OUTST_ECL_STAGE1_LCY", ""
    SeedLink d, "ECL_RATIO_STAGE2", "Ratio", "derivedbase.ECL_RATIO_STAGE2", "derived.ECL_RATIO_STAGE2", "ECL output", "ECL_STAGE2_LCY_PRE_SHOCK / OUTST_ECL_STAGE2_LCY", ""
    SeedLink d, "ECL_RATIO_NPA", "Ratio", "derivedbase.ECL_RATIO_NPA", "derived.ECL_RATIO_NPA", "ECL output", "ECL_STAGE3_LCY / OUTST_ECL_NPA_LCY", ""
    SeedLink d, "ECL_RATIO", "Ratio", "derivedbase.ECL_RATIO", "derived.ECL_RATIO", "ECL output", "ECL_LCY/OUTST_FOR_ECL", ""
    SeedLink d, "RWA_CR_RATIO_STAGE1", "Ratio", "derivedbase.RWA_CR_STAGE1_LCY_PRE_SHOCK/(derivedbase.OUTST_RWA_STAGE1_LCY_PRE_SHOCK-derivedbase.ECL_STAGE1_LCY_PRE_SHOCK)", "derived.RWA_CR_STAGE1_LCY_PRE_SHOCK/(derived.OUTST_RWA_STAGE1_LCY_PRE_SHOCK-derived.ECL_STAGE1_LCY_PRE_SHOCK)", "CAPRWA output, ECL output", "RWA_CR_STAGE1_LCY/ OUTST_RWA_STAGE1_LCY", "RWA over exposure net of ECL; confirm against the system definition."
    SeedLink d, "RWA_CR_RATIO_STAGE2", "Ratio", "derivedbase.RWA_CR_STAGE2_LCY_PRE_SHOCK/(derivedbase.OUTST_RWA_STAGE2_LCY_PRE_SHOCK-derivedbase.ECL_STAGE2_LCY_PRE_SHOCK)", "derived.RWA_CR_STAGE2_LCY_PRE_SHOCK/(derived.OUTST_RWA_STAGE2_LCY_PRE_SHOCK-derived.ECL_STAGE2_LCY_PRE_SHOCK)", "CAPRWA output, ECL output", "RWA_CR_STAGE2_LCY/ (OUTST_RWA_STAGE2_LCY - ECL_STAGE2_LCY)", "RWA over exposure net of ECL; confirm against the system definition."
    SeedLink d, "RWA_CR_RATIO_NPA", "Ratio", "derivedbase.RWA_CR_NPA_LCY_PRE_SHOCK/(derivedbase.OUTST_RWA_NPA_LCY_PRE_SHOCK-derivedbase.ECL_NPA_LCY_PRE_SHOCK)", "derived.RWA_CR_NPA_LCY_PRE_SHOCK/(derived.OUTST_RWA_NPA_LCY_PRE_SHOCK-derived.ECL_NPA_LCY_PRE_SHOCK)", "CAPRWA output, ECL output", "RWA_CR_NPA_LCY/ (OUTST_RWA_NPA_LCY - ECL_NPA_LCY)", "RWA over exposure net of ECL; confirm against the system definition."
    SeedLink d, "IIS_RATIO_STAGE1", "Ratio", "derivedbase.IIS_RATIO_STAGE1", "derived.IIS_RATIO_STAGE1", "ECL output", "", ""
    SeedLink d, "IIS_RATIO_STAGE2", "Ratio", "derivedbase.IIS_RATIO_STAGE2", "derived.IIS_RATIO_STAGE2", "ECL output", "", ""
    SeedLink d, "IIS_RATIO_NPA", "Ratio", "derivedbase.IIS_RATIO_NPA", "derived.IIS_RATIO_NPA", "ECL output", "=IIS_NPA_PRE_SHOCK / OUTST_LCY_NPA_PRE_SHOCK", ""
    SeedLink d, "RSA_LCY", "Base", "figure.RSA", "figure.RSA", "GL / balance sheet (rate sensitive assets)", "", "Enter RSA under Manual figures until a GL extract is wired in."
End Sub

Private Sub SeedLinks2(ByVal d As Collection)
    SeedLink d, "RSL_LCY", "Base", "figure.RSL", "figure.RSL", "GL / balance sheet (rate sensitive liabilities)", "", "Enter RSL under Manual figures until a GL extract is wired in."
    SeedLink d, "NET_RSA_GAP_LCY", "Ratio", "figure.RSA-figure.RSL", "figure.RSA-figure.RSL", "GL / balance sheet", "", ""
    SeedLink d, "IMPACT_OUTST_LCY_STAGE1", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_LCY_STAGE2", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_LCY_NPA", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_IIS_STAGE1", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_IIS_STAGE2", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_IIS_NPA", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_ECL_STAGE1_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_ECL_STAGE2_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_ECL_NPA_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_RWA_STAGE1_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_RWA_STAGE2_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_OUTST_RWA_NPA_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "FX_TOTAL_LONG_POSITION_LCY", "Post-shock base", "figure.FX_LONG_POSITION-figure.FX_SHORT_POSITION", "", "Market risk forex total", "cap_market_risk_forex_total.total_long_position - total_short_position", "As written in the ST Design (long minus short for both rows); confirm."
    SeedLink d, "FX_TOTAL_SHORT_POSITION_LCY", "Post-shock base", "figure.FX_LONG_POSITION-figure.FX_SHORT_POSITION", "", "Market risk forex total", "cap_market_risk_forex_total.total_long_position - total_short_position", "As written in the ST Design (long minus short for both rows); confirm."
    SeedLink d, "IMPACT_NET_OPEN_POSITION", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "OUTST_LCY", "Post-shock base", "derivedbase.OUTST_LCY", "", "ECL output", "CA.reporting_balance_lcy + NCA.reporting_balance_lcy", ""
    SeedLink d, "OUTST_ECL_STAGE1_LCY", "Post-shock base", "derivedbase.OUTST_ECL_STAGE1_LCY_PRE_SHOCK", "", "ECL output", "EO.outst_for_ecl for EO.stage_id=1", ""
    SeedLink d, "OUTST_ECL_STAGE2_LCY", "Post-shock base", "derivedbase.OUTST_ECL_STAGE2_LCY_PRE_SHOCK", "", "ECL output", "EO.outst_for_ecl for EO.stage_id=2", ""
    SeedLink d, "OUTST_ECL_NPA_LCY", "Post-shock base", "derivedbase.OUTST_ECL_NPA_LCY_PRE_SHOCK", "", "ECL output", "EO.outst_for_ecl for EO.stage_id=3", ""
    SeedLink d, "OUTST_FOR_ECL", "Post-shock base", "derivedbase.OUTST_FOR_ECL", "", "ECL output", "EO.outst_for_ecl", ""
    SeedLink d, "OUTST_RWA_STAGE1_LCY", "Post-shock base", "derivedbase.OUTST_RWA_STAGE1_LCY_PRE_SHOCK", "", "CAPRWA output", "RCO.outst_for_ecl for RCO.stage_id=1", ""
    SeedLink d, "OUTST_RWA_STAGE2_LCY", "Post-shock base", "derivedbase.OUTST_RWA_STAGE2_LCY_PRE_SHOCK", "", "CAPRWA output", "RCO.outst_for_ecl for RCO.stage_id=2", ""
    SeedLink d, "OUTST_RWA_NPA_LCY", "Post-shock base", "derivedbase.OUTST_RWA_NPA_LCY_PRE_SHOCK", "", "CAPRWA output", "RCO.outst_for_ecl for RCO.stage_id=3", ""
    SeedLink d, "OUTST_FOR_RWA", "Post-shock base", "derivedbase.OUTST_FOR_RWA", "", "CAPRWA output", "RCO.outst_for_ecl", ""
    SeedLink d, "IMPACT_ECL_STAGE1_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_ECL_STAGE2_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_ECL_NPA_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_ECL_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "ECL_STAGE1_LCY", "Post-shock base", "derivedbase.ECL_STAGE1_LCY_PRE_SHOCK", "", "ECL output", "EO.ecl for EO.stage_id=1", ""
    SeedLink d, "ECL_STAGE2_LCY", "Post-shock base", "derivedbase.ECL_STAGE2_LCY_PRE_SHOCK", "", "ECL output", "EO.ecl for EO.stage_id =2", ""
    SeedLink d, "ECL_NPA_LCY", "Post-shock base", "derivedbase.ECL_NPA_LCY_PRE_SHOCK", "", "ECL output", "EO.ecl for EO.stage_id =3", ""
    SeedLink d, "ECL_LCY", "Post-shock base", "derivedbase.ECL_STAGE1_LCY_PRE_SHOCK+derivedbase.ECL_STAGE2_LCY_PRE_SHOCK+derivedbase.ECL_NPA_LCY_PRE_SHOCK", "", "ECL output", "EO.ecl", ""
    SeedLink d, "IMPACT_CR_RWA_STAGE1_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_CR_RWA_STAGE2_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_CR_RWA_NPA_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_RWA_CR_LCY", "Shock", "", "", "Config formula (shock derivation)", "0", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "RWA_CR_STAGE1_LCY", "Post-shock base", "derivedbase.RWA_CR_STAGE1_LCY_PRE_SHOCK", "", "CAPRWA output", "RWA.RWA_LCY for stage_id=1", ""
    SeedLink d, "RWA_CR_STAGE2_LCY", "Post-shock base", "derivedbase.RWA_CR_STAGE2_LCY_PRE_SHOCK", "", "CAPRWA output", "RWA.RWA_LCY for stage_id=2", ""
End Sub

Private Sub SeedLinks3(ByVal d As Collection)
    SeedLink d, "RWA_CR_NPA_LCY", "Post-shock base", "derivedbase.RWA_CR_NPA_LCY_PRE_SHOCK", "", "CAPRWA output", "RWA.RWA_LCY for stage_id=3", ""
    SeedLink d, "RWA_CR_LCY", "Post-shock base", "derivedbase.RWA_CR_BASE", "", "Capital component", "CAP COMPONENT = NET_RWA_CREDIT_RISK", ""
    SeedLink d, "IMPACT_RWA_MR_EQUITY_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "RWA_MR_EQUITY_LCY", "Post-shock base", "figure.RWA_MR_EQUITY", "", "Market risk RWA", "", "Uses Manual figures until the extract is wired in as a source."
    SeedLink d, "IMPACT_RWA_MR_FOREX_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "RWA_MR_FOREX_LCY", "Post-shock base", "figure.RWA_MR_FOREX", "", "Market risk RWA", "", "Uses Manual figures until the extract is wired in as a source."
    SeedLink d, "IMPACT_RWA_OR_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "RWA_OR_LCY", "Post-shock base", "derivedbase.RWA_OR_BASE", "", "Capital component", "from cap component table", "RWA_OR_BASE uses component code NET_RWA_OPERATIONAL_RISK: confirm the code."
    SeedLink d, "IMPACT_TOTAL_RWA", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "TOTAL_RWA", "Post-shock base", "derivedbase.RWA_CR_BASE+figure.RWA_MR_EQUITY+figure.RWA_MR_FOREX+derivedbase.RWA_OR_BASE", "", "Capital component, Market risk RWA", "cr + market equity + market forex + or rwa", "Uses Manual figures until the extract is wired in as a source. RWA_OR_BASE uses component code NET_RWA_OPERATIONAL_RISK: confirm the code."
    SeedLink d, "MARKET_VALUE_LCY", "Post-shock base", "figure.MARKET_VALUE_FVTPL", "", "Market risk equity (FVTPL)", "", "Uses Manual figures until the extract is wired in as a source."
    SeedLink d, "IMPACT_MARKET_VALUE_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "PROFITS_BEFORE_TAX_LCY", "Post-shock base", "figure.PBT", "", "PBT/PAT table (aod_pbt_pat_data)", "nca.outst for balance sheet type INCOME - nca.outst for balance sheet type expense - take pbt from new config table stresstest.aod_pbt_pat_data", "Uses Manual figures until the extract is wired in as a source."
    SeedLink d, "TAX_RATE", "Post-shock base", "1-figure.PAT/figure.PBT", "", "PBT/PAT table (aod_pbt_pat_data)", "1- ( pat/pbt) from stresstest.aod_pbt_pat_data", "Uses Manual figures until the extract is wired in as a source."
    SeedLink d, "PROFITS_AFTER_TAX_LCY", "Post-shock base", "figure.PAT", "", "PBT/PAT table (aod_pbt_pat_data)", "profits_before_tax_lcy * (1-tax_rate) - take pat from new config table stresstest.aod_pbt_pat_data", "Uses Manual figures until the extract is wired in as a source."
    SeedLink d, "IMPACT_PROFITS_BEFORE_TAX_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_PROFITS_BEFORE_TAX_PCT", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_PROFITS_AFTER_TAX_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_PROFITS_AFTER_TAX_PCT", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "CET1_CAPITAL_LCY", "Post-shock base", "derivedbase.CET1_BASE", "", "Capital component", "from cap component table", ""
    SeedLink d, "AT1_CAPITAL_LCY", "Post-shock base", "derivedbase.AT1_BASE", "", "Capital component", "from cap component table", ""
    SeedLink d, "T2_CAPITAL_LCY", "Post-shock base", "derivedbase.T2_BASE", "", "Capital component", "from cap component table", ""
    SeedLink d, "TOTAL_CAPITAL_LCY", "Post-shock base", "derivedbase.TOTAL_CAPITAL_BASE", "", "Capital component", "from cap component table", ""
    SeedLink d, "REGULATORY_CAR", "Post-shock base", "derivedbase.TOTAL_CAPITAL_BASE/(derivedbase.RWA_CR_BASE+figure.RWA_MR_EQUITY+figure.RWA_MR_FOREX+derivedbase.RWA_OR_BASE)", "", "Capital component, Market risk RWA", "from cap component table", "Uses Manual figures until the extract is wired in as a source. RWA_OR_BASE uses component code NET_RWA_OPERATIONAL_RISK: confirm the code. Check units: the system may hold CAR as a percentage."
    SeedLink d, "CET1_CAR", "Post-shock base", "derivedbase.CET1_BASE/(derivedbase.RWA_CR_BASE+figure.RWA_MR_EQUITY+figure.RWA_MR_FOREX+derivedbase.RWA_OR_BASE)", "", "Capital component, Market risk RWA", "from cap component table", "Uses Manual figures until the extract is wired in as a source. RWA_OR_BASE uses component code NET_RWA_OPERATIONAL_RISK: confirm the code. Check units: the system may hold CAR as a percentage."
    SeedLink d, "IMPACT_CET1_CAPITAL_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_AT1_CAPITAL_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_T2_CAPITAL_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_TOTAL_CAPITAL_LCY", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_NPL_LCY", "Shock", "", "", "Config formula (shock derivation)", "npl for this test element - npl for base test element", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_NPL_PCT", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_REGULATORY_CAR", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_CET1_CAR", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "PRE_LEG_LIQ_TOTAL_ASSETS", "Pre-shock", "derivedbase.LL_ASSETS_PRE_SHOCK", "derived.LL_ASSETS_PRE_SHOCK", "Legal liquidity ALM output", "", "Pre-shock follows the ST Design: the SEC accounts' cashflows, HQLA and assets at the pre-factor amount, outflows and liabilities at the post-factor amount."
    SeedLink d, "IMPACT_LEG_LIQ_TOTAL_ASSETS", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "LEG_LIQ_TOTAL_ASSETS", "Post-shock base", "derivedbase.LL_ASSETS_PRE_SHOCK", "", "Legal liquidity ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id =LL (5)) and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_", ""
    SeedLink d, "PRE_LEG_LIQ_TOTAL_LIAB", "Pre-shock", "derivedbase.LL_LIABILITIES_PRE_SHOCK", "derived.LL_LIABILITIES_PRE_SHOCK", "Legal liquidity ALM output", "", "Pre-shock follows the ST Design: the SEC accounts' cashflows, HQLA and assets at the pre-factor amount, outflows and liabilities at the post-factor amount."
    SeedLink d, "IMPACT_LEG_LIQ_TOTAL_LIAB", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "LEG_LIQ_TOTAL_LIAB", "Post-shock base", "derivedbase.LL_LIABILITIES_PRE_SHOCK", "", "Legal liquidity ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id =LL (5)) and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_", ""
    SeedLink d, "IMPACT_LEGAL_LIQUIDITY_RATIO", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
End Sub

Private Sub SeedLinks4(ByVal d As Collection)
    SeedLink d, "LEGAL_LIQUIDITY_RATIO", "Post-shock base", "derivedbase.LL_ASSETS_PRE_SHOCK/derivedbase.LL_LIABILITIES_PRE_SHOCK", "", "Legal liquidity ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id =LL (5)) and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_", ""
    SeedLink d, "PRE_LCR_HQLA", "Pre-shock", "derivedbase.LCR_HQLA_PRE_SHOCK", "derived.LCR_HQLA_PRE_SHOCK", "LCR ALM output", "", "Pre-shock follows the ST Design: the SEC accounts' cashflows, HQLA and assets at the pre-factor amount, outflows and liabilities at the post-factor amount."
    SeedLink d, "IMPACT_LCR_HQLA", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "LCR_HQLA", "Post-shock base", "derivedbase.LCR_HQLA_PRE_SHOCK", "", "LCR ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id = id for ""LCR"" in alm_framework_categories) and category_id in (select category_id from alm.ALM_PORTFOLIO_SE", ""
    SeedLink d, "PRE_LCR_OUTFLOW", "Pre-shock", "derivedbase.LCR_OUTFLOW_PRE_SHOCK", "derived.LCR_OUTFLOW_PRE_SHOCK", "LCR ALM output", "", "Pre-shock follows the ST Design: the SEC accounts' cashflows, HQLA and assets at the pre-factor amount, outflows and liabilities at the post-factor amount."
    SeedLink d, "IMPACT_LCR_OUTFLOW", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "LCR_OUTFLOW", "Post-shock base", "derivedbase.LCR_OUTFLOW_PRE_SHOCK", "", "LCR ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id = LCR) and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_fr", ""
    SeedLink d, "PRE_LCR_INFLOW", "Pre-shock", "derivedbase.LCR_INFLOW_PRE_SHOCK", "derived.LCR_INFLOW_PRE_SHOCK", "LCR ALM output", "", "Pre-shock follows the ST Design: the SEC accounts' cashflows, HQLA and assets at the pre-factor amount, outflows and liabilities at the post-factor amount."
    SeedLink d, "IMPACT_LCR_INFLOW", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "LCR_INFLOW", "Post-shock base", "derivedbase.LCR_INFLOW_PRE_SHOCK", "", "LCR ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id = LCR) and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_fr", ""
    SeedLink d, "IMPACT_LCR_NET_OUTFLOW", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "IMPACT_LCR", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "LCR", "Post-shock base", "derivedbase.LCR_HQLA_PRE_SHOCK/(derivedbase.LCR_OUTFLOW_PRE_SHOCK-MIN(derivedbase.LCR_INFLOW_PRE_SHOCK,0.75*derivedbase.LCR_OUTFLOW_PRE_SHOCK))", "", "LCR ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id = LCR) and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_fr", "ST Design LCR formula with the 75% inflow cap."
    SeedLink d, "PRE_NSFR_ASF", "Pre-shock", "derivedbase.NSFR_ASF_PRE_SHOCK", "derived.NSFR_ASF_PRE_SHOCK", "NSFR ALM output", "", "Category codes ASF / RSF come from the ST Design; confirm them against the NSFR file."
    SeedLink d, "NSFR_IMPACT_ASF", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "NSFR_TOTAL_ASF", "Post-shock base", "derivedbase.NSFR_ASF_PRE_SHOCK", "", "NSFR ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id =NSFR and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_fra", ""
    SeedLink d, "PRE_NSFR_RSF", "Pre-shock", "derivedbase.NSFR_RSF_PRE_SHOCK", "derived.NSFR_RSF_PRE_SHOCK", "NSFR ALM output", "", "Category codes ASF / RSF come from the ST Design; confirm them against the NSFR file."
    SeedLink d, "NSFR_IMPACT_RSF", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "NSFR_TOTAL_RSF", "Post-shock base", "derivedbase.NSFR_RSF_PRE_SHOCK", "", "NSFR ALM output", "alm.ALM_FRAMEWORK_CF_CATEGORYWISE_SUMMARY.cashflow_element_lcy_post_factor where AOD= and bank= and alm_framework_id = (select framework_id from alm_frameworks where alm_frameowrk_category_id =NSFR and category_id in (select category_id from alm.ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY where alm_fra", ""
    SeedLink d, "IMPACT_NSFR", "Shock", "", "", "Config formula (shock derivation)", "", "Calculated by the manual formula from the shock parameters; no input file value."
    SeedLink d, "NSFR", "Post-shock base", "derivedbase.NSFR_ASF_PRE_SHOCK/derivedbase.NSFR_RSF_PRE_SHOCK", "", "NSFR ALM output", "", ""
End Sub

' ------------------------------------------------------ figures from files ----
'
' A figure can now name the file it comes from (FILE, SHEET, SUM_COLUMN, WHERE),
' so the inputs that are not full sources yet - the forex and equity market-risk
' tables, the PBT/PAT table, RSA/RSL - are read from the same backend exports as
' everything else instead of being typed. "Refresh figures from files" sums the
' column over the rows that match, and records how many rows it used and when.
' A figure that cannot be read is cleared, never left at an old value, so a stale
' number can never pass a reconciliation.

' The columns the ST Design names, filled only where the cell is still blank.
Private Sub SeedFigureColumns(ByVal ws As Worksheet)
    Dim fh As Object, r As Long, lastRow As Long, nm As String, col As String
    Set fh = HeaderMap(ws, VS_FIG_FIRST_COL, VS_FIG_COLS)
    If Not fh.Exists("FIGURE") Or Not fh.Exists("SUM_COLUMN") Then Exit Sub
    lastRow = LastDataRow(ws, fh("FIGURE"))
    For r = VS_FIRST_ROW To lastRow
        nm = SafeUpperText(ws.Cells(r, fh("FIGURE")).Value2)
        col = ""
        Select Case nm
            Case "FX_LONG_POSITION": col = "TOTAL_LONG_POSITION"
            Case "FX_SHORT_POSITION": col = "TOTAL_SHORT_POSITION"
            Case "MARKET_VALUE_FVTPL": col = "MARKET_VALUE_LCY"
            Case "PBT": col = "PBT"
            Case "PAT": col = "PAT"
        End Select
        If Len(col) > 0 And Len(SafeText(ws.Cells(r, fh("SUM_COLUMN")).Value2)) = 0 Then ws.Cells(r, fh("SUM_COLUMN")).Value2 = col
    Next r
End Sub

Public Sub JKB_RefreshFiguresFromFiles()
    Dim ws As Worksheet, fh As Object, r As Long, lastRow As Long, path As String, k As Variant
    Dim st As Object, books As Object, n As Long, bad As Long, msg As String, v As Double, rowsUsed As Long, er As String
    If JKB_Busy Then Exit Sub
    Set ws = GetWorksheetSafe(ThisWorkbook, VS_SHEET)
    If ws Is Nothing Then
        UiNotice "Figures from files", "There is no " & VS_SHEET & " sheet yet.", , "Click Value sources under Base and pre-shock on the home screen. It sets the sheet up."
        Exit Sub
    End If
    Set fh = HeaderMap(ws, VS_FIG_FIRST_COL, VS_FIG_COLS)
    For Each k In Array("FIGURE", "VALUE", "AS_OF_DATE", "ENTITY_CODE", "FILE", "SHEET", "SUM_COLUMN", "WHERE", "LOADED")
        If Not fh.Exists(CStr(k)) Then
            UiNotice "Figures from files", "The Manual figures table has no " & k & " column yet.", , "Run JKB_ApplyValueSourceSetup once (Alt+F8) to add it."
            Exit Sub
        End If
    Next k

    Set st = CaptureState(): JKB_Busy = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set books = NewMap()
    On Error GoTo Failed
    lastRow = LastDataRow(ws, fh("FIGURE"))
    For r = VS_FIRST_ROW To lastRow
        path = SafeText(ws.Cells(r, fh("FILE")).Value2)
        If Len(SafeText(ws.Cells(r, fh("FIGURE")).Value2)) > 0 And Len(path) > 0 Then
            msg = FigureFromFile(books, path, SafeText(ws.Cells(r, fh("SHEET")).Value2), _
                                 SafeText(ws.Cells(r, fh("SUM_COLUMN")).Value2), SafeText(ws.Cells(r, fh("WHERE")).Value2), _
                                 DateKey(ws.Cells(r, fh("AS_OF_DATE")).Value2), SafeUpperText(ws.Cells(r, fh("ENTITY_CODE")).Value2), _
                                 v, rowsUsed)
            If Len(msg) = 0 Then
                ws.Cells(r, fh("VALUE")).Value2 = v
                ws.Cells(r, fh("LOADED")).Value2 = format$(rowsUsed, "#,##0") & " row(s) from " & FileLeafOf(path) & ", " & format$(Now, "dd-mmm hh:mm")
                ws.Cells(r, fh("LOADED")).Interior.Color = UI_OK_BG
                n = n + 1
            Else
                ws.Cells(r, fh("VALUE")).ClearContents
                ws.Cells(r, fh("LOADED")).Value2 = "Not read: " & msg
                ws.Cells(r, fh("LOADED")).Interior.Color = UI_BAD_BG
                bad = bad + 1
            End If
        End If
    Next r
Done:
    On Error Resume Next
    For Each k In books.keys
        books(k).Close SaveChanges:=False
    Next k
    RestoreState st: JKB_Busy = False
    VS_ResetCache
    On Error GoTo 0
    If Len(er) > 0 Then
        UiProblem "Figures from files", "Reading the figure files stopped.", er
    ElseIf n + bad = 0 Then
        UiNotice "Figures from files", "No figure names a file yet.", , "Fill FILE and SUM_COLUMN (and WHERE if needed) on the Manual figures table."
    Else
        UiNotice "Figures from files", n & " figure(s) read from their files.", _
                 IIf(bad > 0, bad & " could not be read and were cleared. The LOADED column says why.", ""), , (bad > 0)
    End If
    Exit Sub
Failed:
    er = Err.description
    Resume Done
End Sub

Private Function ResolveFigurePath(ByVal p As String) As String
    p = Trim$(p)
    If Left$(p, 2) = "\\" Or Mid$(p, 2, 1) = ":" Or Left$(p, 1) = "/" Then
        ResolveFigurePath = p
    Else
        ResolveFigurePath = JoinPath(ThisWorkbook.path, p)
    End If
End Function

' Sums one column of one sheet over the rows that match. Returns "" on success,
' otherwise why it could not.
Private Function FigureFromFile(ByVal books As Object, ByVal path As String, ByVal sheetName As String, _
                                ByVal sumCol As String, ByVal whereText As String, ByVal wantDate As String, _
                                ByVal wantEntity As String, ByRef valueOut As Double, ByRef rowsOut As Long) As String
    Dim full As String, wb As Workbook, sh As Worksheet, a As Variant, r As Long, c As Long, hr As Long
    Dim h As Object, sc As Long, dc As Long, ec As Long, terms As Collection, t As Variant, ok As Boolean
    Dim dates As Object, dk As String, x As Variant, why As String, cellText As String, k As Variant

    valueOut = 0: rowsOut = 0
    If Len(sumCol) = 0 Then FigureFromFile = "SUM_COLUMN is blank": Exit Function
    full = ResolveFigurePath(path)
    If Len(Dir$(full)) = 0 Then FigureFromFile = "file not found (" & full & ")": Exit Function
    If books.Exists(UCase$(full)) Then
        Set wb = books(UCase$(full))
    Else
        On Error Resume Next
        Set wb = Workbooks.Open(full, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
        On Error GoTo 0
        If wb Is Nothing Then FigureFromFile = "the file could not be opened": Exit Function
        Set books(UCase$(full)) = wb
    End If
    If Len(sheetName) > 0 Then Set sh = GetWorksheetSafe(wb, sheetName) Else Set sh = wb.Worksheets(1)
    If sh Is Nothing Then FigureFromFile = "sheet '" & sheetName & "' not found": Exit Function
    a = sh.UsedRange.Value2
    If Not IsArray(a) Then FigureFromFile = "the sheet is empty": Exit Function

    ' The header row is the first of the top 30 rows that holds the column to sum.
    For r = 1 To IIf(UBound(a, 1) < 30, UBound(a, 1), 30)
        For c = 1 To UBound(a, 2)
            If NormalHeader(a(r, c)) = NormalHeader(sumCol) Then hr = r: Exit For
        Next c
        If hr > 0 Then Exit For
    Next r
    If hr = 0 Then FigureFromFile = "column " & sumCol & " not found in the first 30 rows": Exit Function
    Set h = NewMap()
    For c = 1 To UBound(a, 2)
        If Len(NormalHeader(a(hr, c))) > 0 Then If Not h.Exists(NormalHeader(a(hr, c))) Then h(NormalHeader(a(hr, c))) = c
    Next c
    sc = h(NormalHeader(sumCol))
    For Each k In Array("AS_OF_DATE", "AOD", "REPORT_DATE", "REPORTING_DATE")
        If dc = 0 Then If h.Exists(CStr(k)) Then dc = h(CStr(k))
    Next k
    For Each k In Array("ENTITY_CODE", "BANK_CODE")
        If ec = 0 Then If h.Exists(CStr(k)) Then ec = h(CStr(k))
    Next k

    Set terms = ParseFigureWhere(whereText, h, why)
    If Len(why) > 0 Then FigureFromFile = why: Exit Function

    Set dates = NewMap()
    For r = hr + 1 To UBound(a, 1)
        ok = True
        If dc > 0 Then
            dk = DateKey(a(r, dc))
            If Len(wantDate) > 0 Then
                ok = (dk = wantDate)
            ElseIf Len(dk) > 0 Then
                dates(dk) = True
            End If
        End If
        If ok And ec > 0 And Len(wantEntity) > 0 Then ok = (SafeUpperText(a(r, ec)) = wantEntity)
        If ok Then
            For Each t In terms
                cellText = SafeUpperText(a(r, CLng(t(0))))
                If t(1) = "IN" Then
                    If Not t(2).Exists(cellText) Then ok = False
                Else
                    If t(2).Exists(cellText) Then ok = False
                End If
                If Not ok Then Exit For
            Next t
        End If
        If ok Then
            x = a(r, sc)
            If Not IsError(x) Then
                If Not IsEmpty(x) Then
                    If IsNumeric(x) Then valueOut = valueOut + CDbl(x): rowsOut = rowsOut + 1
                End If
            End If
        End If
    Next r
    If dates.count > 1 Then
        FigureFromFile = "the file holds " & dates.count & " as-of dates; set AS_OF_DATE on this row"
        valueOut = 0: rowsOut = 0
    ElseIf rowsOut = 0 Then
        FigureFromFile = "no row matched" & IIf(Len(wantDate) > 0, " for " & wantDate, "")
    End If
End Function

' WHERE in the same shape the metric rules use:
'     CLASSIFICATION = 'FVTPL' AND CURRENCY IN ('USD', 'EUR') AND STATUS <> 'CLOSED'
' Returns Array(column, "IN" | "NOTIN", set of upper-cased values) per term.
Private Function ParseFigureWhere(ByVal whereText As String, ByVal h As Object, ByRef why As String) As Collection
    Dim out As Collection, part As Variant, m As Object, col As String, op As String, vals As Object, raw As String, v As Variant
    Set out = New Collection
    Set ParseFigureWhere = out
    If Len(Trim$(whereText)) = 0 Then Exit Function
    For Each part In Split(Replace(Replace(whereText, " and ", " AND "), " And ", " AND "), " AND ")
        Set m = RegexExecute("^\s*([A-Za-z0-9_ ]+?)\s*(<>|!=|=|\bNOT\s+IN\b|\bIN\b)\s*(.+?)\s*$", CStr(part))
        If m.count = 0 Then why = "WHERE not understood: " & Trim$(CStr(part)): Exit Function
        col = NormalHeader(m(0).SubMatches(0))
        op = UCase$(Replace(m(0).SubMatches(1), " ", ""))
        If Not h.Exists(col) Then why = "WHERE column " & col & " not found": Exit Function
        raw = Trim$(m(0).SubMatches(2))
        If Left$(raw, 1) = "(" And Right$(raw, 1) = ")" Then raw = Mid$(raw, 2, Len(raw) - 2)
        Set vals = NewMap()
        For Each v In Split(raw, ",")
            v = Trim$(CStr(v))
            If Len(v) >= 2 And Left$(v, 1) = "'" And Right$(v, 1) = "'" Then v = Mid$(v, 2, Len(v) - 2)
            vals(UCase$(Trim$(CStr(v)))) = True
        Next v
        If op = "=" Or op = "IN" Then op = "IN" Else op = "NOTIN"
        out.Add Array(h(col), op, vals)
    Next part
End Function

' ------------------------------------------------ reconciliation columns ----

' Derived_Values in each output book: the difference between the input-file
' value and the system value, for base and pre-shock, and whether they agree
' within 0.01% of the system value. The element sheets already show the shock
' and post-shock differences; this closes the chain at its first two steps.
Public Sub VS_WriteDifferences(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal n As Long)
    Dim tolB As String, tolP As String
    If n <= 0 Then Exit Sub
    On Error Resume Next
    ws.Cells(DV_HEADER_ROW, 15).Resize(1, 3).Value2 = Array("Base difference (input files - system)", _
        "Pre-shock difference (input files - system)", "Reconciles")
    ' IF, not OR: Excel evaluates every argument of OR, and the system side of a
    ' row with no input-file value may be #N/A.
    tolB = "IF(RC15="""",TRUE,ABS(RC15)<=MAX(ABS(RC12)*0.0001,0.000001))"
    tolP = "IF(RC16="""",TRUE,ABS(RC16)<=MAX(ABS(RC14)*0.0001,0.000001))"
    ws.Range(ws.Cells(DV_FIRST_ROW, 15), ws.Cells(lastRow, 15)).FormulaR1C1 = "=IF(AND(ISNUMBER(RC11),ISNUMBER(RC12)),RC11-RC12,"""")"
    ws.Range(ws.Cells(DV_FIRST_ROW, 16), ws.Cells(lastRow, 16)).FormulaR1C1 = "=IF(AND(ISNUMBER(RC13),ISNUMBER(RC14)),RC13-RC14,"""")"
    ws.Range(ws.Cells(DV_FIRST_ROW, 17), ws.Cells(lastRow, 17)).FormulaR1C1 = _
        "=IF(AND(RC15="""",RC16=""""),""Not comparable"",IF(AND(" & tolB & "," & tolP & "),""Matches"",""Differs""))"
    With ws.Range(ws.Cells(DV_HEADER_ROW, 15), ws.Cells(DV_HEADER_ROW, 17))
        .Font.Bold = True: .Font.Color = UI_WHITE: .Interior.Color = UI_NAVY
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With
    ws.Range(ws.Cells(DV_FIRST_ROW, 15), ws.Cells(lastRow, 16)).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range(ws.Cells(DV_HEADER_ROW, 15), ws.Cells(lastRow, 17)).Borders.Color = UI_LINE
    With ws.Range(ws.Cells(DV_FIRST_ROW, 17), ws.Cells(lastRow, 17)).FormatConditions
        .Delete
        .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Differs""").Interior.Color = UI_BAD_BG
        .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Matches""").Interior.Color = UI_OK_BG
    End With
    ws.columns("O:P").ColumnWidth = 18
    ws.columns("Q").ColumnWidth = 15
    If ws.Cells(DV_HEADER_ROW, 17).Comment Is Nothing Then
        ws.Cells(DV_HEADER_ROW, 17).AddComment "Matches: within 0.01% of the system value. Not comparable: one side is missing."
    End If
    Err.Clear
End Sub

' ------------------------------------------------------- one switch (part 3) ----
' Value sources is the only place that says "input files or system". A metric takes
' the switch of the Value sources row named after it, else of the first row whose
' link is exactly derived.<metric> (pre-shock) or derivedbase.<metric> (base), else
' the default at C6 (base) or C7 (pre-shock).
Public Function VS_MetricUse(ByVal metric As String, ByVal phase As String) As String
    Dim ln As Variant, k As Variant, want As String, m As String, isBase As Boolean
    EnsureLoaded
    isBase = (UCase$(phase) = "BASE")
    m = SafeUpperText(metric)
    If LinkFor(m, ln) Then
        If isBase Then VS_MetricUse = ln(1) Else VS_MetricUse = ln(3)
        Exit Function
    End If
    If isBase Then want = "DERIVEDBASE." & m Else want = "DERIVED." & m
    If Not mLinks Is Nothing Then
        For Each k In mLinks.keys
            ln = mLinks(k)
            If isBase Then
                If Squash(CStr(ln(0))) = want Then VS_MetricUse = ln(1): Exit Function
            Else
                If Squash(CStr(ln(2))) = want Then VS_MetricUse = ln(3): Exit Function
            End If
        Next k
    End If
    VS_MetricUse = VS_DefaultUse(phase)
End Function

' Moves the two old switches out of the way, once, without losing a choice someone made.
' A metric's own "Value to use" (Pre_Shock_Metrics column J) is copied into the blank
' switches of its Value sources row when it differs from the default there; a metric
' with no row of its own is written to the log instead. The old Pre_Shock default (F11)
' was filled in automatically, so it is only logged. Both are then cleared.
Private Function ConsolidateSwitches(ByVal vs As Worksheet) As Long
    Dim ws As Worksheet, h As Object, r As Long, lastRow As Long, metric As String, t As String, use As String
    Dim linkRow As Long, moved As Long, unmatched As String, oldDefault As String, c As Range
    On Error Resume Next
    VS_ResetCache
    EnsureLoaded
    Set h = HeaderMap(vs, 1, VS_LINK_COLS)
    Set ws = PsMetricsSheet()
    If Not ws Is Nothing And h.Exists("BASE_SOURCE") And h.Exists("PRE_SHOCK_SOURCE") Then
        For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
            metric = SafeUpperText(ws.Cells(r, 1).Value2)
            t = SafeText(ws.Cells(r, PS_METRIC_USE_COL).Value2)
            If Len(metric) > 0 And Len(t) > 0 And InStr(1, t, "Config_ValueSources", vbTextCompare) = 0 Then
                use = UseFromText(t, "")
                linkRow = MetricRow(metric)
                If Len(use) = 0 Then
                    ' Not one of the three choices; nothing to carry.
                ElseIf linkRow = 0 Then
                    If use <> mPreDefault Then AddText unmatched, metric & " (" & use & ")"
                Else
                    Set c = vs.Cells(linkRow, h("BASE_SOURCE"))
                    If Len(SafeText(c.Value2)) = 0 And use <> mBaseDefault Then c.Value2 = use: moved = moved + 1
                    Set c = vs.Cells(linkRow, h("PRE_SHOCK_SOURCE"))
                    If Len(SafeText(c.Value2)) = 0 And use <> mPreDefault Then c.Value2 = use: moved = moved + 1
                End If
            End If
        Next r
        ws.Range(ws.Cells(PS_METRIC_FIRST_ROW, PS_METRIC_USE_COL), ws.Cells(PS_METRIC_LAST_ROW, PS_METRIC_USE_COL)).Validation.Delete
        ws.Range(ws.Cells(PS_METRIC_FIRST_ROW, PS_METRIC_USE_COL), ws.Cells(PS_METRIC_LAST_ROW, PS_METRIC_USE_COL)).ClearContents
        ws.columns(PS_METRIC_USE_COL).Hidden = True
    End If
    Set ws = PreShockSheet()
    If Not ws Is Nothing Then
        oldDefault = SafeText(ws.Range("F11").Value2)
        If InStr(1, oldDefault, "Config_ValueSources", vbTextCompare) > 0 Then oldDefault = ""
        ws.Range("F11").Validation.Delete
        ws.Range("F11").Value2 = "Set on " & VS_SHEET
    End If
    VS_ResetCache
    If moved > 0 Then LogIssue LOG_LEVEL_INFO, "Value sources", moved & " switch(es) moved from " & PS_METRICS_SHEET & _
                               " column J to their rows on " & VS_SHEET & ".", VS_SHEET
    If Len(unmatched) > 0 Then LogIssue LOG_LEVEL_WARN, "Value sources", "These metrics had their own choice on " & _
                               PS_METRICS_SHEET & " but no row on " & VS_SHEET & ", so they now follow the default at C7: " & _
                               unmatched, VS_SHEET
    If Len(oldDefault) > 0 Then LogIssue LOG_LEVEL_INFO, "Value sources", "The old " & PRE_SHOCK_SHEET & " default was """ & _
                               oldDefault & """. The defaults now live at C6 (base) and C7 (pre-shock) on " & VS_SHEET & ".", VS_SHEET
    ConsolidateSwitches = moved
    Err.Clear
End Function

' The Value sources row a metric belongs to, by the rule VS_MetricUse follows (pre-shock side).
Private Function MetricRow(ByVal metric As String) As Long
    Dim ln As Variant, k As Variant
    If LinkFor(metric, ln) Then MetricRow = ln(4): Exit Function
    If mLinks Is Nothing Then Exit Function
    For Each k In mLinks.keys
        ln = mLinks(k)
        If Squash(CStr(ln(2))) = "DERIVED." & metric Then MetricRow = ln(4): Exit Function
    Next k
End Function

' ----------------------------------------------- one results table (part 3) ----
' Derived_Values on the master carries the comparison Pre_Shock_Results used to hold:
' base and pre-shock differences (input files minus system) and Reconciles, in
' columns Q:S after the sixteen columns the store is read back from. Same wording and
' tolerance as each output book's Derived_Values O:Q.
Public Sub VS_WriteMasterDifferences(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim a As Variant, out As Variant, i As Long, n As Long, bd As Variant, pd As Variant, ok As Boolean
    Const C0 As Long = 17
    On Error Resume Next
    ws.Cells(DV_HEADER_ROW, C0).Resize(1, 3).Value2 = Array("Base difference", "Pre-shock difference", "Reconciles")
    With ws.Cells(DV_HEADER_ROW, C0).Resize(1, 3)
        .Font.Bold = True: .Font.Color = UI_WHITE: .Interior.Color = UI_NAVY
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With
    If lastRow < DV_FIRST_ROW Then Exit Sub
    n = lastRow - DV_FIRST_ROW + 1
    a = ws.Range(ws.Cells(DV_FIRST_ROW, 1), ws.Cells(lastRow, 14)).Value2
    ReDim out(1 To n, 1 To 3)
    For i = 1 To n
        bd = VsDiff(a(i, 8), a(i, 13)): pd = VsDiff(a(i, 9), a(i, 14))
        out(i, 1) = bd: out(i, 2) = pd
        If IsEmpty(bd) And IsEmpty(pd) Then
            out(i, 3) = "Not comparable"
        Else
            ok = True
            If Not IsEmpty(bd) Then ok = ok And Abs(bd) <= VsTolerance(a(i, 13))
            If Not IsEmpty(pd) Then ok = ok And Abs(pd) <= VsTolerance(a(i, 14))
            If ok Then out(i, 3) = "Matches" Else out(i, 3) = "Differs"
        End If
    Next i
    ws.Cells(DV_FIRST_ROW, C0).Resize(n, 3).Value2 = out
    ws.Cells(DV_FIRST_ROW, C0).Resize(n, 2).NumberFormat = "#,##0.00"
    With ws.Cells(DV_FIRST_ROW, C0 + 2).Resize(n, 1)
        .FormatConditions.Delete
        .FormatConditions.Add(xlCellValue, xlEqual, "=""Differs""").Interior.Color = UI_BAD_BG
        .FormatConditions.Add(xlCellValue, xlEqual, "=""Matches""").Interior.Color = UI_OK_BG
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(DV_HEADER_ROW, C0), ws.Cells(lastRow, C0 + 2)).Borders.Color = UI_LINE
    ws.columns(C0).ColumnWidth = 17: ws.columns(C0 + 1).ColumnWidth = 17: ws.columns(C0 + 2).ColumnWidth = 15
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(DV_HEADER_ROW, 1), ws.Cells(lastRow, C0 + 2)).AutoFilter
    Err.Clear
End Sub

Private Function VsDiff(ByVal inputVal As Variant, ByVal systemVal As Variant) As Variant
    If VarType(inputVal) = vbString Or VarType(systemVal) = vbString Then Exit Function
    If IsEmpty(inputVal) Or IsEmpty(systemVal) Or IsError(inputVal) Or IsError(systemVal) Then Exit Function
    If Not IsNumeric(inputVal) Or Not IsNumeric(systemVal) Then Exit Function
    VsDiff = CDbl(inputVal) - CDbl(systemVal)
End Function

Private Function VsTolerance(ByVal systemVal As Variant) As Double
    VsTolerance = Abs(CDbl(systemVal)) * 0.0001
    If VsTolerance < 0.000001 Then VsTolerance = 0.000001
End Function

' Counts off the master store, for the status line and the review console:
' rows, rows with a base / pre-shock figure from the input files, and rows that differ.
Public Sub VS_ResultCounts(ByRef results As Long, ByRef baseOk As Long, ByRef preOk As Long, ByRef diffs As Long)
    Dim ws As Worksheet, a As Variant, r As Long, lastRow As Long
    results = 0: baseOk = 0: preOk = 0: diffs = 0
    On Error GoTo Done
    Set ws = DerivedSheet()
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    If lastRow < DV_FIRST_ROW Then Exit Sub
    a = ws.Range(ws.Cells(DV_FIRST_ROW, 1), ws.Cells(lastRow, 19)).Value2
    For r = 1 To UBound(a, 1)
        If Len(SafeText(a(r, 6))) > 0 Then
            results = results + 1
            If Not IsEmpty(a(r, 8)) And Not IsError(a(r, 8)) Then If IsNumeric(a(r, 8)) Then baseOk = baseOk + 1
            If Not IsEmpty(a(r, 9)) And Not IsError(a(r, 9)) Then If IsNumeric(a(r, 9)) Then preOk = preOk + 1
            If SafeText(a(r, 19)) = "Differs" Then diffs = diffs + 1
        End If
    Next r
Done:
End Sub
