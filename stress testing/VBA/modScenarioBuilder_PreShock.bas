Attribute VB_Name = "modScenarioBuilder_PreShock"
Option Explicit

' ============================================================================
'  ECL pre-shock testing - one sheet, one source of truth.
'
'  Everything the independent check needs lives on a single worksheet, laid out
'  as five stacked blocks:
'
'    1  Source      the uploaded ECL output, its detected sheet, and the master
'                   on/off switch for independent checking.
'    2  Formulas    one row per pre-shock metric. THIS is where a formula is
'                   stored, and it is the only place. The metric name is the same
'                   string as the OUTPUT_ROW_LABEL in Config_ElementTypeRules,
'                   which is how the config sheets refer to it; block 2 reports
'                   which element types consume each metric so the link is visible.
'    3  Dimensions  logical filter dimension -> candidate ECL column(s).
'    4  Test cases  one row per test case / test element, carrying the system
'                   filter from the upload and an editable override.
'    5  Results     per test case and metric: the system's pre-shock figure, the
'                   figure rebuilt independently from the ECL output, and the
'                   difference.
'
'  The flow the sheet implements: take the uploaded ECL output, filter it per test
'  case, apply the block-2 formula to the filtered rows to arrive at a pre-shock
'  value, compare that against the system's own pre-shock figure (block 5), and
'  hand the value to the generator so it lands in the Independent / manual column
'  of the generated element sheet.
'
'  This replaces the former Base_Testing + Pre_Shock_Config pair, which held two
'  copies of the metric table, two copies of the dimension map and two copies of
'  the filter overrides. EnsurePreShockSheet migrates any data from those sheets
'  on first run and then retires them.
' ============================================================================

' ---------------------------------------------------------------------------
'  ONE BLOCK, ONE SHEET.
'
'  This used to be five differently-shaped blocks stacked on a single worksheet:
'  test cases at row 26, metric rules at 279, dimensions at 342, results at 395.
'  It was unusable - you could not see a rule and its result at the same time,
'  every block needed its own scroll, and a layout change silently moved every
'  offset underneath the code that read them.
'
'  Each block now has its own sheet, and every one of those sheets has the SAME
'  shape: title at row 3, what-it-is-for at 4, live status at 6, header at 7,
'  data from 8. A sheet you have not opened before reads the way the last one did.
' ---------------------------------------------------------------------------
Public Const PRE_SHOCK_SHEET As String = "Pre_Shock"
Public Const PS_SOURCES_SHEET As String = "Pre_Shock_Sources"
Public Const PS_METRICS_SHEET As String = "Pre_Shock_Metrics"
Public Const PS_FIELDS_SHEET As String = "Pre_Shock_Fields"
Public Const PS_CASES_SHEET As String = "Pre_Shock_Cases"
Public Const PS_RESULTS_SHEET As String = "Pre_Shock_Results"
Private Const LEGACY_COMBINED_SHEET As String = "ECL_Pre_Shock"
Private Const LEGACY_BASE_SHEET As String = "Base_Testing"
Private Const LEGACY_CONFIG_SHEET As String = "Pre_Shock_Config"

' Bumped whenever the layout below changes, so an existing workspace is rebuilt
' rather than silently read at the wrong offsets.
Private Const PS_LAYOUT_VERSION As String = "v7"

' --- the one layout, shared by every pre-shock sheet ------------------------
Public Const PS_TITLE_ROW As Long = 3
Public Const PS_ABOUT_ROW As Long = 4
Public Const PS_STATUS_ROW As Long = 6
Public Const PS_HDR_ROW As Long = 7
Public Const PS_FIRST_ROW As Long = 8

' --- the guided control panel, on the home sheet only ----------------------
Private Const PS_SRC_FIRST_ROW As Long = 9
Private Const PS_SRC_LAST_ROW As Long = 14

' --- per-sheet block bounds -------------------------------------------------
' Each is simply "the one layout" on its own sheet; the last-row caps are
' generous and exist only so a runaway loop cannot walk a million rows.
Public Const PS_SRC_LIST_HEADER_ROW As Long = PS_HDR_ROW
Public Const PS_SRC_LIST_ROW As Long = PS_FIRST_ROW
Public Const PS_SRC_LIST_LAST_ROW As Long = PS_FIRST_ROW + 29

Public Const PS_CASE_HEADER_ROW As Long = PS_HDR_ROW
Public Const PS_CASE_FIRST_ROW As Long = PS_FIRST_ROW
Public Const PS_CASE_LAST_ROW As Long = PS_FIRST_ROW + 2999

Public Const PS_METRIC_HEADER_ROW As Long = PS_HDR_ROW
Public Const PS_METRIC_FIRST_ROW As Long = PS_FIRST_ROW
Public Const PS_METRIC_LAST_ROW As Long = PS_FIRST_ROW + 499
Public Const PS_METRIC_COLS As Long = 10
' Which number derived.<metric> actually resolves to. Appended rather than slotted
' in beside Enabled where it reads best: every existing reader addresses this
' sheet by fixed column number, and renumbering nine of them to gain one column
' of tidiness is precisely the kind of silent misread this workspace was split up
' to stop happening.
Public Const PS_METRIC_USE_COL As Long = 10

' The three answers to "where should this figure come from".
Public Const USE_DERIVED As String = "Manual file only"
Public Const USE_SYSTEM As String = "System output"
Public Const USE_DERIVED_ELSE As String = "Manual else system"

Public Const PS_MAP_HEADER_ROW As Long = PS_HDR_ROW
Public Const PS_MAP_FIRST_ROW As Long = PS_FIRST_ROW
Public Const PS_MAP_LAST_ROW As Long = PS_FIRST_ROW + 299
Public Const PS_MAP_COLS As Long = 7
' Which dimensions a figure is broken down by, so that "why is the pre-shock that
' much" has an answer. Appended rather than slotted in where it reads best: every
' reader of this sheet addresses it by fixed column number.
Public Const PS_MAP_BREAK_COL As Long = 7

' The breakdown is bounded on purpose. Five dimensions by eight measures is forty
' dictionary bumps on every one of ninety-five thousand rows, per test case, per
' source - minutes of work to answer a question nobody asked about most of it.
Public Const MAX_BREAKDOWN_DIMS As Long = 4
Public Const MAX_BREAKDOWN_MEASURES As Long = 2
' A guard against a dimension that turns out to be an account number. A breakdown
' with ninety thousand values in it is not an explanation of anything.
Public Const MAX_BREAKDOWN_VALUES As Long = 4000
Private Const BRK_PREFIX As String = "BRK#"

' The sheet the breakdowns land on.
Public Const PS_BREAKDOWN_SHEET As String = "Pre_Shock_Breakdown"
Public Const PB_COLS As Long = 12

Public Const PS_RESULT_HEADER_ROW As Long = PS_HDR_ROW
Public Const PS_RESULT_FIRST_ROW As Long = PS_FIRST_ROW

' The same derived figures are also written to their own sheet, so base and pre-shock can
' be read, filtered and referenced without scrolling through configuration.
' The store uses the same layout as every other sheet in the workspace, in the
' tool and inside each generated workbook alike.
Public Const DERIVED_SHEET As String = "Derived_Values"
Public Const DV_HEADER_ROW As Long = PS_HDR_ROW
Public Const DV_FIRST_ROW As Long = PS_FIRST_ROW
' 1-12 are the derived figures and their provenance; 13-16 are the system's own
' numbers beside them and the record of which of the two the configuration read.
Public Const DV_COLS As Long = 16
Public Const DV_COLS_LEGACY As Long = 12

' ---------------------------------------------------------------------------
'  Sources
'
'  There is more than one input extract and there will be more still: ECL carries the
'  outstanding / IIS / ECL measures, CAPRWA carries OUTSTANDING_FOR_RWA and RWA_LCY, and
'  the ALM extracts will carry the liquidity measures. They share the same account key and
'  most of the same dimension fields, so they are treated as one registry of sources rather
'  than as one hard-coded "the ECL file".
'
'  mSources: source key -> { Data, Headers, Path, Sheet, HeaderRow, Date1904, BaseFilter }
'  mActiveSource names the source whose arrays the engine helpers are currently reading,
'  because those helpers (EvalFilter, aggregation) walk a single table at a time.
' ---------------------------------------------------------------------------
Private mSources As Object
Private mActiveSource As String
' Filter expression -> the account numbers ECL says it selects. One resolution
' per distinct condition per run, however many sources borrow it.
Private mJoinCache As Object
' Validated date keys are reusable while a registered source is unchanged.
' Store/Register and the start of each reconciliation invalidate this cache.
Private mSourceDateCache As Object

' The table the engine is currently walking. Kept as module state rather than threaded
' through every helper, because the filter evaluator is called once per row per clause and
' the indirection would cost real time on a 95,000-row extract.
Private mEclData As Variant
Private mEclHeaders As Object
Private mEclPath As String
Private mEclSheet As String
Private mEclHeaderRow As Long
Private mEclDate1904 As Boolean
Private mPreShockMeasureFields As Object
' GREATEST(a, b) measures: source|spec -> the columns compared on each row.
Private mGreatest As Object
' Every column name the extract carries, whether or not its data was loaded, so
' the condition builder can still offer any field in the file.
Private mEclAllFields As Object

' Overrides recovered from the retired sheets, held until the next test-case sync can
' re-apply them. (VBA only accepts module-level declarations in this header block.)
Private mMigratedOverrides As Object
' Why the last LoadEclOutput failed, or "". It logs a failure rather than
' raising it, so a caller loading several files reads this to say which failed.
Private mLoadFailure As String

' ---- speed (performance patch) -------------------------------------------
' The stored record mEclData was last copied from. ActivateSource skips the
' copy while the engine still holds that very table: copying a whole extract was
' the dearest thing a run did, dozens of times per test case. Anything that
' assigns mEclData directly calls ForgetActiveTable.
Private mActiveRecord As Object
Private mLent As Boolean
' Each table put into the store gets a fresh stamp.
Private mTableStamp As Long
' Stamp -> what has been worked out about that table: normalised filter columns
' and row masks per condition, shared by every test case that asks the same.
Private mMaskStore As Object
Private mMaskBytes As Double
Private Const MEMO_BYTE_LIMIT As Double = 64000000#
' Metric formula -> its arithmetic, stage and row selector (pure, regex-bound).
Private mFormulaSplits As Object

' ---------------------------------------------------------------------------
'  The breakdown: what a figure is MADE OF.
'
'  Resolved once per source activation and then read on every matching row, so
'  these are plain arrays rather than dictionaries - a dictionary probe per
'  dimension per row is a million probes nobody needs to pay for.
' ---------------------------------------------------------------------------
Private mBreakdownOn As Boolean
Private mBrkDimCount As Long
Private mBrkDimName() As String
Private mBrkDimCol() As Long
Private mBrkMeasureCount As Long
Private mBrkMeasureName() As String
Private mBrkMeasureCol() As Long
' One slot per dimension VALUE, so the row loop never builds a string.
Private mBrkMap() As Object
Private mBrkSlotCount As Long
Private mBrkSlotDim() As Long
Private mBrkSlotVal() As String
Private mBrkSum() As Double

' ---------------------------------------------------------------------------
'  Dimension-value cache
'
'  The filter builder used to reopen the ECL workbook and then walk every row once per
'  logical dimension just to populate its pickers. On a real extract that is minutes of
'  waiting before the window appears.
'
'  The values are now collected ONCE, while the extract is already open during upload, and
'  written to a very hidden sheet. The builder reads that sheet, so opening it costs no
'  file I/O and no scan at all. This is the approach ALM Rule Studio v9 uses, and it is
'  what stopped its editor hanging on the large LCR output.
' ---------------------------------------------------------------------------
Private Const CACHE_SHEET As String = "_ECL_Cache"
Private Const BUILDER_SHEET As String = "_Builder_Src"
Private Const CACHE_FIRST_ROW As Long = 10
Private Const CACHE_MAX_VALUES As Long = 400
' Enough rows to see a dimension's whole vocabulary without walking the extract.
Private Const CACHE_SAMPLE_ROWS As Long = 60000
' Bumped when the cache sheet's own layout changes, so a cache written by an earlier build is
' discarded rather than misread.
Private Const CACHE_LAYOUT As String = "c2"

' ===================== sheet construction and migration =====================

' --- the six sheets -------------------------------------------------------

Public Function PreShockSheet() As Worksheet
    Set PreShockSheet = GetWorksheetSafe(ThisWorkbook, PRE_SHOCK_SHEET)
End Function
Public Function PsSourcesSheet() As Worksheet
    Set PsSourcesSheet = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET)
End Function
Public Function PsMetricsSheet() As Worksheet
    Set PsMetricsSheet = GetWorksheetSafe(ThisWorkbook, PS_METRICS_SHEET)
End Function
Public Function PsFieldsSheet() As Worksheet
    Set PsFieldsSheet = GetWorksheetSafe(ThisWorkbook, PS_FIELDS_SHEET)
End Function
Public Function PsCasesSheet() As Worksheet
    Set PsCasesSheet = GetWorksheetSafe(ThisWorkbook, PS_CASES_SHEET)
End Function
Public Function PsResultsSheet() As Worksheet
    Set PsResultsSheet = GetWorksheetSafe(ThisWorkbook, PS_RESULTS_SHEET)
End Function

Public Function PsSheetNames() As Variant
    PsSheetNames = Array(PRE_SHOCK_SHEET, PS_SOURCES_SHEET, PS_METRICS_SHEET, _
                         PS_FIELDS_SHEET, PS_CASES_SHEET, PS_RESULTS_SHEET, _
                         DERIVED_SHEET, PS_BREAKDOWN_SHEET)
End Function

' Kept as the old name because a hundred call sites use it, but it now builds the
' whole six-sheet workspace rather than one crowded sheet.
Public Sub EnsurePreShockSheet()
    EnsurePreShockWorkspace
End Sub

Public Sub EnsurePreShockWorkspace()
    Dim ws As Worksheet, fresh As Boolean, carried As Object, legacy As Worksheet, nm As Variant

    Set ws = PreShockSheet()
    If ws Is Nothing Then fresh = True
    If Not fresh Then
        If SafeText(ws.Range("N1").Value2) <> PS_LAYOUT_VERSION Then fresh = True
    End If
    If Not fresh Then
        ' A workspace is only complete if every one of its sheets is present.
        For Each nm In PsSheetNames()
            If Not SheetExistsSafe(CStr(nm)) Then fresh = True
        Next nm
    End If
    ' An existing workspace still gets metrics it has never seen.
    '
    ' Seeding used to happen only when the whole workspace was rebuilt, so a
    ' workbook that already had one - which is every workbook in use - would
    ' never acquire the liquidity, capital or max-ECL metrics at all. They were
    ' defined, compiled, and silently absent from the only sheet that matters.
    If Not fresh Then AppendMissingMetrics PsMetricsSheet(): Exit Sub

    ' Everything the user owns is lifted out BY HEADER TEXT before anything is
    ' rebuilt, from wherever it currently lives - the old single sheet included -
    ' so a layout change never costs an edited formula or a filter override.
    modBasePreShock.KeepAsOfSetting
    Set legacy = GetWorksheetSafe(ThisWorkbook, LEGACY_COMBINED_SHEET)
    If Not legacy Is Nothing Then
        Set carried = HarvestExisting(legacy)
    ElseIf Not PreShockSheet() Is Nothing Then
        Set carried = HarvestFromWorkspace()
    End If

    BuildWorkspaceSheets
    SeedMetricDefaults PsMetricsSheet()
    SeedDimensionDefaults PsFieldsSheet()
    If Not carried Is Nothing Then RestoreHarvested PreShockSheet(), carried
    MigrateLegacySheets PreShockSheet()

    ' The single combined sheet has done its job; keeping it would leave two
    ' copies of every rule and no way to tell which one is live.
    If Not legacy Is Nothing Then
        On Error Resume Next
        Application.DisplayAlerts = False
        legacy.Delete
        Application.DisplayAlerts = True
        Err.Clear
        On Error GoTo 0
    End If

    ApplyValidation PreShockSheet()
    PreShockSheet().Range("N1").Value2 = PS_LAYOUT_VERSION
    PreShockSheet().columns("N").hidden = True
    RefreshStatusLine PreShockSheet()
    On Error Resume Next
    StylePreShockSheet PreShockSheet()
    Err.Clear
    On Error GoTo 0
    modBasePreShock.ApplyBasePreShockLayout
End Sub

Private Function SheetExistsSafe(ByVal nm As String) As Boolean
    SheetExistsSafe = Not GetWorksheetSafe(ThisWorkbook, nm) Is Nothing
End Function

' Creates the six sheets in reading order and writes each one's fixed furniture.
Private Sub BuildWorkspaceSheets()
    Dim nm As Variant, ws As Worksheet, prev As Worksheet, anchor As Worksheet
    Set anchor = GetWorksheetSafe(ThisWorkbook, SHEET_FORMATTING)
    Set prev = anchor
    For Each nm In PsSheetNames()
        Set ws = GetWorksheetSafe(ThisWorkbook, CStr(nm))
        If ws Is Nothing Then
            If prev Is Nothing Then
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
            Else
                Set ws = ThisWorkbook.Worksheets.Add(After:=prev)
            End If
            ws.name = CStr(nm)
        End If
        On Error Resume Next
        If Not prev Is Nothing Then ws.Move After:=prev
        Err.Clear
        On Error GoTo 0
        Set prev = ws
    Next nm

    WriteHomeLayout PreShockSheet()
    WriteSourcesLayout PsSourcesSheet()
    WriteMetricsLayout PsMetricsSheet()
    WriteFieldsLayout PsFieldsSheet()
    WriteCasesLayout PsCasesSheet()
    WriteResultsLayout PsResultsSheet()
    WriteBreakdownLayout BreakdownSheet(True)
End Sub

' The breakdown sheet before anything has been broken down, so it explains
' itself rather than being an empty sheet with a name.
Private Sub WriteBreakdownLayout(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    PsChrome ws, "Why the pre-shock is what it is", _
        "Run the tests and this fills with one row per test case, dimension and value: what the filter selected beside what the " & _
        "whole portfolio holds of the same thing. Choose the dimensions in the 'Break down by' column on " & PS_FIELDS_SHEET & "."
    ws.Cells(PS_HDR_ROW, 1).Resize(1, 17).value = _
        Array("As-of date", "Entity", "Test case", "Test element", "Source", "Dimension", "Value", _
              "Rows selected", "Rows in portfolio", "Selected", "In portfolio", "Share selected", _
              "Selected (2nd)", "In portfolio (2nd)", "Measure", "Measure (2nd)", "Filter applied")
End Sub

' Lifts everything the user owns off the sheet by locating each block from its HEADER TEXT
' rather than a fixed row, so it keeps working across layout changes.
Private Function HarvestExisting(ByVal ws As Worksheet) As Object
    Dim res As Object, lastRow As Long, r As Long, hdr As Long, x As Object, key As String
    Set res = NewMap()
    Set res("Metrics") = NewMap()
    Set res("Dims") = NewMap()
    Set res("Cases") = NewMap()
    res("EclPath") = "": res("EclSheet") = "": res("Enabled") = "": res("DateMode") = ""
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row

    ' Source settings, found by their label wherever they sit.
    For r = 1 To WorksheetFunction.Min(lastRow, 40)
        key = SafeUpperText(ws.Cells(r, 1).Value2)
        Select Case True
            Case key Like "ECL SOURCE*": res("EclPath") = SafeText(ws.Cells(r, 2).Value2)
            Case key Like "DETECTED SHEET*": res("EclSheet") = SafeText(ws.Cells(r, 2).Value2)
            Case key Like "*INDEPENDENT ECL CHECK*": res("Enabled") = SafeText(ws.Cells(r, 2).Value2)
        End Select
        If SafeUpperText(ws.Cells(r, 5).Value2) Like "AS-OF DATE MATCHING*" Then res("DateMode") = SafeText(ws.Cells(r, 6).Value2)
    Next r

    ' Read the metric block by COLUMN HEADING, not by position. The Source column was added
    ' later, so a sheet written by an earlier version has Enabled where Source now sits -
    ' harvesting by position would silently record Enabled="Yes" as the source name.
    hdr = FindHeaderRow(ws, "Pre-shock metric", lastRow)
    If hdr = 0 Then hdr = FindHeaderRow(ws, "Derived metric", lastRow)
    If hdr > 0 Then
        Dim mc As Object
        Set mc = HeaderColumns(ws, hdr, 12)
        For r = hdr + 1 To lastRow
            key = SafeUpperText(ws.Cells(r, 1).Value2)
            If Len(key) = 0 Then Exit For
            Set x = NewMap()
            x("Source") = HeaderValue(ws, r, mc, "Source")
            x("Enabled") = HeaderValue(ws, r, mc, "Enabled")
            x("Formula") = HeaderValue(ws, r, mc, "Derivation formula")
            If Len(CStr(x("Formula"))) = 0 Then x("Formula") = HeaderValue(ws, r, mc, "Pre-shock formula")
            x("Stage") = HeaderValue(ws, r, mc, "Stage")
            x("Notes") = HeaderValue(ws, r, mc, "Notes")
            Set res("Metrics")(key) = x
        Next r
    End If

    hdr = FindHeaderRow(ws, "Logical dimension", lastRow)
    If hdr > 0 Then
        For r = hdr + 1 To lastRow
            key = SafeUpperText(ws.Cells(r, 1).Value2)
            If Len(key) = 0 Then Exit For
            Set x = NewMap()
            x("Candidates") = SafeText(ws.Cells(r, 2).Value2)
            x("Mode") = SafeText(ws.Cells(r, 3).Value2)
            x("Notes") = SafeText(ws.Cells(r, 6).Value2)
            Set res("Dims")(key) = x
        Next r
    End If

    hdr = FindHeaderRow(ws, "Test case", lastRow, 2)
    If hdr > 0 Then
        For r = hdr + 1 To lastRow
            key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
            If key = "|" Then Exit For
            Set x = NewMap()
            x("Enabled") = SafeText(ws.Cells(r, 1).Value2)
            x("Override") = SafeText(ws.Cells(r, 7).Value2)
            x("Notes") = SafeText(ws.Cells(r, 12).Value2)
            Set res("Cases")(key) = x
        Next r
    End If
    Set HarvestExisting = res
End Function

Private Function FindHeaderRow(ByVal ws As Worksheet, ByVal caption As String, ByVal lastRow As Long, Optional ByVal col As Long = 1) As Long
    Dim r As Long
    For r = 1 To lastRow
        If StrComp(SafeText(ws.Cells(r, col).Value2), caption, vbTextCompare) = 0 Then FindHeaderRow = r: Exit Function
    Next r
End Function

' Upper-cased heading -> column index, for reading a block whose columns may have moved
' between layout versions.
Private Function HeaderColumns(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal cols As Long) As Object
    Dim c As Long, t As String, d As Object
    Set d = NewMap()
    For c = 1 To cols
        t = SafeUpperText(ws.Cells(hdrRow, c).Value2)
        If Len(t) > 0 Then
            If Not d.Exists(t) Then d(t) = c
        End If
    Next c
    Set HeaderColumns = d
End Function

Private Function HeaderValue(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Object, ByVal caption As String) As String
    Dim k As String
    k = UCase$(Trim$(caption))
    If cols Is Nothing Then Exit Function
    If Not cols.Exists(k) Then Exit Function
    HeaderValue = SafeText(ws.Cells(r, CLng(cols(k))).Value2)
End Function

Private Sub RestoreHarvested(ByVal ws As Worksheet, ByVal carried As Object)
    Dim r As Long, key As String, x As Object, n As Long
    If Len(CStr(carried("EclPath"))) > 0 Then
        ws.Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2 = carried("EclPath")
        ws.Range("B" & (PS_SRC_FIRST_ROW + 5)).Value2 = carried("EclSheet")
    End If
    If Len(CStr(carried("Enabled"))) > 0 Then ws.Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2 = carried("Enabled")
    If Len(CStr(carried("DateMode"))) > 0 Then ws.Range("F" & (PS_SRC_FIRST_ROW + 3)).Value2 = carried("DateMode")

    ' Each block goes back onto its own sheet.
    Set ws = PsMetricsSheet()
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        key = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(key) = 0 Then Exit For
        If carried("Metrics").Exists(key) Then
            Set x = carried("Metrics")(key)
            If Len(CStr(x("Source"))) > 0 Then ws.Cells(r, 2).Value2 = x("Source")
            If Len(CStr(x("Enabled"))) > 0 Then ws.Cells(r, 3).Value2 = x("Enabled")
            If Len(CStr(x("Formula"))) > 0 Then ws.Cells(r, 4).Value2 = x("Formula")
            If Len(CStr(x("Stage"))) > 0 Then ws.Cells(r, 5).Value2 = x("Stage")
            If Len(CStr(x("Notes"))) > 0 Then ws.Cells(r, 9).Value2 = x("Notes")
            carried("Metrics").Remove key
            n = n + 1
        End If
    Next r
    Set ws = PsFieldsSheet()
    For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        key = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(key) = 0 Then Exit For
        If carried("Dims").Exists(key) Then
            Set x = carried("Dims")(key)
            If Len(CStr(x("Candidates"))) > 0 Then ws.Cells(r, 2).Value2 = x("Candidates")
            If Len(CStr(x("Mode"))) > 0 Then ws.Cells(r, 3).Value2 = x("Mode")
            If Len(CStr(x("Notes"))) > 0 Then ws.Cells(r, 6).Value2 = x("Notes")
            If x.Exists("Break") Then
                If Len(CStr(x("Break"))) > 0 Then ws.Cells(r, PS_MAP_BREAK_COL).Value2 = x("Break")
            End If
            n = n + 1
        End If
    Next r
    ' Case overrides are re-applied by the next test-case sync, which rebuilds that block.
    Set mMigratedOverrides = NewMap()
    Dim ck As Variant
    For Each ck In carried("Cases").keys
        If Len(CStr(carried("Cases")(ck)("Override"))) > 0 Then mMigratedOverrides(CStr(ck)) = carried("Cases")(ck)("Override")
    Next ck
    If n > 0 Or mMigratedOverrides.count > 0 Then
        LogIssue LOG_LEVEL_INFO, "ECL pre-shock", "Sheet rebuilt for the new layout; " & n & " setting(s) and " & _
            mMigratedOverrides.count & " filter override(s) carried over.", PRE_SHOCK_SHEET
    End If
End Sub

' One line at the top that says where things stand, because the detail is hundreds of rows
' further down and nobody scrolls to find out whether anything ran.
Public Sub RefreshStatusLine(ByVal ws As Worksheet)
    Dim s As String, results As Long, baseOk As Long, preOk As Long, diffs As Long
    Dim r As Long, lastRow As Long, a As Variant
    On Error GoTo Done
    Dim stored As Long
    stored = PersistedDerivedCount()
    If Not ValueCacheReady() Then
        s = "Not started.  STEP 1 - click Upload ECL to load an extract."
        If stored > 0 Then s = s & "  (" & format$(stored, "#,##0") & " value(s) from an earlier run are still stored on " & DERIVED_SHEET & " and will be used.)"
    ElseIf Not PreShockEnabled() Then
        s = "Loaded (" & CacheSummary() & ").  Next: click Run on the home screen."
    Else
        ' One results table: counted off Derived_Values.
        modValueSources.VS_ResultCounts results, baseOk, preOk, diffs
        If results = 0 Then
            s = "Loaded (" & CacheSummary() & ").  Next: click Run on the home screen."
        Else
            s = results & " comparisons  |  base rebuilt " & baseOk & "  |  pre-shock rebuilt " & preOk & _
                "  |  " & diffs & " differ from the system  |  " & format$(stored, "#,##0") & _
                " value(s) stored on " & DERIVED_SHEET & " and read by the configuration as derived.<metric>"
        End If
    End If
Done:
    On Error Resume Next
    ' The status line is written to the home sheet whichever sheet was read for it.
    PreShockSheet().Range("A" & PS_STATUS_ROW).Value2 = s
    On Error GoTo 0
End Sub

' --- the six layouts -------------------------------------------------------
'
' Each writes only its own sheet's furniture. Nothing here knows the row offsets
' of any other block, which is the whole reason the split was worth doing.

Private Sub PsChrome(ByVal ws As Worksheet, ByVal title As String, ByVal about As String)
    ws.Range("A1").Value2 = "JKB"
    ws.Cells(PS_TITLE_ROW, 1).Value2 = title
    ws.Cells(PS_ABOUT_ROW, 1).Value2 = about
End Sub

Private Sub WriteHomeLayout(ByVal ws As Worksheet)
    PsChrome ws, "Base and pre-shock derivation", _
        "BASE is the bank's whole-portfolio total for a measure, unfiltered. PRE-SHOCK is that same total restricted to the test case's " & _
        "filter condition, so it is identical across Moderate / Medium / Severe and differs per test case. Both are rebuilt here from the " & _
        "uploaded extracts, stored on " & DERIVED_SHEET & ", and read by the configuration as derived.<metric> - not copied from the system output."

    ws.Cells(PS_HDR_ROW, 1).Value2 = "THE THREE STEPS"
    ws.Range("A" & PS_SRC_FIRST_ROW).Value2 = "STEP 1"
    ws.Range("B" & PS_SRC_FIRST_ROW).Value2 = "Upload the extracts      (buttons: Upload ECL, Upload CAPRWA)"
    ws.Range("A" & (PS_SRC_FIRST_ROW + 1)).Value2 = "STEP 2"
    ws.Range("B" & (PS_SRC_FIRST_ROW + 1)).Value2 = "Turn independent checking on, and choose date matching"
    ws.Range("A" & (PS_SRC_FIRST_ROW + 2)).Value2 = "STEP 3"
    ws.Range("B" & (PS_SRC_FIRST_ROW + 2)).Value2 = "Run the tests             (button: Run tests)"
    ' The one switch that decides what derived.<metric> resolves to, for every
    ' metric that does not name its own answer on its rule row.
    ws.Range("E" & (PS_SRC_FIRST_ROW + 2)).Value2 = "Values come from"
    If Len(SafeText(ws.Range("F" & (PS_SRC_FIRST_ROW + 2)).Value2)) = 0 Then _
        ws.Range("F" & (PS_SRC_FIRST_ROW + 2)).Value2 = "Set on " & VS_SHEET
    ws.Range("H" & (PS_SRC_FIRST_ROW + 2)).Value2 = "Explain the figures"
    If Len(SafeText(ws.Range("I" & (PS_SRC_FIRST_ROW + 2)).Value2)) = 0 Then _
        ws.Range("I" & (PS_SRC_FIRST_ROW + 2)).Value2 = "Yes"
    ws.Range("K" & (PS_SRC_FIRST_ROW + 2)).Value2 = "(breaks every figure down by the dimensions marked on " & PS_FIELDS_SHEET & ")"
    ws.Range("A" & (PS_SRC_FIRST_ROW + 3)).Value2 = "Independent checking"
    ws.Range("E" & (PS_SRC_FIRST_ROW + 3)).Value2 = "As-of date matching"
    ws.Range("H" & (PS_SRC_FIRST_ROW + 3)).Value2 = "As-of dates in file"
    ws.Range("A" & (PS_SRC_FIRST_ROW + 4)).Value2 = "Detection notes"
    If Len(SafeText(ws.Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2)) = 0 Then ws.Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2 = "No"
    ' Strict is the safe default: a figure rebuilt from a different reporting date is not a
    ' valid check. Ignore exists for trimmed or sample extracts, where the caveat is accepted.
    If Len(SafeText(ws.Range("F" & (PS_SRC_FIRST_ROW + 3)).Value2)) = 0 Then ws.Range("F" & (PS_SRC_FIRST_ROW + 3)).Value2 = "Strict"
    ws.Range("I" & (PS_SRC_FIRST_ROW + 3)).NumberFormat = "@"

    WriteHomeMap ws
End Sub

' A map of the workspace on the home sheet, so the five other sheets announce
' themselves instead of having to be discovered along the tab strip.
Private Sub WriteHomeMap(ByVal ws As Worksheet)
    Dim r As Long
    r = PS_SRC_LAST_ROW + 2
    ws.Cells(r, 1).Value2 = "WHERE EVERYTHING LIVES"
    r = r + 1
    HomeMapRow ws, r, PS_SOURCES_SHEET, "Which extract is loaded, what it supplies, how many rows, which reporting date.": r = r + 1
    HomeMapRow ws, r, PS_METRICS_SHEET, "One row per derived metric: the source it comes from and the formula that derives it.": r = r + 1
    HomeMapRow ws, r, PS_FIELDS_SHEET, "Logical dimension to physical column, and what values each one actually holds.": r = r + 1
    HomeMapRow ws, r, PS_CASES_SHEET, "One row per test case and element, with the filter that defines its pre-shock population.": r = r + 1
    HomeMapRow ws, r, DERIVED_SHEET, "Base and pre-shock from the input files and from the system, the difference, and Reconciles.": r = r + 1
    HomeMapRow ws, r, DERIVED_SHEET, "The store. Every base and pre-shock figure the tool holds, and what the config reads.": r = r + 1
    HomeMapRow ws, r, PS_BREAKDOWN_SHEET, "Why each figure is what it is: what the filter took, against what the portfolio holds.": r = r + 1
End Sub

Private Sub HomeMapRow(ByVal ws As Worksheet, ByVal r As Long, ByVal sheetName As String, ByVal what As String)
    ws.Cells(r, 1).Value2 = sheetName
    ws.Cells(r, 2).Value2 = what
    On Error Resume Next
    ws.Hyperlinks.Add anchor:=ws.Cells(r, 1), address:="", SubAddress:="'" & sheetName & "'!A1"
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub WriteSourcesLayout(ByVal ws As Worksheet)
    Dim si As Long, sk As Variant
    PsChrome ws, "Source extracts", _
        "One row per input the derivation can draw on. ECL carries the outstanding, IIS and ECL measures; CAPRWA carries the RWA ones. " & _
        "A metric rule names the source it comes from, so adding a source is a row here plus its metric rows."
    ws.Cells(PS_HDR_ROW, 1).Resize(1, 8).value = _
        Array("Source", "What it supplies", "File", "Sheet", "Rows", "Base scope filter", "As-of date(s)", "Status")
    si = 0
    For Each sk In SourceKeys()
        ws.Cells(PS_SRC_LIST_ROW + si, 1).Value2 = CStr(sk)
        ws.Cells(PS_SRC_LIST_ROW + si, 2).Value2 = SourceLabel(CStr(sk))
        If Len(SafeText(ws.Cells(PS_SRC_LIST_ROW + si, 8).Value2)) = 0 Then ws.Cells(PS_SRC_LIST_ROW + si, 8).Value2 = "Not loaded"
        si = si + 1
    Next sk
End Sub

Private Sub WriteMetricsLayout(ByVal ws As Worksheet)
    PsChrome ws, "Derived metric rules", _
        "THIS is where a derivation formula is stored, and the only place. The metric name is the same string as the OUTPUT_ROW_LABEL " & _
        "in Config_ElementTypeRules, which is how a config row refers to it as derived.<metric>.  " & _
        "Whether derived.<metric> reads the input files or the system is set on " & VS_SHEET & "."
    ws.Cells(PS_METRIC_HEADER_ROW, 1).Resize(1, PS_METRIC_COLS).value = _
        Array("Derived metric", "Source", "Enabled", "Derivation formula", "Stage", "Purpose", _
              "Availability", "Used by element types", "Notes", "(set on " & VS_SHEET & ")")
End Sub

Private Sub WriteFieldsLayout(ByVal ws As Worksheet)
    PsChrome ws, "Dimension mapping", _
        "A filter condition names a logical dimension; an extract names a physical column. This is the translation, and it is where " & _
        "you find out whether a dimension a test case filters on exists in the file you uploaded.  'Break down by' marks the " & _
        "dimensions every pre-shock figure is explained against on " & PS_BREAKDOWN_SHEET & " - at most " & MAX_BREAKDOWN_DIMS & " of them."
    ws.Cells(PS_MAP_HEADER_ROW, 1).Resize(1, PS_MAP_COLS).value = _
        Array("Logical dimension", "Candidate field(s)", "Normalization", "Availability", "Distinct values", "Notes", "Break down by")
End Sub

Private Sub WriteCasesLayout(ByVal ws As Worksheet)
    PsChrome ws, "Test cases and their filters", _
        "One row per test case and element. The filter here is what turns the bank's base into this test case's pre-shock population - " & _
        "the same for every severity, different for every test case. Double-click a filter cell to open the builder."
    ws.Cells(PS_CASE_HEADER_ROW, 1).Resize(1, 12).value = _
        Array("Enabled", "Test case", "Test element", "Element type", "Source", "System filter", "Override filter", "Effective filter", "Mapping status", "Rows matched", "Last result", "Notes")
End Sub

Private Sub WriteResultsLayout(ByVal ws As Worksheet)
    PsChrome ws, "Derivation results", _
        "The system's own figure beside the one rebuilt independently from the extract, for base and for pre-shock, with the difference. " & _
        "A difference here is the whole point of the exercise."
    ws.Cells(PS_RESULT_HEADER_ROW, 1).Resize(1, 12).value = _
        Array("As-of date", "Entity", "Test case", "Test element", "Metric", _
              "BASE system", "BASE derived", "BASE diff", _
              "PRE-SHOCK system", "PRE-SHOCK derived", "PRE-SHOCK diff", "Status")
End Sub

' Lifts the user's edits out of an already-split workspace, so a later layout
' change is as safe as the first one was.
Private Function HarvestFromWorkspace() As Object
    Dim res As Object, r As Long, x As Object, key As String, ws As Worksheet, lastRow As Long
    Set res = NewMap()
    Set res("Metrics") = NewMap()
    Set res("Dims") = NewMap()
    Set res("Cases") = NewMap()
    res("EclPath") = "": res("EclSheet") = "": res("Enabled") = "": res("DateMode") = ""

    Set ws = PreShockSheet()
    If Not ws Is Nothing Then
        res("Enabled") = SafeText(ws.Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2)
        res("DateMode") = SafeText(ws.Range("F" & (PS_SRC_FIRST_ROW + 3)).Value2)
        res("EclPath") = SafeText(ws.Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2)
        res("EclSheet") = SafeText(ws.Range("B" & (PS_SRC_FIRST_ROW + 5)).Value2)
    End If

    Set ws = PsMetricsSheet()
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
        For r = PS_METRIC_FIRST_ROW To lastRow
            key = SafeUpperText(ws.Cells(r, 1).Value2)
            If Len(key) > 0 Then
                Set x = NewMap()
                x("Source") = SafeText(ws.Cells(r, 2).Value2)
                x("Enabled") = SafeText(ws.Cells(r, 3).Value2)
                x("Formula") = SafeText(ws.Cells(r, 4).Value2)
                x("Stage") = SafeText(ws.Cells(r, 5).Value2)
                x("Notes") = SafeText(ws.Cells(r, 9).Value2)
                Set res("Metrics")(key) = x
            End If
        Next r
    End If

    Set ws = PsFieldsSheet()
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
        For r = PS_MAP_FIRST_ROW To lastRow
            key = SafeUpperText(ws.Cells(r, 1).Value2)
            If Len(key) > 0 Then
                Set x = NewMap()
                x("Candidates") = SafeText(ws.Cells(r, 2).Value2)
                x("Mode") = SafeText(ws.Cells(r, 3).Value2)
                x("Notes") = SafeText(ws.Cells(r, 6).Value2)
                x("Break") = SafeText(ws.Cells(r, PS_MAP_BREAK_COL).Value2)
                Set res("Dims")(key) = x
            End If
        Next r
    End If

    Set ws = PsCasesSheet()
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.count, 2).End(xlUp).row
        For r = PS_CASE_FIRST_ROW To lastRow
            key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
            If key <> "|" Then
                Set x = NewMap()
                x("Enabled") = SafeText(ws.Cells(r, 1).Value2)
                x("Override") = SafeText(ws.Cells(r, 7).Value2)
                x("Notes") = SafeText(ws.Cells(r, 12).Value2)
                Set res("Cases")(key) = x
            End If
        Next r
    End If
    Set HarvestFromWorkspace = res
End Function

' Seeded once. User edits are never overwritten - a row is only written when its
' metric name cell is still empty.
' Each row names the SOURCE that supplies it. ECL carries outstanding / IIS / ECL; CAPRWA
' carries the RWA measures, which no ECL extract can provide - which is why the RWA rows
' could only ever copy the system's own numbers before.
Private Sub SeedMetricDefaults(ByVal ws As Worksheet)
    Dim m As Variant, i As Long, r As Long
    ' Split across two statements: VBA allows at most 25 line continuations in one statement
    ' and the combined list is longer than that.
    Dim at As Long
    at = PS_METRIC_FIRST_ROW
    m = SeedMetricsEcl():      WriteMetricSeeds ws, m, at: at = at + UBound(m) + 1
    m = SeedMetricsCapRwa():   WriteMetricSeeds ws, m, at: at = at + UBound(m) + 1
    ' The liquidity, capital and max-ECL metrics. Each carries its own row
    ' selector in the formula, because the three extracts they read are not
    ' staged and "which rows" has to be answered some other way.
    m = modScenarioBuilder_Multi.SeedMetricsMaxEcl():     WriteMetricSeeds ws, m, at: at = at + UBound(m) + 1
    m = modScenarioBuilder_Multi.SeedMetricsLiquidity():  WriteMetricSeeds ws, m, at: at = at + UBound(m) + 1
    m = modScenarioBuilder_Multi.SeedMetricsCapital():    WriteMetricSeeds ws, m, at: at = at + UBound(m) + 1
    m = modValueSources.VS_SeedMetrics():                WriteMetricSeeds ws, m, at
    modScenarioBuilder_Multi.UpgradeCapitalMetricDefaults ws
End Sub

' Adds any seeded metric this sheet has not got, at the first free row.
'
' Matched on the metric NAME, so a row the user has edited is never touched and
' running this twice never duplicates anything. This is what lets a new metric
' reach a workbook that was built before it existed.
Private Sub AppendMissingMetrics(ByVal ws As Worksheet)
    Dim have As Object, r As Long, nm As String, freeRow As Long, added As Long
    Dim sets_ As Variant, m As Variant, i As Long, si As Long

    If ws Is Nothing Then Exit Sub
    Set have = NewMap()
    freeRow = 0
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        nm = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(nm) = 0 Then
            If freeRow = 0 Then freeRow = r
        Else
            have(nm) = True
        End If
    Next r
    If freeRow = 0 Then Exit Sub

    sets_ = Array(SeedMetricsEcl(), SeedMetricsCapRwa(), _
                  modScenarioBuilder_Multi.SeedMetricsMaxEcl(), _
                  modScenarioBuilder_Multi.SeedMetricsLiquidity(), _
                  modScenarioBuilder_Multi.SeedMetricsCapital(), _
                  modValueSources.VS_SeedMetrics())
    For si = LBound(sets_) To UBound(sets_)
        m = sets_(si)
        For i = LBound(m) To UBound(m)
            nm = SafeUpperText(m(i)(0))
            If Not have.Exists(nm) Then
                If freeRow > PS_METRIC_LAST_ROW Then Exit Sub
                ws.Cells(freeRow, 1).Value2 = m(i)(0)
                ws.Cells(freeRow, 2).Value2 = m(i)(1)
                ws.Cells(freeRow, 3).Value2 = "Yes"
                ws.Cells(freeRow, 4).Value2 = m(i)(2)
                ws.Cells(freeRow, 5).Value2 = m(i)(3)
                ws.Cells(freeRow, 6).Value2 = m(i)(4)
                ws.Cells(freeRow, 9).Value2 = m(i)(5)
                have(nm) = True
                freeRow = freeRow + 1
                added = added + 1
            End If
        Next i
    Next si
End Sub

Private Sub WriteMetricSeeds(ByVal ws As Worksheet, ByVal m As Variant, ByVal firstRow As Long)
    Dim i As Long, r As Long
    For i = LBound(m) To UBound(m)
        r = firstRow + i
        If r > PS_METRIC_LAST_ROW Then Exit For
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then
            ws.Cells(r, 1).Value2 = m(i)(0)
            ws.Cells(r, 2).Value2 = m(i)(1)
            ws.Cells(r, 3).Value2 = "Yes"
            ws.Cells(r, 4).Value2 = m(i)(2)
            ws.Cells(r, 5).Value2 = m(i)(3)
            ws.Cells(r, 6).Value2 = m(i)(4)
            ws.Cells(r, 9).Value2 = m(i)(5)
        End If
    Next i
End Sub

Private Function SeedMetricsCapRwa() As Variant
    SeedMetricsCapRwa = Array( _
      Array("OUTST_RWA_STAGE1_LCY_PRE_SHOCK", "CAPRWA", "SUM(OUTSTANDING_FOR_RWA) WHERE STAGE=1", "1", "Stage 1 RWA exposure", "From the CAPRWA extract"), _
      Array("OUTST_RWA_STAGE2_LCY_PRE_SHOCK", "CAPRWA", "SUM(OUTSTANDING_FOR_RWA) WHERE STAGE=2", "2", "Stage 2 RWA exposure", ""), _
      Array("OUTST_RWA_NPA_LCY_PRE_SHOCK", "CAPRWA", "SUM(OUTSTANDING_FOR_RWA) WHERE STAGE=3", "3", "Stage 3 RWA exposure", ""), _
      Array("RWA_CR_STAGE1_LCY_PRE_SHOCK", "CAPRWA", "SUM(RWA_LCY) WHERE STAGE=1", "1", "Stage 1 credit RWA", ""), _
      Array("RWA_CR_STAGE2_LCY_PRE_SHOCK", "CAPRWA", "SUM(RWA_LCY) WHERE STAGE=2", "2", "Stage 2 credit RWA", ""), _
      Array("RWA_CR_NPA_LCY_PRE_SHOCK", "CAPRWA", "SUM(RWA_LCY) WHERE STAGE=3", "3", "Stage 3 credit RWA", ""), _
      Array("OUTST_FOR_RWA", "CAPRWA", "SUM(OUTSTANDING_FOR_RWA)", "ALL", "Filtered RWA exposure total", ""), _
      Array("RWA_CR_LCY", "CAPRWA", "SUM(RWA_LCY)", "ALL", "Filtered credit RWA total", ""))
End Function

Private Function SeedMetricsEcl() As Variant
    'Split in two: VBA allows at most 25 line continuations in a single statement.
    Dim a As Variant, b As Variant, o As Variant, i As Long, n As Long
    a = Array( _
      Array("OUTST_LCY_STAGE1_PRE_SHOCK", "ECL", "SUM(OUTSTANDING_LCY) WHERE STAGE=1", "1", "Stage 1 outstanding", "Performing exposure before shock"), _
      Array("OUTST_LCY_STAGE2_PRE_SHOCK", "ECL", "SUM(OUTSTANDING_LCY) WHERE STAGE=2", "2", "Stage 2 outstanding", "Under-performing exposure before shock"), _
      Array("OUTST_LCY_NPA_PRE_SHOCK", "ECL", "SUM(OUTSTANDING_LCY) WHERE STAGE=3", "3", "Stage 3 / NPA outstanding", "Existing NPA exposure before shock"), _
      Array("IIS_STAGE1_PRE_SHOCK", "ECL", "SUM(INTEREST_IN_SUSPENSE_LCY) WHERE STAGE=1", "1", "Stage 1 interest in suspense", ""), _
      Array("IIS_STAGE2_PRE_SHOCK", "ECL", "SUM(INTEREST_IN_SUSPENSE_LCY) WHERE STAGE=2", "2", "Stage 2 interest in suspense", ""), _
      Array("IIS_NPA_PRE_SHOCK", "ECL", "SUM(INTEREST_IN_SUSPENSE_LCY) WHERE STAGE=3", "3", "Stage 3 interest in suspense", ""), _
      Array("OUTST_ECL_STAGE1_LCY_PRE_SHOCK", "ECL", "SUM(OUTST_FOR_ECL|OUTSTANDING_LCY) WHERE STAGE=1", "1", "Stage 1 ECL exposure", "Falls back to outstanding when OUTST_FOR_ECL is absent"), _
      Array("OUTST_ECL_STAGE2_LCY_PRE_SHOCK", "ECL", "SUM(OUTST_FOR_ECL|OUTSTANDING_LCY) WHERE STAGE=2", "2", "Stage 2 ECL exposure", ""), _
      Array("OUTST_ECL_NPA_LCY_PRE_SHOCK", "ECL", "SUM(OUTST_FOR_ECL|OUTSTANDING_LCY) WHERE STAGE=3", "3", "Stage 3 ECL exposure", ""), _
      Array("ECL_STAGE1_LCY_PRE_SHOCK", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL) WHERE STAGE=1", "1", "Stage 1 ECL", "Uses the first available ECL amount field"), _
      Array("ECL_STAGE2_LCY_PRE_SHOCK", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL) WHERE STAGE=2", "2", "Stage 2 ECL", ""))
    b = Array( _
      Array("ECL_NPA_LCY_PRE_SHOCK", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL) WHERE STAGE=3", "3", "Stage 3 ECL", ""), _
      Array("ECL_RATIO_STAGE1", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL)/SUM(OUTST_FOR_ECL|OUTSTANDING_LCY) WHERE STAGE=1", "1", "Stage 1 ECL ratio", ""), _
      Array("ECL_RATIO_STAGE2", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL)/SUM(OUTST_FOR_ECL|OUTSTANDING_LCY) WHERE STAGE=2", "2", "Stage 2 ECL ratio", ""), _
      Array("ECL_RATIO_NPA", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL)/SUM(OUTST_FOR_ECL|OUTSTANDING_LCY) WHERE STAGE=3", "3", "Stage 3 ECL ratio", ""), _
      Array("ECL_RATIO", "ECL", "SUM(ECL|OVERRIDDEN_ECL|CALCULATED_ECL)/SUM(OUTST_FOR_ECL|OUTSTANDING_LCY)", "ALL", "Total ECL ratio", ""), _
      Array("IIS_RATIO_STAGE1", "ECL", "SUM(INTEREST_IN_SUSPENSE_LCY)/SUM(OUTSTANDING_LCY) WHERE STAGE=1", "1", "Stage 1 IIS ratio", ""), _
      Array("IIS_RATIO_STAGE2", "ECL", "SUM(INTEREST_IN_SUSPENSE_LCY)/SUM(OUTSTANDING_LCY) WHERE STAGE=2", "2", "Stage 2 IIS ratio", ""), _
      Array("IIS_RATIO_NPA", "ECL", "SUM(INTEREST_IN_SUSPENSE_LCY)/SUM(OUTSTANDING_LCY) WHERE STAGE=3", "3", "Stage 3 IIS ratio", ""), _
      Array("OUTST_LCY", "ECL", "SUM(OUTSTANDING_LCY)", "ALL", "Filtered outstanding total", ""), _
      Array("OUTST_FOR_ECL", "ECL", "SUM(OUTST_FOR_ECL|OUTSTANDING_LCY)", "ALL", "Filtered ECL exposure total", ""))
    ReDim o(0 To UBound(a) + UBound(b) + 1)
    For i = 0 To UBound(a): o(n) = a(i): n = n + 1: Next i
    For i = 0 To UBound(b): o(n) = b(i): n = n + 1: Next i
    SeedMetricsEcl = o
End Function

Private Sub SeedDimensionDefaults(ByVal ws As Worksheet)
    Dim m As Variant, i As Long, r As Long
    m = Array( _
      Array("REPORT CLASSIFICATION", "JCB_IND_REPORT_CLASSIFICATION_NAME|JCB_IND_REPORT_CLASSIFICATION_CODE", "TEXT", "Direct / indirect JKB report classification"), _
      Array("ECO SECTOR CLASSIFICATION", "JCB_ECO_SECTOR_CLASSIFICATION_NAME|JCB_ECO_SECTOR_CLASSIFICATION_CODE", "TEXT", "JKB economic sector classification"), _
      Array("DATA SOURCE", "SOURCE_NAME|SOURCE_CODE", "TEXT", "Source / portfolio family"), _
      Array("IFRS STAGE", "STAGE_ID|DEAL_STAGE_ID", "STAGE", "Stage 1 / 2 / 3"), _
      Array("IFRS MODEL", "IFRS_SEGMENT_CODE|IFRS_SEGMENT2_CODE|IFRS_SEGMENT3_CODE", "TEXT", "IFRS segmentation / model"), _
      Array("ACCOUNT/DEAL NUMBER", "ACCOUNT_NUMBER|FACILITY_CODE", "TEXT", "Account or facility identifier"), _
      Array("CURRENCY", "ISO_CODE|CURRENCY_CODE", "TEXT", "Currency"), _
      Array("COUNTERPARTY CLASSIFICATION", "COUNTERPARTY_CLASSIFICATION_NAME|COUNTERPARTY_CLASSIFICATION_CODE", "TEXT", "Credit counterparty classification"), _
      Array("COUNTERPARTY CLASSIFICATION (ECL)", "COUNTERPARTY_CLASSIFICATION_NAME|COUNTERPARTY_CLASSIFICATION_CODE", "TEXT", "Preset to ECL counterparty classification"), _
      Array("COUNTERPARTY CLASSIFICATION (ALM)", "COUNTERPARTY_CLASSIFICATION_NAME|COUNTERPARTY_CLASSIFICATION_CODE", "TEXT", "Preset candidate; verify taxonomy equivalence for ALM scenarios"), _
      Array("COUNTERPARTY CODE", "CUSTOMER_CODE|COUNTERPARTY_ID|CUSTOMER_ID", "TEXT", "Customer / counterparty identifier"), _
      Array("SECTOR", "SECTOR_NAME|SECTOR_CODE", "SECTOR", "Detailed sector"), _
      Array("CBJ_TREE_ECO_SECTOR", "CBJ_TREE_ECO_SECTOR_CODE|CBJ_TREE_ECO_SECTOR_NAME", "UNDERSCORE_TEXT", "CBJ sector tree"), _
      Array("ACCOUNTING TYPE", "ACCOUNTING_TYPE_NAME|ACCOUNTING_TYPE_CODE", "TEXT", "Accounting classification"), _
      Array("GOVT FLAG", "GOV_FLAG|FLG_GOVT_PSE", "YESNO", "Government / public-sector indicator"), _
      Array("PRODUCT", "COA_NAME|COA_CODE|PRODUCT_NAME|PRODUCT_CODE", "TEXT", "Product / COA mapping"), _
      Array("FLAG WHOLESALE IMPORTERS", "FLAG_WHOLESALE_IMPORTERS", "YESNO", "Map when a source field is available"), _
      Array("BALANCE SHEET TYPE", "", "TEXT", "ALM / balance-sheet dimension; intentionally not guessed from ECL"), _
      Array("RULE IDENTIFIER", "", "TEXT", "ALM rule dimension; requires an ALM source, not ECL"), _
      Array("BALANCESHEET LINE", "", "TEXT", "Balance-sheet line; requires an explicit bank mapping"))
    For i = LBound(m) To UBound(m)
        r = PS_MAP_FIRST_ROW + i
        If r > PS_MAP_LAST_ROW Then Exit For
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then
            ws.Cells(r, 1).Value2 = m(i)(0)
            ws.Cells(r, 2).Value2 = m(i)(1)
            ws.Cells(r, 3).Value2 = m(i)(2)
            ws.Cells(r, 6).Value2 = m(i)(3)
        End If
    Next i
    SeedBreakdownColumn ws
End Sub

' The 'Break down by' answer for every mapped dimension that has not got one.
'
' Separate from the seeding above, and deliberately: that only writes rows it
' created, so on an existing workspace - which is every workspace after the first
' - the new column stayed blank on every row and the breakdown had nothing to
' break down by. A blank here means "not yet decided", and this decides it.
Private Sub SeedBreakdownColumn(ByVal ws As Worksheet)
    Dim r As Long, lastRow As Long, nm As String
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    For r = PS_MAP_FIRST_ROW To lastRow
        nm = SafeText(ws.Cells(r, 1).Value2)
        If Len(nm) > 0 Then
            If Len(SafeText(ws.Cells(r, PS_MAP_BREAK_COL).Value2)) = 0 Then
                ws.Cells(r, PS_MAP_BREAK_COL).Value2 = IIf(DefaultBreakdownDim(nm), "Yes", "No")
            End If
        End If
    Next r
    Err.Clear
End Sub

' The four dimensions a pre-shock figure is most often questioned on. Seeded so
' the breakdown answers something useful before anyone configures it, and it is
' four because that is the cap - seeding more would mean silently ignoring some.
Private Function DefaultBreakdownDim(ByVal nm As String) As Boolean
    Select Case SafeUpperText(nm)
        Case "IFRS STAGE", "SECTOR", "PRODUCT", "DATA SOURCE"
            DefaultBreakdownDim = True
    End Select
End Function

' Carries forward everything a user could have edited on the two sheets this one
' replaces, then hides them so nothing is silently lost but nothing is edited twice.
Private Sub MigrateLegacySheets(ByVal ws As Worksheet)
    Dim old As Worksheet, r As Long, target As Long, key As String, v As String, moved As Long
    Dim overrides As Object, formulas As Object, cand As Object, k As Variant

    Set overrides = NewMap(): Set formulas = NewMap(): Set cand = NewMap()

    Set old = GetWorksheetSafe(ThisWorkbook, LEGACY_CONFIG_SHEET)
    If Not old Is Nothing Then
        ' legacy test-case block: rows 9..120, override filter in column G
        For r = 9 To 120
            key = SafeUpperText(old.Cells(r, 2).Value2) & "|" & SafeUpperText(old.Cells(r, 3).Value2)
            v = SafeText(old.Cells(r, 7).Value2)
            If key <> "|" And Len(v) > 0 Then overrides(key) = v
        Next r
        ' legacy metric block: rows 125..145, formula in column C
        For r = 125 To 145
            key = SafeUpperText(old.Cells(r, 1).Value2)
            v = SafeText(old.Cells(r, 3).Value2)
            If Len(key) > 0 And Len(v) > 0 Then formulas(key) = v
        Next r
        ' legacy dimension block: rows 150..180, candidates in column B
        For r = 150 To 180
            key = SafeUpperText(old.Cells(r, 1).Value2)
            v = SafeText(old.Cells(r, 2).Value2)
            If Len(key) > 0 Then cand(key) = v
        Next r
    End If

    Set old = GetWorksheetSafe(ThisWorkbook, LEGACY_BASE_SHEET)
    If Not old Is Nothing Then
        If Len(SafeText(ws.Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2)) = 0 Then
            ws.Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2 = SafeText(old.Range("B6").Value2)
            ws.Range("B" & (PS_SRC_FIRST_ROW + 5)).Value2 = SafeText(old.Range("B7").Value2)
            If Len(SafeText(old.Range("B20").Value2)) > 0 Then ws.Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2 = SafeText(old.Range("B20").Value2)
        End If
        ' legacy hidden filter registry: X48.. key, AB.. override
        For r = 48 To 400
            key = SafeUpperText(old.Cells(r, 24).Value2)
            v = SafeText(old.Cells(r, 28).Value2)
            If Len(key) > 0 And Len(v) > 0 Then
                key = Replace(key, " | ", "|")
                If Not overrides.Exists(key) Then overrides(key) = v
            End If
        Next r
    End If

    ' apply migrated metric formulas and dimension candidates, each on its own sheet
    Set ws = PsMetricsSheet()
    For target = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        key = SafeUpperText(ws.Cells(target, 1).Value2)
        If Len(key) = 0 Then Exit For
        If formulas.Exists(key) Then
            If SafeText(ws.Cells(target, 4).Value2) <> CStr(formulas(key)) Then
                ws.Cells(target, 4).Value2 = CStr(formulas(key)): moved = moved + 1
            End If
            formulas.Remove key
        End If
    Next target
    ' any legacy metric the defaults do not cover is appended rather than dropped
    target = NextFreeRow(ws, 1, PS_METRIC_FIRST_ROW, PS_METRIC_LAST_ROW)
    For Each k In formulas.keys
        If target > PS_METRIC_LAST_ROW Then Exit For
        ws.Cells(target, 1).Value2 = CStr(k)
        ws.Cells(target, 2).Value2 = "ECL"
        ws.Cells(target, 3).Value2 = "Yes"
        ws.Cells(target, 4).Value2 = CStr(formulas(k))
        ws.Cells(target, 5).Value2 = "ALL"
        ws.Cells(target, 9).Value2 = "Migrated from " & LEGACY_CONFIG_SHEET
        target = target + 1: moved = moved + 1
    Next k

    Set ws = PsFieldsSheet()
    For target = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        key = SafeUpperText(ws.Cells(target, 1).Value2)
        If Len(key) = 0 Then Exit For
        If cand.Exists(key) Then
            If Len(CStr(cand(key))) > 0 And SafeText(ws.Cells(target, 2).Value2) <> CStr(cand(key)) Then
                ws.Cells(target, 2).Value2 = CStr(cand(key)): moved = moved + 1
            End If
        End If
    Next target

    ' overrides are re-applied by SyncPreShockTestCases, so park them for that pass
    Set mMigratedOverrides = overrides

    On Error Resume Next
    Set old = GetWorksheetSafe(ThisWorkbook, LEGACY_CONFIG_SHEET)
    If Not old Is Nothing Then old.Visible = xlSheetVeryHidden
    Set old = GetWorksheetSafe(ThisWorkbook, LEGACY_BASE_SHEET)
    If Not old Is Nothing Then old.Visible = xlSheetVeryHidden
    On Error GoTo 0
    If moved > 0 Then LogIssue LOG_LEVEL_INFO, "Pre-shock", moved & " setting(s) migrated from the previous Base_Testing / Pre_Shock_Config sheets.", PRE_SHOCK_SHEET
End Sub

Private Function NextFreeRow(ByVal ws As Worksheet, ByVal col As Long, ByVal firstRow As Long, ByVal lastRow As Long) As Long
    Dim r As Long
    For r = firstRow To lastRow
        If Len(SafeText(ws.Cells(r, col).Value2)) = 0 Then NextFreeRow = r: Exit Function
    Next r
    NextFreeRow = lastRow + 1
End Function

' Each sheet's own dropdowns. The `ws` argument is ignored - it is kept so the
' call sites did not all have to change - because validation now belongs to the
' sheet that owns the column.
Private Sub ApplyValidation(ByVal ws As Worksheet)
    On Error Resume Next
    ListValidation PreShockSheet(), "B" & (PS_SRC_FIRST_ROW + 3), "Yes,No"
    ListValidation PreShockSheet(), "F" & (PS_SRC_FIRST_ROW + 3), "Strict,Ignore"
    ListValidation PsMetricsSheet(), "C" & PS_METRIC_FIRST_ROW & ":C" & PS_METRIC_LAST_ROW, "Yes,No"
    ListValidation PsMetricsSheet(), "E" & PS_METRIC_FIRST_ROW & ":E" & PS_METRIC_LAST_ROW, "ALL,1,2,3"
    ListValidation PsFieldsSheet(), "C" & PS_MAP_FIRST_ROW & ":C" & PS_MAP_LAST_ROW, "TEXT,STAGE,YESNO,SECTOR,UNDERSCORE_TEXT"
    ListValidation PsFieldsSheet(), Chr$(64 + PS_MAP_BREAK_COL) & PS_MAP_FIRST_ROW & ":" & _
                   Chr$(64 + PS_MAP_BREAK_COL) & PS_MAP_LAST_ROW, "Yes,No"
    ListValidation PreShockSheet(), "I" & (PS_SRC_FIRST_ROW + 2), "Yes,No"
    ListValidation PsCasesSheet(), "A" & PS_CASE_FIRST_ROW & ":A" & PS_CASE_LAST_ROW, "Yes,No"
    Err.Clear
    On Error GoTo 0
End Sub

Private Function UseChoices() As String
    UseChoices = USE_DERIVED & "," & USE_SYSTEM & "," & USE_DERIVED_ELSE
End Function

Private Sub ListValidation(ByVal ws As Worksheet, ByVal addr As String, ByVal items As String)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    With ws.Range(addr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=items
    End With
    Err.Clear
End Sub

' ========================= block 4: test case sync ==========================

Public Sub SyncPreShockTestCasesFromCurrentSource()
    Dim a As Variant, h As Object, hr As Long, stats As Object, idx As Object
    EnsurePreShockSheet
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    SyncPreShockTestCases idx
End Sub

' Rewrites the test-case block from the current upload while preserving every
' user-owned column (enabled flag, override filter, notes).
Public Sub SyncPreShockTestCases(ByVal idx As Object)
    Dim ws As Worksheet, keep As Object, x As Object, r As Long, outRow As Long
    Dim tk As Variant, ek As Variant, tc As Object, el As Object
    Dim key As String, typ As String, sysFilter As String, isEcl As Boolean
    Dim block As Variant, n As Long

    EnsurePreShockSheet
    Set ws = PsCasesSheet()
    Set keep = NewMap()

    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
        If key <> "|" Then
            Set x = NewMap()
            x("Enabled") = SafeText(ws.Cells(r, 1).Value2)
            x("Override") = SafeText(ws.Cells(r, 7).Value2)
            x("Notes") = SafeText(ws.Cells(r, 12).Value2)
            Set keep(key) = x
        End If
    Next r
    ' first run after the merge: carry the overrides the old sheets held
    If Not mMigratedOverrides Is Nothing Then
        For Each tk In mMigratedOverrides.keys
            If Not keep.Exists(CStr(tk)) Then
                Set x = NewMap(): x("Enabled") = "Yes": x("Override") = CStr(mMigratedOverrides(tk)): x("Notes") = ""
                Set keep(CStr(tk)) = x
            ElseIf Len(CStr(keep(CStr(tk))("Override"))) = 0 Then
                keep(CStr(tk))("Override") = CStr(mMigratedOverrides(tk))
            End If
        Next tk
        Set mMigratedOverrides = Nothing
    End If

    n = PS_CASE_LAST_ROW - PS_CASE_FIRST_ROW + 1
    ReDim block(1 To n, 1 To 12)
    outRow = 0
    For Each tk In idx.keys
        Set tc = idx(tk)
        For Each ek In tc("Elements").keys
            outRow = outRow + 1
            If outRow > n Then
                Err.Raise vbObjectError + 735, , "The pre-shock test-case block holds " & n & " rows and this upload needs more. " & _
                    "Raise PS_CASE_LAST_ROW (and move the results block down) before generating."
            End If
            Set el = tc("Elements")(ek)
            typ = SafeText(el("ElementType"))
            isEcl = IsEclPreShockType(typ)
            sysFilter = PreShockCleanFilter(SafeText(el("FilterExpression")))
            key = SafeUpperText(tc("TestCaseCode")) & "|" & SafeUpperText(el("ScenarioCode"))

            block(outRow, 1) = "Yes"
            block(outRow, 2) = tc("TestCaseCode")
            block(outRow, 3) = el("ScenarioCode")
            block(outRow, 4) = typ
            ' Base needs no filter, so ECL can test every test case. The element-type hint
            ' only says whether the SHOCK itself is an ECL-driven one.
            block(outRow, 5) = IIf(isEcl, "ECL (shock + base)", "ECL (base only)")
            block(outRow, 6) = sysFilter
            block(outRow, 7) = ""
            block(outRow, 9) = "Upload ECL to validate"
            block(outRow, 12) = tc("TestCaseName") & _
                IIf(Len(SafeText(el("ScenarioName"))) > 0 And SafeText(el("ScenarioName")) <> SafeText(tc("TestCaseName")), " | " & SafeText(el("ScenarioName")), "")
            If keep.Exists(key) Then
                If Len(CStr(keep(key)("Enabled"))) > 0 Then block(outRow, 1) = keep(key)("Enabled")
                block(outRow, 7) = keep(key)("Override")
                If Len(CStr(keep(key)("Notes"))) > 0 Then block(outRow, 12) = keep(key)("Notes")
            End If
            ' Effective filter is what the engine actually runs: the override when present.
            block(outRow, 8) = IIf(Len(CStr(block(outRow, 7))) > 0, block(outRow, 7), sysFilter)
        Next ek
    Next tk

    ws.Range("A" & PS_CASE_FIRST_ROW & ":L" & PS_CASE_LAST_ROW).ClearContents
    If outRow > 0 Then WriteLiteralValues ws.Cells(PS_CASE_FIRST_ROW, 1).Resize(n, 12), block
    RefreshMetricConsumers PsMetricsSheet()
    If Not mEclHeaders Is Nothing Then RefreshPreShockAvailability
End Sub

' ================== block 2 <-> config sheet cross-reference ================

' Fills "Used by element types" so the link between a stored pre-shock formula and
' the configuration rows that consume it is visible on the sheet, and flags metrics
' nothing refers to.
Public Sub RefreshMetricConsumers(ByVal ws As Worksheet)
    Dim rules As Object, rl As Collection, ri As Object, k As Variant
    Dim byMetric As Object, r As Long, metric As String, s As String
    On Error GoTo Done
    Set rules = LoadRulesByElementType(ThisWorkbook)
    Set byMetric = NewMap()
    For Each k In rules.keys
        Set rl = rules(k)
        For Each ri In rl
            metric = SafeUpperText(ri("OutputRowLabel"))
            If Len(metric) > 0 Then
                If Not byMetric.Exists(metric) Then byMetric(metric) = CStr(k) Else byMetric(metric) = CStr(byMetric(metric)) & "; " & CStr(k)
            End If
        Next ri
    Next k
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        metric = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(metric) = 0 Then Exit For
        If byMetric.Exists(metric) Then
            s = CStr(byMetric(metric))
            If Len(s) > 240 Then s = Left$(s, 237) & "..."
            ws.Cells(r, 8).Value2 = s
        Else
            ws.Cells(r, 8).Value2 = "Not referenced by any enabled config row"
        End If
    Next r
Done:
End Sub

' ======================== availability / status pass ========================

' Refreshes the three "can the loaded ECL file actually do this?" columns.
'
' Each block is its own procedure so a failure names the block it came from instead of
' surfacing as one opaque error over a 70-line routine.
Public Sub RefreshPreShockAvailability()
    Dim stage As String
    On Error GoTo Failed
    stage = "preparing the workspace"
    EnsurePreShockSheet
    stage = "dimension availability"
    RefreshDimensionAvailability PsFieldsSheet()
    stage = "formula availability"
    RefreshFormulaAvailability PsMetricsSheet()
    stage = "metric consumers"
    RefreshMetricConsumers PsMetricsSheet()
    If mEclHeaders Is Nothing Then Exit Sub
    stage = "test-case filter status"
    RefreshCaseStatus PsCasesSheet()
    Exit Sub
Failed:
    Err.Raise Err.Number, "RefreshPreShockAvailability", "While refreshing " & stage & ": " & Err.description
End Sub

Private Sub RefreshDimensionAvailability(ByVal ws As Worksheet)
    Dim r As Long, candidates As String, cols As Collection, seen As Object, rr As Long, col As Long
    Dim keep As String, switched As Boolean, er As Long, erText As String, noEcl As Boolean
    ' The Fields sheet describes the ECL extract, so it is checked against that
    ' extract, whichever file was loaded last.
    If SourceLoaded("ECL") Then
        If SafeUpperText(mActiveSource) <> "ECL" Then
            keep = mActiveSource
            switched = ActivateSource("ECL")
        End If
    ElseIf AnySourceLoaded() Then
        noEcl = True
    End If
    On Error GoTo Restore
    For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then Exit For
        candidates = SafeText(ws.Cells(r, 2).Value2)
        If Len(candidates) = 0 Then
            ws.Cells(r, 4).Value2 = "Needs explicit mapping": ws.Cells(r, 5).Value2 = ""
        ElseIf mEclHeaders Is Nothing Or noEcl Then
            ws.Cells(r, 4).Value2 = "Upload ECL to validate": ws.Cells(r, 5).Value2 = ""
        Else
            Set cols = ResolveFilterColumns(candidates)
            If cols.count = 0 Then
                ws.Cells(r, 4).Value2 = "Missing in ECL": ws.Cells(r, 5).Value2 = ""
            Else
                ws.Cells(r, 4).Value2 = "Available"
                ' Counted once per loaded table and remembered.
                ws.Cells(r, 5).Value2 = DistinctValueCount(CLng(cols.item(1)))
            End If
        End If
    Next r
Restore:
    er = Err.Number: erText = Err.Description
    On Error GoTo 0
    If switched And Len(keep) > 0 Then ActivateSource keep
    If er <> 0 Then Err.Raise er, "RefreshDimensionAvailability", erText
End Sub

Private Sub RefreshFormulaAvailability(ByVal ws As Worksheet)
    Dim r As Long, formula As String, missing As String, matches As Object, m As Object, fld As String
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then Exit For
        formula = SafeText(ws.Cells(r, 4).Value2)
        missing = ""
        If Len(formula) = 0 Then
            ws.Cells(r, 7).Value2 = "No formula"
        ElseIf mEclHeaders Is Nothing Then
            ws.Cells(r, 7).Value2 = "Upload the source to validate"
        Else
            Set matches = RegexExecute("(?:SUM|AVG|MIN|MAX|COUNT)\s*\(\s*(GREATEST\s*\([^()]*\)|[A-Za-z0-9_ |*]+)\s*\)", formula)
            For Each m In matches
                fld = Trim$(CStr(m.SubMatches(0)))
                If fld <> "*" Then
                    If MeasureColumnFor(fld) = 0 Then AppendUnsupported missing, fld
                End If
            Next m
            If Len(missing) = 0 Then ws.Cells(r, 7).Value2 = "Ready" Else ws.Cells(r, 7).Value2 = "Missing: " & missing
        End If
    Next r
End Sub

Private Sub RefreshCaseStatus(ByVal ws As Worksheet)
    Dim r As Long, expr As String, node As Object, unsupported As String, fm As Object, caseStatus As String
    Set fm = LoadEclFilterMap()
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
        expr = EffectiveFilterAt(ws, r)
        unsupported = "": caseStatus = "Filter maps - base and pre-shock testable"
        If Len(expr) = 0 Then
            ' Still useful: base is unfiltered, so it is tested regardless.
            caseStatus = "No filter condition - base only"
        Else
            On Error Resume Next
            Set node = CompileFilter(expr, fm, unsupported)
            If Err.Number <> 0 Then caseStatus = "Filter not understood: " & Err.description: Err.Clear
            On Error GoTo 0
            If Len(unsupported) > 0 Then caseStatus = "Unmapped field(s): " & unsupported & " - base only"
        End If
        ws.Cells(r, 9).Value2 = caseStatus
    Next r
End Sub

' Resolves the filter a row will actually run with - the override when present, otherwise
' the system filter from the upload - and writes it back into the Effective filter column
' so the sheet always shows what the engine used.
'
' The result goes through a local: inside a Function, the function's own name on the RIGHT
' of an expression is a recursive call, not the return value, which here would mean calling
' a two-argument function with none (run-time error 450).
Private Function EffectiveFilterAt(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim ov As String, effective As String, current As Variant
    ov = SafeText(ws.Cells(r, 7).Value2)
    If Len(ov) > 0 Then
        effective = PreShockCleanFilter(ov)
    Else
        effective = PreShockCleanFilter(SafeText(ws.Cells(r, 6).Value2))
    End If
    ' Written only when it differs: this runs for every test case on every run,
    ' and each write is a recalculation and a repaint.
    current = ws.Cells(r, 8).Value2
    If VarType(current) <> vbString Then
        ws.Cells(r, 8).Value2 = effective
    ElseIf current <> effective Then
        ws.Cells(r, 8).Value2 = effective
    End If
    EffectiveFilterAt = effective
End Function

Private Function IsEclPreShockType(ByVal elementType As String) As Boolean
    Select Case SafeUpperText(elementType)
        Case "SPECIFIED PORTFOLIO SEGMENT MOVES FROM PERFORMING TO NPA", "SPECIFIC CUSTOMERS BECOMES NPA", _
             "INCREASE IN NPA PERCENTAGE OF SEGMENT", "EXCHANGE RATES CHANGING"
            IsEclPreShockType = True
    End Select
End Function

Private Function PreShockCleanFilter(ByVal s As String) As String
    s = Replace(s, "_x000D_", " "): s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Replace(s, " OR(", " OR ("): s = Replace(s, " AND(", " AND (")
    PreShockCleanFilter = Trim$(s)
End Function

' ========================== reading the sheet back ==========================

' The enabled, non-empty formula rows. This is the only place metric formulas are read.
' metric -> { Formula, Stage, Source }. Source decides which extract the formula is run
' against, so one rule table can span ECL, CAPRWA and whatever comes next.
Private Function LoadMetricMap() As Object
    Dim d As Object, ws As Worksheet, r As Long, x As Object, key As String, formula As String, src As String
    Set d = NewMap(): EnsurePreShockSheet: Set ws = PsMetricsSheet()
    For r = PS_METRIC_FIRST_ROW To PS_METRIC_LAST_ROW
        key = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(key) = 0 Then Exit For
        If SafeUpperText(ws.Cells(r, 3).Value2) <> "NO" Then
            formula = SafeText(ws.Cells(r, 4).Value2)
            src = NormalizeSourceKey(SafeText(ws.Cells(r, 2).Value2))
            If Len(formula) > 0 Then
                Set x = NewMap()
                x("Formula") = formula
                x("Stage") = SafeUpperText(ws.Cells(r, 5).Value2)
                x("Source") = src
                x("Use") = modValueSources.VS_MetricUse(key, "PRE")
                x("UseBase") = modValueSources.VS_MetricUse(key, "BASE")
                Set d(key) = x
            End If
        End If
    Next r
    Set LoadMetricMap = d
End Function

' ---------------------------------------------------------------------------
'  Where a figure should come from.
'
'  The whole point of this tool is that the independent column is rebuilt from
'  the raw extract rather than copied from the system's own output. But there are
'  two honest reasons to want the system's number in there instead: a metric
'  whose formula has not been written yet, and a metric no extract can supply.
'  Before this, both were served by quietly writing systemoutput.X into the
'  config - which looks identical to a real check and proves nothing.
'
'  So the choice is stated, per metric, on the rule row, and it travels into the
'  generated workbook beside the figure. Both numbers are always computed and
'  always shown; the setting only decides which one derived.<metric> points at.
' ---------------------------------------------------------------------------
Public Function NormalizeUse(ByVal s As String) As String
    Dim t As String
    t = SafeUpperText(s)
    If Len(t) = 0 Then NormalizeUse = DefaultValueSource(): Exit Function
    If InStr(t, "ELSE") > 0 Then NormalizeUse = USE_DERIVED_ELSE: Exit Function
    If Left$(t, 3) = "SYS" Then NormalizeUse = USE_SYSTEM: Exit Function
    If Left$(t, 3) = "DER" Then NormalizeUse = USE_DERIVED: Exit Function
    If Left$(t, 3) = "MAN" Then NormalizeUse = USE_DERIVED: Exit Function
    NormalizeUse = DefaultValueSource()
End Function

' The whole-workspace default, for flipping every metric at once without editing
' five hundred rows. A metric row that names its own choice always wins.
Public Function DefaultValueSource() As String
    ' One switch: the pre-shock default on Config_ValueSources (C7).
    DefaultValueSource = modValueSources.VS_DefaultUse("PRE")
End Function

' Given both numbers and the setting, which one derived.<metric> resolves to.
Public Function PickValue(ByVal use As String, ByVal derivedVal As Variant, ByVal hasDerived As Boolean, _
                          ByVal systemVal As Variant, ByRef usedOut As String) As Variant
    Select Case use
        Case USE_SYSTEM
            usedOut = USE_SYSTEM
            PickValue = systemVal
        Case USE_DERIVED_ELSE
            If hasDerived Then
                usedOut = USE_DERIVED
                PickValue = derivedVal
            Else
                usedOut = USE_SYSTEM
                PickValue = systemVal
            End If
        Case Else
            usedOut = USE_DERIVED
            PickValue = derivedVal
    End Select
End Function

' Maps whatever sits in the Source cell onto a registered source key. Anything unrecognised
' falls back to ECL rather than silently addressing a source that does not exist - which is
' what happened when a v2 sheet was read at v3 offsets and every rule came back sourced
' from "YES", leaving every ECL-backed metric blank.
Public Function NormalizeSourceKey(ByVal s As String) As String
    Dim k As Variant, t As String
    t = SafeUpperText(s)
    If Len(t) = 0 Then NormalizeSourceKey = "ECL": Exit Function
    For Each k In SourceKeys()
        If SafeUpperText(k) = t Then NormalizeSourceKey = SafeUpperText(k): Exit Function
    Next k
    NormalizeSourceKey = "ECL"
End Function

' The metrics one source supplies, so aggregates can be built per source in one pass.
Private Function MetricsForSource(ByVal metricMap As Object, ByVal sourceKey As String) As Object
    Dim d As Object, k As Variant
    Set d = NewMap()
    For Each k In metricMap.keys
        If SafeUpperText(metricMap(k)("Source")) = SafeUpperText(sourceKey) Then Set d(CStr(k)) = metricMap(k)
    Next k
    Set MetricsForSource = d
End Function

Private Function LoadEclFilterMap() As Object
    Dim d As Object, ws As Worksheet, r As Long, x As Object, key As String, k As Variant
    Set d = NewMap(): EnsurePreShockSheet: Set ws = PsFieldsSheet()
    For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        key = SafeUpperText(ws.Cells(r, 1).Value2)
        If Len(key) = 0 Then Exit For
        Set x = NewMap(): x("Candidates") = SafeText(ws.Cells(r, 2).Value2): x("Mode") = SafeUpperText(ws.Cells(r, 3).Value2)
        If Len(x("Mode")) = 0 Then x("Mode") = "TEXT"
        Set d(key) = x
    Next r
    ' Every physical ECL field is addressable directly too; logical dimensions stay preferred.
    If Not mEclHeaders Is Nothing Then
        For Each k In mEclHeaders.keys
            If Not d.Exists(SafeUpperText(CStr(k))) Then
                Set x = NewMap(): x("Candidates") = CStr(k): x("Mode") = GuessEclNormalization(CStr(k))
                Set d(SafeUpperText(CStr(k))) = x
            End If
        Next k
    End If
    Set LoadEclFilterMap = d
End Function

Private Function GuessEclNormalization(ByVal fieldName As String) As String
    Dim s As String: s = SafeUpperText(fieldName)
    If InStr(1, s, "STAGE", vbTextCompare) > 0 Then GuessEclNormalization = "STAGE": Exit Function
    If InStr(1, s, "FLAG", vbTextCompare) > 0 Or Left$(s, 4) = "FLG_" Then GuessEclNormalization = "YESNO": Exit Function
    GuessEclNormalization = "TEXT"
End Function

' testcase|element -> effective filter, for every enabled ECL-backed row
Private Function LoadCaseFilters() As Object
    Dim d As Object, ws As Worksheet, r As Long, key As String
    Set d = NewMap(): EnsurePreShockSheet: Set ws = PsCasesSheet()
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        key = SafeUpperText(ws.Cells(r, 2).Value2) & "|" & SafeUpperText(ws.Cells(r, 3).Value2)
        If key = "|" Then Exit For
        If SafeUpperText(ws.Cells(r, 1).Value2) <> "NO" Then d(key) = EffectiveFilterAt(ws, r)
    Next r
    Set LoadCaseFilters = d
End Function

Private Function CaseRowFor(ByVal ws As Worksheet, ByVal tcCode As String, ByVal elCode As String) As Long
    Dim r As Long
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
        If SafeUpperText(ws.Cells(r, 2).Value2) = SafeUpperText(tcCode) And _
           SafeUpperText(ws.Cells(r, 3).Value2) = SafeUpperText(elCode) Then CaseRowFor = r: Exit Function
    Next r
End Function

Public Function PreShockEnabled() As Boolean
    ' Always on: turning it off turned off what the tool is for.
    PreShockEnabled = True
End Function

Public Function IgnoreAsOfDate() As Boolean
    ' Set on the Inputs sheet (Pre_Shock_Sources J8); see modBasePreShock.
    IgnoreAsOfDate = (modBasePreShock.AsOfMatchingSetting() = "IGNORE")
End Function

' Setters, so nothing outside this module has to know which cell a switch lives in.
' A caller that hard-codes an address keeps working after a layout change and
' silently writes to the wrong cell, which is exactly what happened to the harness.
Public Sub SetAsOfMatching(ByVal mode As String)
    EnsurePreShockSheet
    modBasePreShock.WriteAsOfMatching mode
End Sub

Public Sub SetIndependentChecking(ByVal onOff As String)
    Dim ws As Worksheet
    EnsurePreShockSheet
    Set ws = PreShockSheet()
    If ws Is Nothing Then Exit Sub
    ws.Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2 = onOff
End Sub


' ========================= loading the ECL output ==========================
'
'  The ECL extract is a third-party export: the sheet set, the column set, the
'  column order and the header row all vary between runs. Nothing here assumes a
'  fixed shape. Every sheet is scored independently and a sheet that cannot be
'  read is skipped rather than failing the workbook - only "no sheet at all
'  qualified" is an error, and it says what it looked for and what it found.

Public Sub UploadECLSource()
    UploadSource "ECL"
End Sub

Public Sub UploadCAPRWASource()
    UploadSource "CAPRWA"
End Sub

' Pick any number of extracts at once; each one is identified by its contents.
'
' The per-source buttons remain, because testing ONE framework should not mean
' loading five files. This is for the ordinary case of having all of them.
Public Sub UploadAllSources()
    Dim v As Variant, path As String, key As String, ok As Long, report As String, completedFiles As Long
    Dim why As String, taken As Object
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Choose the extracts to load - any number, in any order"
        .filters.Clear
        .filters.Add "Excel files", "*.xlsx; *.xlsm; *.xlsb; *.xls"
        .AllowMultiSelect = True
        If Not .Show = -1 Then Exit Sub
        ProgressStart "Load source extracts", "Identifying and loading the selected input files", .SelectedItems.count
        Set taken = NewMap()
        For Each v In .SelectedItems
            path = CStr(v)
            ProgressStep "Loading " & FileLeaf(path), IIf(completedFiles > 0, 1, 0)
            key = IdentifyInputFile(path, why)
            If Len(key) = 0 Then
                ' Said, with the reason, here and in Build_Log.
                report = report & vbCrLf & "   " & FileLeaf(path) & "  ->  not loaded: " & why
                LogIssue LOG_LEVEL_WARN, "Load inputs", "Not loaded: " & why, path
            Else
                On Error Resume Next
                LoadEclOutput path, False, key
                If Err.Number <> 0 Then
                    report = report & vbCrLf & "   " & FileLeaf(path) & "  ->  " & key & " FAILED: " & Err.description
                    Err.Clear
                ElseIf Len(mLoadFailure) > 0 Then
                    ' LoadEclOutput logs its own failures rather than raising them.
                    report = report & vbCrLf & "   " & FileLeaf(path) & "  ->  " & key & " FAILED: " & mLoadFailure
                Else
                    report = report & vbCrLf & "   " & FileLeaf(path) & "  ->  " & key & "  (" & _
                             format$(SourceRowCount(key), "#,##0") & " rows)"
                    ok = ok + 1
                    ' Two files that are the same input: the later one is used, as
                    ' it always was, but that is now said.
                    If taken.Exists(key) Then
                        why = "Loaded as " & key & " in place of " & FileLeaf(CStr(taken(key))) & _
                              ", which was chosen too and looked like the same input."
                        report = report & ", in place of " & FileLeaf(CStr(taken(key)))
                        LogIssue LOG_LEVEL_WARN, "Load inputs", why, path
                    End If
                    taken(key) = path
                End If
                On Error GoTo 0
            End If
            completedFiles = completedFiles + 1
        Next v
    End With
    ProgressDone ok & " source extract(s) loaded."
    UiNotice "Load inputs", ok & " input file(s) loaded.", report, _
             IIf(ok > 0, "Check the scenarios and test cases (step 2), then run the tests.", "Choose the input files again."), (ok = 0)
End Sub

' Which source an unopened file is, from its own columns, or "".
Public Function IdentifySourceKey(ByVal path As String) As String
    Dim why As String
    IdentifySourceKey = IdentifyInputFile(path, why)
End Function

' The same, and when the file is not an input, why not.
Public Function IdentifyInputFile(ByVal path As String, ByRef why As String) As String
    Dim wb As Workbook, pick As Object, oldSec As Variant, opened As Boolean
    why = ""
    oldSec = Application.AutomationSecurity
    On Error Resume Next: Application.AutomationSecurity = 3: On Error GoTo Failed
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    opened = True
    Set pick = PickInputSheet(wb)
    IdentifyInputFile = CStr(pick("Key"))
    why = CStr(pick("Why"))
Finish:
    On Error Resume Next
    If opened Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSec
    Err.Clear
    Exit Function
Failed:
    why = "could not be " & IIf(opened, "read", "opened") & ": " & Err.description
    IdentifyInputFile = ""
    Resume Finish
End Function

' The one check every load path makes: Load inputs, Load a folder, the upload
' button for each source and the harness. So they can never disagree about a
' file. The best-scoring sheet is used (a tie goes to a sheet named CURRENT,
' otherwise to the first), and the input is named from that sheet's own header
' row.
'   Found   a sheet looks like one of the extracts
'   Sheet, HeaderRow, Score, Report   as ChooseEclSheet gives them
'   Key     ECL, CAPRWA, LL, LCR, NSFR or CAP; "" when it is none of them
'   Why     why Key is empty
'   Tied    the other sheets that scored the same as the one used, if any
'   Warn    what a load of this sheet should say in Build_Log, or "": the
'           sheet is hidden, another sheet looked just as likely, or a sheet
'           could not be read, so the choice may be the wrong one
Public Function PickInputSheet(ByVal wb As Workbook) As Object
    Dim res As Object, ws As Worksheet, why As String, rep As String, warn As String
    Set res = ChooseEclSheet(wb)
    res("Key") = ""
    res("Why") = ""
    If CBool(res("Found")) Then
        Set ws = res("Sheet")
        On Error Resume Next
        res("Key") = SourceKeyOfSheet(ws, CLng(res("HeaderRow")), why)
        If Err.Number <> 0 Then res("Key") = "": why = "could not be read (" & Err.description & ")": Err.Clear
        On Error GoTo 0
        If Len(why) > 0 Then res("Why") = "sheet '" & ws.name & "' " & why
        If ws.Visible <> xlSheetVisible Then warn = "sheet '" & ws.name & "' is hidden"
        If Len(CStr(res("Tied"))) > 0 Then
            If Len(warn) > 0 Then warn = warn & "; "
            warn = warn & CStr(res("Tied")) & " looked just as likely as '" & ws.name & "'"
        End If
        If Len(CStr(res("Unreadable"))) > 0 Then
            If Len(warn) > 0 Then warn = warn & "; "
            warn = warn & CStr(res("Unreadable")) & " could not be read, so could not be compared"
        End If
    Else
        rep = Trim$(Replace(Replace(CStr(res("Report")), vbCrLf, "; "), "  ", " "))
        If Len(rep) > 400 Then rep = Left$(rep, 400) & "..."
        res("Why") = "no sheet looks like an ECL, CAPRWA, ALM (LL, LCR, NSFR) or capital component extract (" & rep & ")"
    End If
    res("Warn") = warn
    Set PickInputSheet = res
End Function

' Which input a sheet is, from its header row, or "" with why not.
'
' The order: an ALM cashflow extract, then the capital component report, then
' RWA columns, then ECL. The CAPRWA extract carries a stage as well, so it is
' tested before ECL. ECL needs a stage, an outstanding AND an ECL column. Both
' rules are the ones Load a folder has always used, proven on the bank's
' files; a looser ECL rule would let another staged extract replace the ECL.
'
' LL and LCR are the SAME 61-column schema, so the columns alone cannot separate
' them - the file itself says which it is, in ALM_FRAMEWORK_NAME, and that is
' read rather than guessed from the file name. A file called LCR_group holding
' legal liquidity would otherwise be loaded into the wrong slot and every
' liquidity figure would be quietly wrong.
Public Function SourceKeyOfSheet(ByVal ws As Worksheet, ByVal hr As Long, Optional ByRef why As String) As String
    Const FRAMEWORK_ROWS As Long = 5000
    Dim a As Variant, h As Object, c As Long, lastCol As Long, lastRow As Long, v As String
    Dim fw As String, col As Long, r As Long
    why = ""
    If ws Is Nothing Or hr < 1 Then why = "has no header row": Exit Function
    lastCol = ws.Cells(hr, ws.columns.count).End(xlToLeft).Column
    If lastCol > 4096 Then lastCol = 4096
    If lastCol < 2 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = ws.Cells(hr, 1).Value2
        lastCol = 1
    Else
        a = ws.Range(ws.Cells(hr, 1), ws.Cells(hr, lastCol)).Value2
    End If
    Set h = NewMap()
    For c = 1 To lastCol
        v = CanonicalEclHeader(a(1, c))
        If Len(v) > 0 Then If Not h.Exists(v) Then h(v) = c
        v = NormalHeader(a(1, c))
        If Len(v) > 0 Then If Not h.Exists(v) Then h(v) = c
    Next c

    If h.Exists("CASHFLOW_AMOUNT_LCY_PRE_FACTOR") Or h.Exists("CASHFLOW_AMOUNT_LCY_POST_FACTOR") Then
        If Not h.Exists("ALM_FRAMEWORK_NAME") Then
            why = "is an ALM cashflow extract with no ALM_FRAMEWORK_NAME column, so it cannot say whether it is LL, LCR or NSFR"
            Exit Function
        End If
        col = CLng(h("ALM_FRAMEWORK_NAME"))
        lastRow = ws.Cells(ws.rows.count, col).End(xlUp).row
        If lastRow > hr + FRAMEWORK_ROWS Then lastRow = hr + FRAMEWORK_ROWS
        If lastRow = hr + 1 Then
            fw = SafeUpperText(ws.Cells(hr + 1, col).Value2)
            If fw = "NULL" Then fw = ""
        ElseIf lastRow > hr + 1 Then
            a = ws.Range(ws.Cells(hr + 1, col), ws.Cells(lastRow, col)).Value2
            For r = 1 To UBound(a, 1)
                fw = SafeUpperText(a(r, 1))
                If Len(fw) > 0 And fw <> "NULL" Then Exit For
                fw = ""
            Next r
        End If
        If Len(fw) = 0 Then
            why = "is an ALM cashflow extract, but ALM_FRAMEWORK_NAME is empty in its first " & FRAMEWORK_ROWS & _
                  " rows, so it cannot say whether it is LL, LCR or NSFR"
        ElseIf InStr(1, fw, "NSFR", vbTextCompare) > 0 Or InStr(1, fw, "STABLE FUNDING", vbTextCompare) > 0 Then
            SourceKeyOfSheet = "NSFR"
        ElseIf InStr(1, fw, "LCR", vbTextCompare) > 0 Then
            SourceKeyOfSheet = "LCR"
        Else
            SourceKeyOfSheet = "LL"
        End If
        Exit Function
    End If
    If h.Exists("CAPITAL_ELEMENT") And h.Exists("REPORT_BALANCE") Then SourceKeyOfSheet = "CAP": Exit Function
    If h.Exists("OUTSTANDING_FOR_RWA") Or h.Exists("RWA_LCY") Then SourceKeyOfSheet = "CAPRWA": Exit Function
    If h.Exists("STAGE") And h.Exists("OUTSTANDING_LCY") And h.Exists("ECL") Then SourceKeyOfSheet = "ECL": Exit Function
    If h.Exists("REPORT_BALANCE") Then
        why = "has REPORT_BALANCE but no CAPITAL_ELEMENT, which the capital component file needs"
    ElseIf h.Exists("STAGE") Then
        why = "has a stage column, but the ECL output needs both an outstanding and an ECL column, and it has no RWA column"
    Else
        why = "has none of the columns that mark an ECL, CAPRWA, ALM (LL, LCR, NSFR) or capital component extract in its header row (row " & hr & ")"
    End If
End Function

Private Function FileLeaf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then FileLeaf = Mid$(p, i + 1) Else FileLeaf = p
End Function

' One per source, because Shape.OnAction silently REFUSES a macro name with an
' argument - it reads back empty and the button does nothing, with no error.
Public Sub UploadLLSource()
    UploadSource "LL"
End Sub

Public Sub UploadLCRSource()
    UploadSource "LCR"
End Sub

Public Sub UploadCAPSource()
    UploadSource "CAP"
End Sub

' One upload path for every extract. The source key decides which columns are looked for
' and which metric rules can then be evaluated.
Public Sub UploadSource(ByVal sourceKey As String)
    Dim chosen As Variant
    If JKB_Busy Then Exit Sub
    EnsurePreShockSheet
    chosen = Application.GetOpenFilename("Excel workbooks (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", , _
             "Upload " & sourceKey & " output")
    If VarType(chosen) = vbBoolean Then Exit Sub
    LoadEclOutput CStr(chosen), True, sourceKey
End Sub

' Loads an ECL extract from a known path. Split out from the dialog so the whole
' path can be exercised without one.
Public Sub LoadEclOutput(ByVal chosen As String, Optional ByVal interactive As Boolean = True, _
                         Optional ByVal sourceKey As String = "ECL")
    Dim wb As Workbook, state As Object, opened As Boolean, er As String, oldSecurity As Variant
    Dim pick As Object, rows As Long, ws As Worksheet, stage As String, srcRow As Long
    Dim looks As String
    mLoadFailure = ""
    On Error GoTo Failed
    srcRow = SourceRow(sourceKey)
    stage = "preparing"
    Set state = CaptureState(): JKB_Busy = True
    Application.ScreenUpdating = False: Application.EnableEvents = False
    ReturnActiveTable
    mEclData = Empty: Set mEclHeaders = Nothing: Set mPreShockMeasureFields = Nothing
    ForgetActiveTable
    oldSecurity = Application.AutomationSecurity
    On Error Resume Next: Application.AutomationSecurity = 3: On Error GoTo Failed
    stage = "opening the ECL workbook"
    Set wb = Workbooks.Open(CStr(chosen), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    opened = True

    stage = "identifying the ECL detail sheet"
    Set pick = PickInputSheet(wb)
    If Not CBool(pick("Found")) Then
        Err.Raise vbObjectError + 710, , "No ECL detail sheet could be identified in this workbook." & vbCrLf & vbCrLf & _
            "A usable sheet needs a stage field, an outstanding field and an ECL amount field somewhere in its first 50 rows. " & _
            "Sheet names, column order and extra sheets do not matter." & vbCrLf & vbCrLf & "What was found:" & vbCrLf & CStr(pick("Report"))
    End If

    Set ws = pick("Sheet")
    ' Loaded as asked, but a file that looks like a different input is said, and
    ' so is a sheet choice that may be the wrong one.
    looks = CStr(pick("Key"))
    If looks <> SafeUpperText(sourceKey) Then
        LogIssue LOG_LEVEL_WARN, "Load inputs", "Loaded as " & sourceKey & " from sheet '" & ws.name & "', but its columns look like " & _
                 IIf(Len(looks) > 0, "the " & looks & " extract", "none of the inputs: " & CStr(pick("Why"))) & ".", CStr(chosen)
    End If
    If Len(CStr(pick("Warn"))) > 0 Then
        LogIssue LOG_LEVEL_WARN, "Load inputs", "Sheet '" & ws.name & "' was read as the " & sourceKey & " extract, but " & _
                 CStr(pick("Warn")) & ".", CStr(chosen)
    End If
    stage = "reading sheet '" & ws.name & "'"
    LoadEclSheet ws, CLng(pick("HeaderRow")), wb.date1904, CStr(chosen)
    rows = 0: If IsArray(mEclData) Then rows = UBound(mEclData, 1) - 1

    stage = "recording the source details"
    ' Keep the loaded table under its own source key so several extracts can be held at once.
    StoreActiveSource sourceKey

    With PreShockSheet()
        .Range("C" & PS_SRC_FIRST_ROW).Value2 = LoadedSourceSummary()
        .Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2 = "Yes"
        .Range("I" & (PS_SRC_FIRST_ROW + 3)).NumberFormat = "@"
        .Range("I" & (PS_SRC_FIRST_ROW + 3)).Value2 = EclDateSummary()
        .Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2 = CStr(pick("Report"))
        .Range("B" & (PS_SRC_FIRST_ROW + 5)).Value2 = ws.name
        .Range("C" & (PS_SRC_FIRST_ROW + 1)).Value2 = "Checking is ON"
        .Range("C" & (PS_SRC_FIRST_ROW + 2)).Value2 = "Ready - click Run tests"
    End With
    ' The registry itself lives on its own sheet now.
    If srcRow > 0 Then
        With PsSourcesSheet()
            .Cells(srcRow, 3).Value2 = CStr(chosen)
            .Cells(srcRow, 4).Value2 = ws.name
            .Cells(srcRow, 5).Value2 = rows
            .Cells(srcRow, 7).NumberFormat = "@"
            .Cells(srcRow, 7).Value2 = EclDateSummary()
            .Cells(srcRow, 8).Value2 = "Loaded " & format$(Now, "dd-mmm hh:mm")
        End With
    End If
    LogIssue LOG_LEVEL_INFO, "ECL pre-shock", "Loaded " & rows & " ECL rows from sheet '" & ws.name & "'.", CStr(chosen)

    stage = "closing the ECL workbook"
    If opened Then wb.Close SaveChanges:=False: opened = False
    On Error Resume Next: Application.AutomationSecurity = oldSecurity: On Error GoTo Failed
    stage = "caching dimension values for the filter builder"
    BuildValueCache
    stage = "checking field availability"
    RefreshPreShockAvailability
    stage = "styling the sheet"
    StylePreShockSheet PreShockSheet()
    RestoreState state: JKB_Busy = False
    If interactive Then RefreshEclTests
    Exit Sub
Failed:
    er = "Error " & Err.Number & " while " & stage & ": " & Err.description
    mLoadFailure = er
    On Error Resume Next
    If opened Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    LogIssue LOG_LEVEL_ERROR, "ECL pre-shock", er, CStr(chosen)
    PreShockSheet().Range("C" & PS_SRC_FIRST_ROW).Value2 = "Failed - " & er
    RestoreState state: JKB_Busy = False
    On Error GoTo 0
    If interactive Then UiProblem "Load ECL output", "The ECL file could not be loaded.", er
End Sub

' Scores every sheet and returns the best candidate plus a readable account of the rest.
Private Function ChooseEclSheet(ByVal wb As Workbook) As Object
    Dim res As Object, ws As Worksheet, hr As Long, score As Long
    Dim best As Long, bestHr As Long, bestWs As Worksheet, report As String, why As String
    Dim scores As Object, k As Variant, tied As String, unreadable As String
    Set res = NewMap(): Set scores = NewMap()
    For Each ws In wb.Worksheets
        hr = 0: score = 0: why = ""
        On Error Resume Next
        score = ScoreEclSheet(ws, hr, why)
        If Err.Number <> 0 Then
            why = "not readable (" & Err.description & ")": score = 0: Err.Clear
            If Len(unreadable) > 0 Then unreadable = unreadable & ", "
            unreadable = unreadable & "'" & ws.name & "'"
        End If
        On Error GoTo 0
        If Len(report) > 0 Then report = report & vbCrLf
        report = report & "  " & ws.name & ": " & IIf(score >= 30, "usable (score " & score & ", header row " & hr & ")", "skipped - " & why)
        If score >= 30 Then
            scores(ws.name) = score
            ' A tie is broken towards the current-period sheet; otherwise first best wins.
            If score > best Or (score = best And InStr(1, ws.name, "CURRENT", vbTextCompare) > 0) Then
                best = score: bestHr = hr: Set bestWs = ws
            End If
        End If
    Next ws
    ' Sheets that scored the same as the one used are named, so a load can say so.
    If Not bestWs Is Nothing Then
        For Each k In scores.keys
            If CLng(scores(k)) = best And StrComp(CStr(k), bestWs.name, vbTextCompare) <> 0 Then
                If Len(tied) > 0 Then tied = tied & ", "
                tied = tied & "'" & CStr(k) & "'"
            End If
        Next k
    End If
    res("Tied") = tied
    res("Unreadable") = unreadable
    res("Found") = (Not bestWs Is Nothing)
    Set res("Sheet") = bestWs
    res("HeaderRow") = bestHr
    res("Score") = best
    res("Report") = report
    Set ChooseEclSheet = res
End Function

' Which family of extract a header map belongs to, or "" for none of them.
'
' Each family is named by the columns its metrics actually READ, so a file that
' matches can be aggregated and a file that does not would have produced blanks.
Public Function SourceFamilyOf(ByVal h As Object) As String
    If h Is Nothing Then Exit Function
    ' ALM first: an ALM extract carries an account number and a stage-shaped
    ' nothing, and its cashflow columns belong to no other family.
    If h.Exists("CASHFLOW_AMOUNT_LCY_PRE_FACTOR") Or h.Exists("CASHFLOW_AMOUNT_LCY_POST_FACTOR") Then
        SourceFamilyOf = "ALM": Exit Function
    End If
    If h.Exists("CAPITAL_ELEMENT") And h.Exists("REPORT_BALANCE") Then
        SourceFamilyOf = "CAP": Exit Function
    End If
    If h.Exists("STAGE") Then
        If h.Exists("OUTSTANDING_LCY") Or h.Exists("ECL") Then SourceFamilyOf = "ECL": Exit Function
        If h.Exists("OUTSTANDING_FOR_RWA") Or h.Exists("RWA_LCY") Then SourceFamilyOf = "CAPRWA": Exit Function
    End If
End Function

Private Function ScoreEclSheet(ByVal ws As Worksheet, ByRef bestRow As Long, ByRef why As String) As Long
    Dim a As Variant, r As Long, c As Long, h As Object, s As String
    Dim score As Long, best As Long, lastCol As Long, lastRow As Long, lastCell As Range, missing As String
    Set lastCell = ws.Cells.Find(what:="*", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then why = "empty sheet": Exit Function
    lastRow = lastCell.row
    lastCol = ws.Cells.Find(what:="*", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If lastCol > 4096 Then lastCol = 4096
    If lastRow < 2 Then why = "no data rows": Exit Function
    a = ws.Range(ws.Cells(1, 1), ws.Cells(WorksheetFunction.Min(50, lastRow), lastCol)).Value2
    If Not IsArray(a) Then why = "no usable header area": Exit Function
    For r = 1 To UBound(a, 1)
        Set h = NewMap()
        For c = 1 To UBound(a, 2)
            s = CanonicalEclHeader(a(r, c))
            If Len(s) > 0 Then If Not h.Exists(s) Then h(s) = c
        Next c
        score = 0
        ' A row scores for the family it belongs to. Scoring every sheet against
        ' the ECL trio is what made an ALM file - 223,000 rows of perfectly good
        ' cashflow - look like "no recognisable columns".
        Select Case SourceFamilyOf(h)
            Case "ECL"
                score = 30
                If h.Exists("OUTST_FOR_ECL") Then score = score + 3
                If h.Exists("INTEREST_IN_SUSPENSE_LCY") Then score = score + 2
                If h.Exists("JCB_IND_REPORT_CLASSIFICATION_NAME") Then score = score + 2
            Case "CAPRWA"
                score = 30
                If h.Exists("OUTSTANDING_FOR_RWA") Then score = score + 3
                If h.Exists("RWA_LCY") Then score = score + 3
            Case "ALM"
                score = 30
                If h.Exists("ALM_FRAMEWORK_NAME") Then score = score + 5
                If h.Exists("BUCKET_DISPLAY_NAME") Then score = score + 3
                If h.Exists("COA_BALANCESHEET_CATEGORY") Then score = score + 3
            Case "CAP"
                score = 30
                If h.Exists("CAP_COMPONENT_CODE") Then score = score + 3
                If h.Exists("BRANCH_CODE") Then score = score + 2
        End Select
        If h.Exists("AS_OF_DATE") Then score = score + 3
        If score > best Then best = score: bestRow = r
    Next r
    If best < 30 Then
        why = "no sheet here looks like an ECL, CAPRWA, ALM or capital component extract"
        bestRow = 0
        Exit Function
    End If
    ScoreEclSheet = best
End Function

Public Function CanonicalEclHeader(ByVal v As Variant) As String
    Dim s As String: s = NormalHeader(v)
    ' Deliberately conservative. Guessing an alias wrong would silently rebuild a
    ' pre-shock figure from the wrong column, which is worse than reporting it missing.
    Select Case s
        Case "AOD", "ASOFDATE", "ASOF_DATE", "AS_OF_DT", "REPORT_DATE", "REPORTING_DATE": s = "AS_OF_DATE"
        Case "STAGE", "STAGE_ID", "DEAL_STAGE_ID", "IFRS_STAGE", "IFRS9_STAGE", "IFRS_STAGE_ID", "CURRENT_STAGE_ID": s = "STAGE"
        Case "OUTSTANDING", "OUTSTANDING_AMOUNT", "OUTSTANDING_LCY", "OUTST_LCY", "CURRENT_OUTSTANDING_LCY", "OUTSTANDING_BALANCE_LCY": s = "OUTSTANDING_LCY"
        Case "OUTST_FOR_ECL", "OUTSTANDING_FOR_ECL", "ECL_EXPOSURE", "EXPOSURE_FOR_ECL": s = "OUTST_FOR_ECL"
        Case "ECL", "ECL_LCY", "EXPECTED_CREDIT_LOSS", "ECL_AMOUNT": s = "ECL"
        Case "IIS", "IIS_LCY", "INTEREST_IN_SUSPENSE", "INTEREST_IN_SUSPENSE_LCY": s = "INTEREST_IN_SUSPENSE_LCY"
    End Select
    CanonicalEclHeader = s
End Function

' Keeps every physical column addressable by its own name as well as by the
' canonical name, so a formula or filter can name any field the extract happens
' to carry without the tool needing to know about it in advance.
' Loads the columns the derivation actually consumes, and only those.
'
' The extract is 227 columns wide and the checks use about thirty of them.
' Reading the whole used range meant materialising sixteen million variants -
' a quarter of a gigabyte - to sum a dozen. Now the header row is read first,
' the wanted columns are resolved from the configuration, and each is read on its
' own. Same numbers, a fraction of the work.
'
' Every field NAME is still recorded, so the condition builder can offer any
' column the extract carries; only the DATA of unused columns is skipped. A
' condition written against a field that was not loaded is caught at compile time
' and reported as a line-level dependency, exactly as a missing field always was.
Private Sub LoadEclSheet(ByVal ws As Worksheet, ByVal hr As Long, ByVal date1904 As Boolean, ByVal path As String)
    Dim lastCell As Range, lastRow As Long, lastCol As Long, c As Long, k As String, raw As String
    Dim wanted As Object, cols() As Long, nWanted As Long, i As Long, n As Long
    Dim colData As Variant, r As Long, headerText() As String

    Set lastCell = ws.Cells.Find(what:="*", After:=ws.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Err.Raise vbObjectError + 711, , "The selected ECL sheet is empty."
    lastRow = lastCell.row
    lastCol = ws.Cells(hr, ws.columns.count).End(xlToLeft).Column
    If lastRow <= hr Then Err.Raise vbObjectError + 715, , "Sheet '" & ws.name & "' has a header row but no data rows below it."

    ' The header row alone, which is cheap and tells us everything we need to
    ' decide what else to read.
    ReDim headerText(1 To lastCol)
    colData = ws.Range(ws.Cells(hr, 1), ws.Cells(hr, lastCol)).Value2
    For c = 1 To lastCol
        If IsArray(colData) Then headerText(c) = SafeText(colData(1, c)) Else headerText(c) = SafeText(colData)
    Next c
    Set mEclAllFields = NewMap()
    For c = 1 To lastCol
        If Len(headerText(c)) > 0 Then mEclAllFields(headerText(c)) = c
    Next c

    Set wanted = WantedEclColumns(headerText, lastCol)
    ReDim cols(1 To WorksheetFunction.Max(wanted.count, 1))
    For c = 1 To lastCol
        If wanted.Exists(CStr(c)) Then
            nWanted = nWanted + 1
            cols(nWanted) = c
        End If
    Next c
    If nWanted = 0 Then Err.Raise vbObjectError + 717, , "Sheet '" & ws.name & "' carries none of the fields the derivation needs."

    n = lastRow - hr + 1
    ReturnActiveTable
    ReDim mEclData(1 To n, 1 To nWanted)
    ForgetActiveTable
    Set mEclHeaders = NewMap()
    For i = 1 To nWanted
        c = cols(i)
        ProgressDetail "column " & i & " of " & nWanted & " - " & headerText(c)
        colData = ws.Range(ws.Cells(hr, c), ws.Cells(lastRow, c)).Value2
        If IsArray(colData) Then
            For r = 1 To n
                mEclData(r, i) = colData(r, 1)
            Next r
        Else
            mEclData(1, i) = colData
        End If
        raw = NormalHeader(headerText(c))
        If Len(raw) > 0 Then If Not mEclHeaders.Exists(raw) Then mEclHeaders(raw) = i
        k = CanonicalEclHeader(headerText(c))
        If Len(k) > 0 Then If Not mEclHeaders.Exists(k) Then mEclHeaders(k) = i
    Next i

    ' Identified by what the sheet IS, not by which button was pressed.
    '
    ' Five different extracts load through here now and they have nothing in
    ' common: an ALM cashflow file has no stage and no ECL, a capital component
    ' file has neither and no account number either. Demanding the ECL trio of all
    ' of them rejected three of the five. Demanding nothing would accept a
    ' spreadsheet of somebody's lunch order, so each family states its own
    ' fingerprint and a sheet has to match one of them.
    If Len(SourceFamilyOf(mEclHeaders)) = 0 Then
        Err.Raise vbObjectError + 712, , "Sheet '" & ws.name & "' does not look like any extract this tool reads. " & _
            "An ECL or CAPRWA extract needs a stage column with outstanding or RWA; an ALM extract needs " & _
            "a cashflow pre-factor amount; a capital component file needs a capital element and a report balance."
    End If
    mEclPath = path: mEclSheet = ws.name: mEclHeaderRow = hr: mEclDate1904 = date1904
End Sub

' Which physical columns to read: the ones the metric formulas sum, the ones the
' dimension map points at, the ones any test-case filter names, and the handful
' the engine always needs.
Private Function WantedEclColumns(ByRef headerText() As String, ByVal lastCol As Long) As Object
    Dim want As Object, byName As Object, c As Long, k As Variant, ws As Worksheet, r As Long
    Dim m As Object, matches As Object, mt As Object, fld As String, p As Variant

    Set want = NewMap()
    Set byName = NewMap()
    For c = 1 To lastCol
        If Len(headerText(c)) > 0 Then
            byName(NormalHeader(headerText(c))) = c
            byName(CanonicalEclHeader(headerText(c))) = c
        End If
    Next c

    If SourceFamilyOf(byName) = "ECL" Then
        For c = 1 To lastCol
            If Len(headerText(c)) > 0 Then want(CStr(c)) = True
        Next c
    End If

    ' Always: the fields the engine itself reads.
    '
    ' BRANCH_CODE is in this list because the entity scope is applied to every
    ' source: Jordan is everything except branch 800 and Cyprus is only branch
    ' 800, and a column that is not read cannot be filtered on.
    For Each k In Array("STAGE", "OUTSTANDING_LCY", "ECL", "AS_OF_DATE", "OUTSTANDING_FOR_ECL", "IIS", _
                        "OUTSTANDING_FOR_RWA", "RWA_LCY", "IFRS_STAGE_ID", "ACCOUNT_NUMBER", "COA_CODE", _
                        "BRANCH_CODE", "BANK_CODE", "BANK_ID", "CAP_CONSOLIDATION_ID")
        TakeColumn want, byName, CStr(k)
    Next k

    ' The ALM cashflow extracts and the capital component file.
    For Each k In Array("CASHFLOW_AMOUNT_LCY_PRE_FACTOR", "CASHFLOW_AMOUNT_LCY_POST_FACTOR", _
                        "CASHFLOW_AMOUNT_LCY_WITH_SEGMENT", "CASHFLOW_AMOUNT_LCY_WITHOUT_SEGMENT", _
                        "CASHFLOW_AMOUNT_LCY_WITH_BUCKET", "CASHFLOW_AMOUNT_LCY_WITHOUT_BUCKET", _
                        "ALM_FRAMEWORK_NAME", "ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY", _
                        "ALM_PORTFOLIO_SEGMENTATION_RULE_CODE", "ALM_PORTFOLIO_SEGMENTATION_RULE_NAME", _
                        "COA_BALANCESHEET_CATEGORY", "BUCKET_DISPLAY_NAME", "BALANCESHEET_LINE_NAME", _
                        "RULE_IDENTIFIER_CODE", "COUNTERPARTY_CLASSIFICATION_CODE", "PRODUCT_TYPE", _
                        "CURRENCY_NAME", "ALM_FACTOR_PCT", _
                        "CAPITAL_ELEMENT", "REPORT_BALANCE", "SIGN", "CAP_COMPONENT_CODE", "CAP_COMPONENT_NAME")
        TakeColumn want, byName, CStr(k)
    Next k

    ' Whatever the metric rules sum.
    Set m = LoadMetricMap()
    For Each k In m.keys
        Set matches = RegexExecute("(?:SUM|AVG|MIN|MAX|COUNT)\s*\(\s*(GREATEST\s*\([^()]*\)|[A-Za-z0-9_ |*]+)\s*\)", CStr(m(k)("Formula")))
        For Each mt In matches
            fld = Trim$(CStr(mt.SubMatches(0)))
            If fld <> "*" And Len(fld) > 0 Then
                For Each p In Split(GreatestFlatten(fld), "|")
                    TakeColumn want, byName, Trim$(CStr(p))
                Next p
            End If
        Next mt
    Next k

    ' Whatever the dimension map points at - these are what filters are written on.
    Set ws = PsFieldsSheet()
    If Not ws Is Nothing Then
        For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
            If Len(SafeText(ws.Cells(r, 1).Value2)) = 0 Then Exit For
            For Each p In Split(SafeText(ws.Cells(r, 2).Value2), "|")
                TakeColumn want, byName, Trim$(CStr(p))
            Next p
        Next r
    End If

    ' And whatever the test-case filters actually name, so a hand-written filter
    ' on an unmapped field still works.
    Set ws = PsCasesSheet()
    If Not ws Is Nothing Then
        For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
            If Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then Exit For
            TakeFilterFields want, byName, SafeText(ws.Cells(r, 6).Value2)
            TakeFilterFields want, byName, SafeText(ws.Cells(r, 7).Value2)
        Next r
    End If

    Set WantedEclColumns = want
End Function

Private Sub TakeColumn(ByVal want As Object, ByVal byName As Object, ByVal fieldName As String)
    Dim k As String
    If Len(Trim$(fieldName)) = 0 Then Exit Sub
    k = NormalHeader(fieldName)
    If byName.Exists(k) Then want(CStr(byName(k))) = True
    k = CanonicalEclHeader(fieldName)
    If byName.Exists(k) Then want(CStr(byName(k))) = True
End Sub

' Pulls the field names out of a filter without compiling it - the compiler needs
' a loaded table, and this runs before there is one.
Private Sub TakeFilterFields(ByVal want As Object, ByVal byName As Object, ByVal expr As String)
    Dim matches As Object, mt As Object
    If Len(Trim$(expr)) = 0 Then Exit Sub
    Set matches = RegexExecute("([A-Za-z0-9_ ]+?)\s+(?:NOT\s+)?IN\s*\(", PreShockCleanFilter(expr))
    For Each mt In matches
        TakeColumn want, byName, Trim$(CStr(mt.SubMatches(0)))
    Next mt
    Set matches = RegexExecute("([A-Za-z0-9_ ]+?)\s*(?:>=|<=|<>|=|>|<)\s*[0-9']", PreShockCleanFilter(expr))
    For Each mt In matches
        TakeColumn want, byName, Trim$(CStr(mt.SubMatches(0)))
    Next mt
End Sub

Private Function EclDateSummary() As String
    Dim d As Object, r As Long, k As Variant, s As String, col As Long
    If mEclHeaders Is Nothing Then Exit Function
    If Not mEclHeaders.Exists("AS_OF_DATE") Then EclDateSummary = "(no AS_OF_DATE column)": Exit Function
    col = mEclHeaders("AS_OF_DATE"): Set d = NewMap()
    For r = 2 To UBound(mEclData, 1)
        On Error Resume Next
        d(DateKey(mEclData(r, col), mEclDate1904)) = True
        Err.Clear
        On Error GoTo 0
        If d.count > 12 Then Exit For
    Next r
    For Each k In d.keys
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(k)
    Next k
    EclDateSummary = s
End Function

' Re-reads the saved ECL file when the in-memory copy is gone (for example after the
' workbook was reopened). Raises only if it genuinely cannot be reloaded.
Private Sub EnsureEclLoaded(Optional ByVal fromRegistry As Boolean = True)
    Dim ws As Worksheet, path As String, sheetName As String, wb As Workbook, pick As Object, hr As Long, score As Long, dummy As String
    If Not IsEmpty(mEclData) Then Exit Sub
    EnsurePreShockSheet
    ' Held in memory already: use it rather than reading the file again.
    If fromRegistry And SourceLoaded("ECL") Then
        If ActivateSource("ECL") Then Exit Sub
    End If
    ' The input registry is where every load records the file. The overview cell
    ' is the older place and, after an interactive load, holds the detection notes.
    If fromRegistry Then EclSavedLocation path, sheetName
    If Len(path) = 0 Then
        With PreShockSheet()
            path = SafeText(.Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2)
            sheetName = SafeText(.Range("B" & (PS_SRC_FIRST_ROW + 5)).Value2)
        End With
    End If
    If Len(path) = 0 Then Err.Raise vbObjectError + 713, , "No ECL output has been loaded yet. Click Load inputs on the home screen."
    If Not LooksLikeFilePath(path) Then Err.Raise vbObjectError + 713, , "No input files are loaded in this session. Click Load inputs on the home screen."
    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 716, , "The saved ECL file is no longer at " & path & ". Upload it again."
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    On Error GoTo CloseAndFail
    Set ws = GetWorksheetSafe(wb, sheetName)
    If Not ws Is Nothing Then
        hr = 0: score = ScoreEclSheet(ws, hr, dummy)
        If score < 30 Then Set ws = Nothing
    End If
    If ws Is Nothing Then
        Set pick = ChooseEclSheet(wb)
        If Not CBool(pick("Found")) Then Err.Raise vbObjectError + 714, , "The saved ECL file no longer contains a usable detail sheet. Upload it again."
        Set ws = pick("Sheet"): hr = CLng(pick("HeaderRow"))
    End If
    LoadEclSheet ws, hr, wb.date1904, path
    wb.Close False
    Exit Sub
CloseAndFail:
    Dim er As String: er = Err.description
    On Error Resume Next: wb.Close False: On Error GoTo 0
    Err.Raise vbObjectError + 714, , er
End Sub

' ====================== running the independent checks ======================

Public Sub RefreshEclTests()
    Dim a As Variant, h As Object, hr As Long, idx As Object, stats As Object, cache As Object, state As Object
    Dim er As String, notes As Collection, msg As String, repointed As Long
    If JKB_Busy Then Exit Sub
    On Error GoTo Failed
    EnsurePreShockSheet
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    SyncPreShockTestCases idx
    Set state = CaptureState(): JKB_Busy = True
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Set cache = PrepareBaseTestCache(idx, Nothing, a, h, notes)
    With PreShockSheet()
        .Range("C" & (PS_SRC_FIRST_ROW + 2)).value = "Last run " & format$(Now, "dd-mmm hh:mm") & " - " & cache.count & " derived values"
    End With
    ' Having derived the figures, make the configuration actually use them. Without this the
    ' independent column keeps copying the system output, which is the whole complaint: the
    ' numbers were being rebuilt and then thrown away.
    repointed = MigratePreShockConfigTokens(False)
    StylePreShockSheet PreShockSheet()
    RestoreState state: JKB_Busy = False
    msg = "Base and pre-shock rebuilt from the uploaded extract(s)." & vbCrLf & vbCrLf & _
          cache.count & " derived value(s), stored on the " & DERIVED_SHEET & " sheet."
    If repointed > 0 Then
        msg = msg & vbCrLf & vbCrLf & repointed & " configuration row(s) were switched from copying the system output" & vbCrLf & _
              "to reading the derived value. Generated sheets will now reference " & DERIVED_SHEET & "."
    End If
    If Not notes Is Nothing Then
        If notes.count > 0 Then msg = msg & vbCrLf & vbCrLf & notes.count & " test case(s) could not be checked - see the Results and Mapping status columns."
    End If
    MsgBox msg, vbInformation, "JKB Stress Testing - Run tests"
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    LogIssue LOG_LEVEL_ERROR, "ECL pre-shock", er, PRE_SHOCK_SHEET
    RestoreState state: JKB_Busy = False
    On Error GoTo 0
    UiProblem "Run tests", "The independent checks could not run.", er
End Sub

' Builds the testcase/metric -> independent value cache used by the generator.
'
' This never raises. An ECL problem is a problem with the INDEPENDENT CHECK, not with
' the configured manual calculation, so it must not stop a reconciliation run: the
' affected rows simply fall back to their configured manual formula. Everything that
' went wrong is returned in `notes` and logged under its own "ECL pre-shock" category,
' separate from the configuration errors that do stop a run.
' Derives BASE and PRE-SHOCK independently from the ECL output and compares both against
' what the system produced.
'
'   BASE      the bank's whole-portfolio total for a measure, with NO filter applied.
'             It is the same for every test case at a given as-of date, so it is computed
'             once and reused. It is what the Base column of a generated sheet holds.
'
'   PRE-SHOCK the same measure restricted to the test case's filter condition. It depends
'             only on the filter, so it is identical across Moderate / Medium / Severe and
'             differs per test case / element.
'
' Shock and post-shock are left alone: the configured manual formulas already derive those
' from the pre-shock figures, which is exactly what this feeds them.
'
' Never raises. An ECL problem degrades to "use the configured formula" and is reported
' through `notes`, separately from configuration errors that stop a run.
Public Function PrepareBaseTestCache(ByVal idx As Object, ByVal selections As Object, _
                                     ByRef sourceData As Variant, ByVal sourceHeaders As Object, _
                                     Optional ByRef notes As Collection) As Object
    Dim result As Object, aggregateCache As Object, metricMap As Object, filterMap As Object
    Dim caseFilters As Object, ws As Worksheet
    Dim tk As Variant, ek As Variant, tc As Object, el As Object, mk As Variant
    Dim expr As String, preAgg As Object, baseAgg As Object, metric As String
    Dim preEcl As Variant, baseEcl As Variant, preFound As Boolean, baseFound As Boolean
    Dim preSys As Variant, baseSys As Variant, status As String, j As Long
    Dim sev As Variant, selected As Boolean, key As String, matched As Double, caseRow As Long
    Dim block As Collection, line As Variant, out As Variant, i As Long
    Dim sevRow As Long, col As Long, tested As Long
    Dim baseBySource As Object, preBySource As Object, srcK As Variant, metricSrc As String
    Dim srcMeasure As Object, srcFilterMap As Object, liveSources As Collection, rec As Object
    Dim matchedBySource As Object, matchRows As Double
    Dim matchSummary As String, metricStatus As String
    Dim wk As Variant
    Dim useMode As String, usedBase As String, usedPre As String
    Dim baseOut As Variant, preOut As Variant
    Dim srcBrkDims As Object, brkRows As Collection, completedCases As Long
    Dim caseRows As Object, screenWas As Boolean, screenTaken As Boolean, disabledWhy As String

    ' Every run derives current evidence. Saved figures are snapshots, never inputs.
    Set result = NewMap()
    Set mJoinCache = Nothing
    ClearSourceDateCache
    If notes Is Nothing Then Set notes = New Collection
    Set PrepareBaseTestCache = result

    On Error GoTo Disabled
    EnsurePreShockSheet
    Set ws = PsResultsSheet()
    If Not PreShockEnabled() Then
        Note notes, "Independent checking is switched off on the " & PRE_SHOCK_SHEET & " sheet (row 9)."
        Exit Function
    End If
    ' Reopen a remembered extract when needed; never use an unvalidated prior result.
    If Not AnySourceLoaded() Then
        If result.count = 0 Then EnsureEclLoaded False
    End If
    On Error GoTo 0

    Set aggregateCache = NewMap()
    modValueSources.VS_ResetCache
    Set metricMap = LoadMetricMap()
    Set caseFilters = LoadCaseFilters()
    ' Each test case's row on the Cases sheet, read once rather than scanned
    ' cell by cell for every element.
    Set caseRows = CaseRowIndex(PsCasesSheet())
    Set block = New Collection
    sev = SeverityList()

    ' Per-source scaffolding is built ONCE, not once per test case per source. Rebuilding the
    ' measure map and the field map inside the loop meant re-reading the rule block and the
    ' whole 227-column header set several hundred times for no new information; with two
    ' sources loaded that alone accounted for minutes of the run.
    Set srcMeasure = NewMap(): Set srcFilterMap = NewMap(): Set liveSources = New Collection
    Set srcBrkDims = NewMap(): Set brkRows = New Collection
    For Each srcK In SourceKeys()
        If SourceLoaded(CStr(srcK)) Then
            If ActivateSource(CStr(srcK)) Then
                liveSources.Add CStr(srcK)
                Set srcMeasure(CStr(srcK)) = BuildPreShockMeasureMap(MetricsForSource(metricMap, CStr(srcK)))
                Set srcFilterMap(CStr(srcK)) = LoadEclFilterMap()
                ' Which dimensions and measures this source will be broken down
                ' by, resolved against the columns it actually carries. Once per
                ' source, never per row.
                Set mPreShockMeasureFields = srcMeasure(CStr(srcK))
                PrepareBreakdown
                Set srcBrkDims(CStr(srcK)) = SnapshotBreakdownPlan()
            End If
        End If
    Next srcK
    If liveSources.count = 0 Then
        If result.count > 0 Then
            Note notes, "No extract is loaded in this session, so the " & result.count & " value(s) already stored on " & _
                        DERIVED_SHEET & " were used. Load the input files and click Run on the home screen to rebuild them."
        Else
            Note notes, "No source extract is loaded and nothing is stored on " & DERIVED_SHEET & _
                        ", so nothing could be rebuilt independently."
        End If
        Exit Function
    End If

    ' Nothing on screen needs to follow the rebuild; the per-case status cells
    ' below no longer repaint one at a time.
    screenWas = Application.ScreenUpdating: screenTaken = True
    Application.ScreenUpdating = False
    For Each tk In idx.keys
        Set tc = idx(tk): selected = True
        If Not selections Is Nothing Then selected = IsSelected(tc, selections)
        If Not selected Then GoTo NextCase
        ProgressDetail "Rebuilding base and pre-shock: " & CStr(tc("TestCaseCode"))

        ' Base and pre-shock are built once per SOURCE. Each metric is then read from the
        ' aggregate of the source its rule names, so ECL supplies the outstanding / IIS /
        ' ECL measures and CAPRWA supplies the RWA ones in the same pass.
        Set baseBySource = NewMap()
        For Each srcK In liveSources
            If ActivateSource(CStr(srcK)) Then
                ProgressDetail CStr(tc("TestCaseCode")) & " | Base from " & CStr(srcK)
                Set mPreShockMeasureFields = srcMeasure(CStr(srcK))
                RestoreBreakdownPlan srcBrkDims(CStr(srcK))
                ' One aggregate per source AND row selector. A metric with no
                ' selector uses the "" slice, which is the source's whole
                ' population and exactly what was built before.
                For Each wk In SlicesForSource(metricMap, CStr(srcK))
                    On Error Resume Next
                    Set baseAgg = Nothing
                    ' Base: no test-case condition, so everything is local.
                    Set baseAgg = GetEclAggregate(aggregateCache, CStr(tc("Date")), "", _
                                      modScenarioBuilder_Multi.ApplyEntityScope(CStr(wk)), srcFilterMap(CStr(srcK)))
                    If Err.Number <> 0 Then
                        Note notes, CStr(srcK) & ": base could not be built - " & Err.description
                        Err.Clear
                    Else
                        Set baseBySource(SliceKey(CStr(srcK), CStr(wk))) = baseAgg
                    End If
                    On Error GoTo 0
                Next wk
            End If
        Next srcK

        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            key = SafeUpperText(tc("TestCaseCode")) & "|" & SafeUpperText(el("ScenarioCode"))
            caseRow = CaseRowFromIndex(caseRows, CStr(tc("TestCaseCode")), CStr(el("ScenarioCode")))

            expr = ""
            If caseFilters.Exists(key) Then expr = CStr(caseFilters(key))

            status = "": Set preBySource = NewMap()
            If Len(Trim$(expr)) = 0 Then
                status = "Missing filter definition - pre-shock not tested"
            Else
                For Each srcK In liveSources
                    If ActivateSource(CStr(srcK)) Then
                        ProgressDetail CStr(tc("TestCaseCode")) & " / " & CStr(el("ScenarioCode")) & " | Filtered pre-shock from " & CStr(srcK)
                        Set mPreShockMeasureFields = srcMeasure(CStr(srcK))
                        RestoreBreakdownPlan srcBrkDims(CStr(srcK))
                        For Each wk In SlicesForSource(metricMap, CStr(srcK))
                            On Error Resume Next
                            Set preAgg = Nothing
                            ' The test case's condition on one side, this
                            ' source's own metric selector and entity rule on
                            ' the other. Only the first may need ECL.
                            Set preAgg = GetEclAggregate(aggregateCache, CStr(tc("Date")), expr, _
                                             modScenarioBuilder_Multi.ApplyEntityScope(CStr(wk)), _
                                             srcFilterMap(CStr(srcK)))
                            If Err.Number <> 0 Then
                                If Len(status) = 0 Then status = CStr(srcK) & ": " & Err.description
                                Err.Clear
                            Else
                                Set preBySource(SliceKey(CStr(srcK), CStr(wk))) = preAgg
                            End If
                            On Error GoTo 0
                        Next wk
                    End If
                Next srcK
            End If
            ' The row count is per source: the same filter over ECL and over CAPRWA matches
            ' different numbers of rows, and reporting one source's count against the other's
            ' figure is simply wrong.
            ' Counted on each source's UNSLICED aggregate only. A source now has
            ' one aggregate per metric row selector, and adding those up would
            ' report the same rows once per selector - a population of 223,000
            ' rows presented as 900,000.
            Set matchedBySource = NewMap()
            matched = 0: matchSummary = ""
            For Each srcK In liveSources
                If preBySource.Exists(SliceKey(CStr(srcK), "")) Then
                    Set preAgg = preBySource(SliceKey(CStr(srcK), ""))
                    matchRows = AggVal(preAgg, "ROWS#1") + AggVal(preAgg, "ROWS#2") + AggVal(preAgg, "ROWS#3") + AggVal(preAgg, "ROWS#0")
                    matchedBySource(CStr(srcK)) = matchRows
                    matched = matched + matchRows
                    If Len(matchSummary) > 0 Then matchSummary = matchSummary & ", "
                    matchSummary = matchSummary & CStr(srcK) & " " & format$(matchRows, "#,##0")
                    ' A joined source says how much of the population it held.
                    ' Without this, a figure covering a third of the accounts
                    ' looks exactly like a figure covering all of them.
                    If AggVal(preAgg, "JOIN#WANTED") > 0 Then
                        matchSummary = matchSummary & " (joined via ECL: " & _
                            format$(AggVal(preAgg, "JOIN#FOUND"), "#,##0") & " of " & _
                            format$(AggVal(preAgg, "JOIN#WANTED"), "#,##0") & " account(s) found here)"
                        If AggVal(preAgg, "JOIN#FOUND") = 0 Then
                            Note notes, tc("TestCaseCode") & " / " & el("ScenarioCode") & ": none of the " & _
                                format$(AggVal(preAgg, "JOIN#WANTED"), "#,##0") & " account(s) this filter selects in ECL " & _
                                "appear in " & CStr(srcK) & ", so it contributes nothing to this test case."
                        End If
                    End If
                End If
            Next srcK

            If preBySource.count = 0 Then
                SetCaseResult PsCasesSheet(), caseRow, IIf(Len(status) > 0, status, "Pre-shock not tested"), 0
                If Len(status) > 0 Then Note notes, tc("TestCaseCode") & " / " & el("ScenarioCode") & ": " & status
            Else
                SetCaseResult PsCasesSheet(), caseRow, "Matched " & matchSummary & " row(s)", matched
                If matched = 0 Then Note notes, tc("TestCaseCode") & " / " & el("ScenarioCode") & ": the filter matched no rows"
            End If

            ' Why this test case's population is the size it is: what the filter
            ' selected, per dimension and value, beside what the whole portfolio
            ' holds of the same thing. The share between them is the answer.
            ' The unsliced aggregate again: the breakdown explains the test case's
            ' population, which is one thing per source, not one per metric.
            For Each srcK In liveSources
                If preBySource.Exists(SliceKey(CStr(srcK), "")) Then
                    ' Not IIf: it evaluates BOTH branches, and asking a Dictionary for
                    ' a key it has not got silently ADDS it - so the tidy-looking
                    ' version of this line quietly grows the map it was reading.
                    Set baseAgg = Nothing
                    If baseBySource.Exists(SliceKey(CStr(srcK), "")) Then Set baseAgg = baseBySource(SliceKey(CStr(srcK), ""))
                    CollectBreakdown brkRows, tc, el, CStr(srcK), _
                                     preBySource(SliceKey(CStr(srcK), "")), baseAgg, _
                                     srcBrkDims(CStr(srcK)), expr
                End If
            Next srcK

            ' One row per metric defined in the formulas block, whether or not the element
            ' type happens to configure it - the point is to test the numbers, not the config.
            For Each mk In metricMap.keys
                metric = CStr(mk)
                baseFound = False: preFound = False
                baseEcl = Empty: preEcl = Empty
                ' Read each metric from the aggregate of the source its rule names.
                metricSrc = SliceKey(SafeUpperText(metricMap(metric)("Source")), MetricWhere(metricMap(metric)))
                If baseBySource.Exists(metricSrc) Then baseEcl = MetricFromAggregate(baseBySource(metricSrc), metricMap(metric), baseFound)
                If preBySource.Exists(metricSrc) Then preEcl = MetricFromAggregate(preBySource(metricSrc), metricMap(metric), preFound)

                ' The system's own figures: base row for base, any severity row for pre-shock
                ' (the filter is severity-independent, so they agree by construction).
                baseSys = Empty: preSys = Empty
                col = 0
                If sourceHeaders.Exists(metric) Then col = CLng(sourceHeaders(metric))
                If col > 0 Then
                    baseSys = sourceData(tc("BaseRow"), col)
                    sevRow = 0
                    For j = 0 To 2
                        If el("SeverityRows").Exists(sev(j)) Then sevRow = el("SeverityRows")(sev(j)): Exit For
                    Next j
                    If sevRow > 0 Then preSys = sourceData(sevRow, col)
                End If

                ' The cache now carries the whole derived record, not just the pre-shock
                ' scalar: the generator writes base, pre-shock, source and provenance into the
                ' output workbook so `derived.X` can point at a cell a reviewer can trace.
                matchRows = 0
                If preBySource.Exists(metricSrc) Then
                    Set preAgg = preBySource(metricSrc)
                    matchRows = AggVal(preAgg, "ROWS#1") + AggVal(preAgg, "ROWS#2") + AggVal(preAgg, "ROWS#3") + AggVal(preAgg, "ROWS#0")
                End If
                metricStatus = ResultStatus(baseFound, preFound, status, metricSrc)
                If matchRows = 0 And Len(expr) > 0 Then metricStatus = "No eligible source rows - pre-shock not tested"

                ' Which figure derived.<metric> will resolve to. Both are always
                ' kept; the setting only decides which one the configuration
                ' reads, and the answer travels into the generated workbook so a
                ' reviewer can see that a row was taken from the system on purpose
                ' rather than because the derivation quietly failed.
                useMode = CStr(metricMap(metric)("Use"))
                usedBase = "": usedPre = ""
                baseOut = PickValue(CStr(metricMap(metric)("UseBase")), baseEcl, baseFound, baseSys, usedBase)
                preOut = PickValue(useMode, preEcl, preFound, preSys, usedPre)

                Set rec = NewMap()
                    ' The record carries its own identity so the store can be rewritten from
                    ' the cache alone - including rows carried over from an earlier run that
                    ' this selection did not recompute.
                    rec("Date") = tc("Date")
                    rec("Entity") = tc("EntityCode")
                    rec("EntityID") = tc("EntityID")
                    rec("TestCase") = tc("TestCaseCode")
                    rec("Element") = el("ScenarioCode")
                    rec("Metric") = metric
                    rec("Source") = metricSrc
                    rec("HasBase") = baseFound: rec("BaseDerived") = baseEcl
                    rec("HasPre") = preFound: rec("PreDerived") = preEcl
                    rec("BaseSystem") = baseSys: rec("PreSystem") = preSys
                    rec("Use") = useMode: rec("UseBase") = metricMap(metric)("UseBase")
                    rec("UsedBase") = usedBase: rec("UsedPre") = usedPre
                    ' Base and Pre are what the configuration reads. Everything
                    ' above is what it was chosen from.
                    rec("Base") = baseOut
                    rec("Pre") = preOut
                    rec("Rows") = matchRows
                    rec("Filter") = expr
                    rec("Status") = metricStatus
                    Set result(BaseCacheKey(tc, el, metric)) = rec
                    If baseFound Or preFound Then tested = tested + 1

                block.Add Array(tc("Date"), tc("EntityCode") & " (ID " & tc("EntityID") & ")", _
                                tc("TestCaseCode"), el("ScenarioCode"), metric, _
                                baseSys, IIf(baseFound, baseEcl, Empty), Delta(baseSys, baseEcl, baseFound), _
                                preSys, IIf(preFound, preEcl, Empty), Delta(preSys, preEcl, preFound), _
                                metricStatus)
            Next mk
        Next ek
NextCase:
        completedCases = completedCases + 1
        If ReconStrictMode And idx.count > 0 Then ProgressDetail "Source calculations: " & completedCases & " of " & idx.count & " test cases", 30# * completedCases / idx.count
    Next tk

    ' one block write rather than a cell at a time
    ws.Range("A" & PS_RESULT_FIRST_ROW & ":L" & ws.rows.count).ClearContents
    ' Merged into Derived_Values (differences and Reconciles in Q:S).
    ws.Cells(PS_RESULT_FIRST_ROW, 1).Value2 = "Merged into " & DERIVED_SHEET & _
        ": base and pre-shock from the input files and from the system, their differences and Reconciles (columns Q:S)."
    PreShockSheet().Range("C" & (PS_SRC_FIRST_ROW + 2)).Value2 = tested & " of " & block.count & " comparisons rebuilt"

    WriteDerivedValuesSheet result
    WriteBreakdownSheet brkRows
    RefreshPreShockAvailability
    RefreshStatusLine ws
    If screenTaken Then Application.ScreenUpdating = screenWas
    Set PrepareBaseTestCache = result
    Exit Function

Disabled:
    ' Any failure to get the ECL side going is recorded and the run continues without it.
    ' Kept first: On Error Resume Next below clears Err, and the log used to
    ' record "Independent checking unavailable:" with no reason after it.
    disabledWhy = Err.description
    Note notes, disabledWhy
    On Error Resume Next
    If screenTaken Then Application.ScreenUpdating = screenWas
    LogIssue LOG_LEVEL_WARN, "ECL pre-shock", "Independent checking unavailable: " & disabledWhy, PRE_SHOCK_SHEET
    On Error GoTo 0
    Set PrepareBaseTestCache = result
End Function

' ================== pointing the configuration at derived values =============
'
' The configuration used to say, for every pre-shock row:
'
'     MANUAL_FORMULA_DEFAULT = systemoutput.OUTST_LCY_STAGE1_PRE_SHOCK
'
' which copies the system's own number into the column that is supposed to check it. The
' difference column is then zero by construction and the row proves nothing.
'
' This rewrites those rows to
'
'     MANUAL_FORMULA_DEFAULT = derived.OUTST_LCY_STAGE1_PRE_SHOCK
'
' so the independent column carries the figure this tool rebuilt from the uploaded extract.
' Only rows whose metric the tool can actually derive are touched, and only where the formula
' is the degenerate self-reference; anything hand-written is left alone.
Public Function MigratePreShockConfigTokens(Optional ByVal announce As Boolean = True) As Long
    Dim lo As ListObject, h As Object, hdr As Variant, c As Long, a As Variant, r As Long
    Dim metrics As Object, label As String, cur As String, t As String, n As Long, colF As Long, colL As Long
    Dim changed As Collection, k As Variant

    On Error GoTo Failed
    Set metrics = LoadMetricMap()
    If metrics.count = 0 Then
        If announce Then UiNotice "Use derived values", "There are no derived metric rules on " & PRE_SHOCK_SHEET & ", so there is nothing to link the configuration to."
        Exit Function
    End If

    Set lo = ThisWorkbook.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set h = NewMap()
    hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    If Not h.Exists("OUTPUT_ROW_LABEL") Then Exit Function
    If Not h.Exists("MANUAL_FORMULA_DEFAULT") Then Exit Function
    colL = CLng(h("OUTPUT_ROW_LABEL")): colF = CLng(h("MANUAL_FORMULA_DEFAULT"))

    Set changed = New Collection
    a = lo.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        label = SafeUpperText(a(r, colL))
        If Len(label) > 0 Then
            If metrics.Exists(label) Then
                cur = SafeText(a(r, colF))
                t = SafeUpperText(cur)
                t = Replace(Replace(Replace(t, "@", ""), " ", ""), vbTab, "")
                If Len(t) = 0 Or t = "SYSTEMOUTPUT." & label Or t = "SYSTEM." & label Then
                    lo.DataBodyRange.Cells(r, colF).Value2 = "derived." & label
                    n = n + 1
                    If changed.count < 12 Then changed.Add label
                End If
            End If
        End If
    Next r

    MigratePreShockConfigTokens = n
    If n > 0 Then
        LogIssue LOG_LEVEL_INFO, "Configuration", n & " pre-shock row(s) now read their independent value from " & _
            DERIVED_SHEET & " instead of copying the system output.", SHEET_RULES
    End If
    If announce Then
        If n = 0 Then
            UiNotice "Use derived values", "Nothing to change.", "Every pre-shock row that can be derived already uses a derived value."
        Else
            t = ""
            For Each k In changed
                t = t & vbCrLf & "  derived." & CStr(k)
            Next k
            If n > changed.count Then t = t & vbCrLf & "  ... and " & (n - changed.count) & " more"
            UiNotice "Use derived values", n & " configuration row(s) now use the derived value instead of copying the system output:" & t, _
                     "Generated workbooks refer to the " & DERIVED_SHEET & " sheet, so each figure shows its input file, the filter used and the rows it matched."
        End If
    End If
    Exit Function
Failed:
    On Error Resume Next
    LogIssue LOG_LEVEL_WARN, "Configuration", "Could not repoint pre-shock rows at derived values: " & Err.description, SHEET_RULES
    If announce Then UiProblem "Use derived values", "The configuration could not be updated.", Err.description
End Function

' ========================= folder-upload support ============================
'
' The folder loader opens each workbook once and decides what it is; these let it
' route the result through the same registry, cache and sheet the single-file
' upload uses, so there is one path in and not two.

Public Function ScoreSheetForUpload(ByVal ws As Worksheet, ByRef hrOut As Long, ByRef whyOut As String) As Long
    ScoreSheetForUpload = ScoreEclSheet(ws, hrOut, whyOut)
End Function

' The same reader the single-file upload uses. One path in, so a folder load and
' a file load can never read the extract differently.
Public Sub LoadSheetForUpload(ByVal ws As Worksheet, ByVal hr As Long, ByVal date1904 As Boolean, ByVal path As String)
    LoadEclSheet ws, hr, date1904, path
End Sub

Public Sub StoreActiveSourceFor(ByVal sourceKey As String)
    StoreActiveSource sourceKey
End Sub

Public Sub BuildValueCacheFor(ByVal sourceKey As String)
    mActiveSource = sourceKey
    BuildValueCache
End Sub

Public Sub RecordSourceRow(ByVal sourceKey As String, ByVal path As String, ByVal sheetName As String)
    Dim ws As Worksheet, r As Long, rows As Long
    EnsurePreShockSheet
    Set ws = PsSourcesSheet()
    If ws Is Nothing Then Exit Sub
    r = SourceRow(sourceKey)
    If r = 0 Then Exit Sub
    rows = 0
    If IsArray(mEclData) Then rows = UBound(mEclData, 1) - 1
    ws.Cells(r, 3).Value2 = path
    ws.Cells(r, 4).Value2 = sheetName
    ws.Cells(r, 5).Value2 = rows
    ws.Cells(r, 7).NumberFormat = "@"
    ws.Cells(r, 7).Value2 = EclDateSummary()
    ws.Cells(r, 8).Value2 = "Loaded " & format$(Now, "dd-mmm hh:mm")
    ' The ECL path is also remembered on the home sheet, which is where a later
    ' session reloads it from when nothing has been uploaded yet.
    If SafeUpperText(sourceKey) = "ECL" Then
        With PreShockSheet()
            .Range("B" & (PS_SRC_FIRST_ROW + 4)).Value2 = path
            .Range("B" & (PS_SRC_FIRST_ROW + 5)).Value2 = sheetName
            .Range("B" & (PS_SRC_FIRST_ROW + 3)).Value2 = "Yes"
            .Range("C" & PS_SRC_FIRST_ROW).Value2 = LoadedSourceSummary()
        End With
    End If
End Sub

' ============================ trace support =================================
'
' The tracer needs the same three things the derivation engine uses: the loaded
' table, the compiled filter, and the row evaluator. Exposing them here keeps one
' implementation rather than a second, subtly different one in the tracer - which
' is the only way a trace can be trusted to show the rows the FIGURE actually used.

Public Function CurrentSourceData() As Variant
    CurrentSourceData = mEclData
End Function

Public Function CurrentSourceHeaders() As Object
    Set CurrentSourceHeaders = mEclHeaders
End Function

Public Function ActivateSourceForTrace(ByVal sourceKey As String) As Boolean
    On Error GoTo Failed
    If SourceLoaded(sourceKey) Then
        ActivateSourceForTrace = ActivateSource(sourceKey)
        Exit Function
    End If
    ' Not in memory: fall back to the saved path, which is how a trace works in a
    ' session where nothing has been uploaded yet.
    If SafeUpperText(sourceKey) = "ECL" Then
        EnsureEclLoaded
        ActivateSourceForTrace = Not IsEmpty(mEclData)
        Exit Function
    End If
    UiNotice "Trace", "The " & sourceKey & " input file is not loaded in this session.", _
             "A trace reads the actual rows, so the file has to be loaded.", "Load it on " & PS_SOURCES_SHEET & " and try again."
    Exit Function
Failed:
    UiProblem "Trace", "The " & sourceKey & " input file could not be opened.", Err.description
End Function

Public Function CompileTraceFilter(ByVal expr As String, ByRef needs As String) As Object
    Set CompileTraceFilter = CompileFilter(expr, LoadEclFilterMap(), needs)
End Function

Public Function EvalTraceFilter(ByVal node As Object, ByRef data As Variant, ByVal r As Long) As Boolean
    EvalTraceFilter = EvalFilter(node, r)
End Function

Public Function SourceForMetric(ByVal metric As String) As String
    Dim m As Object
    Set m = LoadMetricMap()
    If m.Exists(SafeUpperText(metric)) Then
        SourceForMetric = CStr(m(SafeUpperText(metric))("Source"))
    Else
        SourceForMetric = "ECL"
    End If
End Function

' The physical column a metric sums, so the trace pivot totals the right thing.
Public Function MeasureForMetric(ByVal metric As String) As String
    Dim m As Object, matches As Object, fld As String
    MeasureForMetric = "OUTSTANDING_LCY"
    If Len(metric) = 0 Then Exit Function
    Set m = LoadMetricMap()
    If Not m.Exists(SafeUpperText(metric)) Then Exit Function
    Set matches = RegexExecute("(?:SUM|AVG|MIN|MAX)\s*\(\s*([A-Za-z0-9_ |*]+)\s*\)", CStr(m(SafeUpperText(metric))("Formula")))
    If matches.count > 0 Then
        fld = Trim$(CStr(matches(0).SubMatches(0)))
        If InStr(fld, "|") > 0 Then fld = Trim$(Split(fld, "|")(0))
        If fld <> "*" And Len(fld) > 0 Then MeasureForMetric = SafeUpperText(fld)
    End If
End Function

Public Function FilterForCase(ByVal testCase As String, ByVal element As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = PsCasesSheet()
    If ws Is Nothing Then Exit Function
    r = CaseRowFor(ws, testCase, element)
    If r = 0 Then Exit Function
    FilterForCase = EffectiveFilterAt(ws, r)
End Function

' ======================== the derived-values store ==========================
'
' Base and pre-shock are the two numbers the whole exercise turns on, so they get their own
' sheet rather than living only inside a configuration block. One row per test case, element
' and metric: what the base is, what the filter reduced it to, which extract it came from and
' how many rows matched. The generator reads the same figures and reproduces the relevant
' slice inside each output workbook, which is what `derived.X` in a config formula points at.
Public Function DerivedSheet(Optional ByVal createIfMissing As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, DERIVED_SHEET)
    If ws Is Nothing And createIfMissing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = DERIVED_SHEET
    End If
    Set DerivedSheet = ws
End Function

' Writes the whole cache out, so the store is a faithful picture of every derived figure the
' tool currently holds - freshly computed and carried over alike.
Private Sub WriteDerivedValuesSheet(ByVal cache As Object)
    Dim ws As Worksheet, out As Variant, k As Variant, rec As Object, i As Long, j As Long, hdr As Variant
    On Error GoTo Failed
    Set ws = DerivedSheet(True)
    ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB"
    ws.Cells(PS_TITLE_ROW, 1).Value2 = "Base and pre-shock results"
    ws.Cells(PS_ABOUT_ROW, 1).Value2 = "Rebuilt from the uploaded extracts, independently of the system output. " & _
                            "Base is the unfiltered total; pre-shock is that total restricted to the test case's filter. " & _
                            "These are the figures a generated sheet's independent column reads through derived.<metric>."
    ws.Cells(PS_STATUS_ROW, 1).Value2 = "Written " & format$(Now, "dd-mmm-yyyy hh:mm") & "  |  " & LoadedSourceSummary() & _
                            "  |  saved evidence; each generation rebuilds from loaded source extracts"

    hdr = Array("As-of date", "Entity", "Entity ID", "Test case", "Element", "Metric", "Source", _
                "Base (derived)", "Pre-shock (derived)", "Rows matched", "Filter applied", "Status", _
                "Base (system)", "Pre-shock (system)", "Value to use", "Which one was used")
    For j = 0 To UBound(hdr): ws.Cells(DV_HEADER_ROW, j + 1).Value2 = hdr(j): Next j
    ' Text, so "2025-12-31" stays "2025-12-31" and not the serial number 46022 - the key this
    ' store is looked up by is built from that exact string.
    ws.Range(ws.Cells(DV_FIRST_ROW, 1), ws.Cells(ws.rows.count, 1)).NumberFormat = "@"

    If cache.count > 0 Then
        ReDim out(1 To cache.count, 1 To DV_COLS)
        i = 0
        For Each k In cache.keys
            Set rec = cache(k)
            i = i + 1
            out(i, 1) = rec("Date"): out(i, 2) = rec("Entity"): out(i, 3) = rec("EntityID")
            out(i, 4) = rec("TestCase"): out(i, 5) = rec("Element"): out(i, 6) = rec("Metric")
            out(i, 7) = rec("Source")
            ' Columns 8 and 9 are the DERIVED figures, always, whatever the
            ' configuration chose to read. A store whose meaning changed with a
            ' setting would not be a record of anything.
            If CBool(rec("HasBase")) Then out(i, 8) = RecOr(rec, "BaseDerived", "Base")
            If CBool(rec("HasPre")) Then out(i, 9) = RecOr(rec, "PreDerived", "Pre")
            out(i, 10) = rec("Rows"): out(i, 11) = rec("Filter"): out(i, 12) = rec("Status")
            out(i, 13) = RecOr(rec, "BaseSystem", "")
            out(i, 14) = RecOr(rec, "PreSystem", "")
            out(i, 15) = RecOr(rec, "Use", "")
            out(i, 16) = RecOr(rec, "UsedPre", "")
        Next k
        WriteLiteralValues ws.Cells(DV_FIRST_ROW, 1).Resize(cache.count, DV_COLS), out
    End If
    StyleDerivedSheet ws, cache.count
    modValueSources.VS_WriteMasterDifferences ws, DV_FIRST_ROW + cache.count - 1
    Exit Sub
Failed:
    On Error Resume Next
    LogIssue LOG_LEVEL_WARN, "Derived values", "The derived-values sheet could not be written: " & Err.description, DERIVED_SHEET
End Sub

' ===================== the breakdown store ==================================

Public Function BreakdownSheet(Optional ByVal createIfMissing As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_BREAKDOWN_SHEET)
    If ws Is Nothing And createIfMissing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = PS_BREAKDOWN_SHEET
    End If
    Set BreakdownSheet = ws
End Function

Public Function BreakdownCount() As Long
    Dim ws As Worksheet, lastRow As Long
    On Error Resume Next
    Set ws = BreakdownSheet()
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    If lastRow >= PS_FIRST_ROW Then BreakdownCount = lastRow - PS_FIRST_ROW + 1
    Err.Clear
End Function

Private Sub WriteBreakdownSheet(ByVal rows As Collection)
    Dim ws As Worksheet, out As Variant, line As Variant, i As Long, j As Long, hdr As Variant
    On Error GoTo Failed
    If rows Is Nothing Then Exit Sub
    Set ws = BreakdownSheet(True)
    ws.Cells.Clear
    ws.Range("A1").Value2 = "JKB"
    ws.Cells(PS_TITLE_ROW, 1).Value2 = "Why the pre-shock is what it is"
    ws.Cells(PS_ABOUT_ROW, 1).Value2 = "What each test case's filter selected, per dimension and value, beside what the whole " & _
        "portfolio holds of the same thing. The share between them is the answer: a filter that takes 31% of one sector and all " & _
        "of another explains its own figure. Choose the dimensions on " & PS_FIELDS_SHEET & ", in the 'Break down by' column."
    ws.Cells(PS_STATUS_ROW, 1).Value2 = "Written " & format$(Now, "dd-mmm-yyyy hh:mm") & "  |  " & rows.count & " row(s)  |  " & _
        IIf(BreakdownsEnabled(), "breakdowns are on", "BREAKDOWNS ARE OFF - switch them on beside 'Values come from' on " & PRE_SHOCK_SHEET)

    hdr = Array("As-of date", "Entity", "Test case", "Test element", "Source", "Dimension", "Value", _
                "Rows selected", "Rows in portfolio", "Selected", "In portfolio", "Share selected", _
                "Selected (2nd)", "In portfolio (2nd)", "Measure", "Measure (2nd)", "Filter applied")
    For j = 0 To UBound(hdr): ws.Cells(PS_HDR_ROW, j + 1).Value2 = hdr(j): Next j
    ws.Range(ws.Cells(PS_FIRST_ROW, 1), ws.Cells(ws.rows.count, 1)).NumberFormat = "@"

    If rows.count > 0 Then
        ReDim out(1 To rows.count, 1 To UBound(hdr) + 1)
        i = 0
        For Each line In rows
            i = i + 1
            For j = 0 To UBound(hdr): out(i, j + 1) = line(j): Next j
        Next line
        WriteLiteralValues ws.Cells(PS_FIRST_ROW, 1).Resize(rows.count, UBound(hdr) + 1), out
        ws.Range(ws.Cells(PS_FIRST_ROW, 8), ws.Cells(PS_FIRST_ROW + rows.count - 1, 11)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(PS_FIRST_ROW, 12), ws.Cells(PS_FIRST_ROW + rows.count - 1, 12)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(PS_FIRST_ROW, 13), ws.Cells(PS_FIRST_ROW + rows.count - 1, 14)).NumberFormat = "#,##0.00"
    Else
        ws.Cells(PS_FIRST_ROW, 1).Value2 = IIf(BreakdownsEnabled(), _
            "Nothing was broken down. Mark at least one dimension 'Yes' in the 'Break down by' column on " & PS_FIELDS_SHEET & ".", _
            "Breakdowns are switched off on " & PRE_SHOCK_SHEET & ".")
    End If
    StyleOne ws, UBound(hdr) + 1, UBound(hdr) + 1
    Exit Sub
Failed:
    On Error Resume Next
    LogIssue LOG_LEVEL_WARN, "Breakdown", "The breakdown sheet could not be written: " & Err.description, PS_BREAKDOWN_SHEET
End Sub

' A cache record that came from an older store has fewer keys than one built this
' session. Asking a Dictionary for a key it does not hold ADDS it as Empty, which
' quietly grows the record and hides the difference, so every optional read goes
' through here instead.
Private Function RecOr(ByVal rec As Object, ByVal key As String, ByVal fallbackKey As String) As Variant
    If rec.Exists(key) Then RecOr = rec(key): Exit Function
    If Len(fallbackKey) > 0 Then
        If rec.Exists(fallbackKey) Then RecOr = rec(fallbackKey)
    End If
End Function

' The as-of date always reads back as "yyyy-mm-dd", whatever Excel decided to store.
'
' Writing the text "2025-12-31" into a General cell makes Excel quietly convert it to the
' serial number 46022. Reading that straight back produced the key "46022|..." instead of
' "2025-12-31|..." - so every stored figure silently failed to match and the output fell back
' to copying the system value. The cell is now formatted as text on write, and anything that
' still arrives as a number is converted here.
Private Function DvDateText(ByVal v As Variant) As String
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then DvDateText = format$(v, "yyyy-mm-dd"): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then
            DvDateText = format$(CDate(CDbl(v) + IIf(ThisWorkbook.date1904, 1462, 0)), "yyyy-mm-dd")
            Exit Function
        End If
    End If
    DvDateText = SafeText(v)
End Function

' How many figures the store currently holds, without parsing them all.
Public Function PersistedDerivedCount() As Long
    Dim ws As Worksheet, lastRow As Long
    On Error Resume Next
    Set ws = DerivedSheet()
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    If lastRow >= DV_FIRST_ROW Then PersistedDerivedCount = lastRow - DV_FIRST_ROW + 1
    Err.Clear
End Function

' Reads the store back into the cache shape, so a session that has not re-uploaded the
' extracts still generates with real independent figures.
Public Function LoadPersistedDerived() As Object
    Dim ws As Worksheet, d As Object, a As Variant, r As Long, rec As Object, lastRow As Long, key As String
    Set d = NewMap()
    Set LoadPersistedDerived = d
    On Error GoTo Done
    Set ws = DerivedSheet()
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.count, 6).End(xlUp).row
    If lastRow < DV_FIRST_ROW Then Exit Function
    If SafeUpperText(ws.Cells(DV_HEADER_ROW, 3).Value2) <> "ENTITY ID" Then Exit Function
    ' A store written by an earlier build has twelve columns, not sixteen. Reading
    ' sixteen out of it would hand back four columns of whatever happened to be
    ' beside it, so the width is read off the sheet rather than assumed.
    Dim wide As Boolean, nCols As Long
    wide = (SafeUpperText(ws.Cells(DV_HEADER_ROW, 13).Value2) = "BASE (SYSTEM)")
    nCols = IIf(wide, DV_COLS, DV_COLS_LEGACY)
    a = ws.Range(ws.Cells(DV_FIRST_ROW, 1), ws.Cells(lastRow, nCols)).Value2
    For r = 1 To UBound(a, 1)
        If Len(SafeText(a(r, 6))) > 0 Then
            Set rec = NewMap()
            rec("Date") = DvDateText(a(r, 1)): rec("Entity") = SafeText(a(r, 2)): rec("EntityID") = SafeText(a(r, 3))
            rec("TestCase") = SafeText(a(r, 4)): rec("Element") = SafeText(a(r, 5))
            rec("Metric") = SafeUpperText(a(r, 6)): rec("Source") = SafeText(a(r, 7))
            rec("HasBase") = IsNumeric(a(r, 8)): rec("BaseDerived") = a(r, 8)
            rec("HasPre") = IsNumeric(a(r, 9)): rec("PreDerived") = a(r, 9)
            rec("Rows") = a(r, 10): rec("Filter") = SafeText(a(r, 11)): rec("Status") = SafeText(a(r, 12))
            rec("BaseSystem") = Empty: rec("PreSystem") = Empty
            rec("Use") = USE_DERIVED: rec("UsedBase") = USE_DERIVED: rec("UsedPre") = USE_DERIVED
            If wide Then
                rec("BaseSystem") = a(r, 13): rec("PreSystem") = a(r, 14)
                If Len(SafeText(a(r, 15))) > 0 Then rec("Use") = NormalizeUse(SafeText(a(r, 15)))
                If Len(SafeText(a(r, 16))) > 0 Then rec("UsedPre") = SafeText(a(r, 16))
                rec("UsedBase") = rec("UsedPre")
            End If
            ' What the configuration reads, re-derived from the stored setting so a
            ' session that never re-ran the tests still honours it.
            rec("Use") = modValueSources.VS_MetricUse(CStr(rec("Metric")), "PRE")
            rec("UseBase") = modValueSources.VS_MetricUse(CStr(rec("Metric")), "BASE")
            rec("Base") = PickValue(CStr(rec("UseBase")), a(r, 8), CBool(rec("HasBase")), rec("BaseSystem"), CStr(rec("UsedBase")))
            rec("Pre") = PickValue(CStr(rec("Use")), a(r, 9), CBool(rec("HasPre")), rec("PreSystem"), CStr(rec("UsedPre")))
            key = UCase$(rec("Date") & "|" & rec("EntityID") & "|" & rec("Entity") & "|" & _
                         rec("TestCase") & "|" & rec("Element") & "|" & rec("Metric"))
            Set d(key) = rec
        End If
    Next r
Done:
End Function

Private Sub StyleDerivedSheet(ByVal ws As Worksheet, ByVal n As Long)
    Dim lastRow As Long
    On Error Resume Next
    lastRow = DV_FIRST_ROW + n - 1
    ' The store wears the same chrome as the rest of the workspace.
    StyleOne ws, DV_COLS, DV_COLS
    With ws.Range(ws.Cells(DV_HEADER_ROW, 1), ws.Cells(DV_HEADER_ROW, DV_COLS))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(52, 64, 84)
        .HorizontalAlignment = xlCenter
    End With
    If n > 0 Then
        ws.Range(ws.Cells(DV_FIRST_ROW, 8), ws.Cells(lastRow, 9)).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(DV_FIRST_ROW, 10), ws.Cells(lastRow, 10)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(DV_HEADER_ROW, 1), ws.Cells(lastRow, DV_COLS)).Borders.Color = RGB(208, 213, 221)
        ws.Range(ws.Cells(DV_HEADER_ROW, 1), ws.Cells(lastRow, DV_COLS)).AutoFilter
    End If
    ws.columns("A:L").AutoFit
    ws.columns("K").ColumnWidth = 60
    ws.rows(DV_HEADER_ROW).RowHeight = 26
    ' Freezing needs a visible window; a headless generation run must not try.
    If Application.Visible Then Sheet1.EnsureFrozenLayout ws, 100, DV_HEADER_ROW, 0
    Err.Clear
End Sub

Private Function Delta(ByVal sys As Variant, ByVal ecl As Variant, ByVal found As Boolean) As Variant
    If Not found Then Exit Function
    If IsNumeric(sys) And IsNumeric(ecl) Then Delta = CDbl(sys) - CDbl(ecl)
End Function

Private Function ResultStatus(ByVal baseFound As Boolean, ByVal preFound As Boolean, ByVal status As String, _
                              Optional ByVal sourceKey As String = "ECL") As String
    If baseFound And preFound Then
        ResultStatus = "Base and pre-shock rebuilt from " & sourceKey
    ElseIf baseFound Then
        ResultStatus = "Base rebuilt" & IIf(Len(status) > 0, " - " & status, " - pre-shock unavailable")
    ElseIf preFound Then
        ResultStatus = "Pre-shock rebuilt - base unavailable"
    Else
        ResultStatus = IIf(Len(status) > 0, status, "Not tested - missing source, invalid measure, empty population or undefined ratio")
    End If
End Function

' One systemic cause - a date mismatch, an unmapped dimension - otherwise produces the same
' sentence once per test case. The detail stays on the sheet, row by row; the summary keeps
' one line per distinct cause with a count, so the real problem is not buried.
Private Sub Note(ByRef notes As Collection, ByVal text As String, Optional ByVal subject As String = "")
    Dim cause As String, i As Long, existing As String
    If notes Is Nothing Then Set notes = New Collection
    cause = NoteCause(text)
    For i = 1 To notes.count
        existing = CStr(notes(i))
        If InStr(1, existing, cause, vbTextCompare) = 1 Then
            ' Collection.Add's Before/After are 1-based and reject 0, so the updated line is
            ' simply re-appended rather than put back at its old index.
            notes.Remove i
            notes.Add cause & " [" & (NoteCount(existing) + 1) & " test cases]"
            Exit Sub
        End If
    Next i
    notes.Add cause & " [1 test case]"
End Sub

' Strips the per-test-case prefix so two reports of the same underlying cause collapse.
Private Function NoteCause(ByVal text As String) As String
    Dim p As Long
    p = InStr(1, text, ": ", vbBinaryCompare)
    If p > 0 And p < 60 Then NoteCause = Mid$(text, p + 2) Else NoteCause = text
End Function

Private Function NoteCount(ByVal s As String) As Long
    Dim p As Long, q As Long
    p = InStrRev(s, " [")
    If p = 0 Then NoteCount = 1: Exit Function
    q = InStr(p, s, " test")
    If q = 0 Then NoteCount = 1: Exit Function
    NoteCount = val(Mid$(s, p + 2, q - p - 2))
    If NoteCount < 1 Then NoteCount = 1
End Function

Private Sub SetCaseResult(ByVal ws As Worksheet, ByVal caseRow As Long, ByVal message As String, ByVal matched As Double)
    If caseRow < PS_CASE_FIRST_ROW Then Exit Sub
    On Error Resume Next
    ws.Cells(caseRow, 10).Value2 = matched
    ws.Cells(caseRow, 11).Value2 = message
    On Error GoTo 0
End Sub

' The independently derived PRE-SHOCK figure for one metric, or Empty.
Public Function BaseTestLookup(ByVal cache As Object, ByVal tc As Object, ByVal el As Object, ByVal metric As String, ByRef found As Boolean) As Variant
    Dim rec As Object
    found = False
    Set rec = DerivedLookup(cache, tc, el, metric)
    If rec Is Nothing Then Exit Function
    If Not CBool(rec("HasPre")) Then Exit Function
    found = True
    BaseTestLookup = rec("Pre")
End Function

' The whole derived record - base, pre-shock, source, filter and status - for one metric.
Public Function DerivedLookup(ByVal cache As Object, ByVal tc As Object, ByVal el As Object, ByVal metric As String) As Object
    Dim k As String
    If cache Is Nothing Then Exit Function
    k = BaseCacheKey(tc, el, SafeUpperText(metric))
    If cache.Exists(k) Then Set DerivedLookup = cache(k)
End Function

Private Function BaseCacheKey(ByVal tc As Object, ByVal el As Object, ByVal metric As String) As String
    BaseCacheKey = UCase$(tc("Date") & "|" & tc("EntityID") & "|" & tc("EntityCode") & "|" & tc("TestCaseCode") & "|" & el("ScenarioCode") & "|" & metric)
End Function

Private Function BuildPreShockMeasureMap(ByVal metricMap As Object) As Object
    Dim d As Object, k As Variant, m As Object, matches As Object, mt As Object, fld As String
    Set d = NewMap()
    For Each k In metricMap.keys
        Set m = metricMap(k)
        Set matches = RegexExecute("(?:SUM|AVG|MIN|MAX|COUNT)\s*\(\s*(GREATEST\s*\([^()]*\)|[A-Za-z0-9_ |*]+)\s*\)", CStr(m("Formula")))
        For Each mt In matches
            fld = Trim$(CStr(mt.SubMatches(0)))
            If fld <> "*" And Len(fld) > 0 Then
                If Not d.Exists(SafeUpperText(fld)) Then d(SafeUpperText(fld)) = MeasureColumnFor(fld)
            End If
        Next mt
    Next k
    Set BuildPreShockMeasureMap = d
End Function

' ============================= Pre-Shock Studio =============================

Public Sub OpenPreShockStudio()
    Dim ws As Worksheet, rowIx As Long, mode As String, target As String, currentText As String, ctx As String, html As String
    Dim fso As Object, folder As String, token As String, valuesPath As String, htmlPath As String, stream As Object, er As String
    On Error GoTo Failed
    EnsurePreShockSheet
    Set ws = ActiveSheet
    If ws.name <> PS_CASES_SHEET And ws.name <> PS_METRICS_SHEET Then Set ws = PsCasesSheet(): ws.Activate
    rowIx = ActiveCell.row
    If rowIx < PS_FIRST_ROW Then rowIx = PS_FIRST_ROW
    ' Which sheet you are on decides what you are editing: a derivation formula on
    ' the metrics sheet, a filter on the test cases sheet. That is the whole benefit
    ' of the split - the tool no longer has to guess from a row number.
    If ws.name = PS_METRICS_SHEET Then
        mode = "formula": target = ws.name & "!D" & rowIx: currentText = SafeText(ws.Cells(rowIx, 4).Value2)
    Else
        If Len(SafeText(ws.Cells(rowIx, 2).Value2)) = 0 Then rowIx = PS_CASE_FIRST_ROW
        mode = "filter": target = ws.name & "!G" & rowIx
        currentText = SafeText(ws.Cells(rowIx, 7).Value2)
        If Len(currentText) = 0 Then currentText = SafeText(ws.Cells(rowIx, 6).Value2)
    End If
    ' Deliberately does NOT load the ECL extract. Everything the builder needs was cached
    ' at upload, so the window opens immediately no matter how large the source is.
    If Not ValueCacheReady() Then
        If Not UiAsk("Filter builder", "Open the filter builder without value suggestions?", _
                     "No ECL values are cached yet, so the builder cannot suggest fields and values. Load the ECL output first to get them.") Then Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject"): folder = fso.GetSpecialFolder(2): token = fso.GetTempName
    valuesPath = fso.BuildPath(folder, "JKB_PreShock_Values_" & token & ".js")
    ' An .hta rather than an .html: a browser page cannot write back to Excel, which is why
    ' the old studio could only offer "copy this and paste it yourself". ALM Rule Studio uses
    ' the same mechanism for the same reason.
    htmlPath = fso.BuildPath(folder, "JKB_Filter_Builder_" & token & ".hta")
    WritePreShockValuesJs valuesPath
    ctx = BuildPreShockStudioContext(ws, rowIx, mode, target, currentText)
    html = PreShockStudioHtml(ctx, fso.GetFileName(valuesPath))
    Dim bridge As String
    bridge = modShared_Bridge.BridgeNewToken()
    html = Replace(html, "__BRIDGE_JS__", modShared_Bridge.BridgeJs(bridge))
    Set stream = fso.CreateTextFile(htmlPath, False, True): stream.Write html: stream.Close: Set stream = Nothing
    On Error Resume Next
    Shell "mshta.exe """ & htmlPath & """", vbNormalFocus
    If Err.Number <> 0 Then
        ' mshta is blocked in some environments. Fall back to the browser: the builder still
        ' works, it just hands the expression back through the clipboard.
        Err.Clear
        modShared_Bridge.BridgeCleanUp bridge
        ThisWorkbook.FollowHyperlink address:=htmlPath, NewWindow:=True
        On Error GoTo Failed
        Exit Sub
    End If
    On Error GoTo Failed
    ' Excel listens while the window is open. It has to: an HTA can no longer
    ' reach back into a running Excel, so the workbook waits and the window asks.
    modShared_Bridge.BridgePump bridge, "ApplyPreShockEditQuiet|TestFilterForBuilder"
    modShared_Bridge.BridgeCleanUp bridge
    Exit Sub
Failed:
    er = Err.description: On Error Resume Next: If Not stream Is Nothing Then stream.Close
    UiProblem "Filter builder", "The filter builder could not open.", er
End Sub

Private Function BuildPreShockStudioContext(ByVal ws As Worksheet, ByVal rowIx As Long, ByVal mode As String, ByVal target As String, ByVal currentText As String) As String
    Dim dims As String, measures As String, r As Long, nm As String, cand As String, avail As String, c As Long, hdr As String
    Dim fws As Worksheet
    ' The mapped dimensions are listed on the Fields sheet. They used to be read
    ' from the sheet being edited (test cases or metrics), so the builder showed
    ' Yes/No or metric names as "dimensions" and could not offer their values.
    Set fws = PsFieldsSheet()
    If Not fws Is Nothing Then
        For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
            nm = SafeText(fws.Cells(r, 1).Value2): If Len(nm) = 0 Then Exit For
            cand = SafeText(fws.Cells(r, 2).Value2): avail = SafeText(fws.Cells(r, 4).Value2)
            If Len(dims) > 0 Then dims = dims & ","
            dims = dims & "{""name"":" & PsJson(nm) & ",""candidates"":" & PsJson(cand) & ",""available"":" & PsJson(avail) & "}"
        Next r
    End If
    ' Measures come from the cache too, so no column sniffing happens at open time.
    Dim meas As Object, mk As Variant
    Set meas = CachedGroup("MEASURE")
    For Each mk In meas.keys
        If Len(measures) > 0 Then measures = measures & ","
        measures = measures & PsJson(CStr(mk))
    Next mk
    ' Every physical column the extract carries, so any field can be filtered on - not only
    ' the logical dimensions someone thought to map in advance.
    Dim flds As String, fieldsMap As Object, fk As Variant
    Set fieldsMap = CachedGroup("FIELD")
    For Each fk In fieldsMap.keys
        If Len(flds) > 0 Then flds = flds & ","
        flds = flds & PsJson(CStr(fk))
    Next fk

    BuildPreShockStudioContext = "{""mode"":" & PsJson(mode) & ",""target"":" & PsJson(target) & ",""current"":" & PsJson(currentText) & _
      ",""testCase"":" & PsJson(SafeText(ws.Cells(rowIx, 2).Value2)) & ",""testElement"":" & PsJson(SafeText(ws.Cells(rowIx, 3).Value2)) & _
      ",""elementType"":" & PsJson(SafeText(ws.Cells(rowIx, 4).Value2)) & ",""source"":" & PsJson(SafeText(ws.Cells(rowIx, 5).Value2)) & _
      ",""dimensions"":[" & dims & "],""measures"":[" & measures & "],""fields"":[" & flds & "]}"
End Function

Private Function PreShockColumnLooksNumeric(ByVal c As Long) As Boolean
    Dim r As Long, lastRow As Long, seen As Long, nums As Long, v As Variant, hdr As String
    hdr = SafeUpperText(mEclData(1, c))
    If Right$(hdr, 3) = "_ID" Or hdr = "ID" Or InStr(1, hdr, "DATE", vbTextCompare) > 0 Or InStr(1, hdr, "CODE", vbTextCompare) > 0 Then Exit Function
    lastRow = WorksheetFunction.Min(UBound(mEclData, 1), 252)
    For r = 2 To lastRow
        v = mEclData(r, c)
        If Len(SafeText(v)) > 0 Then
            seen = seen + 1: If IsNumeric(v) Then nums = nums + 1
        End If
    Next r
    If seen > 0 Then PreShockColumnLooksNumeric = (nums / seen >= 0.8)
End Function

' Serialises the CACHED values. No ECL access, no scan - this is why the window opens at once.
Private Sub WritePreShockValuesJs(ByVal path As String)
    Dim dims As Object, k As Variant, v As Variant, js As String, first As Boolean
    Dim fso As Object, stream As Object
    Set dims = CachedGroup("DIM")
    js = "window.JKB_VALUES={"
    first = True
    For Each k In dims.keys
        If Not first Then js = js & ","
        first = False
        js = js & PsJson(CStr(k)) & ":["
        Dim firstVal As Boolean
        firstVal = True
        For Each v In dims(k)
            If Not firstVal Then js = js & ","
            firstVal = False
            js = js & PsJson(CStr(v))
        Next v
        js = js & "]"
    Next k
    js = js & "};"
    Set fso = CreateObject("Scripting.FileSystemObject"): Set stream = fso.CreateTextFile(path, False, True)
    stream.Write js: stream.Close
End Sub

Private Function PsJson(ByVal s As String) As String
    Dim i As Long, n As Long, out As String
    For i = 1 To Len(s)
        n = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case n
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 0 To 31, 60, 62, 38, 8232, 8233: out = out & "\u" & Right$("0000" & Hex$(n), 4)
            Case Else: out = out & Mid$(s, i, 1)
        End Select
    Next i
    PsJson = """" & out & """"
End Function

' Returns the filter builder source with its two placeholders filled in.
'
' The builder itself lives in VBA\builder.hta and is injected into a very hidden sheet at
' build time. Keeping it out of VBA string concatenation avoids the 1023-character line and
' 25-continuation limits entirely, and means the builder can be edited as a normal file.
Private Function PreShockStudioHtml(ByVal ctx As String, ByVal valuesFile As String) As String
    Dim ws As Worksheet, lastRow As Long, a As Variant, r As Long, sb As String
    Set ws = GetWorksheetSafe(ThisWorkbook, BUILDER_SHEET)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 740, , "The filter builder source is missing from this workbook (sheet " & BUILDER_SHEET & "). Rebuild the tool."
    End If
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < 1 Then Err.Raise vbObjectError + 741, , "The filter builder source is empty."
    a = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value2
    If Not IsArray(a) Then
        sb = CStr(a)
    Else
        For r = 1 To UBound(a, 1)
            sb = sb & CStr(a(r, 1)) & vbCrLf
        Next r
    End If
    sb = Replace(sb, "__VALUES_FILE__", valuesFile)
    sb = Replace(sb, "__CTX__", ctx)
    PreShockStudioHtml = sb
End Function

' ================ ECL aggregation and filter evaluation engine ==============

Private Function GenericAggKey(ByVal op As String, ByVal fieldSpec As String, ByVal stage As String) As String
    GenericAggKey = SafeUpperText(op) & "#" & SafeUpperText(fieldSpec) & "#" & stage
End Function

Private Sub AddGenericAgg(ByVal d As Object, ByVal fieldSpec As String, ByVal stage As String, ByVal v As Variant)
    Dim n As Double, k As String, cnt As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        AddAgg d, GenericAggKey("INVALID", fieldSpec, stage), 1
        Exit Sub
    End If
    If Not IsNumeric(v) Then
        AddAgg d, GenericAggKey("INVALID", fieldSpec, stage), 1
        Exit Sub
    End If
    n = CDbl(v)
    AddAgg d, GenericAggKey("SUM", fieldSpec, stage), n
    AddAgg d, GenericAggKey("COUNT", fieldSpec, stage), 1
    k = GenericAggKey("MIN", fieldSpec, stage)
    If Not d.Exists(k) Or n < CDbl(d(k)) Then d(k) = n
    k = GenericAggKey("MAX", fieldSpec, stage)
    If Not d.Exists(k) Or n > CDbl(d(k)) Then d(k) = n
End Sub

Private Function AggregateFunctionValue(ByVal agg As Object, ByVal op As String, ByVal fieldSpec As String, ByVal stage As String, ByRef found As Boolean) As Double
    Dim stages As Variant, st As Variant, k As String, sm As Double, ct As Double, v As Double, have As Boolean, localCount As Double
    ' "0" is the bucket a source with no stage column uses. Including it in ALL is
    ' what lets one metric engine serve both kinds; a staged extract never writes a
    ' "0" key, so adding it here cannot change an existing ECL or CAPRWA figure.
    stages = IIf(stage = "ALL" Or Len(stage) = 0, Array("1", "2", "3", "0"), Array(stage))
    If SafeUpperText(op) = "COUNT" And Trim$(fieldSpec) = "*" Then
        For Each st In stages
            k = "ROWS#" & CStr(st): If agg.Exists(k) Then ct = ct + CDbl(agg(k)): have = True
        Next st
        found = have: AggregateFunctionValue = ct: Exit Function
    End If
    For Each st In stages
        If agg.Exists(GenericAggKey("INVALID", fieldSpec, CStr(st))) Then
            found = False
            Exit Function
        End If
        Select Case SafeUpperText(op)
            Case "SUM"
                k = GenericAggKey("SUM", fieldSpec, CStr(st)): If agg.Exists(k) Then sm = sm + CDbl(agg(k)): have = True
            Case "COUNT"
                k = GenericAggKey("COUNT", fieldSpec, CStr(st)): If agg.Exists(k) Then sm = sm + CDbl(agg(k)): have = True
            Case "AVG"
                k = GenericAggKey("SUM", fieldSpec, CStr(st)): If agg.Exists(k) Then sm = sm + CDbl(agg(k)): have = True
                k = GenericAggKey("COUNT", fieldSpec, CStr(st)): If agg.Exists(k) Then ct = ct + CDbl(agg(k))
            Case "MIN"
                k = GenericAggKey("MIN", fieldSpec, CStr(st)): If agg.Exists(k) Then v = CDbl(agg(k)): If Not have Or v < sm Then sm = v: have = True
            Case "MAX"
                k = GenericAggKey("MAX", fieldSpec, CStr(st)): If agg.Exists(k) Then v = CDbl(agg(k)): If Not have Or v > sm Then sm = v: have = True
        End Select
    Next st
    found = have
    If SafeUpperText(op) = "AVG" Then
        If ct <> 0 Then
            AggregateFunctionValue = sm / ct
        Else
            AggregateFunctionValue = 0
        End If
    Else
        AggregateFunctionValue = sm
    End If
End Function

Private Function StripPreShockStage(ByVal formula As String, ByRef stage As String) As String
    Dim ignored As String
    StripPreShockStage = SplitMetricFormula(formula, stage, ignored)
End Function

' Splits a metric formula into the bit that is arithmetic over aggregates, the
' stage it wants, and the condition that picks its rows.
'
'     SUM(ECL) WHERE STAGE=2
'         -> "SUM(ECL)", stage 2, no row selector
'     SUM(CASHFLOW_AMOUNT_LCY_PRE_FACTOR) WHERE COA_BALANCESHEET_CATEGORY IN ('Assets')
'         -> "SUM(...)", stage ALL, selector on the balance sheet category
'
' The second form is what the liquidity and capital metrics need. An ALM row is
' not staged, so "which rows" cannot be answered by a stage; it has to be
' answered by a condition, and the condition differs per framework - LCR splits
' on the segmentation rule category, legal liquidity on the balance sheet
' category. Putting it in the formula rather than in code is what makes the
' assets-is-pre-factor / liabilities-is-post-factor policy editable.
' Pure, and asked the same few hundred questions tens of thousands of times a
' run - every metric, for every test case and source - so each is answered once.
Public Function SplitMetricFormula(ByVal formula As String, ByRef stage As String, ByRef whereExpr As String) As String
    Dim parts As Variant
    If mFormulaSplits Is Nothing Then Set mFormulaSplits = NewExactMap()
    If mFormulaSplits.Exists(formula) Then
        parts = mFormulaSplits(formula)
    Else
        parts = Array(SplitMetricFormulaNow(formula, stage, whereExpr), "", "")
        parts(1) = stage: parts(2) = whereExpr
        If mFormulaSplits.count >= 5000 Then mFormulaSplits.RemoveAll
        mFormulaSplits.Add formula, parts
    End If
    stage = CStr(parts(1)): whereExpr = CStr(parts(2))
    SplitMetricFormula = CStr(parts(0))
End Function

Private Function SplitMetricFormulaNow(ByVal formula As String, ByRef stage As String, ByRef whereExpr As String) As String
    Dim s As String, matches As Object, m As Object, p As Long, tail As String
    s = Trim$(formula): stage = "ALL": whereExpr = ""

    Set matches = RegexExecute("\bWHERE\s+STAGE\s*(?:=|IN\s*\()\s*([123])", s)
    If matches.count > 0 Then
        Set m = matches(0)
        stage = CStr(m.SubMatches(0))
        p = m.FirstIndex + 1
        SplitMetricFormulaNow = Trim$(Left$(s, p - 1))
        Exit Function
    End If

    ' Any other WHERE is a row selector, handed to the ordinary filter compiler.
    Set matches = RegexExecute("\bWHERE\b", s)
    If matches.count > 0 Then
        Set m = matches(0)
        p = m.FirstIndex + 1
        tail = Trim$(Mid$(s, p + Len("WHERE")))
        whereExpr = tail
        SplitMetricFormulaNow = Trim$(Left$(s, p - 1))
        Exit Function
    End If

    SplitMetricFormulaNow = s
End Function

' The row selector a metric rule asks for, or "" for none.
Public Function MetricWhere(ByVal m As Object) As String
    Dim stage As String, w As String
    If m Is Nothing Then Exit Function
    SplitMetricFormula CStr(m("Formula")), stage, w
    MetricWhere = w
End Function

' A source and a row selector together name one aggregate. Metrics with no
' selector share their source's ordinary aggregate, so nothing that worked
' before pays for this.
Public Function SliceKey(ByVal src As String, ByVal whereExpr As String) As String
    SliceKey = SafeUpperText(src) & "~" & UCase$(Trim$(whereExpr))
End Function

' The distinct row selectors one source's enabled metrics ask for, always
' including "" - the unsliced population, which the row count and the breakdown
' are reported against.
Private Function SlicesForSource(ByVal metricMap As Object, ByVal src As String) As Collection
    Dim c As New Collection, seen As Object, mk As Variant, w As String
    Set seen = NewMap()
    seen("") = True
    c.Add ""
    For Each mk In metricMap.keys
        If SafeUpperText(metricMap(mk)("Source")) = SafeUpperText(src) Then
            w = MetricWhere(metricMap(mk))
            If Len(Trim$(w)) > 0 Then
                If Not seen.Exists(UCase$(Trim$(w))) Then
                    seen(UCase$(Trim$(w))) = True
                    c.Add w
                End If
            End If
        End If
    Next mk
    Set SlicesForSource = c
End Function

' Two conditions, both of which must hold. Either being empty means "everything",
' so the other one alone is the answer.
Public Function AndTerms(ByVal a As String, ByVal b As String) As String
    Dim x As String, y As String
    x = Trim$(a): y = Trim$(b)
    If Len(x) = 0 Then AndTerms = y: Exit Function
    If Len(y) = 0 Then AndTerms = x: Exit Function
    AndTerms = "( " & x & " ) AND ( " & y & " )"
End Function

' Finds the right-most top-level arithmetic operator, so an expression splits at the
' operator that binds last.
'
' This was written as a single-line If with ElseIf clauses, which VBA does not accept.
' It never showed up because VBA compiles on demand and nothing had reached this
' procedure; the moment ECL evaluation ran it was a hard syntax error.
Private Function FindTopArithmetic(ByVal s As String, ByVal operators As String) As Long
    Dim i As Long, depth As Long, ch As String
    For i = Len(s) To 1 Step -1
        ch = Mid$(s, i, 1)
        If ch = ")" Then
            depth = depth + 1
        ElseIf ch = "(" Then
            depth = depth - 1
        ElseIf depth = 0 Then
            If InStr(1, operators, ch, vbBinaryCompare) > 0 Then
                ' Position 1 is a leading sign, not a binary operator.
                If i > 1 Then FindTopArithmetic = i: Exit Function
            End If
        End If
    Next i
End Function

Private Function EvalPreShockExpression(ByVal agg As Object, ByVal expr As String, ByVal stage As String, ByRef found As Boolean) As Double
    Dim s As String, p As Long, op As String, lf As Boolean, rf As Boolean, lv As Double, rv As Double, fn As String, arg As String, q As Long
    s = TrimOuterParens(Trim$(expr))
    p = FindTopArithmetic(s, "+-")
    If p = 0 Then p = FindTopArithmetic(s, "*/")
    If p > 0 Then
        op = Mid$(s, p, 1): lv = EvalPreShockExpression(agg, Left$(s, p - 1), stage, lf): rv = EvalPreShockExpression(agg, Mid$(s, p + 1), stage, rf)
        found = lf And rf: If Not found Then Exit Function
        Select Case op
            Case "+": EvalPreShockExpression = lv + rv
            Case "-": EvalPreShockExpression = lv - rv
            Case "*": EvalPreShockExpression = lv * rv
            Case "/": If rv = 0 Then found = False Else EvalPreShockExpression = lv / rv
        End Select
        Exit Function
    End If
    If IsNumeric(s) Then found = True: EvalPreShockExpression = CDbl(s): Exit Function
    q = InStr(1, s, "(", vbBinaryCompare)
    If q > 1 And Right$(s, 1) = ")" Then
        fn = SafeUpperText(Left$(s, q - 1)): arg = Trim$(Mid$(s, q + 1, Len(s) - q - 1))
        If fn = "SUM" Or fn = "AVG" Or fn = "MIN" Or fn = "MAX" Or fn = "COUNT" Then EvalPreShockExpression = AggregateFunctionValue(agg, fn, arg, stage, found): Exit Function
    End If
    found = False
End Function

' Whether this source can borrow ECL's fields to resolve a filter.
'
' Only when ECL is loaded, this is not ECL itself, and this source has an
' account number to match on. Anything else and the honest answer is the
' original refusal: a filter this source cannot evaluate.
Private Function CanJoinThroughEcl() As Boolean
    If SafeUpperText(mActiveSource) = "ECL" Then Exit Function
    If Not SourceLoaded("ECL") Then Exit Function
    If mEclHeaders Is Nothing Then Exit Function
    CanJoinThroughEcl = (FirstExistingEclColumn("ACCOUNT_NUMBER") > 0)
End Function

' The account numbers an expression selects, resolved against ECL.
'
' Cached per expression for the whole run: a hundred test cases against three
' joined sources would otherwise re-read a 66,000-row extract three hundred
' times to answer the same hundred questions.
Private Function AccountSetFor(ByVal expr As String, ByVal aod As String) As Object
    Dim keep As String, node As Object, unsupported As String, r As Long
    Dim d As Object, iAcct As Long, acct As String, key As String
    Dim scopedExpr As String, dateOk As Boolean, errorNumber As Long, errorText As String
    Dim matchDates As Boolean, dateInfo As Object, rowDates As Variant
    Dim joinMask As Variant, useMask As Boolean
    Set d = NewMap()
    ' Already resolved this run: answered without switching to ECL and back,
    ' each switch having been a copy of a whole table.
    key = JoinCacheKeyQuiet(expr, aod)
    If Len(key) > 0 Then
        If mJoinCache.Exists(key) Then Set AccountSetFor = mJoinCache(key): Exit Function
    End If
    keep = mActiveSource
    If Not ActivateSource("ECL") Then Set AccountSetFor = d: Exit Function

    On Error GoTo Restore
    matchDates = Not IgnoreAsOfDate()
    If matchDates Then
        If Not mEclHeaders.Exists("AS_OF_DATE") Then Err.Raise vbObjectError + 741, , "ECL: AS_OF_DATE is required for account joins."
        Set dateInfo = SourceDateInventory()
        RequireSourceDate dateInfo, aod
    End If
    scopedExpr = AndTerms(modScenarioBuilder_Multi.ApplyEntityScope(expr), _
                         modScenarioBuilder_Multi.SourceBankFilter(mEclHeaders, "ECL"))
    key = UCase$(aod & "|" & CStr(matchDates) & "|" & Trim$(scopedExpr))
    If mJoinCache Is Nothing Then Set mJoinCache = NewMap()
    If mJoinCache.Exists(key) Then
        Set AccountSetFor = mJoinCache(key)
        ActivateSource keep
        Exit Function
    End If
    If matchDates Then rowDates = dateInfo("Rows")
    iAcct = FirstExistingEclColumn("ACCOUNT_NUMBER")
    If iAcct > 0 Then
        Set node = Nothing
        If Len(Trim$(scopedExpr)) > 0 Then
            Set node = CompileFilter(Trim$(scopedExpr), LoadEclFilterMap(), unsupported)
            ' Anything ECL cannot map either is a genuinely unknown field, and
            ' saying so is better than quietly resolving to every account.
            If Len(unsupported) > 0 Then
                ActivateSource keep
                Err.Raise vbObjectError + 721, , "No loaded extract can map filter field(s): " & unsupported
            End If
        End If
        If Not node Is Nothing Then
            joinMask = FilterMaskForActive(node)
            useMask = IsArray(joinMask)
        End If
        For r = 2 To UBound(mEclData, 1)
            dateOk = True
            If r Mod 20000 = 0 Then ProgressDetail "Resolving ECL account filter: " & format$(r - 1, "#,##0") & " of " & format$(UBound(mEclData, 1) - 1, "#,##0") & " rows"
            If matchDates Then dateOk = (rowDates(r) = aod)
            If Not dateOk Then GoTo NextJoinAccount
            If node Is Nothing Then
                acct = SafeUpperText(mEclData(r, iAcct))
                If Len(acct) > 0 Then If Not d.Exists(acct) Then d(acct) = True
            ElseIf RowSelected(node, joinMask, useMask, r) Then
                acct = SafeUpperText(mEclData(r, iAcct))
                If Len(acct) > 0 Then If Not d.Exists(acct) Then d(acct) = True
            End If
NextJoinAccount:
        Next r
    End If
Restore:
    errorNumber = Err.Number: errorText = Err.description
    ActivateSource keep
    If errorNumber <> 0 Then Err.Raise errorNumber, "AccountSetFor", errorText
    ' Set, not a bare assignment: putting an object into a Dictionary item
    ' without Set raises "wrong number of arguments or invalid property
    ' assignment", which says nothing at all about what actually happened.
    Set mJoinCache(key) = d
    Set AccountSetFor = d
End Function

' Two conditions, deliberately kept apart.
'
'   caseExpr   what the TEST CASE selects. Usually names ECL fields - report
'              classification, IFRS stage, economic sector - and an ALM extract
'              may have none of them. This is the part that gets resolved
'              through ECL into a set of account numbers when it has to be.
'
'   localExpr  what THIS source's own rules say: the metric's row selector and
'              the entity rule. Both are facts about this source's own columns
'              and are always evaluated here.
'
' Merging the two into one expression - which is what the first version did -
' sent COA_BALANCESHEET_CATEGORY to ECL, which has never heard of it, and the
' whole aggregate was refused. The metric's selector is not the test case's
' business and must not travel with it.
Private Function GetEclAggregate(ByVal cache As Object, ByVal aod As String, ByVal caseExpr As String, _
                                 ByVal localExpr As String, ByVal fieldMap As Object) As Object
    Dim key As String, node As Object, agg As Object, unsupported As String, r As Long, st As String, dateOk As Boolean, dateRows As Long
    Dim fieldSpec As Variant, c As Long, matchDates As Boolean, matchRow As Boolean, hasStage As Boolean
    Dim joinMode As Boolean, acctSet As Object, joinHit As Object, iJoinAcct As Long
    Dim expr As String, dateInfo As Object, rowDates As Variant
    Dim specName() As String, specCol() As Long, specSlot() As Long, slotName() As String, nSpec As Long, nSlot As Long
    Dim accSum() As Double, accCnt() As Double, accMin() As Double, accMax() As Double, accInv() As Double
    Dim rowsAt(0 To 3) As Double, rowMask As Variant, useMask As Boolean, lastR As Long, acctKey As String
    Dim stageCol As Variant, stageKnown As Boolean, si As Long, i As Long, k As Long, v As Variant, n As Double
    matchDates = Not IgnoreAsOfDate()
    If matchDates And Not mEclHeaders.Exists("AS_OF_DATE") Then
        ' The supplied capital component file is an undated base snapshot.
        ' Its totals remain usable, while the independent date-proof check
        ' reports REVIEW. Never extend this exemption to filtered pre-shock.
        If mActiveSource = "CAP" And SourceFamilyOf(mEclHeaders) = "CAP" And Len(Trim$(caseExpr)) = 0 Then
            matchDates = False
        Else
            Err.Raise vbObjectError + 741, , mActiveSource & ": AS_OF_DATE is required for date-matched reconciliation."
        End If
    End If
    ' Validate the full source once before a filter or cross-source join can run.
    ' An absent reporting date is then a constant-time refusal for every case.
    If matchDates Then
        Set dateInfo = SourceDateInventory()
        RequireSourceDate dateInfo, aod
    End If
    ' Bank identity stays local to this source and is part of the cache key.
    ' Group ALM extracts also contain EJARA and UFICO; branch scope is not enough.
    localExpr = AndTerms(Trim$(localExpr), modScenarioBuilder_Multi.SourceBankFilter(mEclHeaders, mActiveSource))
    ' The source is part of the key: the same filter run against ECL and against CAPRWA is
    ' two different aggregates.
    key = UCase$(mActiveSource & "|" & aod & "|" & CStr(matchDates) & "|" & Trim$(caseExpr) & "|" & Trim$(localExpr))
    If cache.Exists(key) Then Set GetEclAggregate = cache(key): Exit Function
    If matchDates Then rowDates = dateInfo("Rows")
    ' An empty expression is not an error - it is the BASE: the whole portfolio, unfiltered.
    expr = AndTerms(Trim$(caseExpr), Trim$(localExpr))
    Set node = Nothing
    If Len(Trim$(expr)) > 0 Then
        Set node = CompileFilter(Trim$(expr), fieldMap, unsupported)
    End If
    ' ---- the join -------------------------------------------------------
    '
    ' Most filter conditions name fields that only the ECL extract carries -
    ' report classification, IFRS stage, economic sector. An ALM or capital
    ' extract has none of them, so a credit test case could not be measured
    ' against liquidity at all: the filter simply failed to compile and that
    ' source reported nothing.
    '
    ' When that happens the condition is resolved against ECL instead, down to
    ' the set of account numbers it selects, and THIS source is filtered by
    ' membership of that set. The entity rule is still applied here directly,
    ' because it is a fact about this source's own rows.
    '
    ' The overlap is partial and that is the data, not a fault: of 214,379
    ' accounts in the legal liquidity file, 18,449 exist in ECL - the rest are
    ' GL and trial-balance rows that are not deals at all. So the coverage is
    ' RECORDED on the aggregate and reported, rather than being left to look
    ' like a smaller number than expected.
    If Len(unsupported) > 0 Then
        If CanJoinThroughEcl() Then
            ' ONLY the test case's own condition goes to ECL. The local part
            ' stays here, where its columns are.
            Set acctSet = AccountSetFor(Trim$(caseExpr), aod)
            iJoinAcct = FirstExistingEclColumn("ACCOUNT_NUMBER")
            If iJoinAcct = 0 Then
                Err.Raise vbObjectError + 722, , mActiveSource & " cannot be joined to the ECL extract: it has no account number column."
            End If
            joinMode = True
            unsupported = ""
            Set node = Nothing
            If Len(Trim$(localExpr)) > 0 Then
                Set node = CompileFilter(Trim$(localExpr), fieldMap, unsupported)
                If Len(unsupported) > 0 Then
                    Err.Raise vbObjectError + 723, , mActiveSource & " cannot map its own metric or entity condition: " & unsupported
                End If
            End If
        Else
            Err.Raise vbObjectError + 721, , "ECL source cannot map filter field(s): " & unsupported
        End If
    End If

    Set agg = NewMap()
    ' "Ignore" lets a trimmed or sample extract be used against a different reporting date.
    ' It is off by default because a figure rebuilt from another AOD is not a valid check.
    Set joinHit = NewMap()
    ' By FAMILY, not by whether a stage column happens to be present.
    '
    ' An ALM extract carries STAGE_ID and STAGE_NAME, and on a GL-sourced row -
    ' which is most of them - both are NULL. Testing for the column therefore said
    ' "staged", every row normalised to a blank stage, and every row was dropped:
    ' 212,000 rows in, an aggregate of zero out, and no error anywhere.
    hasStage = (SourceFamilyOf(mEclHeaders) = "ECL" Or SourceFamilyOf(mEclHeaders) = "CAPRWA")
    BeginBreakdown
    ' The measures, resolved once. Totals build up in arrays - one slot per
    ' measure and stage - and become the aggregate's keys once, after the last
    ' row, exactly as AddGenericAgg would have left them. Building about twenty
    ' text keys on every matching row was most of the cost of a run.
    PlanMeasureSlots specName, specCol, specSlot, slotName, nSpec, nSlot
    ReDim accSum(1 To IIf(nSlot > 0, nSlot, 1), 0 To 3): ReDim accCnt(1 To IIf(nSlot > 0, nSlot, 1), 0 To 3)
    ReDim accMin(1 To IIf(nSlot > 0, nSlot, 1), 0 To 3): ReDim accMax(1 To IIf(nSlot > 0, nSlot, 1), 0 To 3)
    ReDim accInv(1 To IIf(nSlot > 0, nSlot, 1), 0 To 3)
    ' The condition, decided for every row at once and shared by every test case
    ' that asks it of this table. A table assigned directly (the regression
    ' tests) has no mask and is evaluated row by row, as before.
    If Not node Is Nothing Then
        rowMask = FilterMaskForActive(node)
        useMask = IsArray(rowMask)
    End If
    lastR = UBound(mEclData, 1)
    ProgressDetail "Rebuilding " & mActiveSource & ": " & format$(lastR - 1, "#,##0") & " source rows"
    For r = 2 To lastR
        If r Mod 20000 = 0 Then ProgressDetail "Rebuilding " & mActiveSource & ": " & format$(r - 1, "#,##0") & " of " & format$(lastR - 1, "#,##0") & " rows"
        dateOk = True
        If matchDates Then
            dateOk = (rowDates(r) = aod)
        End If
        If dateOk Then dateRows = dateRows + 1
        If Not dateOk Then
            matchRow = False
        ElseIf joinMode Then
            ' In the set ECL resolved, and inside the entity, if one is set.
            acctKey = SafeUpperText(mEclData(r, iJoinAcct))
            matchRow = acctSet.Exists(acctKey)
            If matchRow And Not (node Is Nothing) Then
                If useMask Then matchRow = (rowMask(r) <> 0) Else matchRow = EvalFilter(node, r)
            End If
            If matchRow Then
                If Not joinHit.Exists(acctKey) Then joinHit(acctKey) = True
            End If
        ElseIf node Is Nothing Then
            matchRow = True
        ElseIf useMask Then
            matchRow = (rowMask(r) <> 0)
        Else
            matchRow = EvalFilter(node, r)
        End If
        If matchRow Then
            ' A source with no stage column is not an error - a liquidity cashflow
            ' and a capital component are not staged things. Everything from one is
            ' aggregated under stage "0", which no staged extract ever produces, so
            ' the two kinds cannot be confused for one another.
            If hasStage Then
                If Not stageKnown Then stageCol = mEclHeaders("STAGE"): stageKnown = True
                st = NormalizeStage(mEclData(r, stageCol))
                If Len(st) = 0 Then Err.Raise vbObjectError + 742, , mActiveSource & ": invalid or missing IFRS stage at source row " & (r + mEclHeaderRow - 1) & "."
            Else
                st = "0"
            End If
            si = StageSlot(st)
            If si >= 0 Then
                rowsAt(si) = rowsAt(si) + 1
                For i = 1 To nSpec
                    c = specCol(i)
                    If c <> 0 Then
                        If c > 0 Then v = mEclData(r, c) Else v = GreatestOfRow(r, specName(i))
                        k = specSlot(i)
                        If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
                            accInv(k, si) = accInv(k, si) + 1
                        ElseIf Not IsNumeric(v) Then
                            accInv(k, si) = accInv(k, si) + 1
                        Else
                            n = CDbl(v)
                            accSum(k, si) = accSum(k, si) + n
                            accCnt(k, si) = accCnt(k, si) + 1
                            If accCnt(k, si) = 1 Then
                                accMin(k, si) = n: accMax(k, si) = n
                            Else
                                If n < accMin(k, si) Then accMin(k, si) = n
                                If n > accMax(k, si) Then accMax(k, si) = n
                            End If
                        End If
                    End If
                Next i
            End If
            ' WHY the figure is what it is, accumulated in the SAME pass.
            ' A second pass over a 95,000-row extract to answer it would cost more
            ' than the figure did.
            If mBreakdownOn Then BreakdownRow r
        End If
    Next r
    FlushMeasureSlots agg, slotName, nSlot, rowsAt, accSum, accCnt, accMin, accMax, accInv
    FlushBreakdown agg
    ' How much of the joined population this source actually holds. Carried on
    ' the aggregate so the figure can be reported with its own coverage rather
    ' than as a bare number that looks too small.
    If joinMode Then
        agg("JOIN#WANTED") = CDbl(acctSet.count)
        agg("JOIN#FOUND") = CDbl(joinHit.count)
    End If
    cache.Add key, agg: Set GetEclAggregate = agg
End Function
' ---------------------------------------------------------------------------
'  WHY a figure is what it is.
'
'  A pre-shock figure is SUM(measure) over the rows one filter selected. Knowing
'  that it is 436,000 is not knowing anything; knowing that it is 210,000 from
'  TRADE and 140,000 from INDUSTRY is. And because the BASE aggregate - the whole
'  portfolio, unfiltered - is built for the same dimensions and cached, the
'  really useful answer falls out for nothing: the filter took 31% of TRADE and
'  all of INDUSTRY. THAT is why the pre-shock is what it is.
'
'  Deliberately LONG form - one row per dimension and value - and not a cube. A
'  cube of five dimensions is a product and explodes on the first extract with a
'  high-cardinality column in it; this is a sum, and cannot.
'
'  The cost is bounded on purpose: at most MAX_BREAKDOWN_DIMS dimensions and
'  MAX_BREAKDOWN_MEASURES measures, so at most a dozen dictionary bumps per
'  matching row. Everything here can be switched off from the sheet, because a
'  cost you cannot decline is a cost you resent.
' ---------------------------------------------------------------------------
' One row per test case, element, source, dimension and value: what the filter
' selected, beside what the whole portfolio holds of the same thing.
Private Sub CollectBreakdown(ByVal rowsOut As Collection, ByVal tc As Object, ByVal el As Object, _
                             ByVal srcKey As String, ByVal preAgg As Object, ByVal baseAgg As Object, _
                             ByVal plan As Object, ByVal expr As String)
    Dim pre As Object, base As Object, names As Variant, dim_ As Variant, val As Variant
    Dim pv As Object, bv As Object, m1 As String, m2 As String
    Dim preRows As Double, baseRows As Double, preM1 As Double, baseM1 As Double
    Dim preM2 As Double, baseM2 As Double, share As Variant

    If preAgg Is Nothing Then Exit Sub
    If plan Is Nothing Then Exit Sub
    If Not CBool(plan("On")) Then Exit Sub

    Set pre = BreakdownOf(preAgg)
    If pre.count = 0 Then Exit Sub
    Set base = BreakdownOf(baseAgg)

    names = PlanMeasureNames(plan)
    m1 = "": m2 = ""
    If IsArray(names) Then
        If UBound(names) >= 0 Then m1 = CStr(names(0))
        If UBound(names) >= 1 Then m2 = CStr(names(1))
    End If

    For Each dim_ In pre.keys
        For Each val In pre(dim_).keys
            Set pv = pre(dim_)(val)
            Set bv = Nothing
            If base.Exists(dim_) Then
                If base(dim_).Exists(val) Then Set bv = base(dim_)(val)
            End If
            preRows = MapNum(pv, "#ROWS"): baseRows = MapNum(bv, "#ROWS")
            preM1 = MapNum(pv, m1): baseM1 = MapNum(bv, m1)
            preM2 = MapNum(pv, m2): baseM2 = MapNum(bv, m2)
            ' The share is the whole point: 31% of TRADE and all of INDUSTRY says
            ' more about why the figure is what it is than either amount does.
            share = Empty
            If baseM1 <> 0 Then share = preM1 / baseM1
            rowsOut.Add Array(tc("Date"), tc("EntityCode"), tc("TestCaseCode"), el("ScenarioCode"), _
                              srcKey, CStr(dim_), CStr(val), _
                              preRows, baseRows, preM1, baseM1, share, _
                              preM2, baseM2, m1, m2, expr)
        Next val
    Next dim_
End Sub

Private Function MapNum(ByVal d As Object, ByVal k As String) As Double
    If d Is Nothing Then Exit Function
    If Len(k) = 0 Then Exit Function
    If d.Exists(k) Then MapNum = BrkNum(d(k))
End Function

' A measure column can hold text, an error or a blank on any given row, and a
' breakdown that stops on the first of those is worse than no breakdown.
Private Function BrkNum(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    BrkNum = CDbl(v)
End Function

' The hot loop, and it is HOT: this runs on every row every filter matches, four
' times over, across forty aggregates. Two and a half million visits on a real
' extract.
'
' The first version built its dictionary key by concatenation right here -
' prefix, dimension, value, measure - which is four string allocations per visit
' and ten million per run. It turned a four minute pass into one that had not
' finished after twenty-five. Nothing was wrong with it except that it did the
' same work over and over: a dimension has tens of distinct values, not millions.
'
' So a value is resolved to a SLOT once, and from then on this is a dictionary
' lookup and three array additions. The slots are flattened into the aggregate's
' string keys at the end, a few hundred times instead of a few million.
Private Sub BreakdownRow(ByVal r As Long)
    Dim i As Long, j As Long, v As String, slot As Long, d As Object
    For i = 1 To mBrkDimCount
        v = SafeText(mEclData(r, mBrkDimCol(i)))
        If Len(v) = 0 Then v = "(blank)"
        If Len(v) > 120 Then v = Left$(v, 120)
        Set d = mBrkMap(i)
        If d.Exists(v) Then
            slot = d(v)
        Else
            slot = NewBreakdownSlot(i, v)
            If slot = 0 Then GoTo NextDim
            d.Add v, slot
        End If
        mBrkSum(slot, 0) = mBrkSum(slot, 0) + 1
        For j = 1 To mBrkMeasureCount
            mBrkSum(slot, j) = mBrkSum(slot, j) + BrkNum(mEclData(r, mBrkMeasureCol(j)))
        Next j
NextDim:
    Next i
End Sub

' A new value of a dimension. The cap is a guard against a dimension that turns
' out to be an account number: a breakdown with ninety thousand values in it is
' not an explanation, and the run should not slow to a crawl discovering that.
Private Function NewBreakdownSlot(ByVal dimIx As Long, ByVal v As String) As Long
    If mBrkSlotCount >= MAX_BREAKDOWN_VALUES Then Exit Function
    mBrkSlotCount = mBrkSlotCount + 1
    mBrkSlotDim(mBrkSlotCount) = dimIx
    mBrkSlotVal(mBrkSlotCount) = v
    NewBreakdownSlot = mBrkSlotCount
End Function

' Fresh accumulators for one aggregate.
Private Sub BeginBreakdown()
    Dim i As Long
    mBrkSlotCount = 0
    If Not mBreakdownOn Then Exit Sub
    ReDim mBrkMap(1 To mBrkDimCount)
    For i = 1 To mBrkDimCount
        Set mBrkMap(i) = NewMap()
    Next i
    ' VBA cannot preserve values while resizing a multidimensional array's
    ' first dimension. Allocate the existing bounded capacity once instead.
    ' The numeric array is 4,000 x 3 Doubles (96 KB), including row counts.
    ReDim mBrkSlotDim(1 To MAX_BREAKDOWN_VALUES)
    ReDim mBrkSlotVal(1 To MAX_BREAKDOWN_VALUES)
    ReDim mBrkSum(1 To MAX_BREAKDOWN_VALUES, 0 To MAX_BREAKDOWN_MEASURES)
End Sub

' The slots, written into the aggregate under the keys BreakdownOf reads back.
' A few hundred string builds, once, rather than millions in the row loop.
Private Sub FlushBreakdown(ByVal agg As Object)
    Dim s As Long, j As Long, stem As String
    If Not mBreakdownOn Then Exit Sub
    For s = 1 To mBrkSlotCount
        stem = BRK_PREFIX & mBrkDimName(mBrkSlotDim(s)) & "|" & mBrkSlotVal(s) & "|"
        agg(stem & "#ROWS") = mBrkSum(s, 0)
        For j = 1 To mBrkMeasureCount
            agg(stem & mBrkMeasureName(j)) = mBrkSum(s, j)
        Next j
    Next s
    mBrkSlotCount = 0
End Sub

' Resolves the chosen dimensions and measures against the source that is loaded
' NOW. Called once per source activation, never per row.
Private Sub PrepareBreakdown()
    Dim ws As Worksheet, r As Long, lastRow As Long, nm As String, col As Long
    Dim k As Variant, n As Long

    mBrkDimCount = 0: mBrkMeasureCount = 0
    mBreakdownOn = False
    If Not BreakdownsEnabled() Then Exit Sub
    If mEclHeaders Is Nothing Then Exit Sub

    ReDim mBrkDimName(1 To MAX_BREAKDOWN_DIMS)
    ReDim mBrkDimCol(1 To MAX_BREAKDOWN_DIMS)
    ReDim mBrkMeasureName(1 To MAX_BREAKDOWN_MEASURES)
    ReDim mBrkMeasureCol(1 To MAX_BREAKDOWN_MEASURES)

    ' The dimensions marked on the mapping sheet, in its own order, skipping any
    ' the loaded extract does not actually carry.
    On Error Resume Next
    Set ws = PsFieldsSheet()
    On Error GoTo 0
    ' Anything still undecided gets the default answer first, so a workspace that
    ' predates this column is not silently a workspace with no breakdown.
    SeedBreakdownColumn ws
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
        For r = PS_MAP_FIRST_ROW To lastRow
            If mBrkDimCount >= MAX_BREAKDOWN_DIMS Then Exit For
            nm = SafeUpperText(ws.Cells(r, 1).Value2)
            If Len(nm) > 0 Then
                If SafeUpperText(ws.Cells(r, PS_MAP_BREAK_COL).Value2) = "YES" Then
                    col = FirstExistingEclColumn(SafeText(ws.Cells(r, 2).Value2))
                    If col > 0 Then
                        mBrkDimCount = mBrkDimCount + 1
                        mBrkDimName(mBrkDimCount) = nm
                        mBrkDimCol(mBrkDimCount) = col
                    End If
                End If
            End If
        Next r
    End If
    If mBrkDimCount = 0 Then Exit Sub

    ' The measures the enabled metrics for this source actually reference, in the
    ' order they were first named. Two of them: the breakdown is there to explain
    ' a figure, not to reproduce every figure at once.
    If Not mPreShockMeasureFields Is Nothing Then
        For Each k In mPreShockMeasureFields.keys
            If mBrkMeasureCount >= MAX_BREAKDOWN_MEASURES Then Exit For
            col = CLng(mPreShockMeasureFields(k))
            If col > 0 Then
                mBrkMeasureCount = mBrkMeasureCount + 1
                mBrkMeasureName(mBrkMeasureCount) = SafeUpperText(CStr(k))
                mBrkMeasureCol(mBrkMeasureCount) = col
            End If
        Next k
    End If

    mBreakdownOn = (mBrkDimCount > 0)
End Sub

' Reads the breakdown entries back out of an aggregate, as
' dimension -> value -> { Rows, <measure> -> amount }.
Private Function BreakdownOf(ByVal agg As Object) As Object
    Dim out As Object, k As Variant, s As String, parts As Variant, d As Object, v As Object
    Set out = NewMap()
    Set BreakdownOf = out
    If agg Is Nothing Then Exit Function
    For Each k In agg.keys
        s = CStr(k)
        If Left$(s, Len(BRK_PREFIX)) = BRK_PREFIX Then
            parts = Split(Mid$(s, Len(BRK_PREFIX) + 1), "|")
            If UBound(parts) >= 2 Then
                If Not out.Exists(CStr(parts(0))) Then Set out(CStr(parts(0))) = NewMap()
                Set d = out(CStr(parts(0)))
                If Not d.Exists(CStr(parts(1))) Then Set d(CStr(parts(1))) = NewMap()
                Set v = d(CStr(parts(1)))
                v(CStr(parts(2))) = BrkNum(agg(k))
            End If
        End If
    Next k
End Function

' The plan this source was broken down by, kept so it can be restored when the
' pass moves back to that source for the next test case. Resolving it again per
' test case would re-read the mapping sheet a few hundred times for an answer
' that cannot have changed.
Private Function SnapshotBreakdownPlan() As Object
    Dim p As Object, i As Long, dims As Variant, cols As Variant
    Dim mn As Variant, mc As Variant
    Set p = NewMap()
    p("On") = mBreakdownOn
    p("DimCount") = mBrkDimCount
    p("MeasureCount") = mBrkMeasureCount
    If mBrkDimCount > 0 Then
        ReDim dims(1 To mBrkDimCount): ReDim cols(1 To mBrkDimCount)
        For i = 1 To mBrkDimCount
            dims(i) = mBrkDimName(i): cols(i) = mBrkDimCol(i)
        Next i
        p("DimNames") = dims: p("DimCols") = cols
    End If
    If mBrkMeasureCount > 0 Then
        ReDim mn(1 To mBrkMeasureCount): ReDim mc(1 To mBrkMeasureCount)
        For i = 1 To mBrkMeasureCount
            mn(i) = mBrkMeasureName(i): mc(i) = mBrkMeasureCol(i)
        Next i
        p("MeasureNames") = mn: p("MeasureCols") = mc
    End If
    Set SnapshotBreakdownPlan = p
End Function

Private Sub RestoreBreakdownPlan(ByVal p As Object)
    Dim i As Long
    mBreakdownOn = False: mBrkDimCount = 0: mBrkMeasureCount = 0
    If p Is Nothing Then Exit Sub
    mBreakdownOn = CBool(p("On"))
    mBrkDimCount = CLng(p("DimCount"))
    mBrkMeasureCount = CLng(p("MeasureCount"))
    If mBrkDimCount > 0 Then
        ReDim mBrkDimName(1 To mBrkDimCount): ReDim mBrkDimCol(1 To mBrkDimCount)
        For i = 1 To mBrkDimCount
            mBrkDimName(i) = CStr(p("DimNames")(i)): mBrkDimCol(i) = CLng(p("DimCols")(i))
        Next i
    End If
    If mBrkMeasureCount > 0 Then
        ReDim mBrkMeasureName(1 To mBrkMeasureCount): ReDim mBrkMeasureCol(1 To mBrkMeasureCount)
        For i = 1 To mBrkMeasureCount
            mBrkMeasureName(i) = CStr(p("MeasureNames")(i)): mBrkMeasureCol(i) = CLng(p("MeasureCols")(i))
        Next i
    End If
End Sub

Private Function PlanMeasureNames(ByVal p As Object) As Variant
    Dim out As Variant, i As Long, n As Long
    If p Is Nothing Then PlanMeasureNames = Array(): Exit Function
    n = CLng(p("MeasureCount"))
    If n = 0 Then PlanMeasureNames = Array(): Exit Function
    ReDim out(0 To n - 1)
    For i = 1 To n
        out(i - 1) = CStr(p("MeasureNames")(i))
    Next i
    PlanMeasureNames = out
End Function

Public Function BreakdownsEnabled() As Boolean
    ' Off: Trace and the joined input's per-test-case pivots explain a value.
    BreakdownsEnabled = False
End Function

Private Sub AddAgg(ByVal d As Object, ByVal k As String, ByVal v As Double)
    If d.Exists(k) Then d(k) = CDbl(d(k)) + v Else d(k) = v
End Sub
Private Function AggVal(ByVal d As Object, ByVal k As String) As Double
    If d.Exists(k) Then AggVal = CDbl(d(k))
End Function
Private Function FirstExistingEclColumn(ByVal candidates As String) As Long
    Dim p As Variant, k As String
    For Each p In Split(candidates, "|")
        k = CanonicalEclHeader(Trim$(CStr(p))): If mEclHeaders.Exists(k) Then FirstExistingEclColumn = mEclHeaders(k): Exit Function
        k = NormalHeader(Trim$(CStr(p))): If mEclHeaders.Exists(k) Then FirstExistingEclColumn = mEclHeaders(k): Exit Function
    Next p
End Function

Private Function MetricFromAggregate(ByVal agg As Object, ByVal m As Object, ByRef found As Boolean) As Variant
    Dim stage As String, expr As String
    expr = StripPreShockStage(CStr(m("Formula")), stage)
    MetricFromAggregate = EvalPreShockExpression(agg, expr, stage, found)
End Function

Private Function CompileFilter(ByVal expr As String, ByVal fieldMap As Object, ByRef unsupported As String) As Object
    Dim s As String, item As Variant, opPos As Long, opText As String, field As String, listText As String, mode As String
    Dim node As Object, child As Object, values As Object, cols As Collection, children As Collection, parts As Collection
    s = PreShockCleanFilter(expr): s = TrimOuterParens(s)
    If SafeUpperText(s) = "ALL" Or SafeUpperText(s) = "NULL" Then
        Set node = NewMap(): node("Type") = "ALL": Set CompileFilter = node: Exit Function
    End If
    Set parts = SplitTopLevel(s, " OR ")
    If parts.count > 1 Then
        Set node = NewMap(): node("Type") = "OR": Set children = New Collection
        For Each item In parts
            Set child = CompileFilter(CStr(item), fieldMap, unsupported): children.Add child
        Next item
        node.Add "Children", children: Set CompileFilter = node: Exit Function
    End If
    Set parts = SplitTopLevel(s, " AND ")
    If parts.count > 1 Then
        Set node = NewMap(): node("Type") = "AND": Set children = New Collection
        For Each item In parts
            Set child = CompileFilter(CStr(item), fieldMap, unsupported): children.Add child
        Next item
        node.Add "Children", children: Set CompileFilter = node: Exit Function
    End If
    opPos = InStr(1, UCase$(s), " NOT IN (", vbBinaryCompare): If opPos > 0 Then opText = "NOT IN"
    If opPos = 0 Then opPos = InStr(1, UCase$(s), " IN (", vbBinaryCompare): If opPos > 0 Then opText = "IN"
    If opPos = 0 Then Err.Raise vbObjectError + 722, , "Not tested - unsupported filter condition: " & s
    field = SafeUpperText(Trim$(Left$(s, opPos - 1)))
    If opText = "NOT IN" Then listText = Mid$(s, opPos + 9) Else listText = Mid$(s, opPos + 5)
    listText = Trim$(listText): If Left$(listText, 1) = "(" Then listText = Mid$(listText, 2)
    If Right$(listText, 1) = ")" Then listText = Left$(listText, Len(listText) - 1)
    Set node = NewMap(): node("Type") = "ATOM": node("Field") = field: node("Operator") = opText
    If fieldMap.Exists(field) Then
        node("Candidates") = fieldMap(field)("Candidates"): mode = CStr(fieldMap(field)("Mode")): node("Mode") = mode
        Set cols = ResolveFilterColumns(CStr(node("Candidates")))
        If cols.count = 0 Then AppendUnsupported unsupported, field
    Else
        node("Candidates") = "": mode = "TEXT": node("Mode") = mode: Set cols = New Collection: AppendUnsupported unsupported, field
    End If
    Set values = ParseValueList(listText, mode)
    node.Add "Values", values: node.Add "Columns", cols
    Set CompileFilter = node
End Function

Private Sub AppendUnsupported(ByRef s As String, ByVal field As String)
    If InStr(1, "," & s & ",", "," & field & ",", vbTextCompare) = 0 Then
        If Len(s) > 0 Then s = s & ", "
        s = s & field
    End If
End Sub
Private Function ResolveFilterColumns(ByVal candidates As String) As Collection
    Dim c As New Collection, seen As Object, p As Variant, k As String, ix As Long
    Set seen = NewMap()
    For Each p In Split(candidates, "|")
        k = NormalHeader(p)
        If mEclHeaders.Exists(k) Then
            ix = CLng(mEclHeaders(k)): If Not seen.Exists(CStr(ix)) Then c.Add ix: seen(CStr(ix)) = True
        Else
            k = CanonicalEclHeader(p)
            If mEclHeaders.Exists(k) Then ix = CLng(mEclHeaders(k)): If Not seen.Exists(CStr(ix)) Then c.Add ix: seen(CStr(ix)) = True
        End If
    Next p
    Set ResolveFilterColumns = c
End Function
Private Function ParseValueList(ByVal s As String, ByVal mode As String) As Object
    Dim d As Object, p As Variant, v As String: Set d = NewMap()
    For Each p In SplitTopLevel(s, ",")
        v = Trim$(CStr(p))
        If Len(v) >= 2 Then
            If (Left$(v, 1) = "'" And Right$(v, 1) = "'") Or (Left$(v, 1) = """" And Right$(v, 1) = """") Then v = Mid$(v, 2, Len(v) - 2)
        End If
        v = Replace(Replace(v, "''", "'"), """" & """", """")
        If Len(Trim$(v)) = 0 Or SafeUpperText(v) = "NULL" Then Err.Raise vbObjectError + 745, , "Empty or NULL filter-list values require an explicit rule."
        d(NormalizeFilterValue(v, mode)) = True
    Next p
    Set ParseValueList = d
End Function
Private Function FilterQuoteEnd(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long, j As Long, q As String, nextChar As String
    q = Mid$(s, startPos, 1): i = startPos + 1
    Do While i <= Len(s)
        If Mid$(s, i, 1) = q Then
            If i < Len(s) And Mid$(s, i + 1, 1) = q Then
                i = i + 2
            Else
                j = i + 1
                Do While j <= Len(s)
                    If Mid$(s, j, 1) <> " " And Mid$(s, j, 1) <> vbTab Then Exit Do
                    j = j + 1
                Loop
                nextChar = Mid$(s, j, 1)
                If j > Len(s) Or nextChar = "," Or nextChar = ")" Then FilterQuoteEnd = i: Exit Function
                ' Exported labels include unescaped apostrophes, e.g. LC's.
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    Err.Raise vbObjectError + 746, , "Unclosed quoted value in filter condition."
End Function

Private Function SplitTopLevel(ByVal s As String, ByVal sep As String) As Collection
    Dim c As New Collection, depth As Long, i As Long, start As Long, u As String, ch As String
    u = UCase$(s): start = 1: i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "'" Or ch = """" Then
            i = FilterQuoteEnd(s, i) + 1
        Else
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If depth < 0 Then Err.Raise vbObjectError + 747, , "Unbalanced parentheses in filter condition."
            If depth = 0 And Mid$(u, i, Len(sep)) = sep Then
                c.Add Trim$(Mid$(s, start, i - start)): i = i + Len(sep): start = i
            Else
                i = i + 1
            End If
        End If
    Loop
    If depth <> 0 Then Err.Raise vbObjectError + 747, , "Unbalanced parentheses in filter condition."
    c.Add Trim$(Mid$(s, start)): Set SplitTopLevel = c
End Function

Private Function TrimOuterParens(ByVal s As String) As String
    Dim depth As Long, i As Long, whole As Boolean, ch As String
again:
    s = Trim$(s)
    If Len(s) >= 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        depth = 0: whole = True: i = 1
        Do While i <= Len(s)
            ch = Mid$(s, i, 1)
            If ch = "'" Or ch = """" Then
                i = FilterQuoteEnd(s, i)
            Else
                If ch = "(" Then depth = depth + 1
                If ch = ")" Then depth = depth - 1
                If depth = 0 And i < Len(s) Then whole = False: Exit Do
            End If
            i = i + 1
        Loop
        If whole Then s = Mid$(s, 2, Len(s) - 2): GoTo again
    End If
    TrimOuterParens = s
End Function

Private Function EvalFilter(ByVal node As Object, ByVal r As Long) As Boolean
    Dim child As Variant, col As Variant, raw As String, wanted As Object, hit As Boolean, mode As String, hasValue As Boolean
    If node("Type") = "ALL" Then EvalFilter = True: Exit Function
    If node("Type") = "AND" Then
        EvalFilter = True: For Each child In node("Children"): If Not EvalFilter(child, r) Then EvalFilter = False: Exit Function
        Next child: Exit Function
    End If
    If node("Type") = "OR" Then
        For Each child In node("Children"): If EvalFilter(child, r) Then EvalFilter = True: Exit Function
        Next child: Exit Function
    End If
    Set wanted = node("Values"): mode = CStr(node("Mode"))
    For Each col In node("Columns")
        raw = NormalizeFilterValue(mEclData(r, CLng(col)), mode)
        If Len(raw) > 0 And raw <> "NULL" Then
            hasValue = True
            If wanted.Exists(raw) Then hit = True: Exit For
        End If
    Next col
    If node("Operator") = "NOT IN" Then EvalFilter = hasValue And Not hit Else EvalFilter = hit
End Function
Private Function NormalizeFilterValue(ByVal v As Variant, ByVal mode As String) As String
    Dim s As String: s = SafeUpperText(v)
    Select Case mode
        Case "STAGE": s = NormalizeStage(v)
        Case "YESNO"
            If s = "Y" Or s = "YES" Or s = "TRUE" Or s = "1" Then
                s = "YES"
            ElseIf s = "N" Or s = "NO" Or s = "FALSE" Or s = "0" Then
                s = "NO"
            End If
        Case "UNDERSCORE_TEXT": s = Replace(s, "_", " ")
        Case "SECTOR"
            s = Replace(s, "_", " ")
            Do While Len(s) > 0 And Mid$(s, 1, 1) Like "[0-9 ]": s = Mid$(s, 2): Loop
            If InStrRev(s, "(") > 0 And Right$(s, 1) = ")" Then s = Trim$(Left$(s, InStrRev(s, "(") - 1))
    End Select
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormalizeFilterValue = Trim$(s)
End Function
Private Function NormalizeStage(ByVal v As Variant) As String
    Dim s As String, digit As String, tail As String
    s = SafeUpperText(v)
    If s = "NPA" Then NormalizeStage = "3": Exit Function
    If s = "1" Or s = "2" Or s = "3" Then NormalizeStage = s: Exit Function
    s = Replace(s, "STAGE ", "STAGE")
    If Left$(s, 5) <> "STAGE" Or Len(s) < 6 Then Exit Function
    digit = Mid$(s, 6, 1)
    If digit <> "1" And digit <> "2" And digit <> "3" Then Exit Function
    tail = Mid$(s, 7)
    If Len(tail) > 0 Then
        If Left$(tail, 1) <> " " And Left$(tail, 1) <> "-" Then Exit Function
    End If
    NormalizeStage = digit
End Function

' Source date validation is independent of test-case filters. Keep every row's
' canonical date and the distinct-date inventory, including a validation error,
' until that source is replaced or a new reconciliation starts.
Private Sub ClearSourceDateCache(Optional ByVal sourceKey As String = "")
    If Len(sourceKey) = 0 Then
        Set mSourceDateCache = Nothing
    ElseIf Not mSourceDateCache Is Nothing Then
        If mSourceDateCache.Exists(SafeUpperText(sourceKey)) Then mSourceDateCache.Remove SafeUpperText(sourceKey)
    End If
End Sub

Private Function SourceDateInventory() As Object
    Dim info As Object, dates As Object, rowDates() As String
    Dim key As String, r As Long, col As Long, dk As String, summary As String
    Dim errorNumber As Long, errorText As String
    key = SafeUpperText(mActiveSource)
    If mSourceDateCache Is Nothing Then Set mSourceDateCache = NewMap()
    If mSourceDateCache.Exists(key) Then
        Set info = mSourceDateCache(key)
        If info.Exists("ErrorNumber") Then Err.Raise CLng(info("ErrorNumber")), "SourceDateInventory", CStr(info("ErrorText"))
        Set SourceDateInventory = info
        Exit Function
    End If
    Set info = NewMap(): Set dates = NewMap()
    On Error GoTo InvalidSource
    If Not mEclHeaders.Exists("AS_OF_DATE") Then Err.Raise vbObjectError + 741, , mActiveSource & ": AS_OF_DATE is required for date-matched reconciliation."
    col = CLng(mEclHeaders("AS_OF_DATE"))
    ReDim rowDates(1 To UBound(mEclData, 1))
    For r = 2 To UBound(mEclData, 1)
        If r Mod 20000 = 0 Then ProgressDetail "Validating " & mActiveSource & " reporting dates: " & format$(r - 1, "#,##0") & " of " & format$(UBound(mEclData, 1) - 1, "#,##0") & " rows"
        dk = CheckedSourceDate(mEclData(r, col), r)
        rowDates(r) = dk
        If Not dates.Exists(dk) Then
            dates(dk) = True
            If dates.count <= 12 Then
                If Len(summary) > 0 Then summary = summary & ", "
                summary = summary & dk
            ElseIf dates.count = 13 Then
                summary = summary & ", ..."
            End If
        End If
    Next r
    If Len(summary) = 0 Then summary = "(no dated source rows)"
    Set info("Dates") = dates
    info("Rows") = rowDates
    info("Summary") = summary
    Set mSourceDateCache(key) = info
    Set SourceDateInventory = info
    Exit Function
InvalidSource:
    errorNumber = Err.Number: errorText = Err.description
    Set info = NewMap()
    info("ErrorNumber") = errorNumber: info("ErrorText") = errorText
    Set mSourceDateCache(key) = info
    Err.Raise errorNumber, "SourceDateInventory", errorText
End Function

Private Sub RequireSourceDate(ByVal info As Object, ByVal aod As String)
    If info("Dates").Exists(aod) Then Exit Sub
    Err.Raise vbObjectError + 724, , "As-of date mismatch: the " & mActiveSource & " file holds " & CStr(info("Summary")) & " but this test case is as-of " & aod & _
        ". Upload the matching reporting-date extract, or set 'As-of date matching' to Ignore to test against the file as supplied."
End Sub

Private Function CheckedSourceDate(ByVal value As Variant, ByVal SourceRow As Long) As String
    Dim detail As String
    On Error GoTo InvalidDate
    CheckedSourceDate = DateKey(value, mEclDate1904)
    Exit Function
InvalidDate:
    detail = Err.description
    Err.Raise vbObjectError + 743, "CheckedSourceDate", mActiveSource & ": invalid AS_OF_DATE at source row " & (SourceRow + mEclHeaderRow - 1) & ". " & detail
End Function

' ============================ sheet navigation ==============================

' ============================= source registry ==============================

Public Function SourceKeys() As Variant
    ' The order matters only for display. New extracts are added here and to the
    ' Source column of the metric rules; nothing else needs to change.
    SourceKeys = Array("ECL", "CAPRWA", "LL", "LCR", "CAP", "NSFR")
End Function

Private Function SourceLabel(ByVal key As String) As String
    Select Case SafeUpperText(key)
        Case "ECL": SourceLabel = "ECL output - outstanding, IIS, ECL by stage (also carries max ECL and impact max ECL)"
        Case "CAPRWA": SourceLabel = "CAPRWA output - RWA and RWA exposure by stage"
        Case "LL": SourceLabel = "Legal liquidity ALM output - assets and liabilities by bucket"
        Case "LCR": SourceLabel = "LCR ALM output - HQLA, outflow and inflow by bucket"
        Case "CAP": SourceLabel = "Capital component - base capital by element (not deal level)"
        Case "NSFR": SourceLabel = "NSFR ALM output - available and required stable funding"
        Case Else: SourceLabel = key
    End Select
End Function

' The three new extracts carry no IFRS stage, and that is not a fault in them: a
' liquidity cashflow and a capital component are not staged things. They are
' aggregated under a stage of "0", which nothing else ever produces, so a metric
' written for them asks for stage ALL and a metric written for ECL is unaffected.
Public Function SourceIsStaged(ByVal key As String) As Boolean
    Select Case SafeUpperText(key)
        Case "LL", "LCR", "CAP", "NSFR": SourceIsStaged = False
        Case Else: SourceIsStaged = True
    End Select
End Function

Private Sub EnsureSourceStore()
    If mSources Is Nothing Then Set mSources = NewMap()
End Sub

' Points the engine helpers at one source's table. Everything downstream - filter
' evaluation, aggregation, the value cache - reads whichever source is active.
Private Function ActivateSource(ByVal key As String) As Boolean
    Dim s As Object
    EnsureSourceStore
    If Not mSources.Exists(SafeUpperText(key)) Then Exit Function
    Set s = mSources(SafeUpperText(key))
    ' Already holding this very table: keep it. Copying a whole extract is the
    ' dearest thing the engine did, and a run switched source dozens of times
    ' per test case.
    If Not (s Is mActiveRecord) Then TakeTable s
    Set mEclHeaders = s("Headers")
    mEclPath = CStr(s("Path"))
    mEclSheet = CStr(s("Sheet"))
    mEclHeaderRow = CLng(s("HeaderRow"))
    mEclDate1904 = CBool(s("Date1904"))
    mActiveSource = SafeUpperText(key)
    ActivateSource = True
End Function

Private Sub StoreActiveSource(ByVal key As String)
    Dim s As Object
    EnsureSourceStore
    Set s = NewMap()
    KeepActiveAs s, key
    Set s("Headers") = mEclHeaders
    s("Path") = mEclPath
    s("Sheet") = mEclSheet
    s("HeaderRow") = mEclHeaderRow
    s("Date1904") = mEclDate1904
    StampRecord s, mEclData
    DropTableMemo key
    If mSources.Exists(SafeUpperText(key)) Then mSources.Remove SafeUpperText(key)
    mSources.Add SafeUpperText(key), s
    Set mActiveRecord = s
    ClearSourceDateCache key
    Set mJoinCache = Nothing
    mActiveSource = SafeUpperText(key)
End Sub

Public Function SourceLoaded(ByVal key As String) As Boolean
    EnsureSourceStore
    SourceLoaded = mSources.Exists(SafeUpperText(key))
End Function

' ---- what the build's own check needs to see -------------------------------
'
' These exist so the five-source mapping can be proved against the real files in
' the build rather than by opening the workbook and reading numbers off a screen.

' ---- what the joined-input builder borrows ---------------------------------
'
' The joined table is registered as an ordinary source so that the filter
' compiler, the evaluator and the field map all work on it unchanged. Writing a
' second evaluator for it would mean two things that have to agree about what a
' condition means, and they would not stay in agreement.

Public Sub RegisterTable(ByVal key As String, ByRef data As Variant, ByVal headers As Object, _
                         ByVal path As String, ByVal sheetName As String)
    Dim s As Object
    EnsureSourceStore
    Set s = NewMap()
    Set s("Table") = TableFrom(data)
    Set s("Headers") = headers
    s("Path") = path
    s("Sheet") = sheetName
    s("HeaderRow") = 1
    s("Date1904") = False
    StampRecord s, data
    DropTableMemo key
    If mSources.Exists(SafeUpperText(key)) Then mSources.Remove SafeUpperText(key)
    mSources.Add SafeUpperText(key), s
    ClearSourceDateCache key
    Set mJoinCache = Nothing
End Sub

Public Function ActivateFor(ByVal key As String) As Boolean
    ActivateFor = ActivateSource(key)
End Function

Public Function CompileForActive(ByVal expr As String, ByRef unsupported As String) As Object
    Set CompileForActive = CompileFilter(expr, LoadEclFilterMap(), unsupported)
End Function

Public Function EvalForActive(ByVal node As Object, ByVal r As Long) As Boolean
    EvalForActive = EvalFilter(node, r)
End Function

Public Function ActiveData() As Variant
    ActiveData = mEclData
End Function

Public Function ActiveHeaders() As Object
    Set ActiveHeaders = mEclHeaders
End Function

Public Function SourceHeadersFor(ByVal key As String) As Object
    EnsureSourceStore
    If Not mSources.Exists(SafeUpperText(key)) Then Exit Function
    Set SourceHeadersFor = mSources(SafeUpperText(key))("Headers")
End Function

Public Function SourceDataFor(ByVal key As String) As Variant
    EnsureSourceStore
    If Not mSources.Exists(SafeUpperText(key)) Then Exit Function
    SourceDataFor = TableCopyOf(mSources(SafeUpperText(key)))
End Function

Public Function SourceRowCount(ByVal key As String) As Double
    Dim s As Object, d As Variant
    EnsureSourceStore
    If Not mSources.Exists(SafeUpperText(key)) Then Exit Function
    Set s = mSources(SafeUpperText(key))
    ' Stored with the table, so counting rows no longer copies them.
    If s.Exists("Rows") Then SourceRowCount = CDbl(s("Rows")): Exit Function
    d = TableCopyOf(s)
    If IsArray(d) Then SourceRowCount = UBound(d, 1) - 1
End Function

Public Function SourceFamilyFor(ByVal key As String) As String
    Dim s As Object
    EnsureSourceStore
    If Not mSources.Exists(SafeUpperText(key)) Then SourceFamilyFor = "(not loaded)": Exit Function
    Set s = mSources(SafeUpperText(key))
    SourceFamilyFor = SourceFamilyOf(s("Headers"))
End Function

' One metric over its source's whole population, under whatever entity scope is
' currently set. No test-case filter: this is the base figure.
Public Function MetricValueForDiag(ByVal metric As String) As Double
    Dim mm As Object, m As Object, src As String, w As String, stage As String
    Dim agg As Object, cache As Object, found As Boolean, expr As String

    Set mm = LoadMetricMap()
    If Not mm.Exists(SafeUpperText(metric)) Then Exit Function
    Set m = mm(SafeUpperText(metric))
    src = SafeUpperText(m("Source"))
    If Not SourceLoaded(src) Then Exit Function
    If Not ActivateSource(src) Then Exit Function

    SplitMetricFormula CStr(m("Formula")), stage, w
    Set mPreShockMeasureFields = BuildPreShockMeasureMap(MetricsForSource(mm, src))
    Set cache = NewMap()
    expr = modScenarioBuilder_Multi.ApplyEntityScope(w)
    On Error Resume Next
    Set agg = GetEclAggregate(cache, DateKeyOfFirstRow(), "", expr, LoadEclFilterMap())
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    If agg Is Nothing Then Exit Function
    MetricValueForDiag = CDbl(MetricFromAggregate(agg, m, found))
End Function

' One metric, one source, one filter - with the join coverage if it needed one.
Public Function JoinProbeForDiag(ByVal src As String, ByVal metric As String, ByVal expr As String) As String
    Dim mm As Object, m As Object, w As String, stage As String, agg As Object, cache As Object
    Dim found As Boolean, probeCase As String, probeLocal As String, rows_ As Double, s As String

    Set mm = LoadMetricMap()
    If Not mm.Exists(SafeUpperText(metric)) Then JoinProbeForDiag = metric & ": not on the metric sheet": Exit Function
    Set m = mm(SafeUpperText(metric))
    If Not SourceLoaded(src) Then JoinProbeForDiag = src & ": not loaded": Exit Function
    If Not ActivateSource(src) Then JoinProbeForDiag = src & ": could not activate": Exit Function

    SplitMetricFormula CStr(m("Formula")), stage, w
    Set mPreShockMeasureFields = BuildPreShockMeasureMap(MetricsForSource(mm, src))
    probeCase = expr
    probeLocal = modScenarioBuilder_Multi.ApplyEntityScope(w)
    Set cache = NewMap()
    On Error Resume Next
    Set agg = GetEclAggregate(cache, DateKeyOfFirstRow(), probeCase, probeLocal, LoadEclFilterMap())
    If Err.Number <> 0 Then
        JoinProbeForDiag = PadDiag(src & " " & metric, 46) & "REFUSED: " & Err.description
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    rows_ = AggVal(agg, "ROWS#0") + AggVal(agg, "ROWS#1") + AggVal(agg, "ROWS#2") + AggVal(agg, "ROWS#3")
    s = PadDiag(src & " " & metric, 46) & PadDiag(format$(CDbl(MetricFromAggregate(agg, m, found)), "#,##0"), 20) & _
        PadDiag(format$(rows_, "#,##0") & " row(s)", 16)
    If AggVal(agg, "JOIN#WANTED") > 0 Then
        s = s & "joined: " & format$(AggVal(agg, "JOIN#FOUND"), "#,##0") & " of " & _
                format$(AggVal(agg, "JOIN#WANTED"), "#,##0") & " account(s)"
    Else
        s = s & "no join needed"
    End If
    JoinProbeForDiag = s
End Function

Private Function PadDiag(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then PadDiag = Left$(s, n - 1) & " " Else PadDiag = s & Space$(n - Len(s))
End Function

Public Function DescribeMetricsForDiag(ByVal srcFilter As String) As String
    Dim mm As Object, k As Variant, s As String, src As String, w As String, stage As String, bare As String
    Set mm = LoadMetricMap()
    For Each k In mm.keys
        src = SafeUpperText(mm(k)("Source"))
        If InStr(1, "|" & UCase$(srcFilter) & "|", "|" & src & "|") > 0 Then
            bare = SplitMetricFormula(CStr(mm(k)("Formula")), stage, w)
            s = s & "   " & CStr(k) & vbCrLf & _
                    "      source=" & src & "  stage=" & stage & vbCrLf & _
                    "      sums  =" & bare & vbCrLf & _
                    "      where =" & IIf(Len(w) = 0, "(none)", w) & vbCrLf
        End If
    Next k
    If Len(s) = 0 Then s = "   (no metric rows reference LL, LCR or CAP - the seeds did not land)"
    DescribeMetricsForDiag = s
End Function

Public Function TraceMetricForDiag(ByVal metric As String) As String
    Dim mm As Object, m As Object, src As String, w As String, stage As String, s As String
    Dim agg As Object, cache As Object, found As Boolean, expr As String, k As Variant, n As Long
    Set mm = LoadMetricMap()
    If Not mm.Exists(SafeUpperText(metric)) Then TraceMetricForDiag = "   " & metric & ": not on the metric sheet": Exit Function
    Set m = mm(SafeUpperText(metric))
    src = SafeUpperText(m("Source"))
    s = "   " & metric & "  source=" & src & vbCrLf
    If Not SourceLoaded(src) Then TraceMetricForDiag = s & "      source not loaded": Exit Function
    ActivateSource src
    s = s & "      headers seen: " & mEclHeaders.count & vbCrLf
    s = s & "      has CASHFLOW_AMOUNT_LCY_POST_FACTOR: " & mEclHeaders.Exists("CASHFLOW_AMOUNT_LCY_POST_FACTOR") & vbCrLf
    s = s & "      has ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY: " & mEclHeaders.Exists("ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY") & vbCrLf
    s = s & "      has BRANCH_CODE: " & mEclHeaders.Exists("BRANCH_CODE") & vbCrLf
    s = s & "      has STAGE: " & mEclHeaders.Exists("STAGE") & vbCrLf

    SplitMetricFormula CStr(m("Formula")), stage, w
    Set mPreShockMeasureFields = BuildPreShockMeasureMap(MetricsForSource(mm, src))
    s = s & "      measure map: "
    For Each k In mPreShockMeasureFields.keys
        s = s & CStr(k) & "->col" & CStr(mPreShockMeasureFields(k)) & "  "
    Next k
    s = s & vbCrLf

    expr = modScenarioBuilder_Multi.ApplyEntityScope(w)
    s = s & "      expression: " & expr & vbCrLf
    s = s & "      as-of key: " & DateKeyOfFirstRow() & vbCrLf
    Set cache = NewMap()
    On Error Resume Next
    Set agg = GetEclAggregate(cache, DateKeyOfFirstRow(), "", expr, LoadEclFilterMap())
    If Err.Number <> 0 Then
        s = s & "      AGGREGATE RAISED: " & Err.description & vbCrLf
        Err.Clear
        TraceMetricForDiag = s
        Exit Function
    End If
    On Error GoTo 0
    s = s & "      agg keys: " & agg.count & vbCrLf
    For Each k In agg.keys
        If n < 14 Then s = s & "         " & CStr(k) & " = " & format$(agg(k), "#,##0.##") & vbCrLf: n = n + 1
    Next k
    s = s & "      value: " & format$(CDbl(MetricFromAggregate(agg, m, found)), "#,##0") & "   found=" & found
    TraceMetricForDiag = s
End Function

' The as-of date the active source actually holds, so the diag is not rejected
' for a date mismatch against a test case it is not using.
Private Function DateKeyOfFirstRow() As String
    If mEclHeaders Is Nothing Then Exit Function
    If Not mEclHeaders.Exists("AS_OF_DATE") Then Exit Function
    If Not IsArray(mEclData) Then Exit Function
    If UBound(mEclData, 1) < 2 Then Exit Function
    On Error Resume Next
    DateKeyOfFirstRow = DateKey(mEclData(2, mEclHeaders("AS_OF_DATE")), mEclDate1904)
    Err.Clear
End Function

Public Function AnySourceLoaded() As Boolean
    Dim k As Variant
    EnsureSourceStore
    For Each k In SourceKeys()
        If mSources.Exists(SafeUpperText(k)) Then AnySourceLoaded = True: Exit Function
    Next k
End Function

Private Function LoadedSourceSummary() As String
    Dim k As Variant, s As String, n As Long
    EnsureSourceStore
    For Each k In SourceKeys()
        If mSources.Exists(SafeUpperText(k)) Then
            n = n + 1
            If Len(s) > 0 Then s = s & ", "
            s = s & CStr(k)
        End If
    Next k
    If n = 0 Then LoadedSourceSummary = "Nothing loaded yet" Else LoadedSourceSummary = n & " source(s) loaded: " & s
End Function

' The row on the control sheet that describes one source.
Private Function SourceRow(ByVal key As String) As Long
    Dim i As Long, k As Variant
    i = 0
    For Each k In SourceKeys()
        If SafeUpperText(k) = SafeUpperText(key) Then SourceRow = PS_SRC_LIST_ROW + i: Exit Function
        i = i + 1
    Next k
End Function

' ========================= dimension-value cache ============================

Private Function CacheSheet(Optional ByVal createIfMissing As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = GetWorksheetSafe(ThisWorkbook, CACHE_SHEET)
    If ws Is Nothing And createIfMissing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = CACHE_SHEET
    End If
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Visible = xlSheetVeryHidden
        On Error GoTo 0
    End If
    Set CacheSheet = ws
End Function

' Called once per source, while that extract is still open, so the builder never has to scan
' again.
'
' The cache is PER SOURCE. It used to clear the whole sheet on every upload, which meant
' loading CAPRWA second silently erased everything cached from ECL - the builder then had no
' ECL fields to offer and the ECL metrics came back blank. Rows now carry their source key in
' column D, and only that source's rows are replaced.
Private Sub BuildValueCache()
    Dim ws As Worksheet, psh As Worksheet, r As Long, nm As String, cand As String
    Dim cols As Collection, col As Long, rr As Long, seen As Object, k As Variant
    Dim out As Collection, line As Variant, block As Variant, i As Long, c As Long, hdr As String
    Dim keep As Collection, a As Variant, lastRow As Long, srcK As String

    srcK = SafeUpperText(mActiveSource)
    If Len(srcK) = 0 Then srcK = "ECL"
    Set psh = PsFieldsSheet()
    Set ws = CacheSheet(True)
    If SafeText(ws.Range("H1").Value2) <> CACHE_LAYOUT Then
        ws.Cells.Clear
        WriteCacheHeader ws
    End If
    WriteCacheSourceRow ws, srcK

    ' Everything already cached for OTHER sources is carried across untouched.
    Set keep = New Collection
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow >= CACHE_FIRST_ROW Then
        a = ws.Range(ws.Cells(CACHE_FIRST_ROW, 1), ws.Cells(lastRow, 4)).Value2
        For r = 1 To UBound(a, 1)
            If Len(SafeText(a(r, 1))) > 0 Then
                If SafeUpperText(a(r, 4)) <> srcK Then
                    keep.Add Array(SafeText(a(r, 1)), SafeText(a(r, 2)), SafeText(a(r, 3)), SafeText(a(r, 4)))
                End If
            End If
        Next r
    End If

    Set out = New Collection
    For Each line In keep
        out.Add line
    Next line

    ' One pass per mapped dimension. Capped, because a picker showing 50,000 account
    ' numbers helps nobody and costs real time to serialise.
    For r = PS_MAP_FIRST_ROW To PS_MAP_LAST_ROW
        nm = SafeText(psh.Cells(r, 1).Value2)
        If Len(nm) = 0 Then Exit For
        cand = SafeText(psh.Cells(r, 2).Value2)
        If Len(cand) > 0 Then
            Set cols = ResolveFilterColumns(cand)
            If cols.count > 0 Then
                col = CLng(cols.item(1))
                Set seen = NewMap()
                ' Sampled. A picker showing the first few hundred values of a
                ' dimension is as useful as one showing every value, and walking
                ' seventy thousand rows for a dozen distinct codes is time spent
                ' for nothing.
                For rr = 2 To WorksheetFunction.Min(UBound(mEclData, 1), CACHE_SAMPLE_ROWS)
                    If seen.count >= CACHE_MAX_VALUES Then Exit For
                    If Len(SafeText(mEclData(rr, col))) > 0 Then seen(SafeText(mEclData(rr, col))) = True
                Next rr
                For Each k In seen.keys
                    out.Add Array("DIM", nm, CStr(k), srcK)
                Next k
            End If
        End If
    Next r

    ' Numeric columns, for the metric side of the builder. Detected once here rather than
    ' re-sniffed every time the window opens.
    For c = 1 To UBound(mEclData, 2)
        hdr = SafeText(mEclData(1, c))
        If Len(hdr) > 0 Then
            If PreShockColumnLooksNumeric(c) Then out.Add Array("MEASURE", hdr, "", srcK)
        End If
    Next c

    ' Every field name the extract carries - including the columns whose data was
    ' not loaded - so the builder can offer any field in the file.
    If Not mEclAllFields Is Nothing Then
        For Each k In mEclAllFields.keys
            out.Add Array("FIELD", CStr(k), "", srcK)
        Next k
    Else
        For c = 1 To UBound(mEclData, 2)
            hdr = SafeText(mEclData(1, c))
            If Len(hdr) > 0 Then out.Add Array("FIELD", hdr, "", srcK)
        Next c
    End If

    ws.Range(ws.Cells(CACHE_FIRST_ROW, 1), ws.Cells(ws.rows.count, 4)).ClearContents
    If out.count > 0 Then
        ReDim block(1 To out.count, 1 To 4)
        i = 0
        For Each line In out
            i = i + 1
            block(i, 1) = line(0): block(i, 2) = line(1): block(i, 3) = line(2): block(i, 4) = line(3)
        Next line
        WriteLiteralValues ws.Cells(CACHE_FIRST_ROW, 1).Resize(out.count, 4), block
    End If
End Sub

Private Sub WriteCacheHeader(ByVal ws As Worksheet)
    ws.Range("A1").Value2 = "JKB value cache - one row block per loaded source. Safe to delete; it is rebuilt on upload."
    ws.Range("A2").Value2 = "Source"
    ws.Range("B2").Value2 = "File"
    ws.Range("C2").Value2 = "Sheet"
    ws.Range("D2").Value2 = "Header row"
    ws.Range("E2").Value2 = "Data rows"
    ws.Range("F2").Value2 = "As-of dates"
    ws.Range("G2").Value2 = "Built"
    ws.Range("H1").Value2 = CACHE_LAYOUT
End Sub

' The per-source descriptor block sits at rows 3..8, one row per source key, in registry
' order, so a reload overwrites exactly its own row.
Private Sub WriteCacheSourceRow(ByVal ws As Worksheet, ByVal srcK As String)
    Dim i As Long, k As Variant, r As Long
    r = 0
    i = 0
    For Each k In SourceKeys()
        If SafeUpperText(k) = srcK Then r = 3 + i
        i = i + 1
    Next k
    If r = 0 Then r = 3 + i
    ws.Cells(r, 1).Value2 = srcK
    ws.Cells(r, 2).Value2 = mEclPath
    ws.Cells(r, 3).Value2 = mEclSheet
    ws.Cells(r, 4).Value2 = mEclHeaderRow
    ws.Cells(r, 5).Value2 = UBound(mEclData, 1) - 1
    ws.Cells(r, 6).Value2 = EclDateSummary()
    ws.Cells(r, 7).value = Now
End Sub

' True when a usable cache exists, so the builder can open without touching any extract.
Public Function ValueCacheReady() As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = CacheSheet()
    If ws Is Nothing Then Exit Function
    If SafeText(ws.Range("H1").Value2) <> CACHE_LAYOUT Then Exit Function
    If Len(SafeText(ws.Cells(CACHE_FIRST_ROW, 1).Value2)) = 0 Then Exit Function
    For r = 3 To 8
        If Len(SafeText(ws.Cells(r, 2).Value2)) > 0 Then ValueCacheReady = True: Exit Function
    Next r
End Function

' True when a usable cache exists for one named source.
Public Function SourceCacheReady(ByVal key As String) As Boolean
    Dim ws As Worksheet, r As Long
    Set ws = CacheSheet()
    If ws Is Nothing Then Exit Function
    If SafeText(ws.Range("H1").Value2) <> CACHE_LAYOUT Then Exit Function
    For r = 3 To 8
        If SafeUpperText(ws.Cells(r, 1).Value2) = SafeUpperText(key) Then
            SourceCacheReady = (Len(SafeText(ws.Cells(r, 2).Value2)) > 0)
            Exit Function
        End If
    Next r
End Function

Public Function CacheSummary() As String
    Dim ws As Worksheet, r As Long, s As String, n As Long
    Set ws = CacheSheet()
    If ws Is Nothing Then CacheSummary = "no cached values - upload an extract": Exit Function
    If Not ValueCacheReady() Then CacheSummary = "no cached values - upload an extract": Exit Function
    For r = 3 To 8
        If Len(SafeText(ws.Cells(r, 2).Value2)) > 0 Then
            n = n + 1
            If Len(s) > 0 Then s = s & "; "
            s = s & SafeText(ws.Cells(r, 1).Value2) & " " & format$(val(SafeText(ws.Cells(r, 5).Value2)), "#,##0") & " rows"
        End If
    Next r
    CacheSummary = n & " source(s): " & s
End Function

' Reads the cache back into the shapes the builder needs. sourceKey "" means every source.
Private Function CachedGroup(ByVal kind As String, Optional ByVal sourceKey As String = "") As Object
    Dim ws As Worksheet, a As Variant, r As Long, d As Object, key As String, c As Collection
    Dim want As String, seen As Object, v As String
    want = SafeUpperText(sourceKey)
    Set d = NewMap(): Set seen = NewMap()
    Set ws = CacheSheet()
    If ws Is Nothing Then Set CachedGroup = d: Exit Function
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < CACHE_FIRST_ROW Then Set CachedGroup = d: Exit Function
    a = ws.Range(ws.Cells(CACHE_FIRST_ROW, 1), ws.Cells(lastRow, 4)).Value2
    For r = 1 To UBound(a, 1)
        If SafeUpperText(a(r, 1)) = SafeUpperText(kind) Then
            If Len(want) = 0 Or SafeUpperText(a(r, 4)) = want Then
                key = SafeText(a(r, 2))
                If Len(key) > 0 Then
                    If Not d.Exists(key) Then
                        Set c = New Collection
                        d.Add key, c
                    End If
                    ' Each value once per name, although several inputs cache it.
                    v = SafeText(a(r, 3))
                    If Len(v) > 0 Then
                        If Not seen.Exists(key & ChrW$(1) & v) Then
                            seen.Add key & ChrW$(1) & v, True
                            d(key).Add v
                        End If
                    End If
                End If
            End If
        End If
    Next r
    Set CachedGroup = d
End Function

' ===================== filter editing: write-back and testing =================
'
' The studio used to end at "copy this to the clipboard and paste it into Excel yourself".
' ALM Rule Studio v9 solves the same problem properly: the popup writes back through a
' bridge and can test a rule against real data before you commit it. These two entry points
' bring that to the JKB tool - the popup drops its result in a handoff file, and Excel
' applies it, tests it and reports the match count.

' Applies an edited expression to the cell the studio was opened on, then immediately
' reports how many ECL rows it matches, so a filter is never saved blind.
Public Sub ApplyPreShockEdit(ByVal target As String, ByVal expression As String)
    Dim msg As String
    msg = ApplyPreShockEditQuiet(target, expression)
    If Left$(msg, 4) = "ERR:" Then
        UiProblem "Apply filter", "The filter was not applied.", Mid$(msg, 5)
    Else
        UiNotice "Apply filter", msg
    End If
End Sub

' The same work, with no dialog, for the builder window to call.
'
' A MsgBox here would be a deadlock rather than a message: the window is waiting
' on the answer, Excel is waiting on the dialog, and the dialog is behind the
' window. Anything the bridge can invoke has to return its news instead of
' showing it.
Public Function ApplyPreShockEditQuiet(ByVal target As String, ByVal expression As String) As String
    Dim ws As Worksheet, parts As Variant, r As Long, colLetter As String, matched As Double, er As String
    On Error GoTo Failed
    EnsurePreShockSheet
    parts = Split(target, "!")
    If UBound(parts) < 1 Then Err.Raise vbObjectError + 730, , "The studio returned an unusable target: " & target
    ' The builder names the sheet it was opened on, so the write lands wherever it
    ' came from - a filter on the test cases sheet, a formula on the metrics sheet.
    Set ws = GetWorksheetSafe(ThisWorkbook, CStr(parts(0)))
    If ws Is Nothing Then Set ws = PsCasesSheet()
    colLetter = Left$(CStr(parts(1)), 1)
    r = val(Mid$(CStr(parts(1)), 2))
    If r <= 0 Then Err.Raise vbObjectError + 731, , "The studio returned an unusable target row: " & target
    ws.Range(colLetter & r).Value2 = expression
    If ws.name = PS_CASES_SHEET Then EffectiveFilterAt ws, r
    RefreshPreShockAvailability
    StylePreShockWorkspace
    ThisWorkbook.Activate
    ws.Activate
    Application.Goto ws.Range(colLetter & r), False
    If ws.name = PS_CASES_SHEET Then
        matched = TestFilterMatchCount(EffectiveFilterAt(ws, r), er)
        If Len(er) > 0 Then
            ApplyPreShockEditQuiet = "Filter saved, but it cannot be run yet: " & er
        Else
            ApplyPreShockEditQuiet = "Filter saved. It matches " & format$(matched, "#,##0") & " row(s)."
        End If
    Else
        ApplyPreShockEditQuiet = "Saved."
    End If
    Exit Function
Failed:
    ApplyPreShockEditQuiet = "ERR:The edit could not be applied: " & Err.description
End Function

' Entry point the filter builder calls over the HTA bridge. Returns the match count, or a
' negative number the builder turns into a message. It never shows a dialog: the builder
' owns the conversation while it is open.
Public Function TestFilterForBuilder(ByVal expr As String) As Double
    Dim reason As String, n As Double
    On Error GoTo Failed
    n = TestFilterMatchCount(expr, reason)
    If Len(reason) > 0 Then TestFilterForBuilder = -1 Else TestFilterForBuilder = n
    Exit Function
Failed:
    TestFilterForBuilder = -1
End Function

' Counts the ECL rows a filter expression selects. Returns 0 and a reason when it cannot run.
Public Function TestFilterMatchCount(ByVal expr As String, ByRef reason As String) As Double
    Dim agg As Object, cache As Object
    reason = ""
    On Error GoTo Failed
    EnsureEclLoaded
    Set cache = NewMap()
    Set agg = GetEclAggregate(cache, CurrentEclAod(), expr, "", LoadEclFilterMap())
    TestFilterMatchCount = AggVal(agg, "ROWS#1") + AggVal(agg, "ROWS#2") + AggVal(agg, "ROWS#3")
    Exit Function
Failed:
    reason = Err.description
End Function

' The as-of date the loaded extract actually carries, so a test does not silently compare
' against a date the file does not hold.
Private Function CurrentEclAod() As String
    Dim r As Long, col As Long
    If mEclHeaders Is Nothing Then Exit Function
    If Not mEclHeaders.Exists("AS_OF_DATE") Then Exit Function
    col = mEclHeaders("AS_OF_DATE")
    For r = 2 To UBound(mEclData, 1)
        On Error Resume Next
        CurrentEclAod = DateKey(mEclData(r, col), mEclDate1904)
        On Error GoTo 0
        If Len(CurrentEclAod) > 0 Then Exit Function
    Next r
End Function

' Tests the filter on the row the cursor is on, without saving anything.
Public Sub TestSelectedFilter()
    Dim ws As Worksheet, r As Long, expr As String, reason As String, matched As Double
    On Error GoTo Failed
    EnsurePreShockSheet
    Set ws = PsCasesSheet()
    If ActiveSheet.name <> PS_CASES_SHEET Then ws.Activate
    r = ActiveCell.row
    If r < PS_CASE_FIRST_ROW Or r > PS_CASE_LAST_ROW Or Len(SafeText(ws.Cells(r, 2).Value2)) = 0 Then
        UiNotice "Test filter", "Select a test case row on " & PS_CASES_SHEET & " first.", , "Then click Test filter."
        Exit Sub
    End If
    expr = EffectiveFilterAt(ws, r)
    If Len(expr) = 0 Then UiNotice "Test filter", "That test case has no filter condition.", "Pre-shock cannot be derived for it. Base is still tested.": Exit Sub
    matched = TestFilterMatchCount(expr, reason)
    If Len(reason) > 0 Then
        UiProblem "Test filter", "That filter cannot be run.", reason
    Else
        UiNotice "Test filter", ws.Cells(r, 2).Value2 & " / " & ws.Cells(r, 3).Value2 & " matches " & format$(matched, "#,##0") & _
                 " of " & format$(UBound(mEclData, 1) - 1, "#,##0") & " ECL rows.", expr, , (matched = 0)
    End If
    Exit Sub
Failed:
    UiProblem "Test filter", "The filter could not be tested.", Err.description
End Sub

Public Sub GoPreShock()
    ' The overview is gone; base and pre-shock are set up on Value sources.
    GoValueSources
End Sub
Public Sub GoPreShockSources()
    GoPsSheet PS_SOURCES_SHEET
End Sub
Public Sub GoPreShockFormulas()
    GoPsSheet PS_METRICS_SHEET
End Sub
Public Sub GoPreShockDimensions()
    GoPsSheet PS_FIELDS_SHEET
End Sub
Public Sub GoPreShockCases()
    GoPsSheet PS_CASES_SHEET
End Sub
Public Sub GoPreShockResults()
    ' One results table.
    GoDerivedValues
End Sub
Public Sub GoDerivedValues()
    Dim ws As Worksheet
    On Error GoTo Failed
    Set ws = DerivedSheet()
    If ws Is Nothing Then
        UiNotice "Derived values", "No derived values have been written yet.", _
                 "Base and pre-shock are stored on the " & DERIVED_SHEET & " sheet as soon as they are rebuilt.", _
                 "Load the input files, then click Run on the home screen."
        Exit Sub
    End If
    ThisWorkbook.Activate
    ws.Visible = xlSheetVisible
    ws.Activate
    DressWorkSheet ws
    Application.Goto ws.Range("A1"), True
    Exit Sub
Failed:
    UiProblem "Derived values", DERIVED_SHEET & " could not be opened.", Err.description
End Sub

' Button entry point: point the configuration's pre-shock rows at the derived values.
Public Sub UseDerivedValuesInConfig()
    If JKB_Busy Then Exit Sub
    MigratePreShockConfigTokens True
End Sub
' Navigation is now to a SHEET rather than to a row several hundred lines down.
Private Sub GoPreShockRow(ByVal r As Long)
    GoPsSheet PRE_SHOCK_SHEET
End Sub

Public Sub GoPsSheet(ByVal nm As String)
    Dim ws As Worksheet
    On Error GoTo Failed
    EnsurePreShockSheet
    Set ws = GetWorksheetSafe(ThisWorkbook, nm)
    If ws Is Nothing Then Exit Sub
    ThisWorkbook.Activate
    ws.Visible = xlSheetVisible
    ws.Activate
    DressWorkSheet ws
    Application.Goto ws.Cells(1, 1), True
    If ws.Cells(ws.rows.count, 1).End(xlUp).row >= PS_FIRST_ROW Then ws.Cells(PS_FIRST_ROW, 1).Select
    Exit Sub
Failed:
    UiProblem "Open sheet", nm & " could not be opened.", Err.description
End Sub

' The ECL pre-shock sheet is one long document made of five stacked blocks. Styling keeps
' the blocks visually distinct so the eye can find them without a separate sheet each.
' ============================ appearance ====================================
'
' Every pre-shock sheet is styled from one place, in one pass, so they cannot
' drift apart. The old styler was 135 lines of per-block special cases on a
' single crowded sheet; this is the same look applied six times.

Public Sub StylePreShockSheet(ByVal ws As Worksheet)
    StylePreShockWorkspace
End Sub

Public Sub StylePreShockWorkspace()
    On Error Resume Next
    StyleOne PreShockSheet(), 0, 12
    StyleOne PsSourcesSheet(), 8, 8
    StyleOne PsMetricsSheet(), PS_METRIC_COLS, PS_METRIC_COLS
    StyleOne PsFieldsSheet(), PS_MAP_COLS, PS_MAP_COLS
    StyleOne PsCasesSheet(), 12, 12
    StyleOne PsResultsSheet(), 12, 12
    StyleOne BreakdownSheet(), 17, 17
    StylePsWidths
    ' Built, filled, then dressed - in that order, always. StyleOne gives every
    ' sheet the same bones; the design pass gives each one its own proportions.
    Dim proNm As Variant, proWs As Worksheet
    For Each proNm In PsSheetNames()
        Set proWs = GetWorksheetSafe(ThisWorkbook, CStr(proNm))
        If Not proWs Is Nothing Then Sheet1.ProStyleSheet proWs
    Next proNm
    DressWorkSheets
    Err.Clear
End Sub

' cols = how many columns the table has; 0 means the sheet is a panel, not a table.
Private Sub StyleOne(ByVal ws As Worksheet, ByVal cols As Long, ByVal span As Long)
    Dim lastRow As Long, rng As Range
    If ws Is Nothing Then Exit Sub
    On Error Resume Next

    ws.Cells.Font.name = "Aptos"
    ws.Cells.Font.Size = 9.5
    ws.Cells.Font.Color = RGB(16, 24, 40)
    ws.Cells.Interior.Color = RGB(249, 250, 251)

    ws.Range("A" & PS_TITLE_ROW & ":" & colLetter(span) & PS_TITLE_ROW).UnMerge
    ws.Range("A" & PS_ABOUT_ROW & ":" & colLetter(span) & PS_ABOUT_ROW).UnMerge
    ws.Range("A" & PS_TITLE_ROW & ":" & colLetter(span) & PS_TITLE_ROW).Merge
    ws.Range("A" & PS_ABOUT_ROW & ":" & colLetter(span) & PS_ABOUT_ROW).Merge

    With ws.Range("A" & PS_TITLE_ROW & ":" & colLetter(span) & PS_TITLE_ROW)
        .Interior.Color = RGB(238, 243, 251)
        .Font.Color = RGB(18, 48, 107)
        .Font.Bold = True
        .Font.Size = 15
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(176, 138, 46)
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Range("A" & PS_ABOUT_ROW & ":" & colLetter(span) & PS_ABOUT_ROW)
        .Interior.Color = RGB(249, 250, 251)
        .Font.Color = RGB(102, 112, 133)
        .Font.Size = 9
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.rows(PS_TITLE_ROW).RowHeight = 29
    ws.rows(PS_ABOUT_ROW).RowHeight = 34
    ' Row 1 is the toolbar band and carries TWO strips of controls: the sheet's
    ' actions and the view controls under them. Setting it to 6 here - which this
    ' did - left the second strip drawn over the title until the sheet was next
    ' activated and the ribbon put it back.
    ws.rows(1).RowHeight = 42
    ws.rows(1).Font.Color = RGB(249, 250, 251)
    ws.rows(5).RowHeight = 7

    With ws.Cells(PS_STATUS_ROW, 1)
        .Font.Bold = True
        .Font.Size = 9.5
        .Font.Color = RGB(18, 48, 107)
        .IndentLevel = 1
    End With

    If cols = 0 Then
        StyleHomePanel ws
        Exit Sub
    End If

    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastRow < PS_FIRST_ROW Then lastRow = PS_FIRST_ROW

    With ws.Range(ws.Cells(PS_HDR_ROW, 1), ws.Cells(PS_HDR_ROW, cols))
        .Interior.Color = RGB(52, 64, 84)
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 9
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.rows(PS_HDR_ROW).RowHeight = 30

    Set rng = ws.Range(ws.Cells(PS_FIRST_ROW, 1), ws.Cells(lastRow, cols))
    rng.Interior.Color = vbWhite
    rng.VerticalAlignment = xlTop
    rng.Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
    rng.Borders(xlInsideVertical).Color = RGB(234, 236, 240)
    rng.Borders(xlEdgeBottom).Color = RGB(234, 236, 240)

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(PS_HDR_ROW, 1), ws.Cells(lastRow, cols)).AutoFilter
    PsDataBars ws, cols, lastRow
    PsStatusColours ws, cols, lastRow
    ' The header stays put while a long table scrolls. Freezing needs a visible
    ' window, so it is advisory and never allowed to cost the rest of the styling.
    On Error Resume Next
    If Application.Visible Then
        If ws.Parent.Windows.count > 0 Then
            ws.Activate
            ws.Range("A1").Select
            ActiveWindow.FreezePanes = False
            ws.Cells(PS_FIRST_ROW, 1).Select
            ActiveWindow.FreezePanes = True
            ws.Range("A1").Select
        End If
    End If
    Err.Clear
End Sub

' Data bars on the amount columns.
'
' Every one of these tables is a list of amounts, and the question asked of it is
' always "how big is this one compared with the rest". A column of right-aligned
' digits answers that only after the reader has counted commas in two of them.
Private Sub PsDataBars(ByVal ws As Worksheet, ByVal cols As Long, ByVal lastRow As Long)
    Dim c As Long, h As String, rng As Range, n As Long
    On Error Resume Next
    If lastRow <= PS_FIRST_ROW Then Exit Sub
    ws.Range(ws.Cells(PS_FIRST_ROW, 1), ws.Cells(lastRow, cols)).FormatConditions.Delete
    For c = 1 To cols
        h = SafeUpperText(ws.Cells(PS_HDR_ROW, c).Value2)
        If PsIsAmountHeading(h) Then
            n = n + 1
            ' Two is the limit. A table with a bar in every numeric column is a
            ' bar chart with some numbers in it, and reads as neither.
            If n > 2 Then Exit For
            Set rng = ws.Range(ws.Cells(PS_FIRST_ROW, c), ws.Cells(lastRow, c))
            With rng.FormatConditions.AddDatabar
                .BarColor.Color = RGB(183, 196, 216)
                .BarFillType = xlDataBarFillGradient
                .Direction = xlContext
                .NegativeBarFormat.ColorType = xlDataBarColor
                .NegativeBarFormat.Color.Color = RGB(254, 205, 202)
                .ShowValue = True
                .BarBorder.Type = xlDataBarBorderNone
            End With
        End If
    Next c
    Err.Clear
End Sub

Private Function PsIsAmountHeading(ByVal h As String) As Boolean
    If Len(h) = 0 Then Exit Function
    If InStr(h, "DERIVED") > 0 Then PsIsAmountHeading = True: Exit Function
    If InStr(h, "DIFF") > 0 Then PsIsAmountHeading = True: Exit Function
    If InStr(h, "ROWS MATCHED") > 0 Then PsIsAmountHeading = True
End Function

' A status column is the one thing on the sheet a reader scans rather than reads,
' so it is coloured by what it says instead of being left as grey text among grey
' text. Rules rather than painted cells: they survive sorting and re-running.
Private Sub PsStatusColours(ByVal ws As Worksheet, ByVal cols As Long, ByVal lastRow As Long)
    Dim c As Long, h As String, rng As Range
    On Error Resume Next
    If lastRow <= PS_FIRST_ROW Then Exit Sub
    For c = 1 To cols
        h = SafeUpperText(ws.Cells(PS_HDR_ROW, c).Value2)
        If h = "STATUS" Or h = "AVAILABILITY" Or h = "MAPPING STATUS" Or h = "LAST RESULT" Then
            Set rng = ws.Range(ws.Cells(PS_FIRST_ROW, c), ws.Cells(lastRow, c))
            ' Added good first, then warning, then bad, each promoted to the front
            ' as it goes. "not available" contains "available"; without an explicit
            ' order the greener of the two rules wins and a missing column reads as
            ' a healthy one - which is the worst way for this to be wrong.
            PsRule rng, "Matched", RGB(6, 118, 71), RGB(236, 253, 243)
            PsRule rng, "Ready", RGB(6, 118, 71), RGB(236, 253, 243)
            PsRule rng, "Available", RGB(6, 118, 71), RGB(236, 253, 243)
            PsRule rng, "not", RGB(181, 71, 8), RGB(253, 246, 227)
            PsRule rng, "cannot", RGB(181, 71, 8), RGB(253, 246, 227)
            PsRule rng, "no ", RGB(181, 71, 8), RGB(253, 246, 227)
            PsRule rng, "Break", RGB(180, 35, 24), RGB(254, 243, 242)
        End If
    Next c
    Err.Clear
End Sub

Private Sub PsRule(ByVal rng As Range, ByVal contains As String, ByVal fg As Long, ByVal bg As Long)
    Dim fc As Object
    On Error Resume Next
    Set fc = rng.FormatConditions.Add(Type:=xlTextString, String:=contains, TextOperator:=xlContains)
    If fc Is Nothing Then Exit Sub
    fc.Font.Color = fg
    fc.Interior.Color = bg
    fc.SetFirstPriority
    fc.StopIfTrue = True
    Err.Clear
End Sub

Private Sub StyleHomePanel(ByVal ws As Worksheet)
    Dim lastRow As Long
    On Error Resume Next
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    ws.Cells(PS_HDR_ROW, 1).Font.Bold = True
    ws.Cells(PS_HDR_ROW, 1).Font.Size = 9
    ws.Cells(PS_HDR_ROW, 1).Font.Color = RGB(102, 112, 133)
    With ws.Range("A" & PS_SRC_FIRST_ROW & ":L" & PS_SRC_LAST_ROW)
        .Interior.Color = vbWhite
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(234, 236, 240)
        .Borders.Weight = xlThin
    End With
    With ws.Range("A" & PS_SRC_FIRST_ROW & ":A" & (PS_SRC_FIRST_ROW + 2))
        .Font.Bold = True
        .Font.Color = RGB(18, 48, 107)
    End With
    ws.Range("A" & (PS_SRC_LAST_ROW + 2)).Font.Bold = True
    ws.Range("A" & (PS_SRC_LAST_ROW + 2)).Font.Size = 9
    ws.Range("A" & (PS_SRC_LAST_ROW + 2)).Font.Color = RGB(102, 112, 133)
    If lastRow > PS_SRC_LAST_ROW + 2 Then
        With ws.Range("A" & (PS_SRC_LAST_ROW + 3) & ":L" & lastRow)
            .Interior.Color = vbWhite
            .Borders(xlInsideHorizontal).Color = RGB(234, 236, 240)
        End With
        ws.Range("A" & (PS_SRC_LAST_ROW + 3) & ":A" & lastRow).Font.Bold = True
        ws.Range("B" & (PS_SRC_LAST_ROW + 3) & ":B" & lastRow).Font.Color = RGB(102, 112, 133)
    End If
    Err.Clear
End Sub

Private Sub StylePsWidths()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = PreShockSheet()
    If Not ws Is Nothing Then
        ws.columns("A").ColumnWidth = 26
        ws.columns("B").ColumnWidth = 62
        ws.columns("C:L").ColumnWidth = 14
    End If
    Set ws = PsSourcesSheet()
    If Not ws Is Nothing Then
        ws.columns("A").ColumnWidth = 12: ws.columns("B").ColumnWidth = 44
        ws.columns("C").ColumnWidth = 46: ws.columns("D").ColumnWidth = 18
        ws.columns("E").ColumnWidth = 11: ws.columns("F").ColumnWidth = 30
        ws.columns("G").ColumnWidth = 22: ws.columns("H").ColumnWidth = 20
    End If
    Set ws = PsMetricsSheet()
    If Not ws Is Nothing Then
        ws.columns("A").ColumnWidth = 34: ws.columns("B").ColumnWidth = 11
        ws.columns("C").ColumnWidth = 9: ws.columns("D").ColumnWidth = 52
        ws.columns("E").ColumnWidth = 8: ws.columns("F").ColumnWidth = 30
        ws.columns("G").ColumnWidth = 22: ws.columns("H").ColumnWidth = 40
        ws.columns("I").ColumnWidth = 30
    End If
    Set ws = PsFieldsSheet()
    If Not ws Is Nothing Then
        ws.columns("A").ColumnWidth = 34: ws.columns("B").ColumnWidth = 52
        ws.columns("C").ColumnWidth = 18: ws.columns("D").ColumnWidth = 24
        ws.columns("E").ColumnWidth = 14: ws.columns("F").ColumnWidth = 40
    End If
    Set ws = PsCasesSheet()
    If Not ws Is Nothing Then
        ws.columns("A").ColumnWidth = 8: ws.columns("B").ColumnWidth = 14
        ws.columns("C").ColumnWidth = 16: ws.columns("D").ColumnWidth = 34
        ws.columns("E").ColumnWidth = 18: ws.columns("F").ColumnWidth = 46
        ws.columns("G").ColumnWidth = 46: ws.columns("H").ColumnWidth = 46
        ws.columns("I").ColumnWidth = 28: ws.columns("J").ColumnWidth = 13
        ws.columns("K").ColumnWidth = 34: ws.columns("L").ColumnWidth = 26
    End If
    Set ws = PsResultsSheet()
    If Not ws Is Nothing Then
        ws.columns("A").ColumnWidth = 12: ws.columns("B").ColumnWidth = 20
        ws.columns("C").ColumnWidth = 14: ws.columns("D").ColumnWidth = 16
        ws.columns("E").ColumnWidth = 34
        ws.columns("F:K").ColumnWidth = 17
        ws.columns("L").ColumnWidth = 40
        ws.Range(ws.Cells(PS_FIRST_ROW, 6), ws.Cells(ws.rows.count, 11)).NumberFormat = "#,##0.00"
    End If
    Err.Clear
End Sub

Private Sub StyleBlockHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal cols As Long, ByVal caption As String)
    On Error Resume Next
    With ws.Cells(r - 1, 1)
        .Value2 = caption
        .Font.Bold = True: .Font.Size = 11: .Font.Color = RGB(18, 48, 107)
        .IndentLevel = 0
    End With
    ws.rows(r - 1).RowHeight = 24
    With ws.Cells(r, 1).Resize(1, cols)
        .Interior.Color = RGB(59, 99, 176): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 8.5
        .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.Color = RGB(52, 64, 84)
    End With
    ws.rows(r).RowHeight = 28
    On Error GoTo 0
End Sub


' Bounded regression checks against the production filter and aggregation helpers.
' Restores the active table and never writes to the user's source files.
Public Function ReconEngineRegressionTests() As String
    Dim oldData As Variant, oldHeaders As Object, oldSource As String
    Dim fm As Object, rule As Object, node As Object, agg As Object
    Dim needs As String, found As Boolean, v As Double, n As Long, errorText As String
    Dim a(1 To 4, 1 To 3) As Variant, emptyValue As Variant, dateFailed As Boolean
    oldData = mEclData: Set oldHeaders = mEclHeaders: oldSource = mActiveSource
    On Error GoTo Failed
    a(1, 1) = "CATEGORY": a(1, 2) = "FLAG": a(1, 3) = "AMOUNT"
    a(2, 1) = "A, B": a(2, 2) = "YES": a(2, 3) = 10
    a(3, 1) = "Letters of Credit (LC's)": a(3, 2) = "NO": a(3, 3) = 0
    a(4, 1) = Empty: a(4, 2) = Empty: a(4, 3) = Empty
    ReturnActiveTable
    mEclData = a: Set mEclHeaders = NewMap(): mActiveSource = "REGRESSION"
    ForgetActiveTable
    mEclHeaders("CATEGORY") = 1: mEclHeaders("FLAG") = 2: mEclHeaders("AMOUNT") = 3
    Set fm = NewMap(): Set rule = NewMap(): rule("Candidates") = "CATEGORY": rule("Mode") = "TEXT": Set fm("CATEGORY") = rule
    Set rule = NewMap(): rule("Candidates") = "FLAG": rule("Mode") = "YESNO": Set fm("FLAG") = rule
    Set node = CompileFilter("CATEGORY IN ('A, B')", fm, needs)
    ReconAssert EvalFilter(node, 2), "Quoted comma", n
    Set node = CompileFilter("CATEGORY IN ('Letters of Credit (LC's)')", fm, needs)
    ReconAssert EvalFilter(node, 3), "Supplied apostrophe label", n
    Set node = CompileFilter("CATEGORY NOT IN ('A, B')", fm, needs)
    ReconAssert Not EvalFilter(node, 4), "NULL must not satisfy NOT IN", n
    Set node = CompileFilter("FLAG IN ('NO')", fm, needs)
    ReconAssert Not EvalFilter(node, 4), "Blank flag is not NO", n
    Set node = CompileFilter("CATEGORY IN ('A, B') OR FLAG IN ('NO') AND CATEGORY NOT IN ('A, B')", fm, needs)
    ReconAssert EvalFilter(node, 2) And EvalFilter(node, 3) And Not EvalFilter(node, 4), "AND/OR precedence", n
    Set node = CompileFilter("ALL", fm, needs)
    ReconAssert EvalFilter(node, 4), "Explicit full population", n
    ReconAssert NormalizeStage("Stage10") = "" And NormalizeStage("Stage2 - SICR") = "2", "Stage boundaries", n
    Set agg = NewMap(): AddGenericAgg agg, "AMOUNT", "1", 0
    v = AggregateFunctionValue(agg, "SUM", "AMOUNT", "1", found)
    ReconAssert found And v = 0, "Preserve genuine zero", n
    AddGenericAgg agg, "AMOUNT", "1", emptyValue
    v = AggregateFunctionValue(agg, "SUM", "AMOUNT", "1", found)
    ReconAssert Not found, "Incomplete measure must not produce valid sum", n
    v = EvalPreShockExpression(agg, "10/0", "ALL", found)
    ReconAssert Not found, "Zero denominator unavailable", n
    On Error Resume Next
    errorText = CheckedSourceDate("not-a-date", 2)
    dateFailed = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    ReconAssert dateFailed, "Invalid reporting date refused", n
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: mActiveSource = oldSource
    ForgetActiveTable
    ReconEngineRegressionTests = "PASS: " & n & " production-engine regression checks"
    Exit Function
Failed:
    errorText = Err.description
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: mActiveSource = oldSource
    ForgetActiveTable
    Err.Raise vbObjectError + 748, "ReconEngineRegressionTests", errorText
End Function

Private Sub ReconAssert(ByVal condition As Boolean, ByVal label As String, ByRef count As Long)
    If Not condition Then Err.Raise vbObjectError + 749, , "Regression failed: " & label
    count = count + 1
End Sub

' Exercises the actual aggregate path, including its narrow undated CAP rule.
Public Function ReconCapitalDateRegressionTests() As String
    Dim oldData As Variant, oldHeaders As Object, oldMeasures As Object, oldSource As String, oldMode As Variant, oldDates As Object
    Dim a(1 To 3, 1 To 4) As Variant, fm As Object, cache As Object, agg As Object
    Dim found As Boolean, rejected As Boolean, v As Double, n As Long, errorText As String, ws As Worksheet
    oldData = mEclData: Set oldHeaders = mEclHeaders: Set oldMeasures = mPreShockMeasureFields: oldSource = mActiveSource
    Set ws = PreShockSheet(): oldMode = StrConv(modBasePreShock.AsOfMatchingSetting(), vbProperCase)
    Set oldDates = mSourceDateCache: ClearSourceDateCache
    On Error GoTo Failed
    SetAsOfMatching "Match"
    a(1, 1) = "BANK_ID": a(1, 2) = "REPORT_BALANCE": a(1, 3) = "CAPITAL_ELEMENT": a(1, 4) = "AS_OF_DATE"
    a(2, 1) = 101: a(2, 2) = 10: a(2, 3) = "CET1": a(2, 4) = CDbl(DateSerial(2025, 12, 31))
    a(3, 1) = 101: a(3, 2) = 15: a(3, 3) = "CET1": a(3, 4) = a(2, 4)
    ReturnActiveTable
    mEclData = a: Set mEclHeaders = NewMap(): mActiveSource = "CAP"
    ForgetActiveTable
    mEclHeaders("BANK_ID") = 1: mEclHeaders("REPORT_BALANCE") = 2: mEclHeaders("CAPITAL_ELEMENT") = 3
    Set mPreShockMeasureFields = NewMap(): mPreShockMeasureFields("REPORT_BALANCE") = 2
    Set fm = LoadEclFilterMap(): Set cache = NewMap()
    Set agg = GetEclAggregate(cache, "2025-12-31", "", "", fm)
    v = AggregateFunctionValue(agg, "SUM", "REPORT_BALANCE", "ALL", found)
    ReconAssert found And v = 25, "Undated CAP base is calculable", n
    On Error Resume Next
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2025-12-31", "ALL", "", fm)
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    ReconAssert rejected, "Undated CAP filtered pre-shock refused", n
    mActiveSource = "LCR"
    On Error Resume Next
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2025-12-31", "", "", fm)
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    ReconAssert rejected, "Undated non-CAP source refused", n
    mActiveSource = "CAP": mEclHeaders("AS_OF_DATE") = 4
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2025-12-31", "", "", fm)
    v = AggregateFunctionValue(agg, "SUM", "REPORT_BALANCE", "ALL", found)
    ReconAssert found And v = 25, "Dated CAP retains ordinary date matching", n
    mEclData(2, 4) = "not-a-date"
    ClearSourceDateCache
    On Error Resume Next
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2025-12-31", "", "", fm)
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    ReconAssert rejected, "Malformed dated CAP refused", n
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: Set mPreShockMeasureFields = oldMeasures: mActiveSource = oldSource
    ForgetActiveTable
    SetAsOfMatching SafeText(oldMode)
    Set mSourceDateCache = oldDates
    ReconCapitalDateRegressionTests = "PASS: " & n & " capital date-scope regression checks"
    Exit Function
Failed:
    errorText = Err.description
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: Set mPreShockMeasureFields = oldMeasures: mActiveSource = oldSource
    ForgetActiveTable
    SetAsOfMatching SafeText(oldMode)
    Set mSourceDateCache = oldDates
    Err.Raise vbObjectError + 751, "ReconCapitalDateRegressionTests", errorText
End Function

' Production-path checks for reusable dates and source-specific breakdown plans.
Public Function ReconDateCacheRegressionTests() As String
    Dim oldData As Variant, oldHeaders As Object, oldMeasures As Object, oldSources As Object, oldDates As Object, oldJoin As Object
    Dim oldSource As String, oldPath As String, oldSheet As String, oldHeaderRow As Long, old1904 As Boolean, oldMode As Variant, oldPlan As Object
    Dim a(1 To 3, 1 To 6) As Variant, h As Object, info As Object, again As Object, cache As Object, agg As Object, fm As Object, breakdown As Object
    Dim rowDates As Variant, found As Boolean, amount As Double, n As Long, errorText As String, failureNumber As Long, ws As Worksheet, j As Long
    Dim firstPlan As Object, otherPlan As Object
    oldData = mEclData: Set oldHeaders = mEclHeaders: Set oldMeasures = mPreShockMeasureFields
    Set oldSources = mSources: Set oldDates = mSourceDateCache: Set oldJoin = mJoinCache: Set oldPlan = SnapshotBreakdownPlan()
    oldSource = mActiveSource: oldPath = mEclPath: oldSheet = mEclSheet: oldHeaderRow = mEclHeaderRow: old1904 = mEclDate1904
    Set ws = PreShockSheet(): oldMode = StrConv(modBasePreShock.AsOfMatchingSetting(), vbProperCase)
    On Error GoTo Failed
    Set mSources = NewMap(): ClearSourceDateCache: Set mJoinCache = Nothing
    mBreakdownOn = False
    SetAsOfMatching "Match"
    a(1, 1) = "BANK_ID": a(1, 2) = "REPORT_BALANCE": a(1, 3) = "CAPITAL_ELEMENT"
    a(1, 4) = "AS_OF_DATE": a(1, 5) = "ACCOUNT_NUMBER": a(1, 6) = "BRANCH_CODE"
    a(2, 1) = 101: a(2, 2) = 10: a(2, 3) = "CET1": a(2, 4) = CDbl(DateSerial(2025, 12, 31)): a(2, 5) = "A": a(2, 6) = 1
    a(3, 1) = 101: a(3, 2) = 20: a(3, 3) = "T2": a(3, 4) = CDbl(DateSerial(2024, 12, 31)): a(3, 5) = "B": a(3, 6) = 1
    Set h = NewMap()
    For j = 1 To 6: h(CStr(a(1, j))) = j: Next j
    RegisterTable "DATE_TEST", a, h, "", "Date test": ActivateSource "DATE_TEST"
    Set mPreShockMeasureFields = NewMap(): mPreShockMeasureFields("REPORT_BALANCE") = 2
    Set fm = LoadEclFilterMap()
    Set info = SourceDateInventory(): rowDates = info("Rows")
    ReconAssert info("Dates").count = 2 And rowDates(2) = "2025-12-31" And rowDates(3) = "2024-12-31", "Date inventory and row keys", n
    Set again = SourceDateInventory()
    ReconAssert (info Is again), "Unchanged source reuses validated date inventory", n
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2025-12-31", "", "", fm)
    amount = AggregateFunctionValue(agg, "SUM", "REPORT_BALANCE", "ALL", found)
    ReconAssert found And amount = 10, "Cached date keys retain mixed-date population", n
    On Error Resume Next
    Set agg = GetEclAggregate(cache, "2023-12-31", "UNKNOWN_FIELD IN ('X')", "", fm)
    failureNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    ReconAssert failureNumber = vbObjectError + 724, "Absent date refused before filter compilation", n

    mEclData(3, 4) = CDbl(DateSerial(2026, 12, 31))
    StoreActiveSource "DATE_TEST"
    Set again = SourceDateInventory()
    ReconAssert Not (info Is again), "StoreActiveSource invalidates prior inventory", n
    ReconAssert again("Dates").Exists("2026-12-31") And Not again("Dates").Exists("2024-12-31"), "Stored source replacement supplies current dates", n
    RegisterTable "DATE_TEST", a, h, "", "Replacement": ActivateSource "DATE_TEST"
    Set info = SourceDateInventory()
    ReconAssert info("Dates").Exists("2024-12-31") And Not info("Dates").Exists("2026-12-31"), "RegisterTable invalidates prior inventory", n
    ClearSourceDateCache
    Set again = SourceDateInventory()
    ReconAssert Not (info Is again), "New run invalidates all source inventories", n

    a(3, 4) = "malformed-date"
    RegisterTable "DATE_TEST", a, h, "", "Malformed": ActivateSource "DATE_TEST"
    On Error Resume Next
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2023-12-31", "UNKNOWN_FIELD IN ('X')", "", fm)
    failureNumber = Err.Number: errorText = Err.description: Err.Clear
    On Error GoTo Failed
    ReconAssert failureNumber = vbObjectError + 743 And InStr(errorText, "source row 3") > 0, "Malformed row is never hidden by absent requested date", n
    On Error Resume Next
    Set info = SourceDateInventory()
    failureNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    ReconAssert failureNumber = vbObjectError + 743, "Cached validation error remains explicit", n
    SetAsOfMatching "Ignore"
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2023-12-31", "", "", fm)
    amount = AggregateFunctionValue(agg, "SUM", "REPORT_BALANCE", "ALL", found)
    ReconAssert found And amount = 30, "Explicit Ignore retains all source rows", n
    SetAsOfMatching "Match"
    On Error Resume Next
    Set agg = GetEclAggregate(cache, "2023-12-31", "", "", fm)
    failureNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    ReconAssert failureNumber = vbObjectError + 743, "Changing date policy cannot reuse Ignore aggregate", n
    a(3, 4) = CDbl(DateSerial(2024, 12, 31))
    RegisterTable "DATE_TEST", a, h, "", "Corrected": ActivateSource "DATE_TEST"
    Set info = SourceDateInventory()
    ReconAssert info("Dates").count = 2, "Corrected source clears cached date error", n

    ' Source plans deliberately use different column positions, as CAP and ECL do.
    mBreakdownOn = True: mBrkDimCount = 1: mBrkMeasureCount = 1
    ReDim mBrkDimName(1 To 1): ReDim mBrkDimCol(1 To 1): ReDim mBrkMeasureName(1 To 1): ReDim mBrkMeasureCol(1 To 1)
    mBrkDimName(1) = "CAPITAL_ELEMENT": mBrkDimCol(1) = 3
    mBrkMeasureName(1) = "REPORT_BALANCE": mBrkMeasureCol(1) = 2
    Set firstPlan = SnapshotBreakdownPlan()
    mBrkDimName(1) = "BANK_ID": mBrkDimCol(1) = 1: mBrkMeasureCol(1) = 6
    Set otherPlan = SnapshotBreakdownPlan()
    RestoreBreakdownPlan otherPlan
    RestoreBreakdownPlan firstPlan
    Set cache = NewMap(): Set agg = GetEclAggregate(cache, "2025-12-31", "ALL", "", fm)
    Set breakdown = BreakdownOf(agg)
    ReconAssert breakdown.Exists("CAPITAL_ELEMENT") And Not breakdown.Exists("BANK_ID"), "Source switch restores correct breakdown dimension", n
    ReconAssert CDbl(breakdown("CAPITAL_ELEMENT")("CET1")("REPORT_BALANCE")) = 10, "Source switch restores correct breakdown amount column", n

    RegisterTable "ECL", a, h, "", "Join dates"
    Set agg = AccountSetFor("ALL", "2025-12-31")
    ReconAssert agg.count = 1 And agg.Exists("A"), "Account joins use cached reporting-date keys", n
    ReconDateCacheRegressionTests = "PASS: " & n & " date-cache and source-plan regression checks"
    GoTo Restore
Failed:
    errorText = Err.description
Restore:
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: Set mPreShockMeasureFields = oldMeasures
    ForgetActiveTable
    Set mSources = oldSources: Set mSourceDateCache = oldDates: Set mJoinCache = oldJoin
    mActiveSource = oldSource: mEclPath = oldPath: mEclSheet = oldSheet: mEclHeaderRow = oldHeaderRow: mEclDate1904 = old1904
    RestoreBreakdownPlan oldPlan
    SetAsOfMatching SafeText(oldMode)
    If Len(ReconDateCacheRegressionTests) = 0 Then Err.Raise vbObjectError + 752, "ReconDateCacheRegressionTests", errorText
End Function

Public Function ReconPerformanceRegressionTests() As String
    ReconPerformanceRegressionTests = ReconEngineRegressionTests() & vbLf & ReconCapitalDateRegressionTests() & vbLf & ReconDateCacheRegressionTests() & vbLf & ReconBreakdownCapacityRegressionTests() & vbLf & ReconSpeedEquivalenceTests() & vbLf & ReconTableHandoverTests()
    Debug.Print ReconPerformanceRegressionTests
End Function

Public Function ReconBreakdownCapacityRegressionTests() As String
    Dim oldData As Variant, oldPlan As Object, a(1 To 1301, 1 To 3) As Variant, agg As Object
    Dim r As Long, i As Long, n As Long, checkSlot As Variant, expectedRows As Long, stem As String
    Dim rowsTotal As Double, amountTotal As Double, secondTotal As Double, allPreserved As Boolean, errorText As String
    oldData = mEclData: Set oldPlan = SnapshotBreakdownPlan()
    On Error GoTo Failed
    a(1, 1) = "PRODUCT": a(1, 2) = "AMOUNT": a(1, 3) = "SECOND_AMOUNT"
    For r = 2 To 1301
        a(r, 1) = "P" & format$(r - 1, "0000"): a(r, 2) = CDbl(r - 1): a(r, 3) = -2# * (r - 1)
    Next r
    ReturnActiveTable
    mEclData = a: mBreakdownOn = True: mBrkDimCount = 1: mBrkMeasureCount = 2
    ForgetActiveTable
    ReDim mBrkDimName(1 To 1): ReDim mBrkDimCol(1 To 1)
    ReDim mBrkMeasureName(1 To 2): ReDim mBrkMeasureCol(1 To 2)
    mBrkDimName(1) = "PRODUCT": mBrkDimCol(1) = 1
    mBrkMeasureName(1) = "AMOUNT": mBrkMeasureCol(1) = 2
    mBrkMeasureName(2) = "SECOND_AMOUNT": mBrkMeasureCol(2) = 3
    BeginBreakdown
    For r = 2 To 1301: BreakdownRow r: Next r
    ReconAssert mBrkSlotCount = 1300, "All 1,300 distinct breakdown values retained", n
    ' Revisit earlier values after crossing both former growth boundaries.
    For r = 2 To 601: BreakdownRow r: Next r
    For Each checkSlot In Array(1, 512, 513, 1024, 1025, 1300)
        i = CLng(checkSlot): expectedRows = IIf(i <= 600, 2, 1)
        ReconAssert mBrkSlotVal(i) = "P" & format$(i, "0000") And mBrkSum(i, 0) = expectedRows And mBrkSum(i, 1) = CDbl(i) * expectedRows And mBrkSum(i, 2) = -2# * i * expectedRows, "Preserve breakdown boundary " & i, n
    Next checkSlot
    Set agg = NewMap(): FlushBreakdown agg: allPreserved = True
    For i = 1 To 1300
        stem = BRK_PREFIX & "PRODUCT|P" & format$(i, "0000") & "|"
        expectedRows = IIf(i <= 600, 2, 1)
        If CDbl(agg(stem & "#ROWS")) <> expectedRows Or CDbl(agg(stem & "AMOUNT")) <> CDbl(i) * expectedRows Or CDbl(agg(stem & "SECOND_AMOUNT")) <> -2# * i * expectedRows Then allPreserved = False
        rowsTotal = rowsTotal + CDbl(agg(stem & "#ROWS"))
        amountTotal = amountTotal + CDbl(agg(stem & "AMOUNT"))
        secondTotal = secondTotal + CDbl(agg(stem & "SECOND_AMOUNT"))
    Next i
    ReconAssert allPreserved, "Every flushed breakdown value and measure preserved", n
    ReconAssert rowsTotal = 1900, "Breakdown contributing row counts conserved", n
    ReconAssert amountTotal = 1025950, "Breakdown first measure conserved", n
    ReconAssert secondTotal = -2051900, "Breakdown second measure conserved", n
    BeginBreakdown
    BreakdownRow 1301
    ReconAssert mBrkSlotCount = 1 And mBrkSum(1, 0) = 1 And mBrkSum(1, 1) = 1300, "Next aggregate starts with empty breakdown state (slots=" & mBrkSlotCount & ", rows=" & mBrkSum(1, 0) & ", amount=" & mBrkSum(1, 1) & ")", n
    ReturnActiveTable
    mEclData = oldData: RestoreBreakdownPlan oldPlan
    ForgetActiveTable
    ReconBreakdownCapacityRegressionTests = "PASS: " & n & " breakdown-capacity checks across 1,300 distinct values and 1,900 contributing rows"
    Exit Function
Failed:
    errorText = Err.description
    ReturnActiveTable
    mEclData = oldData: RestoreBreakdownPlan oldPlan
    ForgetActiveTable
    Err.Raise vbObjectError + 753, "ReconBreakdownCapacityRegressionTests", errorText
End Function


' Remove only the selected source, including persisted picker data.
Public Sub ClearLoadedSourceData(ByVal key As String)
    Dim ws As Worksheet, r As Long, k As Variant, selected As Boolean
    key = UCase$(Trim$(key))
    If key <> "ALL" And key <> "ECL" And key <> "CAPRWA" And key <> "LL" And key <> "LCR" And key <> "CAP" And key <> "NSFR" Then Err.Raise 5, , "Unknown input source."
    EnsureSourceStore
    If key = "ALL" Then
        Set mSources = NewMap()
    ElseIf mSources.Exists(key) Then
        mSources.Remove key
    End If
    ReturnActiveTable
    mEclData = Empty: Set mEclHeaders = Nothing: Set mEclAllFields = Nothing
    ForgetActiveTable
    mActiveSource = "": mEclPath = "": mEclSheet = "": mEclHeaderRow = 0: mEclDate1904 = False
    Set mJoinCache = Nothing: Set mSourceDateCache = Nothing: Set mPreShockMeasureFields = Nothing
    Set mMaskStore = Nothing: mMaskBytes = 0: Set mActiveRecord = Nothing
    mBreakdownOn = False: mBrkDimCount = 0: mBrkMeasureCount = 0: mBrkSlotCount = 0
    Erase mBrkMap: Erase mBrkSum: Erase mBrkSlotDim: Erase mBrkSlotVal
    Set ws = GetWorksheetSafe(ThisWorkbook, PS_SOURCES_SHEET)
    If Not ws Is Nothing Then
        For r = PS_SRC_LIST_ROW To PS_SRC_LIST_LAST_ROW
            If key = "ALL" Or SafeUpperText(ws.Cells(r, 1).Value2) = key Then
                ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).ClearContents
                ws.Range(ws.Cells(r, 7), ws.Cells(r, 8)).ClearContents
                If Len(SafeText(ws.Cells(r, 1).Value2)) > 0 Then ws.Cells(r, 8).Value2 = "Not loaded"
            End If
        Next r
    End If
    Set ws = GetWorksheetSafe(ThisWorkbook, CACHE_SHEET)
    If Not ws Is Nothing Then
        For r = 3 To 8
            If key = "ALL" Or SafeUpperText(ws.Cells(r, 1).Value2) = key Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 7)).ClearContents
        Next r
        For r = CACHE_FIRST_ROW To ws.Cells(ws.rows.count, 1).End(xlUp).row
            If key = "ALL" Or SafeUpperText(ws.Cells(r, 4).Value2) = key Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).ClearContents
        Next r
    End If
    Set ws = GetWorksheetSafe(ThisWorkbook, PRE_SHOCK_SHEET)
    If Not ws Is Nothing Then ws.Range("I12:M13").ClearContents
    Set ws = GetWorksheetSafe(ThisWorkbook, "Base_Testing")
    If Not ws Is Nothing Then ws.Range("B6:D8,F7:H8").ClearContents
End Sub

' ---------------------------------------------------------------------------
'  GREATEST(a, b) inside an aggregate: the larger of two columns on the SAME row,
'  summed afterwards. This is how Moody's stress ECL is compared with ECL account
'  by account (ST Design: max(stress ECL, ECL) per account) when the extract does
'  not already carry the maximum as its own column. Each argument may list
'  alternative column names with |, exactly like SUM(ECL|OVERRIDDEN_ECL).
'
'  The measure map holds -1 for a GREATEST spec, and the columns it compares are
'  kept here, per source, because the same spec resolves to different columns in
'  different extracts.
'  (mGreatest is declared at the top of the module, beside mPreShockMeasureFields.)
' ---------------------------------------------------------------------------

Private Function IsGreatestSpec(ByVal spec As String) As Boolean
    IsGreatestSpec = (UCase$(Left$(Replace(Trim$(spec), " ", ""), 9)) = "GREATEST(")
End Function

Private Function GreatestArgs(ByVal spec As String) As Variant
    Dim s As String, p As Long
    s = Trim$(spec)
    p = InStr(1, s, "(")
    s = Mid$(s, p + 1)
    If Right$(s, 1) = ")" Then s = Left$(s, Len(s) - 1)
    GreatestArgs = Split(s, ",")
End Function

' GREATEST(A|B, C) -> A|B|C, so every column it may read is loaded.
Private Function GreatestFlatten(ByVal spec As String) As String
    Dim a As Variant, i As Long, s As String
    If Not IsGreatestSpec(spec) Then GreatestFlatten = spec: Exit Function
    a = GreatestArgs(spec)
    For i = LBound(a) To UBound(a)
        If Len(s) > 0 Then s = s & "|"
        s = s & Trim$(CStr(a(i)))
    Next i
    GreatestFlatten = s
End Function

' The column a measure reads in the active source: 0 when it (or any GREATEST
' argument) is missing, -1 for a resolved GREATEST, otherwise the column number.
Private Function MeasureColumnFor(ByVal spec As String) As Long
    Dim a As Variant, i As Long, cols() As Long, c As Long
    If Not IsGreatestSpec(spec) Then MeasureColumnFor = FirstExistingEclColumn(spec): Exit Function
    If mEclHeaders Is Nothing Then Exit Function
    a = GreatestArgs(spec)
    If UBound(a) < LBound(a) Then Exit Function
    ReDim cols(LBound(a) To UBound(a))
    For i = LBound(a) To UBound(a)
        c = FirstExistingEclColumn(Trim$(CStr(a(i))))
        If c = 0 Then Exit Function
        cols(i) = c
    Next i
    If mGreatest Is Nothing Then Set mGreatest = NewMap()
    mGreatest(UCase$(mActiveSource & "|" & Trim$(spec))) = cols
    MeasureColumnFor = -1
End Function

' The larger of the row's values. Empty when every argument is blank, so the row
' is reported exactly as a blank in an ordinary SUM column is.
Private Function GreatestOfRow(ByVal r As Long, ByVal spec As String) As Variant
    Dim k As String, cols As Variant, i As Long, v As Variant, have As Boolean, best As Double
    If mGreatest Is Nothing Then Exit Function
    k = UCase$(mActiveSource & "|" & Trim$(spec))
    If Not mGreatest.Exists(k) Then Exit Function
    cols = mGreatest(k)
    For i = LBound(cols) To UBound(cols)
        v = mEclData(r, cols(i))
        If Not IsError(v) Then
            If Not IsEmpty(v) Then
                If IsNumeric(v) Then
                    If Not have Then
                        best = CDbl(v): have = True
                    ElseIf CDbl(v) > best Then
                        best = CDbl(v)
                    End If
                End If
            End If
        End If
    Next i
    If have Then GreatestOfRow = best
End Function

' ============================================================================
'  SPEED: the engine's working memory (performance patch, 2026-09-26)
'
'  Three things made a run slow, and none of them changed a number:
'
'   1. Pointing the engine at another source copied that source's whole table,
'      every cell of it, and the run did that dozens of times per test case
'      (base and pre-shock per source, and twice more for every ECL join).
'      ActivateSource now keeps the table it already holds.
'   2. Every row of every aggregate walked the filter tree and normalised the
'      same cell values again. A condition is now decided once per loaded table
'      for all rows at once (a row mask), and every test case that asks the
'      same condition of the same table reuses it.
'   3. Every matching row built about twenty text keys for its totals. Totals
'      now build up in arrays and become the same keys once, after the last row.
'
'  ReconSpeedEquivalenceTests proves the fast paths return exactly what the
'  evaluator and the accumulator they replace return, row for row and bit for
'  bit. SpeedBenchmark times both on a synthetic extract.
' ============================================================================

Private Function NewExactMap() As Object
    Set NewExactMap = CreateObject("Scripting.Dictionary")
    NewExactMap.CompareMode = vbBinaryCompare
End Function

' Anything that assigns mEclData directly calls this, so ActivateSource never
' mistakes an edited table for the stored one.
Private Sub ForgetActiveTable()
    Dim k As Variant
    ' Still on loan means the lent table was overwritten without being handed
    ' back. Its record cannot be trusted, so it is dropped: the input then reads
    ' as not loaded and has to be loaded again, rather than feeding a figure.
    If mLent Then
        mLent = False
        If Not mSources Is Nothing Then
            For Each k In mSources.keys
                If mSources(k) Is mActiveRecord Then mSources.Remove k: Exit For
            Next k
        End If
        Set mActiveRecord = Nothing
        Err.Raise vbObjectError + 795, "ForgetActiveTable", "A loaded input table was replaced without being handed back. Load the inputs again, and report this."
    End If
    Set mActiveRecord = Nothing
End Sub

' A fresh identity for every table put into the source store, and its row count,
' so counting rows no longer copies them.
Private Sub StampRecord(ByVal s As Object, ByRef data As Variant)
    mTableStamp = mTableStamp + 1
    s("Stamp") = mTableStamp
    s("Rows") = 0#
    If IsArray(data) Then s("Rows") = CDbl(UBound(data, 1) - 1)
End Sub

' What was worked out about a table that is being replaced goes with it.
Private Sub DropTableMemo(ByVal key As String)
    Dim old As Object, k As String
    If mSources Is Nothing Or mMaskStore Is Nothing Then Exit Sub
    If Not mSources.Exists(SafeUpperText(key)) Then Exit Sub
    Set old = mSources(SafeUpperText(key))
    If Not old.Exists("Stamp") Then Exit Sub
    k = CStr(old("Stamp"))
    If mMaskStore.Exists(k) Then
        mMaskBytes = mMaskBytes - CDbl(mMaskStore(k)("#BYTES"))
        If mMaskBytes < 0 Then mMaskBytes = 0
        mMaskStore.Remove k
    End If
End Sub

' The stamp of the table the engine holds, or 0 when it cannot be vouched for:
' a table assigned directly (the regression tests do this) is always evaluated
' row by row, exactly as before.
Private Function ActiveTableStamp() As Long
    If mActiveRecord Is Nothing Then Exit Function
    If Not IsArray(mEclData) Then Exit Function
    If mActiveRecord.Exists("Stamp") Then ActiveTableStamp = CLng(mActiveRecord("Stamp"))
End Function

Private Function TableMemo(ByVal stamp As Long) As Object
    Dim k As String, memo As Object
    If mMaskStore Is Nothing Then Set mMaskStore = NewExactMap(): mMaskBytes = 0
    k = CStr(stamp)
    If Not mMaskStore.Exists(k) Then
        Set memo = NewExactMap(): memo("#BYTES") = 0#
        mMaskStore.Add k, memo
    End If
    Set TableMemo = mMaskStore(k)
End Function

' Bounded, so a long session cannot grow Excel past its memory: past the limit
' everything remembered is dropped and worked out again when next asked.
Private Sub RememberInMemo(ByVal memo As Object, ByVal k As String, ByRef v As Variant, ByVal bytes As Double)
    ' Dropped a moment ago: this memo is no longer in the store, so keep nothing.
    If mMaskStore Is Nothing Then Exit Sub
    If mMaskBytes + bytes > MEMO_BYTE_LIMIT Then
        Set mMaskStore = Nothing: mMaskBytes = 0
        Exit Sub
    End If
    memo(k) = v
    memo("#BYTES") = CDbl(memo("#BYTES")) + bytes
    mMaskBytes = mMaskBytes + bytes
End Sub

' One column of the active table, normalised the way filter conditions compare
' it, as a code per row into its list of distinct values. Worked out once per
' column, mode and table instead of once per condition per test case.
Private Sub ColumnCodes(ByVal col As Long, ByVal mode As String, ByVal memo As Object, _
                        ByRef codes() As Long, ByRef vals() As String)
    Dim k As String, seen As Object, r As Long, lastR As Long, v As String, n As Long, pack As Variant
    k = "C|" & col & "|" & mode
    If memo.Exists(k) Then
        pack = memo(k): codes = pack(0): vals = pack(1)
        Exit Sub
    End If
    lastR = UBound(mEclData, 1)
    ReDim codes(1 To lastR)
    ReDim vals(1 To 16)
    Set seen = NewExactMap()
    For r = 2 To lastR
        v = NormalizeFilterValue(mEclData(r, col), mode)
        If seen.Exists(v) Then
            codes(r) = seen(v)
        Else
            n = n + 1
            If n > UBound(vals) Then ReDim Preserve vals(1 To 2 * UBound(vals))
            vals(n) = v
            seen.Add v, n
            codes(r) = n
        End If
    Next r
    If n = 0 Then n = 1
    ReDim Preserve vals(1 To n)
    RememberInMemo memo, k, Array(codes, vals), 4# * lastR + 64# * n
End Sub

' The rows one IN / NOT IN condition selects. Same rule as EvalFilter: a row
' holds a value when any candidate column is neither blank nor NULL; IN wants
' one of the listed values in any candidate column; NOT IN wants a value and
' none of the listed ones.
Private Function AtomMask(ByVal node As Object, ByVal memo As Object) As Byte()
    Dim sig As String, k As Variant, wanted As Object, mode As String, notIn As Boolean
    Dim m() As Byte, hit() As Byte, hasV() As Byte, flagV() As Byte, flagHit() As Byte
    Dim codes() As Long, vals() As String, i As Long, r As Long, lastR As Long, col As Variant
    Set wanted = node("Values"): mode = CStr(node("Mode"))
    notIn = (node("Operator") = "NOT IN")
    sig = "M|" & CStr(node("Operator")) & "|" & mode & "|"
    For Each col In node("Columns"): sig = sig & CStr(col) & ",": Next col
    sig = sig & "|"
    For Each k In wanted.keys: sig = sig & CStr(k) & ChrW$(1): Next k
    If memo.Exists(sig) Then
        AtomMask = memo(sig)
        Exit Function
    End If
    lastR = UBound(mEclData, 1)
    ReDim m(1 To lastR): ReDim hit(1 To lastR): ReDim hasV(1 To lastR)
    For Each col In node("Columns")
        ColumnCodes CLng(col), mode, memo, codes, vals
        ' The condition is decided once per distinct value, not once per row.
        ReDim flagV(1 To UBound(vals)): ReDim flagHit(1 To UBound(vals))
        For i = 1 To UBound(vals)
            If Len(vals(i)) > 0 And vals(i) <> "NULL" Then
                flagV(i) = 1
                If wanted.Exists(vals(i)) Then flagHit(i) = 1
            End If
        Next i
        For r = 2 To lastR
            If hit(r) = 0 Then
                i = codes(r)
                If flagV(i) <> 0 Then
                    hasV(r) = 1
                    If flagHit(i) <> 0 Then hit(r) = 1
                End If
            End If
        Next r
    Next col
    For r = 2 To lastR
        If notIn Then
            If hasV(r) <> 0 And hit(r) = 0 Then m(r) = 1
        ElseIf hit(r) <> 0 Then
            m(r) = 1
        End If
    Next r
    RememberInMemo memo, sig, m, CDbl(lastR) + Len(sig)
    AtomMask = m
End Function

' A compiled filter over every row of the active table: 1 where it holds.
Private Function FilterMask(ByVal node As Object, ByVal memo As Object) As Byte()
    Dim m() As Byte, a() As Byte, child As Variant, r As Long, lastR As Long, t As String
    lastR = UBound(mEclData, 1)
    t = CStr(node("Type"))
    If t = "ALL" Then
        ReDim m(1 To lastR)
        For r = 2 To lastR: m(r) = 1: Next r
    ElseIf t = "AND" Then
        ReDim m(1 To lastR)
        For r = 2 To lastR: m(r) = 1: Next r
        For Each child In node("Children")
            a = FilterMask(child, memo)
            For r = 2 To lastR
                If a(r) = 0 Then m(r) = 0
            Next r
        Next child
    ElseIf t = "OR" Then
        ReDim m(1 To lastR)
        For Each child In node("Children")
            a = FilterMask(child, memo)
            For r = 2 To lastR
                If a(r) <> 0 Then m(r) = 1
            Next r
        Next child
    Else
        m = AtomMask(node, memo)
    End If
    FilterMask = m
End Function

' The rows of the active table a compiled filter selects, as a Byte array (1 =
' selected, index = row, row 1 is the header). Empty when the active table
' cannot be vouched for; callers then evaluate row by row with EvalForActive.
Public Function FilterMaskForActive(ByVal node As Object) As Variant
    Dim stamp As Long
    If node Is Nothing Then Exit Function
    stamp = ActiveTableStamp()
    If stamp = 0 Then Exit Function
    FilterMaskForActive = FilterMask(node, TableMemo(stamp))
End Function

' The join-cache key AccountSetFor would build, worked out without switching the
' engine to ECL. "" whenever that cannot be done with certainty, and the full
' path then runs exactly as before.
Private Function JoinCacheKeyQuiet(ByVal expr As String, ByVal aod As String) As String
    Dim eclHeaders As Object, scopedExpr As String, matchDates As Boolean
    If mJoinCache Is Nothing Or mSources Is Nothing Then Exit Function
    If Not mSources.Exists("ECL") Then Exit Function
    ' The engine must be holding a stored table, which is what switching back
    ' afterwards would have restored.
    If mActiveRecord Is Nothing Then Exit Function
    If Not mSources.Exists(mActiveSource) Then Exit Function
    If Not (mSources(mActiveSource) Is mActiveRecord) Then Exit Function
    On Error GoTo Unknown
    Set eclHeaders = mSources("ECL")("Headers")
    matchDates = Not IgnoreAsOfDate()
    scopedExpr = AndTerms(modScenarioBuilder_Multi.ApplyEntityScope(expr), _
                         modScenarioBuilder_Multi.SourceBankFilter(eclHeaders, "ECL"))
    JoinCacheKeyQuiet = UCase$(aod & "|" & CStr(matchDates) & "|" & Trim$(scopedExpr))
    Exit Function
Unknown:
    JoinCacheKeyQuiet = ""
End Function

Private Function RowSelected(ByVal node As Object, ByRef mask As Variant, ByVal useMask As Boolean, ByVal r As Long) As Boolean
    If useMask Then RowSelected = (mask(r) <> 0) Else RowSelected = EvalFilter(node, r)
End Function

Private Function StageSlot(ByVal st As String) As Long
    Select Case st
        Case "0": StageSlot = 0
        Case "1": StageSlot = 1
        Case "2": StageSlot = 2
        Case "3": StageSlot = 3
        Case Else: StageSlot = -1
    End Select
End Function

' The measures an aggregate totals, in the order the measure map holds them.
' Two specs that name the same aggregate key share one slot, so they add up in
' the same order AddGenericAgg added them.
Private Sub PlanMeasureSlots(ByRef specName() As String, ByRef specCol() As Long, ByRef specSlot() As Long, _
                             ByRef slotName() As String, ByRef nSpec As Long, ByRef nSlot As Long)
    Dim fieldSpec As Variant, slotOf As Object, i As Long, nm As String
    nSpec = 0: nSlot = 0
    If Not mPreShockMeasureFields Is Nothing Then nSpec = mPreShockMeasureFields.count
    ReDim specName(1 To IIf(nSpec > 0, nSpec, 1))
    ReDim specCol(1 To IIf(nSpec > 0, nSpec, 1))
    ReDim specSlot(1 To IIf(nSpec > 0, nSpec, 1))
    ReDim slotName(1 To IIf(nSpec > 0, nSpec, 1))
    If nSpec = 0 Then Exit Sub
    Set slotOf = NewMap()
    For Each fieldSpec In mPreShockMeasureFields.keys
        i = i + 1
        specName(i) = CStr(fieldSpec)
        specCol(i) = CLng(mPreShockMeasureFields(fieldSpec))
        nm = SafeUpperText(specName(i))
        If Not slotOf.Exists(nm) Then
            nSlot = nSlot + 1
            slotOf(nm) = nSlot
            slotName(nSlot) = specName(i)
        End If
        specSlot(i) = CLng(slotOf(nm))
    Next fieldSpec
End Sub

' The totals, written under the keys AggregateFunctionValue reads: the keys
' AddGenericAgg and AddAgg would have written, with the same values. A key
' exists only if something was added to it, exactly as before.
Private Sub FlushMeasureSlots(ByVal agg As Object, ByRef slotName() As String, ByVal nSlot As Long, _
                              ByRef rowsAt() As Double, ByRef accSum() As Double, ByRef accCnt() As Double, _
                              ByRef accMin() As Double, ByRef accMax() As Double, ByRef accInv() As Double)
    Dim si As Variant, s As Long, k As Long, st As String
    For Each si In Array(1, 2, 3, 0)
        s = CLng(si): st = CStr(s)
        If rowsAt(s) > 0 Then agg("ROWS#" & st) = rowsAt(s)
        For k = 1 To nSlot
            If accInv(k, s) > 0 Then agg(GenericAggKey("INVALID", slotName(k), st)) = accInv(k, s)
            If accCnt(k, s) > 0 Then
                agg(GenericAggKey("SUM", slotName(k), st)) = accSum(k, s)
                agg(GenericAggKey("COUNT", slotName(k), st)) = accCnt(k, s)
                agg(GenericAggKey("MIN", slotName(k), st)) = accMin(k, s)
                agg(GenericAggKey("MAX", slotName(k), st)) = accMax(k, s)
            End If
        Next k
    Next si
End Sub

' testcase|element -> the FIRST row that carries it, read in one pass with the
' matching rules CaseRowFor applies (it stops at the first blank test case).
Private Function CaseRowIndex(ByVal ws As Worksheet) As Object
    Dim d As Object, a As Variant, i As Long, k As String
    If ws Is Nothing Then Exit Function
    Set d = NewExactMap()
    a = ws.Range(ws.Cells(PS_CASE_FIRST_ROW, 2), ws.Cells(PS_CASE_LAST_ROW, 3)).Value2
    For i = 1 To UBound(a, 1)
        If Len(SafeText(a(i, 1))) = 0 Then Exit For
        k = KeyPart(SafeUpperText(a(i, 1))) & KeyPart(SafeUpperText(a(i, 2)))
        If Not d.Exists(k) Then d.Add k, PS_CASE_FIRST_ROW + i - 1
    Next i
    Set CaseRowIndex = d
End Function

Private Function CaseRowFromIndex(ByVal index As Object, ByVal tcCode As String, ByVal elCode As String) As Long
    Dim k As String
    ' No Cases sheet: the same "object not set" the sheet scan stopped on.
    If index Is Nothing Then Err.Raise 91
    k = KeyPart(SafeUpperText(tcCode)) & KeyPart(SafeUpperText(elCode))
    If index.Exists(k) Then CaseRowFromIndex = CLng(index(k))
End Function

' Distinct non-blank values in one column of the active table, remembered per
' table. The same count the availability pass used to rebuild on every refresh.
Private Function DistinctValueCount(ByVal col As Long) As Long
    Dim stamp As Long, memo As Object, k As String, seen As Object, rr As Long
    stamp = ActiveTableStamp()
    If stamp > 0 Then
        Set memo = TableMemo(stamp): k = "D|" & col
        If memo.Exists(k) Then DistinctValueCount = CLng(memo(k)): Exit Function
    End If
    Set seen = NewMap()
    For rr = 2 To UBound(mEclData, 1)
        If Len(SafeText(mEclData(rr, col))) > 0 Then seen(SafeText(mEclData(rr, col))) = True
    Next rr
    DistinctValueCount = seen.count
    If stamp > 0 Then RememberInMemo memo, k, seen.count, 16#
End Function

' ============================================================================
'  Proof that the fast paths change nothing.
' ============================================================================

' The aggregate exactly as GetEclAggregate built it before the speed patch:
' EvalFilter on every row, AddGenericAgg for every measure. Tests only.
Private Function SpeedReferenceAggregate(ByVal aod As String, ByVal caseExpr As String, _
                                         ByVal localExpr As String, ByVal fieldMap As Object) As Object
    Dim agg As Object, node As Object, unsupported As String, expr As String, r As Long, st As String
    Dim matchDates As Boolean, rowDates As Variant, dateOk As Boolean, matchRow As Boolean, hasStage As Boolean
    Dim fieldSpec As Variant, c As Long, joinMode As Boolean, acctSet As Object, joinHit As Object, iJoinAcct As Long
    Dim dateInfo As Object
    matchDates = Not IgnoreAsOfDate()
    If matchDates And Not mEclHeaders.Exists("AS_OF_DATE") Then
        If mActiveSource = "CAP" And SourceFamilyOf(mEclHeaders) = "CAP" And Len(Trim$(caseExpr)) = 0 Then
            matchDates = False
        Else
            Err.Raise vbObjectError + 741, , mActiveSource & ": AS_OF_DATE is required for date-matched reconciliation."
        End If
    End If
    If matchDates Then
        Set dateInfo = SourceDateInventory()
        RequireSourceDate dateInfo, aod
        rowDates = dateInfo("Rows")
    End If
    localExpr = AndTerms(Trim$(localExpr), modScenarioBuilder_Multi.SourceBankFilter(mEclHeaders, mActiveSource))
    expr = AndTerms(Trim$(caseExpr), Trim$(localExpr))
    If Len(Trim$(expr)) > 0 Then Set node = CompileFilter(Trim$(expr), fieldMap, unsupported)
    If Len(unsupported) > 0 Then
        If Not CanJoinThroughEcl() Then Err.Raise vbObjectError + 721, , "ECL source cannot map filter field(s): " & unsupported
        Set acctSet = SpeedReferenceAccounts(Trim$(caseExpr), aod)
        iJoinAcct = FirstExistingEclColumn("ACCOUNT_NUMBER")
        If iJoinAcct = 0 Then Err.Raise vbObjectError + 722, , mActiveSource & " cannot be joined to the ECL extract: it has no account number column."
        joinMode = True: unsupported = "": Set node = Nothing
        If Len(Trim$(localExpr)) > 0 Then
            Set node = CompileFilter(Trim$(localExpr), fieldMap, unsupported)
            If Len(unsupported) > 0 Then Err.Raise vbObjectError + 723, , mActiveSource & " cannot map its own metric or entity condition: " & unsupported
        End If
    End If
    Set agg = NewMap(): Set joinHit = NewMap()
    hasStage = (SourceFamilyOf(mEclHeaders) = "ECL" Or SourceFamilyOf(mEclHeaders) = "CAPRWA")
    For r = 2 To UBound(mEclData, 1)
        dateOk = True
        If matchDates Then dateOk = (rowDates(r) = aod)
        If Not dateOk Then
            matchRow = False
        ElseIf joinMode Then
            matchRow = acctSet.Exists(SafeUpperText(mEclData(r, iJoinAcct)))
            If matchRow And Not (node Is Nothing) Then matchRow = EvalFilter(node, r)
            If matchRow Then If Not joinHit.Exists(SafeUpperText(mEclData(r, iJoinAcct))) Then joinHit(SafeUpperText(mEclData(r, iJoinAcct))) = True
        ElseIf node Is Nothing Then
            matchRow = True
        Else
            matchRow = EvalFilter(node, r)
        End If
        If matchRow Then
            If hasStage Then
                st = NormalizeStage(mEclData(r, mEclHeaders("STAGE")))
                If Len(st) = 0 Then Err.Raise vbObjectError + 742, , mActiveSource & ": invalid or missing IFRS stage at source row " & (r + mEclHeaderRow - 1) & "."
            Else
                st = "0"
            End If
            If st = "1" Or st = "2" Or st = "3" Or st = "0" Then
                AddAgg agg, "ROWS#" & st, 1
                If Not mPreShockMeasureFields Is Nothing Then
                    For Each fieldSpec In mPreShockMeasureFields.keys
                        c = CLng(mPreShockMeasureFields(fieldSpec))
                        If c > 0 Then
                            AddGenericAgg agg, CStr(fieldSpec), st, mEclData(r, c)
                        ElseIf c < 0 Then
                            AddGenericAgg agg, CStr(fieldSpec), st, GreatestOfRow(r, CStr(fieldSpec))
                        End If
                    Next fieldSpec
                End If
            End If
        End If
    Next r
    If joinMode Then
        agg("JOIN#WANTED") = CDbl(acctSet.count)
        agg("JOIN#FOUND") = CDbl(joinHit.count)
    End If
    Set SpeedReferenceAggregate = agg
End Function

' The ECL accounts a condition selects, the way AccountSetFor resolved them
' before the speed patch: ECL copied in, every row evaluated, the table copied back.
Private Function SpeedReferenceAccounts(ByVal expr As String, ByVal aod As String) As Object
    Dim keep As String, d As Object, node As Object, unsupported As String, r As Long, iAcct As Long
    Dim scopedExpr As String, matchDates As Boolean, rowDates As Variant, acct As String
    Dim dateInfo As Object
    Set d = NewMap(): keep = mActiveSource
    ReturnActiveTable
    If Not ActivateSource("ECL") Then Set SpeedReferenceAccounts = d: Exit Function
    matchDates = Not IgnoreAsOfDate()
    If matchDates Then
        If Not mEclHeaders.Exists("AS_OF_DATE") Then Err.Raise vbObjectError + 741, , "ECL: AS_OF_DATE is required for account joins."
        Set dateInfo = SourceDateInventory()
        RequireSourceDate dateInfo, aod
        rowDates = dateInfo("Rows")
    End If
    scopedExpr = AndTerms(modScenarioBuilder_Multi.ApplyEntityScope(expr), _
                         modScenarioBuilder_Multi.SourceBankFilter(mEclHeaders, "ECL"))
    iAcct = FirstExistingEclColumn("ACCOUNT_NUMBER")
    If Len(Trim$(scopedExpr)) > 0 Then Set node = CompileFilter(Trim$(scopedExpr), LoadEclFilterMap(), unsupported)
    If Len(unsupported) > 0 Then Err.Raise vbObjectError + 721, , "No loaded extract can map filter field(s): " & unsupported
    For r = 2 To UBound(mEclData, 1)
        If matchDates Then If rowDates(r) <> aod Then GoTo NextRow
        If node Is Nothing Then
            acct = SafeUpperText(mEclData(r, iAcct)): If Len(acct) > 0 Then d(acct) = True
        ElseIf EvalFilter(node, r) Then
            acct = SafeUpperText(mEclData(r, iAcct)): If Len(acct) > 0 Then d(acct) = True
        End If
NextRow:
    Next r
    ReturnActiveTable
    ActivateSource keep
    Set SpeedReferenceAccounts = d
End Function

' Deterministic pseudo-random numbers (Park-Miller), so every run of the test
' builds the same synthetic extract.
Private Function SpeedNext(ByRef seed As Double) As Double
    seed = seed * 16807# - Int(seed * 16807# / 2147483647#) * 2147483647#
    SpeedNext = seed / 2147483647#
End Function

Private Function SpeedPick(ByRef seed As Double, ByVal choices As Variant) As Variant
    Dim i As Long
    i = LBound(choices) + Int(SpeedNext(seed) * (UBound(choices) - LBound(choices) + 1))
    If i > UBound(choices) Then i = UBound(choices)
    SpeedPick = choices(i)
End Function

' A synthetic ECL-shaped extract with the awkward values real ones carry:
' mixed stage spellings, sector codes and brackets, apostrophes and commas in
' labels, Y/N/1/0 flags, NULL text, blanks, #N/A, numbers held as text.
Private Function SpeedEclTable(ByVal nRows As Long, ByVal seed As Double, ByVal extraCols As Long) As Variant
    Dim a() As Variant, r As Long, c As Long, heads As Variant
    heads = Array("ACCOUNT_NUMBER", "BANK_ID", "BRANCH_CODE", "AS_OF_DATE", "STAGE", "SPEEDTEST_SECTOR", _
                  "SPEEDTEST_CATEGORY", "SPEEDTEST_CATEGORY2", "SPEEDTEST_FLAG", "OUTSTANDING_LCY", "ECL", "ECL_STRESS")
    ReDim a(1 To nRows + 1, 1 To 12 + extraCols)
    For c = 0 To 11: a(1, c + 1) = heads(c): Next c
    For c = 1 To extraCols: a(1, 12 + c) = "SPEEDTEST_EXTRA_" & c: Next c
    For r = 2 To nRows + 1
        a(r, 1) = "A" & format$(Int(SpeedNext(seed) * (nRows / 3 + 1)), "000000")
        a(r, 2) = SpeedPick(seed, Array(101, 101, 101, "101", 102))
        a(r, 3) = SpeedPick(seed, Array(800, "800", 101, "0101", "", "NULL", 205))
        a(r, 4) = CDbl(DateSerial(2025, 12, 31))
        a(r, 5) = SpeedPick(seed, Array("Stage 1", "Stage2 - SICR", "3", "NPA", "stage 2", 1, 2, 3, "Stage 3 ", "STAGE 1"))
        a(r, 6) = SpeedPick(seed, Array("01 Trade (A)", "02_INDUSTRY", "TRADE", "Real Estate", " 7 Real  estate (B)", Empty, "NULL", "trade"))
        a(r, 7) = SpeedPick(seed, Array("A, B", "Letters of Credit (LC's)", "Loans", "loans", Empty, "NULL", CVErr(xlErrNA), "Guarantees"))
        a(r, 8) = SpeedPick(seed, Array("Loans", Empty, "Guarantees", "A, B"))
        a(r, 9) = SpeedPick(seed, Array("Y", "N", "YES", "no", 1, 0, "TRUE", Empty, "maybe"))
        a(r, 10) = SpeedPick(seed, Array(Round(SpeedNext(seed) * 1000000#, 2), -Round(SpeedNext(seed) * 5000#, 3), 0, _
                                         "1234.5", "n/a", Empty, CVErr(xlErrDiv0), SpeedNext(seed) * 1E+15))
        a(r, 11) = SpeedPick(seed, Array(Round(SpeedNext(seed) * 20000#, 2), 0, Empty, 0.1, 7))
        a(r, 12) = SpeedPick(seed, Array(Round(SpeedNext(seed) * 30000#, 2), Empty, Empty, 3.3))
        For c = 1 To extraCols: a(r, 12 + c) = SpeedPick(seed, Array("X" & (r Mod 17), r * 1.5, Empty)): Next c
    Next r
    SpeedEclTable = a
End Function

Private Function SpeedHeaders(ByRef a As Variant) As Object
    Dim h As Object, c As Long
    Set h = NewMap()
    For c = 1 To UBound(a, 2): h(CStr(a(1, c))) = c: Next c
    Set SpeedHeaders = h
End Function

Private Function SpeedFieldMap() As Object
    Dim fm As Object, spec As Variant, rule As Object
    Set fm = NewMap()
    For Each spec In Array("SPEEDTEST_SECTOR|SPEEDTEST_SECTOR|SECTOR", "SPEEDTEST_SEG|SPEEDTEST_CATEGORY|SPEEDTEST_CATEGORY2|TEXT", _
                           "SPEEDTEST_FLAG|SPEEDTEST_FLAG|YESNO", "STAGE|STAGE|STAGE", "SPEEDTEST_UNDERSCORE|SPEEDTEST_SECTOR|UNDERSCORE_TEXT", _
                           "BRANCH_CODE|BRANCH_CODE|TEXT", "BANK_ID|BANK_ID|TEXT", "BANK_CODE|BANK_CODE|TEXT", "SPEEDTEST_CATEGORY|SPEEDTEST_CATEGORY|TEXT")
        Set rule = NewMap()
        rule("Candidates") = Split(CStr(spec), "|")(1)
        If UBound(Split(CStr(spec), "|")) = 3 Then rule("Candidates") = Split(CStr(spec), "|")(1) & "|" & Split(CStr(spec), "|")(2)
        rule("Mode") = Split(CStr(spec), "|")(UBound(Split(CStr(spec), "|")))
        Set fm(Split(CStr(spec), "|")(0)) = rule
    Next spec
    Set SpeedFieldMap = fm
End Function

Private Function SpeedFilters() As Variant
    SpeedFilters = Array("", "ALL", "SPEEDTEST_SECTOR IN ('TRADE')", _
        "SPEEDTEST_SECTOR NOT IN ('TRADE', 'INDUSTRY')", _
        "SPEEDTEST_SEG IN ('A, B', 'Letters of Credit (LC's)')", _
        "SPEEDTEST_FLAG IN ('YES') AND STAGE IN (2, 3)", _
        "( SPEEDTEST_FLAG IN ('NO') OR SPEEDTEST_SEG NOT IN ('Loans') ) AND SPEEDTEST_SECTOR IN ('REAL ESTATE')", _
        "SPEEDTEST_UNDERSCORE IN ('02 INDUSTRY', 'TRADE')", _
        "BRANCH_CODE NOT IN ('800') AND STAGE NOT IN (1)", _
        "SPEEDTEST_FLAG NOT IN ('YES') OR SPEEDTEST_CATEGORY IN ('Guarantees')", _
        "STAGE IN (3) AND ( SPEEDTEST_SEG IN ('Loans') OR SPEEDTEST_SECTOR IN ('TRADE') ) AND BRANCH_CODE IN ('205', '101')")
End Function

' Same keys, same values, bit for bit.
Private Function SpeedSameAggregate(ByVal want As Object, ByVal got As Object, ByRef why As String) As Boolean
    Dim k As Variant
    If want.count <> got.count Then why = want.count & " keys expected, " & got.count & " built": Exit Function
    For Each k In want.keys
        If Not got.Exists(k) Then why = "missing " & CStr(k): Exit Function
        If VarType(want(k)) <> VarType(got(k)) Then why = "type differs at " & CStr(k): Exit Function
        If CDbl(want(k)) <> CDbl(got(k)) Then why = CStr(k) & ": " & CStr(want(k)) & " expected, " & CStr(got(k)) & " built": Exit Function
    Next k
    SpeedSameAggregate = True
End Function

' Both succeed with the same aggregate, or both fail with the same error.
Private Sub SpeedCompare(ByVal label As String, ByVal aod As String, ByVal caseExpr As String, ByVal localExpr As String, _
                         ByVal fm As Object, ByRef n As Long)
    Dim want As Object, got As Object, wantErr As Long, gotErr As Long, wantText As String, gotText As String, why As String
    Dim keep As String
    keep = mActiveSource
    On Error Resume Next
    Set want = SpeedReferenceAggregate(aod, caseExpr, localExpr, fm)
    wantErr = Err.Number: wantText = Err.description: Err.Clear
    ' A reference that stopped part-way through a join may have left ECL active.
    If mActiveSource <> keep Then ReturnActiveTable: ActivateSource keep
    Set got = GetEclAggregate(NewMap(), aod, caseExpr, localExpr, fm)
    gotErr = Err.Number: gotText = Err.description: Err.Clear
    On Error GoTo 0
    ReconAssert wantErr = gotErr, label & " [" & caseExpr & "]: reference error " & wantErr & " '" & wantText & "', fast path error " & gotErr & " '" & gotText & "'", n
    If wantErr = 0 Then ReconAssert SpeedSameAggregate(want, got, why), label & " [" & caseExpr & "]: " & why, n
End Sub

Public Function ReconSpeedEquivalenceTests() As String
    Dim oldData As Variant, oldHeaders As Object, oldMeasures As Object, oldSources As Object, oldDates As Object, oldJoin As Object
    Dim oldSource As String, oldPath As String, oldSheet As String, oldHeaderRow As Long, old1904 As Boolean
    Dim oldRecord As Object, oldPlan As Object, oldGreatest As Object
    Dim a As Variant, b As Variant, ll As Variant, h As Object, fm As Object, node As Object, mask As Variant
    Dim e As Variant, needs As String, r As Long, n As Long, bad As Long, pass As Long, errorText As String
    Dim aod As String, seed As Double, want As Object, got As Object, f As Variant
    Dim s1 As String, w1 As String, s2 As String, w2 As String, x1 As String, x2 As String, fmLL As Object, rule As Object
    Dim wantErr As Long, gotErr As Long
    oldData = mEclData: Set oldHeaders = mEclHeaders: Set oldMeasures = mPreShockMeasureFields
    Set oldSources = mSources: Set oldDates = mSourceDateCache: Set oldJoin = mJoinCache: Set oldRecord = mActiveRecord
    oldSource = mActiveSource: oldPath = mEclPath: oldSheet = mEclSheet: oldHeaderRow = mEclHeaderRow: old1904 = mEclDate1904
    Set oldPlan = SnapshotBreakdownPlan(): Set oldGreatest = mGreatest
    On Error GoTo Failed
    Set mSources = NewMap(): ClearSourceDateCache: Set mJoinCache = Nothing
    mBreakdownOn = False: mBrkDimCount = 0: mBrkMeasureCount = 0
    aod = "2025-12-31"
    a = SpeedEclTable(2400, 12345#, 0)
    Set h = SpeedHeaders(a)
    Set fm = SpeedFieldMap()

    ' 1. Keeping the table: a switch back to the table already held changes
    '    nothing, and a table assigned directly is never mistaken for it.
    b = SpeedEclTable(50, 777#, 0)
    RegisterTable "SPEED_A", a, h, "", "Speed A": RegisterTable "SPEED_B", b, SpeedHeaders(b), "", "Speed B"
    ActivateSource "SPEED_A": ReconAssert mEclData(2, 1) = a(2, 1) And UBound(mEclData, 1) = UBound(a, 1), "Switch copies the requested table", n
    ActivateSource "SPEED_A": ReconAssert UBound(mEclData, 1) = UBound(a, 1) And mEclData(3, 5) = a(3, 5), "Switch to the table already held keeps it", n
    ActivateSource "SPEED_B": ReconAssert UBound(mEclData, 1) = UBound(b, 1) And mEclData(2, 1) = b(2, 1), "Switch to another table copies it", n
    ReturnActiveTable
    mEclData = a: ForgetActiveTable
    ActivateSource "SPEED_B": ReconAssert UBound(mEclData, 1) = UBound(b, 1), "A directly assigned table is never kept", n
    ReconAssert SourceRowCount("SPEED_A") = 2400 And SourceRowCount("SPEED_B") = 50, "Row counts without copying", n
    ReconAssert Not IsArray(FilterMaskForActive(Nothing)), "No mask without a filter", n
    ReturnActiveTable
    mEclData = a: ForgetActiveTable
    Set node = CompileFilter("SPEEDTEST_SECTOR IN ('TRADE')", fm, needs)
    ReconAssert Not IsArray(FilterMaskForActive(node)), "No mask for a directly assigned table", n

    ' 2. Row masks: every filter, every row, against EvalFilter; then again from memory.
    RegisterTable "ECL", a, h, "", "Speed ECL": ActivateSource "ECL"
    For pass = 1 To 2
        For Each e In SpeedFilters()
            If Len(CStr(e)) > 0 Then
                needs = ""
                Set node = CompileFilter(CStr(e), fm, needs)
                mask = FilterMaskForActive(node)
                ReconAssert IsArray(mask), "Mask built for " & CStr(e), n
                bad = 0
                For r = 2 To UBound(mEclData, 1)
                    If (mask(r) <> 0) <> EvalFilter(node, r) Then bad = bad + 1
                Next r
                ReconAssert bad = 0, "Mask agrees with EvalFilter on every row (pass " & pass & ", " & bad & " differ): " & CStr(e), n
            End If
        Next e
    Next pass

    ' 3. Aggregates: plain measures, a missing one, GREATEST, invalid values,
    '    unfiltered, filtered and with a local selector, under the date rule in force.
    Set mPreShockMeasureFields = NewMap()
    mPreShockMeasureFields("OUTSTANDING_LCY") = CLng(h("OUTSTANDING_LCY"))
    mPreShockMeasureFields("ECL") = CLng(h("ECL"))
    mPreShockMeasureFields("SPEEDTEST_MISSING") = 0
    mPreShockMeasureFields("GREATEST(ECL, ECL_STRESS)") = MeasureColumnFor("GREATEST(ECL, ECL_STRESS)")
    ReconAssert CLng(mPreShockMeasureFields("GREATEST(ECL, ECL_STRESS)")) = -1, "GREATEST resolves", n
    For Each e In SpeedFilters()
        SpeedCompare "ECL aggregate", aod, CStr(e), "", fm, n
        SpeedCompare "ECL aggregate with selector", aod, CStr(e), "SPEEDTEST_FLAG IN ('YES')", fm, n
    Next e
    SpeedCompare "Absent reporting date", "2024-12-31", "SPEEDTEST_SECTOR IN ('TRADE')", "", fm, n
    ' A matched row with no valid stage stops both, with the same message.
    a(5, 5) = "Stage10": a(5, 2) = 101: RegisterTable "ECL", a, h, "", "Speed ECL bad stage": ActivateSource "ECL"
    SpeedCompare "Invalid stage", aod, "", "", fm, n

    ' 4. Joins through ECL: a liquidity-shaped source filtered by an ECL-only field.
    a(5, 5) = "Stage 1": RegisterTable "ECL", a, h, "", "Speed ECL": Set mJoinCache = Nothing
    seed = 4242#
    ReDim ll(1 To 1501, 1 To 5)
    ll(1, 1) = "ACCOUNT_NUMBER": ll(1, 2) = "BANK_CODE": ll(1, 3) = "AS_OF_DATE"
    ll(1, 4) = "CASHFLOW_AMOUNT_LCY_PRE_FACTOR": ll(1, 5) = "COA_BALANCESHEET_CATEGORY"
    For r = 2 To 1501
        ll(r, 1) = SpeedPick(seed, Array(a(2 + Int(SpeedNext(seed) * 2400), 1), "GL" & r, Empty))
        ll(r, 2) = SpeedPick(seed, Array("JKB", "JKB", "EJARA"))
        ll(r, 3) = CDbl(DateSerial(2025, 12, 31))
        ll(r, 4) = SpeedPick(seed, Array(Round(SpeedNext(seed) * 90000#, 2), -12.5, "text", Empty))
        ll(r, 5) = SpeedPick(seed, Array("Assets", "Liabilities"))
    Next r
    RegisterTable "LL", ll, SpeedHeaders(ll), "", "Speed LL": ActivateSource "LL"
    Set mPreShockMeasureFields = NewMap(): mPreShockMeasureFields("CASHFLOW_AMOUNT_LCY_PRE_FACTOR") = 4
    ' The liquidity source's own fields only, as its field map would carry them.
    Set fmLL = NewMap()
    For Each f In Array("ACCOUNT_NUMBER", "BANK_CODE", "COA_BALANCESHEET_CATEGORY")
        Set rule = NewMap(): rule("Candidates") = CStr(f): rule("Mode") = "TEXT": Set fmLL(CStr(f)) = rule
    Next f
    For Each e In Array("SPEEDTEST_SECTOR IN ('TRADE')", "SPEEDTEST_FLAG IN ('YES') AND STAGE IN (2, 3)", "SPEEDTEST_CATEGORY NOT IN ('Loans')")
        ' The workbook's own field map resolves these, so both sides may also
        ' refuse a field; then they must refuse it the same way.
        On Error Resume Next
        Set want = SpeedReferenceAccounts(CStr(e), aod)
        wantErr = Err.Number: Err.Clear
        If mActiveSource <> "LL" Then ReturnActiveTable: ActivateSource "LL"
        Set got = AccountSetFor(CStr(e), aod)
        gotErr = Err.Number: Err.Clear
        On Error GoTo Failed
        If mActiveSource <> "LL" Then ReturnActiveTable: ActivateSource "LL"
        ReconAssert wantErr = gotErr, "Joined accounts fail alike (" & wantErr & " / " & gotErr & "): " & CStr(e), n
        If wantErr = 0 Then
            ReconAssert SpeedSameKeys(want, got), "Joined accounts agree: " & CStr(e), n
            Set got = AccountSetFor(CStr(e), aod)
            ReconAssert SpeedSameKeys(want, got) And mActiveSource = "LL", "Joined accounts from memory, LL still active: " & CStr(e), n
        End If
        SpeedCompare "LL joined through ECL", aod, CStr(e), "COA_BALANCESHEET_CATEGORY IN ('Assets')", fmLL, n
    Next e

    ' 5. Metric formula split: remembered answers equal fresh ones.
    For Each f In Array("SUM(ECL) WHERE STAGE=2", "SUM(CASHFLOW_AMOUNT_LCY_PRE_FACTOR) WHERE COA_BALANCESHEET_CATEGORY IN ('Assets')", _
                        "SUM(OUTSTANDING_LCY)+SUM(IIS)", "  sum(ecl) where stage in (3)", "")
        x1 = SplitMetricFormulaNow(CStr(f), s1, w1)
        x2 = SplitMetricFormula(CStr(f), s2, w2)
        x2 = SplitMetricFormula(CStr(f), s2, w2)
        ReconAssert x1 = x2 And s1 = s2 And w1 = w2, "Formula split remembered: " & CStr(f), n
    Next f

    ReconSpeedEquivalenceTests = "PASS: " & n & " speed-equivalence checks (row masks, table reuse, array totals, joins, formula split)"
    GoTo Restore
Failed:
    errorText = Err.description
Restore:
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: Set mPreShockMeasureFields = oldMeasures
    Set mSources = oldSources: Set mSourceDateCache = oldDates: Set mJoinCache = oldJoin
    mActiveSource = oldSource: mEclPath = oldPath: mEclSheet = oldSheet: mEclHeaderRow = oldHeaderRow: mEclDate1904 = old1904
    Set mGreatest = oldGreatest
    RestoreBreakdownPlan oldPlan
    Set mActiveRecord = oldRecord
    ' The synthetic tables' remembered masks go with them.
    Set mMaskStore = Nothing: mMaskBytes = 0
    If Len(ReconSpeedEquivalenceTests) = 0 Then Err.Raise vbObjectError + 754, "ReconSpeedEquivalenceTests", errorText
End Function

Private Function SpeedSameKeys(ByVal want As Object, ByVal got As Object) As Boolean
    Dim k As Variant
    If want.count <> got.count Then Exit Function
    For Each k In want.keys
        If Not got.Exists(k) Then Exit Function
    Next k
    SpeedSameKeys = True
End Function

' Times the engine as it was against the engine as it is, on a synthetic
' extract shaped like a real one: two sources, one aggregate per test case per
' source, switching source for each, as a run does. Returns a readable line.
Public Function SpeedBenchmark(Optional ByVal nRows As Long = 60000, Optional ByVal nCases As Long = 24, _
                               Optional ByVal extraCols As Long = 30) As String
    Dim oldData As Variant, oldHeaders As Object, oldMeasures As Object, oldSources As Object, oldDates As Object, oldJoin As Object
    Dim oldSource As String, oldRecord As Object, oldPlan As Object, oldGreatest As Object
    Dim a As Variant, b As Variant, h As Object, hb As Object, fm As Object, filters As Variant, i As Long, j As Long
    Dim t0 As Double, tOld As Double, tNew As Double, want As Object, got As Object, why As String, same As Boolean, errorText As String
    Dim wantList As Collection, cache As Object
    oldData = mEclData: Set oldHeaders = mEclHeaders: Set oldMeasures = mPreShockMeasureFields
    Set oldSources = mSources: Set oldDates = mSourceDateCache: Set oldJoin = mJoinCache: Set oldRecord = mActiveRecord
    oldSource = mActiveSource: Set oldPlan = SnapshotBreakdownPlan(): Set oldGreatest = mGreatest
    On Error GoTo Failed
    Set mSources = NewMap(): ClearSourceDateCache: Set mJoinCache = Nothing
    mBreakdownOn = False: mBrkDimCount = 0: mBrkMeasureCount = 0
    a = SpeedEclTable(nRows, 9876#, extraCols): Set h = SpeedHeaders(a)
    b = SpeedEclTable(nRows \ 2, 5555#, extraCols): Set hb = SpeedHeaders(b)
    RegisterTable "ECL", a, h, "", "Benchmark ECL": RegisterTable "CAPRWA", b, hb, "", "Benchmark second source"
    Set fm = SpeedFieldMap(): filters = SpeedFilters()
    Set wantList = New Collection

    ' As it was: every switch copies the table, every row is evaluated and
    ' totalled through text keys.
    t0 = Timer
    For i = 1 To nCases
        For j = 1 To 2
            ReturnActiveTable
            ActivateSource IIf(j = 1, "ECL", "CAPRWA")
            SpeedCopyCost
            SpeedMeasures
            wantList.Add SpeedReferenceAggregate("2025-12-31", CStr(filters(i Mod (UBound(filters) + 1))), "", fm)
        Next j
    Next i
    tOld = Timer - t0: If tOld < 0 Then tOld = tOld + 86400

    ' As it is.
    Set mMaskStore = Nothing: mMaskBytes = 0
    ClearSourceDateCache
    t0 = Timer: same = True
    For i = 1 To nCases
        For j = 1 To 2
            ActivateSource IIf(j = 1, "ECL", "CAPRWA")
            SpeedMeasures
            Set got = GetEclAggregate(NewMap(), "2025-12-31", CStr(filters(i Mod (UBound(filters) + 1))), "", fm)
            If same Then same = SpeedSameAggregate(wantList((i - 1) * 2 + j), got, why)
        Next j
    Next i
    tNew = Timer - t0: If tNew < 0 Then tNew = tNew + 86400

    SpeedBenchmark = "Speed benchmark: " & format$(nRows, "#,##0") & " + " & format$(nRows \ 2, "#,##0") & " rows x " & (12 + extraCols) & _
                     " columns, " & nCases & " test cases x 2 sources. Before " & format$(tOld, "0.00") & " s, after " & format$(tNew, "0.00") & _
                     " s (" & IIf(tNew > 0, format$(tOld / tNew, "0.0") & "x faster", "instant") & "). Results " & IIf(same, "identical", "DIFFER: " & why) & "."
    GoTo Restore
Failed:
    errorText = Err.description
Restore:
    ReturnActiveTable
    mEclData = oldData: Set mEclHeaders = oldHeaders: Set mPreShockMeasureFields = oldMeasures
    Set mSources = oldSources: Set mSourceDateCache = oldDates: Set mJoinCache = oldJoin
    mActiveSource = oldSource: Set mGreatest = oldGreatest
    RestoreBreakdownPlan oldPlan
    Set mActiveRecord = oldRecord
    Set mMaskStore = Nothing: mMaskBytes = 0
    If Len(SpeedBenchmark) = 0 Then SpeedBenchmark = "Speed benchmark failed: " & errorText
End Function

Private Sub SpeedMeasures()
    Set mPreShockMeasureFields = NewMap()
    mPreShockMeasureFields("OUTSTANDING_LCY") = FirstExistingEclColumn("OUTSTANDING_LCY")
    mPreShockMeasureFields("ECL") = FirstExistingEclColumn("ECL")
    mPreShockMeasureFields("GREATEST(ECL, ECL_STRESS)") = MeasureColumnFor("GREATEST(ECL, ECL_STRESS)")
End Sub

' Where the ECL file was loaded from, as the input registry recorded it.
Private Sub EclSavedLocation(ByRef path As String, ByRef sheetName As String)
    Dim ws As Worksheet, r As Long
    path = "": sheetName = ""
    r = SourceRow("ECL")
    If r = 0 Then Exit Sub
    Set ws = PsSourcesSheet()
    If ws Is Nothing Then Exit Sub
    path = SafeText(ws.Cells(r, 3).Value2)
    sheetName = SafeText(ws.Cells(r, 4).Value2)
    If Not LooksLikeFilePath(path) Then path = "": sheetName = ""
End Sub

' Text that could be a file name. Notes and messages are not, and handing them
' to Dir$ fails with "bad file name" rather than a message anyone can act on.
Private Function LooksLikeFilePath(ByVal s As String) As Boolean
    If Len(s) = 0 Or Len(s) > 259 Then Exit Function
    If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then Exit Function
    If InStr(s, "\") = 0 And InStr(s, "/") = 0 Then Exit Function
    LooksLikeFilePath = (InStr(s, "*") = 0 And InStr(s, "?") = 0 And InStr(s, "|") = 0 And InStr(s, "<") = 0 And InStr(s, ">") = 0)
End Function

' ======================= one copy of each input table =======================
'
' The table the engine is using is LENT to mEclData by its record, not copied
' (clsJkbTable.ExchangeAt): at any moment exactly one of the two holds it. Every
' direct assignment to mEclData hands the lent table back first
' (ReturnActiveTable), so the record never loses it.

Private Sub ReturnActiveTable()
    If Not mLent Then Exit Sub
    mLent = False
    If Not mActiveRecord Is Nothing Then
        mActiveRecord("Table").ExchangeAt VarPtr(mEclData)
        mEclData = Empty
    End If
    Set mActiveRecord = Nothing
End Sub

' Makes s the table in use: handed over when it can be, copied as before when not.
Private Sub TakeTable(ByVal s As Object)
    ReturnActiveTable
    If s.Exists("Table") Then
        mEclData = Empty
        s("Table").ExchangeAt VarPtr(mEclData)
        mLent = True
    Else
        mEclData = s("Data")
    End If
    Set mActiveRecord = s
End Sub

' The freshly loaded table in mEclData becomes s's own, still in use: no copy.
Private Sub KeepActiveAs(ByVal s As Object, ByVal key As String)
    Dim replacing As Boolean
    If mLent Then
        If mSources.Exists(SafeUpperText(key)) Then replacing = (mSources(SafeUpperText(key)) Is mActiveRecord)
        ' The table on loan is another record's; that record keeps a copy, as it did before.
        If Not replacing And Not mActiveRecord Is Nothing Then mActiveRecord("Table").SetCopy mEclData
        mLent = False
    End If
    Set s("Table") = New clsJkbTable
    mLent = True
    Set mActiveRecord = s
End Sub

Private Function TableFrom(ByRef data As Variant) As Object
    Dim t As clsJkbTable
    Set t = New clsJkbTable
    t.SetCopy data
    Set TableFrom = t
End Function

Private Function TableCopyOf(ByVal s As Object) As Variant
    If mLent And (s Is mActiveRecord) Then TableCopyOf = mEclData: Exit Function
    If s.Exists("Table") Then TableCopyOf = s("Table").CopyOut() Else TableCopyOf = s("Data")
End Function

' One cell of a stored table, without copying the table.
Public Function SourceCellFor(ByVal key As String, ByVal r As Long, ByVal c As Long) As Variant
    Dim s As Object
    EnsureSourceStore
    If Not mSources.Exists(SafeUpperText(key)) Then Exit Function
    Set s = mSources(SafeUpperText(key))
    If mLent And (s Is mActiveRecord) Then
        SourceCellFor = mEclData(r, c)
    ElseIf s.Exists("Table") Then
        SourceCellFor = s("Table").Cell(r, c)
    Else
        SourceCellFor = s("Data")(r, c)
    End If
End Function

' The benchmark's picture of the old switch: one full copy of the table in use.
Private Sub SpeedCopyCost()
    Dim t As Variant
    t = mEclData
End Sub

Private Function HandSameTable(ByRef x As Variant, ByRef y As Variant) As Boolean
    Dim r As Long, c As Long
    If Not IsArray(x) Or Not IsArray(y) Then Exit Function
    If UBound(x, 1) <> UBound(y, 1) Or UBound(x, 2) <> UBound(y, 2) Then Exit Function
    For r = LBound(x, 1) To UBound(x, 1)
        For c = LBound(x, 2) To UBound(x, 2)
            If VarType(x(r, c)) <> VarType(y(r, c)) Then Exit Function
            If Not IsEmpty(x(r, c)) And Not IsError(x(r, c)) Then
                If x(r, c) <> y(r, c) Then Exit Function
            End If
        Next c
    Next r
    HandSameTable = True
End Function

' Proves the hand-over keeps every table intact through switches, stores,
' direct assignments and re-registrations, and that a bypass fails loudly.
Public Function ReconTableHandoverTests() As String
    Dim oldSources As Object, oldDates As Object, oldJoin As Object, oldSource As String
    Dim a As Variant, b As Variant, i As Long, n As Long, errorText As String, gotErr As Long
    ReturnActiveTable
    Set oldSources = mSources: Set oldDates = mSourceDateCache: Set oldJoin = mJoinCache: oldSource = mActiveSource
    On Error GoTo Failed
    Set mSources = NewMap(): ClearSourceDateCache: Set mJoinCache = Nothing
    a = SpeedEclTable(300, 111#, 3): b = SpeedEclTable(40, 222#, 0)
    RegisterTable "HAND_A", a, SpeedHeaders(a), "", "Hand A": RegisterTable "HAND_B", b, SpeedHeaders(b), "", "Hand B"

    ActivateSource "HAND_A"
    ReconAssert mLent And HandSameTable(mEclData, a), "The table in use is handed over intact", n
    For i = 1 To 5
        ActivateSource "HAND_B": ActivateSource "HAND_A"
    Next i
    ReconAssert HandSameTable(mEclData, a), "Ten switches leave the table intact", n
    ReconAssert HandSameTable(SourceDataFor("HAND_B"), b) And HandSameTable(SourceDataFor("HAND_A"), a), "Stored and lent tables read back intact", n
    ReconAssert SourceCellFor("HAND_A", 3, 2) = a(3, 2) And SourceCellFor("HAND_B", 2, 1) = b(2, 1), "Single cells read without copying", n

    ReturnActiveTable
    mEclData = b: ForgetActiveTable
    ReconAssert Not mLent, "A direct assignment ends the loan", n
    ActivateSource "HAND_A"
    ReconAssert HandSameTable(mEclData, a), "A direct assignment never overwrites a stored table", n

    StoreActiveSource "HAND_C"
    ReconAssert HandSameTable(SourceDataFor("HAND_C"), a) And HandSameTable(SourceDataFor("HAND_A"), a), "Storing a lent table under a new name keeps both", n
    ActivateSource "HAND_A"
    ReconAssert HandSameTable(mEclData, a), "The original still activates intact", n

    ActivateSource "HAND_B"
    mEclData(2, 1) = "changed"
    StoreActiveSource "HAND_B"
    ReconAssert SourceCellFor("HAND_B", 2, 1) = "changed" And HandSameTable(SourceDataFor("HAND_A"), a), "Re-storing an edited table keeps the edit, and nothing else changes", n

    ActivateSource "HAND_A"
    RegisterTable "HAND_A", b, SpeedHeaders(b), "", "Replacement"
    ActivateSource "HAND_A"
    ReconAssert UBound(mEclData, 1) = UBound(b, 1), "Re-registering a lent table replaces it", n

    ActivateSource "HAND_C"
    On Error Resume Next
    ForgetActiveTable
    gotErr = Err.Number: Err.Clear
    On Error GoTo Failed
    ReconAssert gotErr = vbObjectError + 795 And Not SourceLoaded("HAND_C"), "Bypassing the hand-over fails loudly and drops the table", n

    ReturnActiveTable
    Set mSources = oldSources: Set mSourceDateCache = oldDates: Set mJoinCache = oldJoin
    mEclData = Empty: ForgetActiveTable
    If Len(oldSource) > 0 Then If SourceLoaded(oldSource) Then ActivateSource oldSource
    ReconTableHandoverTests = "PASS: " & n & " table hand-over checks."
    Exit Function
Failed:
    errorText = Err.description
    On Error Resume Next
    ReturnActiveTable
    Set mSources = oldSources: Set mSourceDateCache = oldDates: Set mJoinCache = oldJoin
    mLent = False: mEclData = Empty: Set mActiveRecord = Nothing
    If Len(oldSource) > 0 Then If SourceLoaded(oldSource) Then ActivateSource oldSource
    ReconTableHandoverTests = "FAIL after " & n & " checks: " & errorText
End Function

' ===================== isolation for the regression suites ==================
'
' The fixture suites register their own small tables under the real source
' keys, and some of them expect ONLY their tables to be loaded. Run after a real
' load, they saw the bank's files as well and failed, and they left their
' fixtures in place of the real inputs. IsolateSources sets the loaded inputs
' aside and starts an empty store; RestoreSources puts the inputs back.

Public Function IsolateSources() As Object
    Dim keep As Object
    ReturnActiveTable
    EnsureSourceStore
    Set keep = NewMap()
    Set keep("Sources") = mSources
    Set keep("Dates") = mSourceDateCache
    Set keep("Join") = mJoinCache
    Set keep("AllFields") = mEclAllFields
    Set keep("Greatest") = mGreatest
    keep("Active") = mActiveSource
    Set mSources = NewMap(): Set mSourceDateCache = Nothing: Set mJoinCache = Nothing
    Set mEclAllFields = Nothing: Set mGreatest = Nothing
    ResetWorkingTable
    Set IsolateSources = keep
End Function

Public Sub RestoreSources(ByVal keep As Object)
    If keep Is Nothing Then Exit Sub
    ReturnActiveTable
    Set mSources = keep("Sources"): Set mSourceDateCache = keep("Dates"): Set mJoinCache = keep("Join")
    Set mEclAllFields = keep("AllFields"): Set mGreatest = keep("Greatest")
    ResetWorkingTable
    If Len(CStr(keep("Active"))) > 0 Then
        If SourceLoaded(CStr(keep("Active"))) Then ActivateSource CStr(keep("Active"))
    End If
End Sub

' The working table and everything read from it, emptied. The file details go
' too: a 1904-dated fixture must not leave its date system behind for the next
' suite or for the real inputs.
Private Sub ResetWorkingTable()
    mEclData = Empty: ForgetActiveTable
    Set mEclHeaders = Nothing: Set mPreShockMeasureFields = Nothing: mActiveSource = ""
    mEclPath = "": mEclSheet = "": mEclHeaderRow = 0: mEclDate1904 = False
    Set mMaskStore = Nothing: mMaskBytes = 0
End Sub
