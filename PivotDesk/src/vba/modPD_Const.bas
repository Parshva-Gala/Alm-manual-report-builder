Option Explicit

' ============================================================================
'  Avati - a desk for daily analysis.
'
'  Three things, done from the Desk:
'
'    ADD FILES   point it at them; it works out what each one is
'    PIVOTS      pick the frameworks; each becomes a workbook of live pivots
'    RECONCILE   the outputs against control reports 3 and 6
'
'  The Desk is the whole interface. There used to be a separate console
'  window as well; everything it did now happens on the sheet, so there is
'  one place to look and nothing to open.
'
'  What makes it a desk rather than a report writer: everything it builds is a
'  real PivotTable over one cache, so you can drag a field, add a slicer, change
'  a filter and keep working. Nothing here is a picture of an analysis - it is
'  the analysis, and it is yours to move.
'
'  There is no rule engine, no report form, no cashflow model. Those were other
'  tools. This one stages the data and hands you the pivots.
'
'  EVERY declaration in this module sits above the first procedure. The build
'  linter enforces it, because a module-level declaration placed after one
'  compiles to "Variable not defined" as a modal dialog inside an invisible
'  Excel - a hang, not an error.
' ============================================================================

Public Const TOOL_NAME As String = "Avati"
Public Const TOOL_VERSION As String = "3.1"
Public Const BANK_NAME As String = "MIDBANK  Cairo"

' --- sheets in the desk itself ----------------------------------------------
Public Const SH_HOME As String = "Desk"
Public Const SH_SOURCES As String = "Files"
Public Const SH_RECON As String = "Reconciliation"
Public Const SH_LOG As String = "Activity"
Public Const SH_SETTINGS As String = "_Settings"
Public Const SH_CONFIG As String = "Pivot config"
Public Const SH_FIELDS As String = "Pivot fields"
Public Const SH_BOOKS As String = "Workbooks"
Public Const SH_CHARTS As String = "Chart config"
Public Const SH_GALLERY As String = "Gallery"

' --- sheets in a generated framework workbook -------------------------------
Public Const SH_STAGE As String = "_data"
' The pivot behind the Start here chart. Hidden, not very hidden: a PivotChart
' must be able to reach its pivot.
Public Const SH_CHART As String = "_chart"
Public Const SH_BALSHEET As String = "Balance sheet"
Public Const SH_GUIDE As String = "Start here"

' ============================================================================
'  ONE number format, everywhere.
'
'  Every figure this tool writes or pivots goes through this: thousands
'  separated, no decimals, negatives in brackets and red, an empty cell as a
'  dash. A tool where some sheets carry decimals and others do not is a tool
'  whose totals look like they disagree when they do not.
' ============================================================================
Public Const NUM_FMT As String = "#,##0;[Red](#,##0);-"
Public Const PCT_FMT As String = "0%"

' --- the fields read off an ALM output --------------------------------------
'
' Names, not codes. Where a field has both, only the name is staged: a pivot
' showing COA_MBGL.0020 beside "CASH" is a column of noise on every row, and the
' code is never the thing being read.
Public Const F_RULE_ORDER As String = "ALM_PORTFOLIO_SEGMENTATION_RULE_ORDER"
Public Const F_RULE_CAT As String = "ALM_PORTFOLIO_SEGMENTATION_RULE_CATEGORY"
Public Const F_RULE_NAME As String = "ALM_PORTFOLIO_SEGMENTATION_RULE_NAME"
Public Const F_FACTOR As String = "ALM_FACTOR_PCT"
Public Const F_TYPE As String = "COA_BALANCESHEET_CATEGORY"
Public Const F_LINE As String = "BALANCESHEET_LINE_NAME"
Public Const F_SUBLINE As String = "BALANCESHEET_SUB_LINE_NAME"
Public Const F_COA_NAME As String = "COA_NAME"
Public Const F_COA_CODE As String = "COA_CODE"
Public Const F_CURRENCY As String = "CURRENCY_NAME"
Public Const F_BUCKET As String = "BUCKET_DISPLAY_NAME"
Public Const F_FRAMEWORK As String = "ALM_FRAMEWORK_NAME"
Public Const F_ACCOUNT As String = "ACCOUNT_NUMBER"
Public Const F_AS_OF As String = "AS_OF_DATE"
Public Const F_MATURITY As String = "MATURITY_DATE"

' --- the amount fields ------------------------------------------------------
'
' LCY is in every extract seen. The native-currency pair may be spelled CCY or
' ACY or something else again - nobody is sure - so it is RESOLVED against the
' header rather than assumed, each spelling tried in turn, and whichever was
' found is written onto the Files sheet and the guide. If none is there the tool
' says so and uses LCY, rather than quietly reporting one thing as another.
Public Const F_PRE_LCY As String = "CASHFLOW_AMOUNT_LCY_PRE_FACTOR"
Public Const F_POST_LCY As String = "CASHFLOW_AMOUNT_LCY_POST_FACTOR"

' --- control reports --------------------------------------------------------
Public Const F_CTRL3_AMT As String = "GL_BALANCE_LCY"
Public Const F_CTRL6_AMT As String = "REPORTING_BALANCE_LCY_NET"

' --- frameworks -------------------------------------------------------------
Public Const FW_LCR As String = "LCR"
Public Const FW_NSFR As String = "NSFR"
Public Const FW_ML As String = "MATURITY_LADDER"

' --- staged column order ----------------------------------------------------
'
' Everything downstream names these columns through these constants, so adding
' a field is one entry here and one line in the stager.
Public Const C_RULE_ORDER As Long = 1
Public Const C_RULE_CAT As Long = 2
Public Const C_RULE_NAME As Long = 3
Public Const C_FACTOR As Long = 4
Public Const C_TYPE As Long = 5
Public Const C_LINE As Long = 6
Public Const C_SUBLINE As Long = 7
Public Const C_COA_NAME As Long = 8
Public Const C_CURRENCY As Long = 9
Public Const C_CCYCLASS As Long = 10
Public Const C_BUCKET As Long = 11
Public Const C_PRE As Long = 12
Public Const C_POST As Long = 13
Public Const C_COLS As Long = 13

' What those columns are called on the sheet and in every pivot field list.
' Plain words: the field list is read by a person, not by a schema.
Public Const H_RULE_ORDER As String = "Rule order"
Public Const H_RULE_CAT As String = "Rule category"
Public Const H_RULE_NAME As String = "Rule name"
Public Const H_FACTOR As String = "Factor"

' The caption a data field carries in the pivot. Excel REFUSES a data field
' whose caption equals an existing source column name, and the refusal surfaces
' as "cannot set Orientation" three calls later when the pivot turns out to have
' no data fields at all. So the staged columns are "... amount" and the
' captions are these.
Public Const CAP_PRE As String = "Pre-factor"
Public Const CAP_POST As String = "Post-factor"
Public Const H_TYPE As String = "Type"
Public Const H_LINE As String = "Line"
Public Const H_SUBLINE As String = "Subline"
Public Const H_COA_NAME As String = "COA name"
Public Const H_CURRENCY As String = "Currency"
Public Const H_CCYCLASS As String = "LCY / FCY"
Public Const H_BUCKET As String = "Bucket"
Public Const H_PRE As String = "Pre factor amount"
Public Const H_POST As String = "Post factor amount"

' --- limits -----------------------------------------------------------------
' One Range read per block rather than per row is the difference between a pass
' that takes a minute and one that takes an hour.
Public Const CHUNK_ROWS As Long = 40000

' A sheet per rule name, times the currency split. Past this the workbook stops
' being something Excel opens quickly, so the rest are left out and SAID to be
' left out rather than silently missing.
Public Const MAX_RULE_SHEETS As Long = 120

' However the Pivot config is set, a workbook stops here: past it Excel is
' slow to open the book and nobody scrolls the tabs.
Public Const MAX_BOOK_SHEETS As Long = 250

' An amount that has been through a database and a spreadsheet is not exactly
' equal to another; one unit is well below anything worth acting on.
Public Const TOLERANCE As Double = 1#

' --- verdicts ---------------------------------------------------------------
Public Const V_OK As String = "OK"
Public Const V_CHECK As String = "Check"
Public Const V_BREAK As String = "Break"
Public Const V_NONE As String = "Not supplied"

Public PD_Busy As Boolean
Public PD_Quiet As Boolean
Public PD_Trace As Object

' The clock behind the time-left estimate in Progress_: when the current
' phase started, and how far it had got at the last report.
Private mProgT0 As Single
Private mProgLast As Double
' When the last breadcrumb let Windows in. A build that does not for five
' seconds is shown as "Not Responding", and its status bar stops painting.
Private mYieldAt As Single

' ===================== procedures begin here ================================

Public Function Frameworks() As Variant
    Frameworks = Array(FW_LCR, FW_NSFR, FW_ML)
End Function

Public Function FwLabel(ByVal fw As String) As String
    Select Case UCase$(fw)
        Case UCase$(FW_ML): FwLabel = "Maturity Ladder"
        Case UCase$(FW_NSFR): FwLabel = "NSFR"
        Case UCase$(FW_LCR): FwLabel = "LCR"
        Case Else: FwLabel = fw
    End Select
End Function

' The ladder is split by actual currency, one worksheet each. LCR and NSFR are
' split local against foreign. That is the only structural difference between
' them, and it lives here rather than in four places.
Public Function SplitsByCurrency(ByVal fw As String) As Boolean
    SplitsByCurrency = (StrComp(fw, FW_ML, vbTextCompare) = 0)
End Function

Public Function StageHeadings() As Variant
    StageHeadings = Array( _
        H_RULE_ORDER, H_RULE_CAT, H_RULE_NAME, H_FACTOR, _
        H_TYPE, H_LINE, H_SUBLINE, H_COA_NAME, _
        H_CURRENCY, H_CCYCLASS, H_BUCKET, _
        H_PRE, H_POST)
End Function

Public Function PreNativeSpellings() As Variant
    PreNativeSpellings = Array( _
        "CASHFLOW_AMOUNT_CCY_PRE_FACTOR", "CASHFLOW_AMOUNT_ACY_PRE_FACTOR", _
        "CASHFLOW_AMOUNT_PRE_FACTOR_CCY", "CASHFLOW_AMOUNT_PRE_FACTOR_ACY", _
        "CASHFLOW_AMOUNT_CCY", "CASHFLOW_AMOUNT_ACY", _
        "CASHFLOW_AMT_CCY_PRE_FACTOR", "CASHFLOW_AMT_ACY_PRE_FACTOR")
End Function

Public Function PostNativeSpellings() As Variant
    PostNativeSpellings = Array( _
        "CASHFLOW_AMOUNT_CCY_POST_FACTOR", "CASHFLOW_AMOUNT_ACY_POST_FACTOR", _
        "CASHFLOW_AMOUNT_POST_FACTOR_CCY", "CASHFLOW_AMOUNT_POST_FACTOR_ACY", _
        "CASHFLOW_AMT_CCY_POST_FACTOR", "CASHFLOW_AMT_ACY_POST_FACTOR")
End Function

' Information goes to the Desk as a toast when the Desk is in front - a message
' box for "3 files placed" is a click the reader should not owe. Warnings,
' questions and anything said from another sheet still get a box, because
' there it is the only thing guaranteed to be seen.
Public Sub Tell(ByVal msg As String, ByVal style As Long)
    If PD_Quiet Then Exit Sub
    If (style And vbExclamation) = 0 And (style And vbCritical) = 0 And (style And vbQuestion) = 0 Then
        If modPD_Desk.DeskInFront() Then
            modPD_Desk.Toast msg, "OK"
            Exit Sub
        End If
    End If
    MsgBox msg, style, TOOL_NAME
End Sub

' A result with a verdict: on the Desk it is a toast coloured by the level,
' anywhere else a message box with the matching icon.
Public Sub Notify(ByVal msg As String, ByVal level As String)
    If PD_Quiet Then Exit Sub
    If modPD_Desk.DeskInFront() Then
        modPD_Desk.Toast msg, level
    ElseIf UCase$(level) = UCase$(V_OK) Then
        MsgBox msg, vbInformation, TOOL_NAME
    Else
        MsgBox msg, vbExclamation, TOOL_NAME
    End If
End Sub

' A breadcrumb the headless harness writes and the desk ignores. Every long
' stage calls it, so a run that dies without reaching its own error handler
' still says on disk which stage it was in.
Public Sub Step_(ByVal what As String)
    On Error Resume Next
    Application.StatusBar = TOOL_NAME & "   " & ChrW(183) & "   " & what
    If Not PD_Trace Is Nothing Then PD_Trace.Note what
    LetWindowsIn
    Err.Clear
End Sub

' DoEvents, a few times a second at most: the status bar repaints and Windows
' sees Excel alive. The Desk's buttons are held off meanwhile (PD_Busy), and
' the Working veil covers the Desk.
Private Sub LetWindowsIn()
    If Timer < mYieldAt Then mYieldAt = 0            ' past midnight
    If Timer - mYieldAt < 0.25 Then Exit Sub
    mYieldAt = Timer
    DoEvents
End Sub

' The same breadcrumb with a measured bar in front of it, for the stages that
' know how far through they are:
'
'    Avati   [filled x6][empty x14]  30%   LCR - staged 120,000 of 400,000 rows
Public Sub Progress_(ByVal what As String, ByVal frac As Double)
    Dim n As Long, i As Long, bar As String
    On Error Resume Next
    If frac < 0 Then frac = 0
    If frac > 1 Then frac = 1
    n = CLng(frac * 20)
    ' Built a character at a time: String$ is not to be trusted with a
    ' character outside the ANSI code page.
    For i = 1 To 20
        If i <= n Then bar = bar & ChrW(9632) Else bar = bar & ChrW(9633)
    Next i
    Application.StatusBar = TOOL_NAME & "   " & bar & "  " & Format$(frac, "0%") & "   " & what & TimeLeft(frac)
    If Not PD_Trace Is Nothing Then PD_Trace.Note what
    LetWindowsIn
    Err.Clear
End Sub

' "  -  about 40s left", once a phase has run long enough to say. Progress
' going backwards means a new phase (the next framework, the next pass), and
' the clock starts again; so does midnight, where Timer wraps to zero.
Private Function TimeLeft(ByVal frac As Double) As String
    Dim now_ As Single, took As Double, left_ As Double
    now_ = Timer
    If mProgT0 = 0 Or frac < mProgLast Or now_ < mProgT0 Then mProgT0 = now_
    mProgLast = frac
    took = now_ - mProgT0
    If frac < 0.08 Or frac >= 0.995 Or took < 3 Then Exit Function
    left_ = took * (1 - frac) / frac
    If left_ < 60 Then
        TimeLeft = "   " & ChrW(183) & "   about " & (Int(left_ / 5) + 1) * 5 & "s left"
    Else
        TimeLeft = "   " & ChrW(183) & "   about " & Int(left_ / 60 + 0.5) & " min left"
    End If
End Function
