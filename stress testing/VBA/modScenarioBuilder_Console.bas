Attribute VB_Name = "modScenarioBuilder_Console"
Option Explicit

' ============================================================================
'  Opening the console for THIS tool.
'
'  Stress testing and ALM are different products for different modules. They
'  share the console's machinery - one reader, one renderer, one set of
'  interactions - but nothing of their identity: this one carries its own name,
'  its own accent, its own sheets and its own language.
'
'  The launcher is a declaration and nothing more. Adding a sheet to the tool
'  adds a view to the console with no change here beyond naming it.
' ============================================================================

Public Sub OpenStressConsole()
    OpenStressConsole_Paths
End Sub

' Same thing, returning where it wrote the files, so a harness can read them.
Public Function OpenStressConsole_Paths() As String
    Dim tiles As Collection, flow As Collection, acts As Collection, asOf As String

    On Error GoTo Failed
    EnsurePreShockSheet
    Set tiles = New Collection
    Set flow = New Collection
    Set acts = New Collection

    asOf = LatestDerivedDate()
    BuildStressTiles tiles
    BuildStressFlow flow

    acts.Add Act("Build manual reports", "BuildManualReconPack", True)
    acts.Add Act("Reconciliation", "OpenReconWorkbench", False)
    acts.Add Act("Review workbooks per test case", "GenerateReviewWorkbooks", False)

    ' Its own colour, so a glance tells you which of the two tools you are in.
    OpenStressConsole_Paths = OpenConsole("JKB Stress Testing", "V15", asOf, _
        Array("Overview/" & DERIVED_SHEET, _
              "Derivation/" & PS_METRICS_SHEET, _
              "Derivation/" & PS_FIELDS_SHEET, _
              "Derivation/" & PS_CASES_SHEET, _
              "Inputs/" & PS_SOURCES_SHEET, _
              "Inputs/" & SHEET_LOG), _
        tiles, flow, acts, "#193F89", "#0E2240")
    Exit Function
Failed:
    UiProblem "Review console", "The review console could not be opened.", Err.description
End Function

' What a reviewer opens this tool to know: how much was rebuilt, how much of it
' disagrees with the system, and whether the configuration is reading it.
Private Sub BuildStressTiles(ByVal tiles As Collection)
    Dim ws As Worksheet, r As Long, lastR As Long
    Dim results As Long, baseOk As Long, preOk As Long, diffs As Long, worst As Double
    Dim d As Double, derived As Long, usingDerived As Long

    ' One results table: Derived_Values (Reconciles in column S).
    modValueSources.VS_ResultCounts results, baseOk, preOk, diffs
    derived = PersistedDerivedCount()
    usingDerived = CountConfigUsingDerived()

    tiles.Add Tile("Comparisons", results, "count", "", "One per test case, element and metric.")
    tiles.Add Tile("Rebuilt independently", preOk, "count", IIf(preOk > 0, "Clean", "Check"), _
                   "Pre-shock figures derived from the uploaded extract rather than read from the system.")
    tiles.Add Tile("Disagree with the system", diffs, "count", IIf(diffs > 0, "Break", IIf(preOk > 0, "Clean", "Check")), _
                   "Rows whose Reconciles column on " & DERIVED_SHEET & " says Differs.")
    tiles.Add Tile("Config rows reading derived", usingDerived, "count", IIf(usingDerived > 0, "Clean", "Break"), _
                   derived & " value(s) stored, read as derived.<metric> instead of copying the system output.")
End Sub

' The chain this tool proves: the extract becomes a base, the filter makes it a
' pre-shock, and the configuration reads that rather than the system's own number.
Private Sub BuildStressFlow(ByVal flow As Collection)
    ' Monetary values and rates across overlapping test cases are not additive.
    ' The console workflow strip supplies navigation; field-level evidence is in
    ' Recon_Detail and in the input-based native pivots.
End Sub

Private Function CountConfigUsingDerived() As Long
    Dim lo As ListObject, h As Object, hdr As Variant, c As Long, a As Variant, r As Long
    On Error GoTo Done
    Set lo = ThisWorkbook.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set h = NewMap()
    hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    If Not h.Exists("MANUAL_FORMULA_DEFAULT") Then Exit Function
    a = lo.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        If InStr(1, SafeText(a(r, h("MANUAL_FORMULA_DEFAULT"))), "derived.", vbTextCompare) > 0 Then
            CountConfigUsingDerived = CountConfigUsingDerived + 1
        End If
    Next r
Done:
End Function

Private Function LatestDerivedDate() As String
    Dim ws As Worksheet, r As Long, v As String
    Set ws = DerivedSheet()
    If ws Is Nothing Then Exit Function
    For r = DV_FIRST_ROW To DV_FIRST_ROW + 5
        v = SafeText(ws.Cells(r, 1).text)
        If Len(v) > 0 Then LatestDerivedDate = v: Exit Function
    Next r
End Function

Private Function SafeNumeric(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNumeric(v) Then SafeNumeric = CDbl(v)
    Err.Clear
    On Error GoTo 0
End Function
