Attribute VB_Name = "modScenarioBuilder_Joined"
Option Explicit

' ============================================================================
'  Joined input and native reconciliation pivots.
'
'  The table contains one row per original source record. ECL, CAPRWA, LL,
'  LCR and capital records retain their own amounts and original field names.
'  ECL enriches the other account records with dimensions only; its monetary
'  values are never repeated across cashflow or RWA records.
'
'  Enrichment uses the common account, reporting date, bank/entity and branch
'  fields. Conflicting ECL account dimensions remain explicit review items.
'  ROW_SOURCE and SOURCE_ROW_NUMBER preserve the owning source and lineage.
'
'  Base pivots are provided for each source. Testcase membership comes from
'  the effective configured filter, reporting-date policy and entity scope.
'  Membership columns distinguish eligible, excluded and untested records.
'  Case_Coverage links each definition to its native pivot or coverage notice;
'  Source_Controls compares original input totals with the joined records.
'  Native pivots share one compact, row-preserving reporting projection.
'  Joined_Data retains every original field; source and row identify details.
' ============================================================================

' The full source row count is checked before writing. Exceeding this Excel
' capacity guard raises an error; a truncated joined workbook is not produced.
Private Const MAX_JOINED_ROWS As Long = 1040000
Private Const CHUNK As Long = 20000
Private Const ROW_SOURCE_COL As String = "ROW_SOURCE"
Private Const JOINED_SHEET As String = "Joined_Data"
' Test switch. When True, each native pivot is updated as soon as it is made,
' instead of once its filters are set. Only ReconJoinedPivotBatchRegressionTests
' sets it, to build its own unbatched baseline; the review pack never does.
Private mPivotsImmediate As Boolean

' Which sources contribute, and in which order their columns are laid out.
Private Function JoinOrder() As Variant
    JoinOrder = Array("LL", "LCR", "ECL", "CAPRWA", "CAP")
End Function

' ===================== the entry point ======================================

Public Sub BuildJoinedInput()
    Dim folder As String, path As String
    If Not AnyJoinSourceLoaded() Then
        UiNotice "Joined input", "No input file is loaded yet.", , "Load the input files (step 1 on the home screen), then build the joined input."
        Exit Sub
    End If
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Where should the joined input file be written?"
        If Not .Show = -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    path = BuildJoinedInputTo(folder, True)
    If Len(path) > 0 Then
        UiNotice "Joined input", "The joined input file is ready.", _
                 path & vbCrLf & vbCrLf & "Each source record appears once, with the ECL account dimensions and the amounts from its own source. " & _
                 "Source_Controls and Case_Coverage show differences and records to review.", _
                 "Open it and start with the base pivots, then the test case pivots."
    End If
End Sub

' The headless path: same work, a folder passed in, no dialog.
Public Function BuildJoinedInputTo(ByVal folder As String, Optional ByVal announce As Boolean = False) As String
    Dim wb As Workbook, ws As Worksheet, path As String, checkpoint As String
    Dim heads As Object, layout As Object, nRows As Long, capped As Boolean
    Dim cases As Collection, st As Object, failureText As String

    Set st = CaptureState(): JKB_Busy = True
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set layout = BuildColumnLayout(heads)
    Set wb = Workbooks.Add
    Do While wb.Worksheets.count > 1
        wb.Worksheets(wb.Worksheets.count).Delete
    Loop
    Set ws = wb.Worksheets(1)
    ws.name = JOINED_SHEET

    ProgressDetail "Joined input: preserving source records and enriching ECL dimensions", 35
    nRows = WriteJoinedRows(ws, layout, heads, capped)
    ProgressDetail "Joined input: applying testcase filters to " & format$(nRows, "#,##0") & " records", 55
    Set cases = MarkTestCases(ws, layout, heads, nRows)
    FormatJoinedSheet ws, heads.count + cases.count, nRows
    ProgressDetail "Joined input: reconciling original and joined source totals", 70
    WriteSourceControls wb, ws, heads, nRows
    WriteCaseCoverage wb, cases
    checkpoint = JoinPath(folder, "_Joined_Input_pending.xlsx")
    ProgressDetail "Joined input: preserving the completed source and filter evidence", 71
    wb.SaveAs fileName:=checkpoint, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ProgressDetail "Joined input: creating native base and testcase pivots", 72
    BuildPivots wb, ws, heads, nRows, cases
    modReconWorkbench.EnsureJoinedOverview wb, nRows

    path = JoinPath(folder, "Joined_Input.xlsx")
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    Err.Clear
    On Error GoTo Failed
    ProgressDetail "Joined input: saving the account evidence and native pivots", 79
    UiFinishReportBook wb, "Joined input - account evidence and pivots", "Input files joined at account level, with base and test case pivots"
    wb.SaveAs fileName:=path, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    wb.Close SaveChanges:=False
    On Error Resume Next
    If Len(checkpoint) > 0 Then If Len(Dir$(checkpoint)) > 0 Then Kill checkpoint
    On Error GoTo Failed

    RestoreState st: JKB_Busy = False
    Application.ScreenUpdating = True
    LogJoined nRows, heads.count, cases.count, capped, path
    BuildJoinedInputTo = path
    Exit Function

Failed:
    failureText = Err.description
    If Len(checkpoint) > 0 Then
        If Len(Dir$(checkpoint)) > 0 Then failureText = failureText & " Completed input evidence was retained in " & checkpoint
    End If
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    RestoreState st: JKB_Busy = False
    Application.ScreenUpdating = True
    If announce Then UiProblem "Joined input", "The joined input file could not be built.", failureText
    On Error GoTo 0
    Err.Raise vbObjectError + 781, "BuildJoinedInputTo", failureText
End Function

Private Sub LogJoined(ByVal nRows As Long, ByVal nCols As Long, ByVal nCases As Long, _
                      ByVal capped As Boolean, ByVal path As String)
    On Error Resume Next
    LogIssue IIf(capped, LOG_LEVEL_WARN, LOG_LEVEL_INFO), "BuildJoinedInput", _
             format$(nRows, "#,##0") & " row(s), " & nCols & " column(s), " & _
             nCases & " test case column(s)" & IIf(capped, " - ROW CAP REACHED, the table is truncated", ""), path
    Err.Clear
End Sub

Private Function AnyJoinSourceLoaded() As Boolean
    Dim k As Variant
    For Each k In JoinOrder()
        If modScenarioBuilder_PreShock.SourceLoaded(CStr(k)) Then AnyJoinSourceLoaded = True: Exit Function
    Next k
End Function

' ===================== the column layout ====================================

' One column per distinct physical field name across the loaded extracts.
' Shared names occupy the same column, while ROW_SOURCE identifies ownership.
' Canonical lookup aliases do not create duplicate physical data columns.
Private Function BuildColumnLayout(ByRef headsOut As Object) As Object
    Dim layout As Object, heads As Object, k As Variant, h As Object, f As Variant, nm As String, data As Variant

    Set layout = NewMap()       ' source -> (field name -> column in the output)
    Set heads = NewMap()        ' field name -> column in the output

    heads(ROW_SOURCE_COL) = 1
    For Each k In Array("SOURCE_ROW_NUMBER", "ROW_COUNT", "JOIN_STATUS", "ECL_MATCH_COUNT", "JOIN_KEY_FIELDS", "RECON_STAGE", "BASE_ELIGIBLE", "RECON_BANK")
        heads(CStr(k)) = heads.count + 1
    Next k

    For Each k In JoinOrder()
        If modScenarioBuilder_PreShock.SourceLoaded(CStr(k)) Then
            Set h = modScenarioBuilder_PreShock.SourceHeadersFor(CStr(k))
            Set layout(CStr(k)) = NewMap()
            If Not h Is Nothing Then
                For Each f In h.keys
                    nm = CStr(f)
                    ' The header map holds each column under BOTH its raw name
                    ' and its canonical one. Only the raw spelling becomes an
                    ' output column - the canonical alias would produce a second
                    ' column holding the same numbers under a name the source
                    ' never used, which is exactly the renaming this must avoid.
                    If NormalHeader(modScenarioBuilder_PreShock.SourceCellFor(CStr(k), 1, CLng(h(nm)))) = nm Then
                        If Not heads.Exists(nm) Then heads(nm) = heads.count + 1
                        layout(CStr(k))(nm) = CLng(h(nm))
                    End If
                Next f
            End If
        End If
    Next k

    Set headsOut = heads
    Set BuildColumnLayout = layout
End Function

' A raw header is one that is not merely an alias of another entry pointing at
' the same column under a tidier name.
Private Function IsRawHeader(ByVal h As Object, ByVal nm As String) As Boolean
    Dim canon As String
    canon = modScenarioBuilder_PreShock.CanonicalEclHeader(nm)
    ' If the canonical form of this name is a DIFFERENT key that maps to the
    ' same column, this name is the raw one and the other is the alias.
    If canon = nm Then
        ' It is its own canonical form. Keep it unless a different raw spelling
        ' already claims the column.
        IsRawHeader = True
        Exit Function
    End If
    IsRawHeader = True
End Function

' ===================== writing the rows =====================================

Private Function WriteJoinedRows(ByVal ws As Worksheet, ByVal layout As Object, ByVal heads As Object, _
                                 ByRef capped As Boolean) As Long
    Dim buf() As Variant, nCols As Long, used As Long, outRow As Long, totalRows As Double
    Dim k As Variant, eclData As Variant, eclH As Object, dimensions As Object
    nCols = heads.count: WriteHeaderRow ws, heads
    For Each k In JoinOrder()
        totalRows = totalRows + modScenarioBuilder_PreShock.SourceRowCount(CStr(k))
    Next k
    If totalRows > MAX_JOINED_ROWS Then Err.Raise vbObjectError + 780, , "Joined source rows exceed Excel's row limit. No truncated report was written."
    ReDim buf(1 To CHUNK, 1 To nCols): outRow = 2
    Set dimensions = NewMap(): Set eclH = NewMap()
    If modScenarioBuilder_PreShock.SourceLoaded("ECL") Then
        eclData = modScenarioBuilder_PreShock.SourceDataFor("ECL")
        Set eclH = modScenarioBuilder_PreShock.SourceHeadersFor("ECL")
        Set dimensions = EclDimensions(layout("ECL"))
    End If
    For Each k In JoinOrder()
        ProgressDetail "Joined input: preserving " & CStr(k) & " source records"
        EmitFactSource ws, CStr(k), layout, heads, buf, used, outRow, eclData, eclH, dimensions, CLng(totalRows)
    Next k
    FlushBuf ws, buf, used, outRow, nCols
    WriteJoinedRows = outRow - 2
End Function

Private Function EclDimensions(ByVal sourceMap As Object) As Object
    Dim d As Object, k As Variant
    Set d = NewMap()
    For Each k In sourceMap.keys
        If Not IsMeasureField(CStr(k)) Then d(CStr(k)) = CLng(sourceMap(k))
    Next k
    Set EclDimensions = d
End Function

Private Function IsMeasureField(ByVal field As String) As Boolean
    Dim s As String, k As Variant
    s = SafeUpperText(field)
    For Each k In Array("_CODE", "_NAME", "_ID", "_NUMBER", "_FLAG", "_TYPE", "_DATE", "_DESCRIPTION", "_DESC")
        If Right$(s, Len(CStr(k))) = CStr(k) Then Exit Function
    Next k
    If Left$(s, 4) = "FLG_" Or Left$(s, 5) = "FLAG_" Then Exit Function
    If s = "ECL" Or s = "IIS" Or s = "EAD" Or s = "PD" Or s = "LGD" Or s = "RWA" Or s = "CCF" Then IsMeasureField = True: Exit Function
    ' Actual extract measures whose names carry no AMOUNT/LCY marker.
    ' Keep financial values on their owning source even when two ECL records
    ' have the same account dimensions but different balances or risk inputs.
    For Each k In Array("GUARANTEED_LOANS", "GUARANTEED_LOANS_POST_HC", "ACCRUED_INTEREST", _
                        "TOTAL_INTEREST_ADVANCE", "PREV_TOTAL_INTEREST_ADVANCE", "SEC_PURCHASE_COST", _
                        "FAIR_VALUE_FOR_CURRENT_PERIOD", "GROSS_CARRYING_VALUE", "SEC_NUM_UNITS")
        If s = CStr(k) Then IsMeasureField = True: Exit Function
    Next k
    If Left$(s, 3) = "HC_" Or Left$(s, 4) = "PCT_" Then IsMeasureField = True: Exit Function
    For Each k In Array("COLLATERAL", "_EAD", "_LGD", "_PD")
        If InStr(1, s, CStr(k), vbTextCompare) > 0 Then IsMeasureField = True: Exit Function
    Next k
    For Each k In Array("AMOUNT", "BALANCE", "_LCY", "_FCY", "OUTST", "EXPOSURE", "ECL", "RWA", "RATIO", "_RATE", "PERCENT", "_PCT", "FACTOR", "LOSS", "PROVISION", "COLLATERAL_VALUE")
        If InStr(1, s, CStr(k), vbTextCompare) > 0 Then IsMeasureField = True: Exit Function
    Next k
End Function

Private Function SharedJoinFields(ByVal target As Object, ByVal ecl As Object) As Collection
    Dim fields As New Collection, k As Variant
    If Not target.Exists("ACCOUNT_NUMBER") Or Not ecl.Exists("ACCOUNT_NUMBER") Then Set SharedJoinFields = fields: Exit Function
    fields.Add "ACCOUNT_NUMBER"
    For Each k In Array("AS_OF_DATE", "ENTITY_ID", "ENTITY_CODE", "BANK_CODE", "BRANCH_CODE")
        If target.Exists(CStr(k)) And ecl.Exists(CStr(k)) Then fields.Add CStr(k)
    Next k
    Set SharedJoinFields = fields
End Function

Private Function ScopedAccountKey(ByRef data As Variant, ByVal r As Long, ByVal h As Object, ByVal fields As Collection) As String
    Dim f As Variant, v As String
    If fields.count = 0 Then Exit Function
    For Each f In fields
        v = SafeUpperText(data(r, CLng(h(CStr(f)))))
        If Len(v) = 0 Or v = "NULL" Then ScopedAccountKey = "": Exit Function
        If CStr(f) = "AS_OF_DATE" Then
            On Error GoTo BadKey
            v = DateKey(data(r, CLng(h(CStr(f)))))
            On Error GoTo 0
        End If
        ScopedAccountKey = ScopedAccountKey & KeyPart(v)
    Next f
    Exit Function
BadKey:
    ScopedAccountKey = ""
End Function

Private Function DimensionRowsAgree(ByRef data As Variant, ByVal firstRow As Long, ByVal otherRow As Long, ByVal dimensions As Object) As Boolean
    Dim k As Variant, c As Long
    For Each k In dimensions.keys
        c = CLng(dimensions(k))
        If SafeUpperText(data(firstRow, c)) <> SafeUpperText(data(otherRow, c)) Then Exit Function
    Next k
    DimensionRowsAgree = True
End Function

Private Function ScopedEclIndex(ByRef data As Variant, ByVal eclH As Object, ByVal fields As Collection, ByVal dimensions As Object) As Object
    Dim d As Object, rec As Object, r As Long, key As String
    Set d = NewMap(): Set ScopedEclIndex = d
    If Not IsArray(data) Or fields.count = 0 Then Exit Function
    For r = 2 To UBound(data, 1)
        key = ScopedAccountKey(data, r, eclH, fields)
        If Len(key) > 0 Then
            If Not d.Exists(key) Then
                Set rec = NewMap(): rec("Row") = r: rec("Count") = 1: rec("Conflict") = False
                Set d(key) = rec
            Else
                Set rec = d(key): rec("Count") = CLng(rec("Count")) + 1
                If Not CBool(rec("Conflict")) Then rec("Conflict") = Not DimensionRowsAgree(data, CLng(rec("Row")), r, dimensions)
            End If
        End If
    Next r
End Function

Private Sub EmitFactSource(ByVal ws As Worksheet, ByVal sourceKey As String, ByVal layout As Object, ByVal heads As Object, _
                           ByRef buf() As Variant, ByRef used As Long, ByRef outRow As Long, _
                           ByRef eclData As Variant, ByVal eclH As Object, ByVal dimensions As Object, ByVal totalRows As Long)
    Dim data As Variant, h As Object, sourceMap As Object, fields As Collection, index As Object, rec As Object
    Dim sc() As Long, oc() As Long, nMap As Long, r As Long, k As Variant, key As String, status As String
    Dim c As Long, o As Long, v As Variant, joinFields As String, nCols As Long, nMatch As Long
    If Not modScenarioBuilder_PreShock.SourceLoaded(sourceKey) Then Exit Sub
    data = modScenarioBuilder_PreShock.SourceDataFor(sourceKey)
    Set h = modScenarioBuilder_PreShock.SourceHeadersFor(sourceKey): Set sourceMap = layout(sourceKey)
    PlanCopy sourceMap, heads, sc, oc, nMap
    Set fields = SharedJoinFields(h, eclH)
    Set index = ScopedEclIndex(eclData, eclH, fields, dimensions)
    For Each k In fields
        If Len(joinFields) > 0 Then joinFields = joinFields & " + "
        joinFields = joinFields & CStr(k)
    Next k
    nCols = heads.count
    For r = 2 To UBound(data, 1)
        used = used + 1: ClearRow buf, used, nCols
        CopyRowFast buf, used, data, r, sc, oc, nMap, LBound(data, 2), UBound(data, 2)
        buf(used, heads(ROW_SOURCE_COL)) = sourceKey
        buf(used, heads("SOURCE_ROW_NUMBER")) = r
        buf(used, heads("ROW_COUNT")) = 1
        If h.Exists("BANK_CODE") Then
            buf(used, heads("RECON_BANK")) = SafeUpperText(data(r, CLng(h("BANK_CODE"))))
        ElseIf h.Exists("BANK_ID") Then
            If SafeText(data(r, CLng(h("BANK_ID")))) = "101" Then buf(used, heads("RECON_BANK")) = "JKB"
        End If
        buf(used, heads("JOIN_KEY_FIELDS")) = joinFields
        nMatch = 0
        If sourceKey = "ECL" Then
            status = "ECL source row"
        ElseIf sourceKey = "CAP" Then
            status = "Base-level capital row"
        Else
            key = ScopedAccountKey(data, r, h, fields)
            status = "No matching ECL account"
            If fields.count = 0 Then status = "Account key unavailable"
            If fields.count > 0 And Len(key) = 0 Then status = "Missing or invalid join key"
            If Len(key) > 0 And index.Exists(key) Then
                Set rec = index(key): nMatch = CLng(rec("Count"))
                If CBool(rec("Conflict")) Then
                    status = "Conflicting ECL dimensions - review"
                Else
                    status = "Matched ECL dimensions"
                    If nMatch > 1 Then status = "Matched identical ECL dimensions"
                    For Each k In dimensions.keys
                        c = CLng(dimensions(k)): o = CLng(heads(CStr(k)))
                        v = eclData(CLng(rec("Row")), c)
                        If Len(SafeText(buf(used, o))) = 0 Or SafeUpperText(buf(used, o)) = "NULL" Then
                            buf(used, o) = v
                        ElseIf Len(SafeText(v)) > 0 And SafeUpperText(v) <> "NULL" Then
                            If SafeUpperText(buf(used, o)) <> SafeUpperText(v) Then status = "Source and ECL dimension conflict - review"
                        End If
                    Next k
                    If Not h.Exists("AS_OF_DATE") Or Not eclH.Exists("AS_OF_DATE") Then status = status & "; date key unavailable"
                End If
            End If
        End If
        buf(used, heads("JOIN_STATUS")) = status
        buf(used, heads("ECL_MATCH_COUNT")) = nMatch
        If h.Exists("STAGE") Then
            o = CLng(heads(NormalHeader(data(1, CLng(h("STAGE"))))))
            buf(used, heads("RECON_STAGE")) = StageKey(buf(used, o))
        Else
            buf(used, heads("RECON_STAGE")) = "Unstaged"
        End If
        outRow = outRow + 1
        If used >= CHUNK Then
            ProgressDetail "Joined input: preserving " & sourceKey & " records " & format$(r - 1, "#,##0") & " of " & format$(UBound(data, 1) - 1, "#,##0"), 35# + 20# * (outRow - 2) / totalRows
            FlushBuf ws, buf, used, outRow, nCols
        End If
    Next r
End Sub

Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal heads As Object)
    Dim f As Variant, a() As Variant, n As Long
    n = heads.count
    ReDim a(1 To 1, 1 To n)
    For Each f In heads.keys
        a(1, CLng(heads(f))) = CStr(f)
    Next f
    ws.Range(ws.Cells(1, 1), ws.Cells(1, n)).Value2 = a
End Sub

' Copies the owning source record without overwriting a populated target cell.
' Source and target column arrays are prepared once per source to avoid
' rebuilding a dictionary key array for every record.
Private Sub CopyRowFast(ByRef buf() As Variant, ByVal used As Long, ByRef data As Variant, ByVal r As Long, _
                        ByRef srcCol() As Long, ByRef outCol() As Long, ByVal n As Long, _
                        ByVal loC As Long, ByVal hiC As Long)
    Dim i As Long, c As Long, o As Long, v As Variant
    For i = 1 To n
        o = outCol(i)
        If IsEmpty(buf(used, o)) Then
            c = srcCol(i)
            If c >= loC And c <= hiC Then
                v = data(r, c)
                If Not IsEmpty(v) Then buf(used, o) = v
            End If
        End If
    Next i
End Sub

' The two arrays above, built once per source.
Private Sub PlanCopy(ByVal map As Object, ByVal heads As Object, _
                     ByRef srcCol() As Long, ByRef outCol() As Long, ByRef n As Long)
    Dim f As Variant
    n = map.count
    If n = 0 Then
        ReDim srcCol(1 To 1): ReDim outCol(1 To 1)
        Exit Sub
    End If
    ReDim srcCol(1 To n)
    ReDim outCol(1 To n)
    n = 0
    For Each f In map.keys
        n = n + 1
        srcCol(n) = CLng(map(f))
        outCol(n) = CLng(heads(CStr(f)))
    Next f
End Sub

Private Sub ClearRow(ByRef buf() As Variant, ByVal used As Long, ByVal nCols As Long)
    Dim c As Long
    For c = 1 To nCols
        buf(used, c) = Empty
    Next c
End Sub

Private Sub FlushBuf(ByVal ws As Worksheet, ByRef buf() As Variant, ByRef used As Long, _
                     ByVal outRow As Long, ByVal nCols As Long)
    Dim first As Long
    If used = 0 Then Exit Sub
    first = outRow - used
    ws.Range(ws.Cells(first, 1), ws.Cells(first + used - 1, nCols)).Value2 = buf
    used = 0
End Sub

' ===================== one column per test case =============================

' Evaluates effective testcase filters using the production filter compiler.
' Each membership column retains excluded, date-mismatch and review states;
' only Yes records enter the testcase's native pivot.
Private Function MarkTestCases(ByVal ws As Worksheet, ByVal layout As Object, ByVal heads As Object, ByVal nRows As Long) As Collection
    Dim Done As New Collection, cases As Object, contexts As Object, k As Variant, spec As Object
    Dim data As Variant, hd As Object, node As Object, baseNode As Object, nodes As Object, failures As Object
    Dim cols As Object, hits As Object, reviews As Object, order As Collection
    Dim col As Long, r As Long, colBuf() As Variant, nm As String, unsupported As String, baseFailure As String
    Dim firstRow As Long, lastRow As Long, n As Long, i As Long, cn As String, dateValue As String, status As String
    Dim ignoreDates As Boolean, rowDates() As String
    Dim baseMask As Variant, caseMask As Variant
    Set MarkTestCases = Done: If nRows = 0 Then Exit Function
    ignoreDates = modScenarioBuilder_PreShock.IgnoreAsOfDate()
    Set cases = EnabledCases(): Set contexts = CaseContexts()
    Set hd = NewMap()
    For Each k In heads.keys
        hd(CStr(k)) = CLng(heads(k))
        nm = modScenarioBuilder_PreShock.CanonicalEclHeader(CStr(k))
        If Len(nm) > 0 Then If Not hd.Exists(nm) Then hd(nm) = CLng(heads(k))
    Next k
    Set nodes = NewMap(): Set cols = NewMap(): Set hits = NewMap(): Set reviews = NewMap(): Set order = New Collection: Set failures = NewMap()
    col = heads.count
    data = ws.Range(ws.Cells(1, 1), ws.Cells(2, heads.count)).Value2
    modScenarioBuilder_PreShock.RegisterTable "JOINED", data, hd, "(joined)", JOINED_SHEET
    modScenarioBuilder_PreShock.ActivateFor "JOINED"
    On Error Resume Next
    Set baseNode = modScenarioBuilder_PreShock.CompileForActive(modScenarioBuilder_Multi.ApplyEntityScope("RECON_BANK IN ('JKB')"), baseFailure)
    If Err.Number <> 0 Then baseFailure = Err.description: Err.Clear
    On Error GoTo 0
    For Each k In cases.keys
        Set spec = cases(k): unsupported = "": Set node = Nothing
        spec("Date") = ""
        If contexts.Exists(CStr(k)) Then spec("Date") = CStr(contexts(k))
        On Error Resume Next
        Set node = modScenarioBuilder_PreShock.CompileForActive(CStr(spec("Filter")), unsupported)
        If Err.Number <> 0 Then unsupported = Err.description: Err.Clear
        On Error GoTo 0
        col = col + 1: cn = CStr(spec("Column"))
        Set nodes(cn) = node: cols(cn) = col: hits(cn) = 0#: reviews(cn) = 0#: failures(cn) = unsupported
        order.Add spec: ws.Cells(1, col).Value2 = cn
    Next k
    firstRow = 2
    Do While firstRow <= nRows + 1
        ProgressDetail "Joined input: testing records " & format$(firstRow - 1, "#,##0") & " to " & format$(WorksheetFunction.Min(firstRow + CHUNK - 2, nRows), "#,##0") & " of " & format$(nRows, "#,##0"), 55# + 15# * (firstRow - 2) / nRows
        lastRow = WorksheetFunction.Min(firstRow + CHUNK - 1, nRows + 1): n = lastRow - firstRow + 1
        data = ChunkWithHeader(ws, heads.count, firstRow, lastRow)
        modScenarioBuilder_PreShock.RegisterTable "JOINED", data, hd, "(joined)", JOINED_SHEET
        modScenarioBuilder_PreShock.ActivateFor "JOINED"
        ReDim rowDates(2 To n + 1)
        If Not ignoreDates And hd.Exists("AS_OF_DATE") Then
            For r = 2 To n + 1
                On Error Resume Next
                rowDates(r) = DateKey(data(r, CLng(hd("AS_OF_DATE"))))
                If Err.Number <> 0 Then Err.Clear: rowDates(r) = "INVALID"
                On Error GoTo 0
            Next r
        End If
        ReDim colBuf(1 To n, 1 To 1)
        ' Each condition is decided for the whole chunk at once by the engine.
        baseMask = Empty
        If Len(baseFailure) = 0 And Not baseNode Is Nothing Then baseMask = modScenarioBuilder_PreShock.FilterMaskForActive(baseNode)
        For r = 2 To n + 1
            colBuf(r - 1, 1) = "Not tested"
            If Len(baseFailure) = 0 And Not baseNode Is Nothing Then
                If RowPasses(baseMask, baseNode, r) Then colBuf(r - 1, 1) = "Yes" Else colBuf(r - 1, 1) = "No"
            End If
        Next r
        ws.Cells(firstRow, CLng(heads("BASE_ELIGIBLE"))).Resize(n, 1).Value2 = colBuf
        For i = 1 To order.count
            Set spec = order(i): cn = CStr(spec("Column")): Set node = nodes(cn)
            caseMask = Empty
            If Len(CStr(failures(cn))) = 0 And Not node Is Nothing Then caseMask = modScenarioBuilder_PreShock.FilterMaskForActive(node)
            ReDim colBuf(1 To n, 1 To 1)
            For r = 2 To n + 1
                status = "No"
                If SafeText(data(r, heads(ROW_SOURCE_COL))) = "CAP" Then
                    ' Capital components support base snapshots, never account-filtered pre-shock.
                    status = "No"
                ElseIf Len(CStr(failures(cn))) > 0 Or node Is Nothing Then
                    status = "Not tested"
                ElseIf InStr(1, SafeText(data(r, heads("JOIN_STATUS"))), "conflict", vbTextCompare) > 0 Then
                    status = "Review"
                ElseIf Not ignoreDates Then
                    dateValue = rowDates(r)
                    If Len(dateValue) = 0 Or dateValue = "INVALID" Or Len(CStr(spec("Date"))) = 0 Or CStr(spec("Date")) = "AMBIGUOUS" Then
                        status = "Not tested"
                    ElseIf dateValue <> CStr(spec("Date")) Then
                        status = "Date mismatch"
                    ElseIf RowPasses(caseMask, node, r) Then
                        status = "Yes"
                    End If
                ElseIf RowPasses(caseMask, node, r) Then
                    status = "Yes"
                End If
                colBuf(r - 1, 1) = status
                If status = "Yes" Then hits(cn) = CDbl(hits(cn)) + 1
                If status <> "Yes" And status <> "No" Then reviews(cn) = CDbl(reviews(cn)) + 1
            Next r
            ws.Cells(firstRow, CLng(cols(cn))).Resize(n, 1).Value2 = colBuf
        Next i
        firstRow = lastRow + 1
    Loop
    For i = 1 To order.count
        Set spec = order(i): cn = CStr(spec("Column"))
        status = CStr(failures(cn))
        If Len(status) = 0 Then
            status = "Calculated"
            If CDbl(reviews(cn)) > 0 Then status = "Review source dates, join exceptions and mapping"
            If CDbl(hits(cn)) = 0 Then status = "No eligible rows - review coverage"
        End If
        Done.Add Array(cn, CStr(spec("TestCase")), CStr(spec("Element")), CDbl(hits(cn)), CLng(cols(cn)), _
                       CDbl(reviews(cn)), CStr(spec("Filter")), CStr(spec("Date")), status, "Case_" & format$(i, "000"))
    Next i
End Function

Private Function CaseContexts() As Object
    Dim d As Object, a As Variant, h As Object, hr As Long, stats As Object, idx As Object
    Dim tk As Variant, ek As Variant, tc As Object, el As Object, key As String
    Set d = NewMap(): Set CaseContexts = d
    ReadSource ThisWorkbook.Worksheets(SHEET_SOURCE), a, h, hr
    Set idx = IndexSource(a, h, hr, ThisWorkbook.date1904, stats)
    For Each tk In idx.keys
        Set tc = idx(tk)
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek)
            key = SafeUpperText(tc("TestCaseCode")) & "|" & SafeUpperText(el("ScenarioCode"))
            If d.Exists(key) Then
                If CStr(d(key)) <> CStr(tc("Date")) Then d(key) = "AMBIGUOUS"
            Else
                d(key) = CStr(tc("Date"))
            End If
        Next ek
    Next tk
End Function

' Header row on top, then the chunk's own rows - the shape the evaluator expects.
Private Function ChunkWithHeader(ByVal ws As Worksheet, ByVal nCols As Long, _
                                 ByVal firstRow As Long, ByVal lastRow As Long) As Variant
    Dim hdr As Variant, body As Variant, out() As Variant, r As Long, c As Long, n As Long
    n = lastRow - firstRow + 1
    hdr = ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols)).Value2
    body = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, nCols)).Value2
    ReDim out(1 To n + 1, 1 To nCols)
    For c = 1 To nCols
        out(1, c) = hdr(1, c)
    Next c
    If n = 1 Then
        For c = 1 To nCols
            out(2, c) = body(1, c)
        Next c
    Else
        For r = 1 To n
            For c = 1 To nCols
                out(r + 1, c) = body(r, c)
            Next c
        Next r
    End If
    ChunkWithHeader = out
End Function

' The enabled test cases and their conditions, from the case sheet.
Private Function EnabledCases() As Object
    Dim ws As Worksheet, r As Long, d As Object, tc As String, el As String, f As String, x As Object
    Set d = NewMap()
    Set EnabledCases = d
    Set ws = modScenarioBuilder_PreShock.PsCasesSheet()
    If ws Is Nothing Then Exit Function
    For r = PS_CASE_FIRST_ROW To PS_CASE_LAST_ROW
        tc = SafeText(ws.Cells(r, 2).Value2)
        If Len(tc) = 0 Then Exit For
        If SafeUpperText(ws.Cells(r, 1).Value2) <> "NO" Then
            el = SafeText(ws.Cells(r, 3).Value2)
            f = modScenarioBuilder_PreShock.FilterForCase(tc, el)
            If True Then
                Set x = NewMap()
                x("TestCase") = tc
                x("Element") = el
                If Len(Trim$(f)) = 0 Then x("Filter") = "MISSING FILTER DEFINITION" Else x("Filter") = modScenarioBuilder_Multi.ApplyEntityScope("(" & f & ") AND RECON_BANK IN ('JKB')")
                x("Column") = CaseColumnName(tc, el) & "_" & CStr(r)
                Set d(SafeUpperText(tc) & "|" & SafeUpperText(el)) = x
            End If
        End If
    Next r
End Function

' A column name Excel will accept as a pivot field: no spaces at the ends, no
' punctuation that a field name cannot carry, and short enough to read.
Private Function CaseColumnName(ByVal tc As String, ByVal el As String) As String
    Dim s As String
    s = tc
    If Len(el) > 0 And StrComp(el, tc, vbTextCompare) <> 0 Then s = s & "_" & el
    s = Replace(Replace(Replace(s, " ", "_"), "-", "_"), ".", "_")
    s = Replace(Replace(s, "(", ""), ")", "")
    If Len(s) > 60 Then s = Left$(s, 60)
    CaseColumnName = "TC_" & s
End Function

' ===================== the pivots ===========================================

' A source-grouped overview, a base pivot for each source, and testcase pivots
' share one cache. Empty cases receive a coverage notice instead of an
' unfiltered pivot that could be mistaken for their selected population.
Private Sub BuildPivots(ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal heads As Object, _
                        ByVal nRows As Long, ByVal cases As Collection)
    Dim pc As PivotCache, i As Long, spec As Variant
    Dim rowFields As String, colField As String, dataFields As String, filters As String, sourceKey As Variant
    Dim sourceHeads As Object, dateScopeLabel As String, sourceRows As String, sourceCols As String, sourceFilters As String

    If nRows = 0 Then Exit Sub
    Set pc = CreateJoinedPivotCache(wb, dataWs, heads, nRows, cases)

    rowFields = ROW_SOURCE_COL & "|RECON_STAGE|" & PickFirst(heads, "JCB_IND_REPORT_CLASSIFICATION_NAME|BALANCESHEET_LINE_NAME|COA_NAME") & "|" & _
                PickFirst(heads, "COA_CODE|COA_NAME")
    colField = PickFirst(heads, "BUCKET_DISPLAY_NAME|ALM_BUCKET_NAME")
    filters = "BASE_ELIGIBLE|" & PickFirst(heads, "AS_OF_DATE") & "|" & PickFirst(heads, "ALM_FRAMEWORK_NAME") & "|" & _
              PickFirst(heads, "BRANCH_CODE") & "|" & PickFirst(heads, "SOURCE_CODE|DATA_SOURCE") & "|" & _
              PickFirst(heads, "STAGE_ID")
    dataFields = MeasureList(heads)

    ProgressDetail "Joined pivots: building the source-grouped base overview"
    MakePivotSheet wb, pc, "Base", "Bank base | All source reporting dates", _
        "Each source record appears once. Double-click for reporting fields; ROW_SOURCE and SOURCE_ROW_NUMBER locate the complete record in Joined_Data.", _
        rowFields, colField, filters, dataFields, ""
    If pc.recordCount <> nRows Then Err.Raise vbObjectError + 829, "BuildPivots", "Native cache row count " & pc.recordCount & " differs from the complete joined input " & nRows & "."

    For Each sourceKey In JoinOrder()
        If modScenarioBuilder_PreShock.SourceLoaded(CStr(sourceKey)) Then
            Set sourceHeads = modScenarioBuilder_PreShock.SourceHeadersFor(CStr(sourceKey))
            dateScopeLabel = "Source snapshot; date not supplied"
            If Not sourceHeads Is Nothing Then
                If sourceHeads.Exists("AS_OF_DATE") Then dateScopeLabel = "All source reporting dates"
            End If
            sourceRows = Replace(rowFields, ROW_SOURCE_COL & "|", ""): sourceCols = colField
            sourceFilters = filters & "|" & ROW_SOURCE_COL
            If CStr(sourceKey) = "CAP" Then
                sourceRows = "CAP_COMPONENT_CODE|" & PickFirst(heads, "CAP_COMPONENT_NAME") & "|" & PickFirst(heads, "BRANCH_CODE")
                sourceCols = "": sourceFilters = "BASE_ELIGIBLE|" & ROW_SOURCE_COL & "|" & PickFirst(heads, "AS_OF_DATE")
            End If
            ProgressDetail "Joined pivots: building " & CStr(sourceKey) & " base"
            MakePivotSheet wb, pc, "Base_" & CStr(sourceKey), "Base / " & CStr(sourceKey) & " | " & dateScopeLabel, _
                "One row per original source record. ECL dimensions enrich accounts; amounts belong only to this source.", _
                sourceRows, sourceCols, sourceFilters, _
                SourceMeasureList(CStr(sourceKey), heads), "", CStr(sourceKey)
        End If
    Next sourceKey
    For i = 1 To cases.count
        spec = cases(i)
        ProgressDetail "Joined pivots: testcase " & i & " of " & cases.count & " / " & CStr(spec(1)) & " / " & CStr(spec(2))
        If CDbl(spec(3)) = 0 Then
            WriteEmptyCase wb, CStr(spec(9)), CStr(spec(1)), CStr(spec(2)), "No eligible rows; no all-population pivot substituted."
        Else
        MakePivotSheet wb, pc, CStr(spec(9)), CStr(spec(1)) & " - " & CStr(spec(2)), _
            "The same table, filtered to this test case: " & format$(CDbl(spec(3)), "#,##0") & " row(s) qualify. " & _
            "Change " & CStr(spec(0)) & " in the filters to compare against the rest.", _
            rowFields, colField, filters & "|" & CStr(spec(0)), dataFields, CStr(spec(0))
        End If
    Next i

    On Error Resume Next
    wb.Worksheets("Base").Move Before:=wb.Worksheets(1)
    wb.Worksheets("Base").Activate
    Err.Clear
End Sub

' Copy only reporting fields into the native cache. Every original row remains
' in the same order, with its source identity, source row and account. The full
' wide table remains in Joined_Data; no amounts are grouped or duplicated here.
Private Function CreateJoinedPivotCache(ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal heads As Object, _
                                       ByVal nRows As Long, ByVal cases As Collection) As PivotCache
    Dim selected As Object, lookup As Object, maps As Object, ws As Worksheet, metricWs As Worksheet, mapWs As Worksheet
    Dim field As Variant, spec As Variant, item As Variant, a As Variant, measures As String, expr As String
    Dim r As Long, c As Long, i As Long, physical As String, raw As String, src As Range, pc As PivotCache, proof As PivotTable
    Dim nativeNumber As Long, nativeDescription As String, stage As String
    On Error GoTo Failed
    stage = "select reporting fields"
    Set selected = NewMap(): Set lookup = NewMap(): Set maps = NewMap()
    For Each field In heads.keys
        raw = CStr(field): lookup(raw) = raw
        physical = modScenarioBuilder_PreShock.CanonicalEclHeader(raw)
        If Not lookup.Exists(physical) Then lookup(physical) = raw
    Next field
    Set mapWs = modScenarioBuilder_PreShock.PsFieldsSheet()
    If Not mapWs Is Nothing Then
        a = mapWs.Range(mapWs.Cells(PS_MAP_FIRST_ROW, 1), mapWs.Cells(PS_MAP_LAST_ROW, 2)).Value2
        For r = 1 To UBound(a, 1)
            If Len(SafeText(a(r, 1))) > 0 Then maps(SafeUpperText(a(r, 1))) = SafeText(a(r, 2))
        Next r
    End If
    IncludePivotFields selected, lookup, heads, _
        "ROW_SOURCE|SOURCE_ROW_NUMBER|ROW_COUNT|JOIN_STATUS|ECL_MATCH_COUNT|JOIN_KEY_FIELDS|RECON_STAGE|BASE_ELIGIBLE|RECON_BANK|" & _
        "ACCOUNT_NUMBER|AS_OF_DATE|ENTITY_ID|ENTITY_CODE|BANK_CODE|BANK_ID|BRANCH_CODE|CAP_COMPONENT_CODE|CAP_COMPONENT_NAME|" & _
        "JCB_IND_REPORT_CLASSIFICATION_NAME|BALANCESHEET_LINE_NAME|COA_NAME|COA_CODE|BUCKET_DISPLAY_NAME|ALM_BUCKET_NAME|" & _
        "ALM_FRAMEWORK_NAME|SOURCE_CODE|DATA_SOURCE|STAGE_ID"
    measures = MeasureList(heads, True)
    For Each item In Split(measures, "|")
        IncludePivotFields selected, lookup, heads, Split(CStr(item), ">")(0)
    Next item
    ' Metric expressions also supply fields added dynamically by the native
    ' comparison bridge, including ECL aliases absent from the default layout.
    Set metricWs = modScenarioBuilder_PreShock.PsMetricsSheet()
    If Not metricWs Is Nothing Then
        a = metricWs.Range(metricWs.Cells(PS_METRIC_FIRST_ROW, 4), metricWs.Cells(PS_METRIC_LAST_ROW, 4)).Value2
        For r = 1 To UBound(a, 1)
            expr = SafeText(a(r, 1))
            If Len(expr) > 0 Then IncludePivotExpression selected, lookup, maps, heads, expr
        Next r
    End If
    For i = 1 To cases.count
        spec = cases(i)
        IncludePivotExpression selected, lookup, maps, heads, CStr(spec(6))
        If SafeText(dataWs.Cells(1, CLng(spec(4))).Value2) <> CStr(spec(0)) Then Err.Raise vbObjectError + 828, , "Testcase membership column does not match its coverage definition."
        selected(CStr(spec(0))) = CLng(spec(4))
    Next i
    stage = "copy reporting fields"
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = "Pivot_Source"
    c = 0
    For Each field In selected.keys
        c = c + 1
        ProgressDetail "Joined pivots: preparing field " & c & " of " & selected.count & " / " & CStr(field)
        ws.Cells(1, c).Value2 = CStr(field)
        ws.Cells(2, c).Resize(nRows, 1).Value2 = dataWs.Cells(2, CLng(selected(field))).Resize(nRows, 1).Value2
    Next field
    ws.rows(1).Font.Bold = True: ws.rows(1).Interior.Color = RGB(14, 34, 64): ws.rows(1).Font.Color = vbWhite
    ws.columns.ColumnWidth = 20
    Set src = ws.Range(ws.Cells(1, 1), ws.Cells(nRows + 1, selected.count))
    src.AutoFilter
    stage = "create " & format$(nRows, "#,##0") & " row / " & selected.count & " field native cache"
    ProgressDetail "Joined pivots: creating native cache / " & format$(nRows, "#,##0") & " rows, " & selected.count & " reporting fields"
    ' Microsoft recommends a qualified address string: passing a Range object
    ' can raise an unexpected type mismatch, especially for a large range.
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, sourceData:=src.address(ReferenceStyle:=xlR1C1, External:=True))
    pc.RefreshOnFileOpen = False
    pc.MissingItemsLimit = xlMissingItemsNone
    ' Excel materializes the records when the first PivotTable is created.
    ' Refresh before that point is invalid on the installed native Excel.
    stage = "materialize " & format$(nRows, "#,##0") & " row / " & selected.count & " field native cache"
    Set proof = pc.CreatePivotTable(TableDestination:=ws.Cells(3, selected.count + 3), TableName:="Joined_Record_Control")
    proof.AddDataField proof.PivotFields("ROW_COUNT"), "Complete source records", xlSum
    If pc.recordCount <> nRows Then Err.Raise vbObjectError + 829, , "Native cache row count " & pc.recordCount & " differs from joined input " & nRows & "."
    If CDbl(proof.GetPivotData("Complete source records").Value2) <> nRows Then Err.Raise vbObjectError + 845, , "Native cache source-row sum does not reconcile."
    Set CreateJoinedPivotCache = pc
    Exit Function
Failed:
    nativeNumber = Err.Number: nativeDescription = Err.description
    Err.Raise vbObjectError + 784, "CreateJoinedPivotCache", "Could not " & stage & ". Excel error " & nativeNumber & ": " & nativeDescription
End Function

Private Sub IncludePivotFields(ByVal selected As Object, ByVal lookup As Object, ByVal heads As Object, ByVal candidates As String)
    Dim part As Variant, fieldName As String, raw As String
    For Each part In Split(candidates, "|")
        fieldName = modScenarioBuilder_PreShock.CanonicalEclHeader(Trim$(CStr(part)))
        If lookup.Exists(fieldName) Then
            raw = CStr(lookup(fieldName)): selected(raw) = CLng(heads(raw))
        End If
    Next part
End Sub

Private Sub IncludePivotExpression(ByVal selected As Object, ByVal lookup As Object, ByVal maps As Object, ByVal heads As Object, ByVal expr As String)
    Dim matches As Object, item As Object, token As String
    Set matches = RegexExecute("[A-Za-z_][A-Za-z0-9_]*", expr)
    For Each item In matches
        token = UCase$(CStr(item.value))
        IncludePivotFields selected, lookup, heads, token
        If maps.Exists(token) Then IncludePivotFields selected, lookup, heads, CStr(maps(token))
    Next item
End Sub

' Complete a checkpoint made after row, filter and conservation controls were
' written. Input sources must be loaded so each source keeps its own measures.
' The pending file is retained unchanged if native pivot generation fails.
Public Function ResumeJoinedPivotsTo(ByVal folder As String) As String
    Dim wb As Workbook, ws As Worksheet, coverage As Worksheet, st As Object, heads As Object, layout As Object, cases As Collection
    Dim path As String, pending As String, failure As String, r As Long, c As Long, nRows As Long, col As Long, spec As Variant
    On Error GoTo Failed
    If Not AnyJoinSourceLoaded() Then Err.Raise vbObjectError + 830, , "Load the original source extracts before resuming native pivots."
    pending = JoinPath(folder, "_Joined_Input_pending.xlsx")
    path = JoinPath(folder, "Joined_Input.xlsx")
    Set st = CaptureState(): Application.ScreenUpdating = False: Application.DisplayAlerts = False
    Set wb = Workbooks.Open(pending, UpdateLinks:=0, ReadOnly:=False)
    If wb.ReadOnly Then Err.Raise vbObjectError + 831, , "The pending joined checkpoint is read-only."
    Set ws = wb.Worksheets(JOINED_SHEET): Set coverage = wb.Worksheets("Case_Coverage")
    Set layout = BuildColumnLayout(heads)
    For Each spec In heads.keys
        If SafeText(ws.Cells(1, CLng(heads(spec))).Value2) <> CStr(spec) Then Err.Raise vbObjectError + 832, , "Loaded source columns differ from the pending joined checkpoint."
    Next spec
    nRows = ws.Cells(ws.rows.count, 1).End(xlUp).row - 1
    Set cases = New Collection
    For r = 2 To coverage.Cells(coverage.rows.count, 1).End(xlUp).row
        col = 0
        For c = heads.count + 1 To ws.Cells(1, ws.columns.count).End(xlToLeft).Column
            If SafeText(ws.Cells(1, c).Value2) = SafeText(coverage.Cells(r, 3).Value2) Then col = c: Exit For
        Next c
        If col = 0 Then Err.Raise vbObjectError + 833, , "A pending testcase membership field is missing."
        cases.Add Array(SafeText(coverage.Cells(r, 3).Value2), SafeText(coverage.Cells(r, 1).Value2), SafeText(coverage.Cells(r, 2).Value2), _
                        CDbl(coverage.Cells(r, 5).Value2), col, CDbl(coverage.Cells(r, 6).Value2), SafeText(coverage.Cells(r, 7).Value2), _
                        SafeText(coverage.Cells(r, 8).Value2), SafeText(coverage.Cells(r, 10).Value2), SafeText(coverage.Cells(r, 4).Value2))
    Next r
    BuildPivots wb, ws, heads, nRows, cases
    modReconWorkbench.EnsureJoinedOverview wb, nRows
    UiFinishReportBook wb, "Joined input - account evidence and pivots", "Input files joined at account level, with base and test case pivots"
    wb.SaveAs fileName:=path, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    wb.Close SaveChanges:=False: Set wb = Nothing
    RestoreState st
    ResumeJoinedPivotsTo = path
    Exit Function
Failed:
    failure = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    RestoreState st
    On Error GoTo 0
    Err.Raise vbObjectError + 834, "ResumeJoinedPivotsTo", failure
End Function

Private Sub MakePivotSheet(ByVal wb As Workbook, ByVal pc As PivotCache, ByVal nm As String, _
                           ByVal title As String, ByVal about As String, _
                           ByVal rowFields As String, ByVal colField As String, ByVal filters As String, _
                           ByVal dataFields As String, ByVal yesField As String, Optional ByVal sourceKey As String = "")
    Dim ws As Worksheet, pt As PivotTable, sheetName As String
    On Error Resume Next
    sheetName = SafeSheetName(nm)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.name = sheetName
    If Err.Number <> 0 Then Err.Clear: ws.name = sheetName & "_" & wb.Worksheets.count
    ws.Cells(1, 1).Value2 = title
    ws.Cells(1, 1).Font.Size = 12
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value2 = about
    ws.Cells(2, 1).Font.Size = 8.5
    Err.Clear
    On Error GoTo 0

    modShared_Pivot.LayoutReportSheet ws
    Set pt = modShared_Pivot.AddPivotFrom(pc, ws.Cells(modShared_Pivot.RPT_PIVOT_ROW, 1), nm, _
                                          rowFields, colField, filters, dataFields, Not mPivotsImmediate)
    If pt Is Nothing Then Err.Raise vbObjectError + 782, , "Could not create native pivot: " & nm
    If Len(sourceKey) > 0 Then pt.PivotFields(ROW_SOURCE_COL).CurrentPage = sourceKey
    On Error Resume Next
    pt.PivotFields("BASE_ELIGIBLE").CurrentPage = "Yes"
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 784, , "Base scope is empty or unavailable. Native pivot was not allowed to show all rows."
    End If
    On Error GoTo 0
    ' The case sheets open already filtered to their own case.
    If Len(yesField) > 0 Then
        pt.PivotFields(yesField).CurrentPage = "Yes"
    End If
    On Error Resume Next
    modShared_Pivot.SubtotalOuterOnly pt
    Err.Clear
    On Error GoTo 0
    ApplyJoinedPivotSafety pt, sourceKey
    modShared_Pivot.RemoveEmptyOptionalPageFields pt
    pt.ManualUpdate = False
    On Error Resume Next
    modShared_Pivot.AddSlicers pt, ROW_SOURCE_COL & "|" & colField, 6, _
        modShared_Pivot.RPT_SLICER_Y, modShared_Pivot.RPT_SLICER_W, modShared_Pivot.RPT_SLICER_H
    ws.columns("A:C").ColumnWidth = 28
    modShared_Pivot.FreezeReportSheet ws
    modReconWorkbench.PositionPivotSlicers ws
    ws.Cells(1, 1).Select
    Err.Clear
End Sub

' The amounts worth totalling, whichever of them this set of extracts carries.
Private Function MeasureList(ByVal heads As Object, Optional ByVal includeCapital As Boolean = False) As String
    Dim want As Variant, k As Variant, s As String
    want = Array("ROW_COUNT", "CASHFLOW_AMOUNT_LCY_PRE_FACTOR", "CASHFLOW_AMOUNT_LCY_POST_FACTOR", _
                 "OUTSTANDING_LCY", "OUTSTANDING_FOR_ECL", "OUTST_FOR_ECL", "INTEREST_IN_SUSPENSE_LCY", "IIS", "ECL", "MAX_PRE_FINAL_ECL_MOODY_ECL", _
                 "IMPACT_ECL_LCY_MOODY", "RWA_LCY", "OUTSTANDING_FOR_RWA", "REPORT_BALANCE")
    For Each k In want
        If heads.Exists(CStr(k)) And (CStr(k) <> "REPORT_BALANCE" Or includeCapital) Then
            If Len(s) > 0 Then s = s & "|"
            s = s & CStr(k) & ">" & CStr(k) & " >#,##0"
        End If
    Next k
    MeasureList = s
End Function

' The first of these field names the table actually has.
Private Function PickFirst(ByVal heads As Object, ByVal candidates As String) As String
    Dim p As Variant
    For Each p In Split(candidates, "|")
        If heads.Exists(Trim$(CStr(p))) Then PickFirst = Trim$(CStr(p)): Exit Function
    Next p
End Function

Private Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, o As String
    o = s
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each b In bad
        o = Replace(o, CStr(b), "_")
    Next b
    If Len(o) > 31 Then o = Left$(o, 31)
    If Len(o) = 0 Then o = "Sheet"
    SafeSheetName = o
End Function

Private Sub FormatJoinedSheet(ByVal ws As Worksheet, ByVal nCols As Long, ByVal nRows As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Font.Bold = True
        .Interior.Color = RGB(25, 63, 137)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.rows(1).RowHeight = 32
    ws.rows(1).WrapText = True
    ws.columns.ColumnWidth = 18
    ws.columns(4).ColumnWidth = 39
    ws.Range(ws.Cells(1, 1), ws.Cells(nRows + 1, nCols)).AutoFilter
    ws.Range("A1").EntireRow.Font.Size = 9
    ' Through the book's own window: the person may be in another one by now.
    If UiTryActivate(ws) Then UiSetView ws, 1, 0
    Err.Clear
End Sub


Private Function StageKey(ByVal value As Variant) As String
    Dim s As String
    s = Replace(SafeUpperText(value), "STAGE ", "STAGE")
    If s = "1" Or s = "STAGE1" Then StageKey = "1": Exit Function
    If s = "2" Or s = "STAGE2" Then StageKey = "2": Exit Function
    If s = "3" Or s = "STAGE3" Or s = "NPA" Then StageKey = "3": Exit Function
    StageKey = "Unstaged"
End Function

Private Function SourceMeasureList(ByVal sourceKey As String, ByVal heads As Object) As String
    Dim h As Object, d As Object, k As Variant
    Set h = modScenarioBuilder_PreShock.SourceHeadersFor(sourceKey): Set d = NewMap()
    d("ROW_COUNT") = 1
    For Each k In h.keys
        If heads.Exists(CStr(k)) Then d(CStr(k)) = heads(k)
    Next k
    SourceMeasureList = MeasureList(d, sourceKey = "CAP")
End Function

Private Sub ApplyJoinedPivotSafety(ByVal pt As PivotTable, ByVal sourceKey As String)
    Dim pf As PivotField, i As Long
    If Len(sourceKey) = 0 Or sourceKey = "CAP" Then
        pt.ColumnGrand = False: pt.RowGrand = False
    End If
    If sourceKey = "CAP" Then
        For Each pf In pt.rowFields
            For i = 1 To 12: pf.Subtotals(i) = False: Next i
        Next pf
        For Each pf In pt.dataFields
            If pf.sourceName = "REPORT_BALANCE" Then pf.NumberFormat = "#,##0.########"
        Next pf
        pt.Parent.Range("A2").Value2 = "Capital components are shown separately by branch. Balances, RWA and ratios are not added together; this snapshot has no independent reporting-date proof."
    End If
End Sub

' Updates an already generated pack using its original joined records. No source
' extracts are reread, and the raw source-conservation controls are untouched.
Public Function RepairJoinedReport(ByVal path As String) As String
    Dim wb As Workbook, existing As Workbook, ws As Worksheet, dataWs As Worksheet, coverage As Worksheet
    Dim state As Object, heads As Object, pt As PivotTable, pc As PivotCache, pf As PivotField
    Dim firstCap As Range, lastCap As Range, sourceRange As Range, capFlags As Range, a As Variant
    Dim lastRow As Long, lastCol As Long, capRows As Long, r As Long, c As Long, i As Long
    Dim flagCol As Long, removedYes As Long, removedReview As Long, caseRow As Long, position As Long
    Dim key As String, sourceKey As String, caseSheet As String, er As String
    On Error GoTo Failed
    For Each existing In Application.Workbooks
        If StrComp(existing.FullName, path, vbTextCompare) = 0 Then Err.Raise vbObjectError + 793, , "Close the joined report before applying its presentation update."
    Next existing
    Set state = CaptureState(): Application.ScreenUpdating = False: Application.EnableEvents = False
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=False)
    If wb.ReadOnly Then Err.Raise vbObjectError + 794, , "The joined report is read-only."
    Set dataWs = wb.Worksheets(JOINED_SHEET): Set coverage = wb.Worksheets("Case_Coverage")
    lastRow = dataWs.Cells(dataWs.rows.count, 1).End(xlUp).row
    lastCol = dataWs.Cells(1, dataWs.columns.count).End(xlToLeft).Column
    Set heads = NewMap()
    For c = 1 To lastCol: heads(SafeText(dataWs.Cells(1, c).Value2)) = c: Next c
    Set sourceRange = dataWs.Range(dataWs.Cells(2, heads(ROW_SOURCE_COL)), dataWs.Cells(lastRow, heads(ROW_SOURCE_COL)))
    Set firstCap = sourceRange.Find(what:="CAP", After:=sourceRange.Cells(sourceRange.rows.count, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If Not firstCap Is Nothing Then
        Set lastCap = sourceRange.Find(what:="CAP", After:=sourceRange.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
        capRows = lastCap.row - firstCap.row + 1
        If WorksheetFunction.CountIf(sourceRange, "CAP") <> capRows Then Err.Raise vbObjectError + 795, , "Capital source records are not contiguous; no partial presentation repair was saved."
        For caseRow = 2 To coverage.Cells(coverage.rows.count, 1).End(xlUp).row
            key = SafeText(coverage.Cells(caseRow, 3).Value2)
            If heads.Exists(key) Then
                flagCol = CLng(heads(key)): removedYes = 0: removedReview = 0
                Set capFlags = dataWs.Range(dataWs.Cells(firstCap.row, flagCol), dataWs.Cells(lastCap.row, flagCol))
                a = capFlags.Value2
                For r = 1 To capRows
                    If capRows = 1 Then key = SafeText(a) Else key = SafeText(a(r, 1))
                    If key = "Yes" Then removedYes = removedYes + 1
                    If key <> "Yes" And key <> "No" Then removedReview = removedReview + 1
                Next r
                capFlags.Value2 = "No"
                coverage.Cells(caseRow, 5).Value2 = WorksheetFunction.Max(0, CDbl(coverage.Cells(caseRow, 5).Value2) - removedYes)
                coverage.Cells(caseRow, 6).Value2 = WorksheetFunction.Max(0, CDbl(coverage.Cells(caseRow, 6).Value2) - removedReview)
                If CDbl(coverage.Cells(caseRow, 5).Value2) = 0 Then
                    coverage.Cells(caseRow, 10).Value2 = "No eligible rows - review coverage"
                ElseIf SafeText(coverage.Cells(caseRow, 10).Value2) = "Review source dates, join exceptions and mapping" And CDbl(coverage.Cells(caseRow, 6).Value2) = 0 Then
                    coverage.Cells(caseRow, 10).Value2 = "Calculated"
                End If
            End If
        Next caseRow
    End If
    For Each pc In wb.PivotCaches: pc.Refresh: Next pc
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            sourceKey = ""
            If Left$(ws.name, 5) = "Base_" Then sourceKey = Mid$(ws.name, 6)
            If Len(sourceKey) = 0 Then
                For i = pt.dataFields.count To 1 Step -1
                    If SafeUpperText(pt.dataFields(i).sourceName) = "REPORT_BALANCE" Then pt.dataFields(i).orientation = xlHidden
                Next i
            ElseIf sourceKey = "CAP" Then
                For i = pt.rowFields.count To 1 Step -1: pt.rowFields(i).orientation = xlHidden: Next i
                For i = pt.ColumnFields.count To 1 Step -1
                    If pt.ColumnFields(i).name <> pt.DataPivotField.name Then pt.ColumnFields(i).orientation = xlHidden
                Next i
                position = 0
                For Each a In Array("CAP_COMPONENT_CODE", "CAP_COMPONENT_NAME", "BRANCH_CODE")
                    If heads.Exists(CStr(a)) Then
                        position = position + 1: Set pf = pt.PivotFields(CStr(a))
                        pf.orientation = xlRowField: pf.position = position
                    End If
                Next a
                If Not heads.Exists("CAP_COMPONENT_CODE") Then Err.Raise vbObjectError + 796, , "Capital component codes are required to separate amounts and ratios."
            End If
            ApplyJoinedPivotSafety pt, sourceKey
        Next pt
    Next ws
    For caseRow = 2 To coverage.Cells(coverage.rows.count, 1).End(xlUp).row
        caseSheet = SafeText(coverage.Cells(caseRow, 4).Value2)
        Set ws = wb.Worksheets(caseSheet)
        ws.Range("A2").Value2 = format$(CDbl(coverage.Cells(caseRow, 5).Value2), "#,##0") & " source records qualify. Capital components are base-only; see Case_Coverage for the filter and reporting date."
        If CDbl(coverage.Cells(caseRow, 5).Value2) = 0 Then
            If ws.PivotTables.count > 0 Then ws.PivotTables(1).TableRange2.Clear
            ws.Range("A3").Value2 = "No eligible rows. Capital components are available only in Base_CAP."
        ElseIf ws.PivotTables.count > 0 Then
            ws.PivotTables(1).PivotFields(SafeText(coverage.Cells(caseRow, 3).Value2)).CurrentPage = "Yes"
        End If
    Next caseRow
    wb.Save: wb.Close SaveChanges:=False: Set wb = Nothing
    RestoreState state
    RepairJoinedReport = "PASS: capital components separated; cross-source totals removed; " & capRows & " capital rows excluded from testcase pivots."
    Exit Function
Failed:
    er = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    RestoreState state
    RepairJoinedReport = "ERROR: " & er
End Function

Private Sub WriteEmptyCase(ByVal wb As Workbook, ByVal nm As String, ByVal testCase As String, ByVal element As String, ByVal explanation As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = nm
    ws.Range("A1").Value2 = testCase & " / " & element
    ws.Range("A3").Value2 = explanation
    ws.Range("A5").Value2 = "See Case_Coverage and Joined_Data for the filter, reporting date and source exceptions."
    ws.columns("A:A").ColumnWidth = 100: ws.Range("A1").Font.Bold = True
End Sub

Private Sub WriteCaseCoverage(ByVal wb As Workbook, ByVal cases As Collection)
    Dim ws As Worksheet, i As Long, spec As Variant
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1)): ws.name = "Case_Coverage"
    ws.Range("A1:J1").Value2 = Array("Test case", "Element", "Membership field", "Pivot sheet", "Rows matched", "Review rows", "Filter", "Reporting date", "Date mode", "Status")
    For i = 1 To cases.count
        spec = cases(i)
        ws.Cells(i + 1, 1).Value2 = spec(1): ws.Cells(i + 1, 2).Value2 = spec(2)
        ws.Cells(i + 1, 3).Value2 = spec(0): ws.Cells(i + 1, 4).Value2 = spec(9)
        ws.Cells(i + 1, 5).Value2 = spec(3): ws.Cells(i + 1, 6).Value2 = spec(5)
        ws.Cells(i + 1, 7).Value2 = spec(6): ws.Cells(i + 1, 8).Value2 = spec(7)
        ws.Cells(i + 1, 9).Value2 = IIf(modScenarioBuilder_PreShock.IgnoreAsOfDate(), "Ignore - comparison requires review", "Strict")
        ws.Cells(i + 1, 10).Value2 = spec(8)
        ws.Hyperlinks.Add anchor:=ws.Cells(i + 1, 4), address:="", SubAddress:="'" & CStr(spec(9)) & "'!A1"
    Next i
    ws.rows(1).Font.Bold = True: ws.rows(1).Interior.Color = RGB(14, 34, 64): ws.rows(1).Font.Color = vbWhite
    ws.columns.ColumnWidth = 25: ws.columns(7).ColumnWidth = 70: ws.columns(10).ColumnWidth = 45
    ws.Range("A1:J" & cases.count + 1).AutoFilter
End Sub

Private Sub WriteSourceControls(ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal heads As Object, ByVal nRows As Long)
    Dim ws As Worksheet, sourceKey As Variant, data As Variant, h As Object, fields As String, item As Variant, field As String
    Dim r As Long, outRow As Long, inputSum As Double, invalid As Long, sourceRange As String, amountRange As String
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1)): ws.name = "Source_Controls"
    ws.Range("A1:G1").Value2 = Array("Source", "Measure", "Original input total", "Joined source total", "Variance", "Missing or invalid inputs", "Control")
    outRow = 2
    sourceRange = "'Joined_Data'!" & dataWs.Range(dataWs.Cells(2, heads(ROW_SOURCE_COL)), dataWs.Cells(nRows + 1, heads(ROW_SOURCE_COL))).address
    For Each sourceKey In JoinOrder()
        If modScenarioBuilder_PreShock.SourceLoaded(CStr(sourceKey)) Then
            data = modScenarioBuilder_PreShock.SourceDataFor(CStr(sourceKey)): Set h = modScenarioBuilder_PreShock.SourceHeadersFor(CStr(sourceKey))
            fields = SourceMeasureList(CStr(sourceKey), heads)
            For Each item In Split(fields, "|")
                field = Split(CStr(item), ">")(0): inputSum = 0: invalid = 0
                If field = "ROW_COUNT" Then
                    inputSum = UBound(data, 1) - 1
                ElseIf h.Exists(field) Then
                    For r = 2 To UBound(data, 1)
                        If IsNumeric(data(r, CLng(h(field)))) And Not IsEmpty(data(r, CLng(h(field)))) Then
                            inputSum = inputSum + CDbl(data(r, CLng(h(field))))
                        Else
                            invalid = invalid + 1
                        End If
                    Next r
                End If
                ws.Cells(outRow, 1).Value2 = CStr(sourceKey): ws.Cells(outRow, 2).Value2 = field
                ws.Cells(outRow, 3).Value2 = inputSum: ws.Cells(outRow, 6).Value2 = invalid
                amountRange = "'Joined_Data'!" & dataWs.Range(dataWs.Cells(2, heads(field)), dataWs.Cells(nRows + 1, heads(field))).address
                ws.Cells(outRow, 4).formula = "=SUMIFS(" & amountRange & "," & sourceRange & ",A" & outRow & ")"
                ws.Cells(outRow, 5).formula = "=D" & outRow & "-C" & outRow
                ws.Cells(outRow, 7).formula = "=IF(F" & outRow & ">0,""REVIEW missing values"",IF(ABS(E" & outRow & ")<=0.01,""MATCH"",""DIFFERENCE""))"
                outRow = outRow + 1
            Next item
        End If
    Next sourceKey
    ws.Calculate
    ws.rows(1).Font.Bold = True: ws.rows(1).Interior.Color = RGB(14, 34, 64): ws.rows(1).Font.Color = vbWhite
    ws.columns("A:B").ColumnWidth = 33: ws.columns("C:G").ColumnWidth = 25
    ws.columns("C:E").NumberFormat = "#,##0.00;[Red](#,##0.00)"
    ws.Range("A1:G" & outRow - 1).AutoFilter
End Sub


Public Function ReconJoinedRegressionTests(ByVal folder As String) As String
    Dim ecl(1 To 5, 1 To 7) As Variant, ll(1 To 5, 1 To 6) As Variant, rwa(1 To 3, 1 To 6) As Variant
    Dim h As Object, r As Long, c As Long, hdr As Variant, resultPath As String, wb As Workbook, ws As Worksheet
    Dim headers As Object, a As Variant, sumEcl As Double, sumLl As Double, sumRwa As Double, conflictRows As Long
    If StageKey("Stage 1") <> "1" Or StageKey("Stage 2") <> "2" Or StageKey("NPA") <> "3" Then Err.Raise vbObjectError + 791, , "Joined regression: valid stage labels were rejected."
    If StageKey("Stage10") <> "Unstaged" Or StageKey("Stage20") <> "Unstaged" Or StageKey("Stage30") <> "Unstaged" Then Err.Raise vbObjectError + 792, , "Joined regression: invalid stage prefixes were accepted."
    hdr = Array("ACCOUNT_NUMBER", "AS_OF_DATE", "BANK_CODE", "BRANCH_CODE", "STAGE_ID", "SECTOR", "OUTSTANDING_LCY")
    For c = 1 To 7: ecl(1, c) = hdr(c - 1): Next c
    For r = 2 To 5
        ecl(r, 1) = "A": ecl(r, 2) = CDbl(DateSerial(2025, 12, 31)): ecl(r, 3) = "JKB": ecl(r, 4) = "1": ecl(r, 5) = 1
        ecl(r, 6) = "RETAIL": ecl(r, 7) = (r - 1) * 100
    Next r
    ecl(3, 6) = "CORPORATE": ecl(4, 1) = "B": ecl(5, 2) = CDbl(DateSerial(2024, 12, 31))
    Set h = TestHeaders(ecl): modScenarioBuilder_PreShock.RegisterTable "ECL", ecl, h, "synthetic", "ECL"
    hdr = Array("ACCOUNT_NUMBER", "AS_OF_DATE", "BANK_CODE", "BRANCH_CODE", "CASHFLOW_AMOUNT_LCY_PRE_FACTOR", "STAGE_ID")
    For c = 1 To 6: ll(1, c) = hdr(c - 1): Next c
    For r = 2 To 5
        ll(r, 1) = "A": ll(r, 2) = CDbl(DateSerial(2025, 12, 31)): ll(r, 3) = "JKB": ll(r, 4) = "1": ll(r, 5) = (r - 1) * 10
    Next r
    ll(3, 1) = "B": ll(4, 2) = CDbl(DateSerial(2024, 12, 31)): ll(5, 1) = "Z": ll(5, 3) = "UFICO"
    Set h = TestHeaders(ll): modScenarioBuilder_PreShock.RegisterTable "LL", ll, h, "synthetic", "LL"
    hdr = Array("ACCOUNT_NUMBER", "AS_OF_DATE", "BANK_CODE", "BRANCH_CODE", "RWA_LCY", "IFRS_STAGE_ID")
    For c = 1 To 6: rwa(1, c) = hdr(c - 1): Next c
    For r = 2 To 3
        rwa(r, 1) = IIf(r = 2, "A", "B"): rwa(r, 2) = CDbl(DateSerial(2025, 12, 31)): rwa(r, 3) = "JKB": rwa(r, 4) = "1": rwa(r, 5) = r + 5: rwa(r, 6) = 1
    Next r
    Set h = TestHeaders(rwa): modScenarioBuilder_PreShock.RegisterTable "CAPRWA", rwa, h, "synthetic", "CAPRWA"
    resultPath = BuildJoinedInputTo(folder)
    Set wb = Workbooks.Open(resultPath, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(JOINED_SHEET): a = ws.UsedRange.Value2: Set headers = NewMap()
    For c = 1 To UBound(a, 2): headers(SafeText(a(1, c))) = c: Next c
    If UBound(a, 1) - 1 <> 10 Then Err.Raise vbObjectError + 785, , "Joined regression: source record count changed."
    For r = 2 To UBound(a, 1)
        If IsNumeric(a(r, headers("OUTSTANDING_LCY"))) Then sumEcl = sumEcl + CDbl(a(r, headers("OUTSTANDING_LCY")))
        If IsNumeric(a(r, headers("CASHFLOW_AMOUNT_LCY_PRE_FACTOR"))) Then sumLl = sumLl + CDbl(a(r, headers("CASHFLOW_AMOUNT_LCY_PRE_FACTOR")))
        If IsNumeric(a(r, headers("RWA_LCY"))) Then sumRwa = sumRwa + CDbl(a(r, headers("RWA_LCY")))
        If InStr(1, SafeText(a(r, headers("JOIN_STATUS"))), "Conflicting ECL", vbTextCompare) > 0 Then conflictRows = conflictRows + 1
    Next r
    If sumEcl <> 1000 Or sumLl <> 100 Or sumRwa <> 15 Then Err.Raise vbObjectError + 786, , "Joined regression: measures were multiplied or dropped."
    If conflictRows <> 2 Then Err.Raise vbObjectError + 787, , "Joined regression: conflicting duplicate accounts not flagged."
    If wb.Worksheets("Base_ECL").PivotTables.count <> 1 Or wb.Worksheets("Base_LL").PivotTables.count <> 1 Then Err.Raise vbObjectError + 788, , "Joined regression: native source pivots missing."
    If Abs(CDbl(wb.Worksheets("Base_ECL").PivotTables(1).GetPivotData("OUTSTANDING_LCY ").Value2) - 1000) > 0.01 Then Err.Raise vbObjectError + 789, , "Joined regression: native ECL pivot does not reconcile."
    If Abs(CDbl(wb.Worksheets("Base_LL").PivotTables(1).GetPivotData("CASHFLOW_AMOUNT_LCY_PRE_FACTOR ").Value2) - 60) > 0.01 Then Err.Raise vbObjectError + 790, , "Joined regression: subsidiary cashflow contaminated JKB base."
    wb.Close SaveChanges:=False
    ReconJoinedRegressionTests = "PASS: 10 source records preserved; ECL 1000, LL 100, CAPRWA 15; 2 ambiguous joins exposed; native pivot reconciled."
End Function

Private Function TestHeaders(ByRef data As Variant) As Object
    Dim h As Object, c As Long, key As String
    Set h = NewMap()
    For c = 1 To UBound(data, 2)
        key = NormalHeader(data(1, c)): h(key) = c
        key = modScenarioBuilder_PreShock.CanonicalEclHeader(key)
        If Not h.Exists(key) Then h(key) = c
    Next c
    Set TestHeaders = h
End Function

Public Function ReconJoinedCapitalRegressionTests(ByVal folder As String) As String
    Dim a(1 To 4, 1 To 7) As Variant, hdr As Variant, c As Long, r As Long, i As Long, flagCol As Long
    Dim h As Object, wb As Workbook, ws As Worksheet, coverage As Worksheet, pt As PivotTable, pf As PivotField
    Dim path As String, result As String, fieldName As String, seen As Boolean
    hdr = Array("BANK_CODE", "BRANCH_CODE", "CAPITAL_ELEMENT", "REPORT_BALANCE", "CAP_COMPONENT_CODE", "CAP_COMPONENT_NAME", "BANK_ID")
    For c = 1 To 7: a(1, c) = hdr(c - 1): Next c
    For r = 2 To 4: a(r, 1) = "JKB": a(r, 2) = "1": a(r, 7) = 101: Next r
    a(2, 3) = "CET1": a(2, 4) = 100: a(2, 5) = "NET_CET1": a(2, 6) = "Core capital"
    a(3, 3) = "CAPITAL": a(3, 4) = 0.1: a(3, 5) = "CAR": a(3, 6) = "Capital ratio"
    a(4, 3) = "CAPITAL": a(4, 4) = 2000: a(4, 5) = "TOTAL_RWA": a(4, 6) = "Risk weighted assets"
    Set h = TestHeaders(a): modScenarioBuilder_PreShock.RegisterTable "CAP", a, h, "synthetic", "CAP"
    path = BuildJoinedInputTo(folder)
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=False)
    Set pt = wb.Worksheets("Base_CAP").PivotTables(1)
    If pt.RowGrand Or pt.ColumnGrand Then Err.Raise vbObjectError + 797, , "Capital regression: mixed component grand totals are visible."
    If pt.PivotFields("CAP_COMPONENT_CODE").orientation <> xlRowField Then Err.Raise vbObjectError + 798, , "Capital regression: components are not separated."
    For Each pf In pt.rowFields
        For i = 1 To 12
            If pf.Subtotals(i) Then Err.Raise vbObjectError + 799, , "Capital regression: mixed component subtotal is visible."
        Next i
    Next pf
    If Abs(CDbl(pt.GetPivotData("REPORT_BALANCE ", "CAP_COMPONENT_CODE", "NET_CET1", "CAP_COMPONENT_NAME", "Core capital", "BRANCH_CODE", "1").Value2) - 100) > 0.001 Then Err.Raise vbObjectError + 800, , "Capital regression: native component detail changed."
    Set ws = wb.Worksheets(JOINED_SHEET): Set coverage = wb.Worksheets("Case_Coverage")
    fieldName = SafeText(coverage.Cells(2, 3).Value2)
    For c = 1 To ws.Cells(1, ws.columns.count).End(xlToLeft).Column
        If SafeText(ws.Cells(1, c).Value2) = fieldName Then flagCol = c: Exit For
    Next c
    If flagCol = 0 Then Err.Raise vbObjectError + 801, , "Capital regression: testcase flag missing."
    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
        If SafeText(ws.Cells(r, 1).Value2) = "CAP" Then
            If SafeText(ws.Cells(r, flagCol).Value2) <> "No" Then Err.Raise vbObjectError + 802, , "Capital regression: CAP entered filtered pre-shock."
            ws.Cells(r, flagCol).Value2 = "Yes"
        End If
    Next r
    coverage.Cells(2, 5).Value2 = CDbl(coverage.Cells(2, 5).Value2) + 3
    Set pt = wb.Worksheets("Base").PivotTables(1)
    Set pf = pt.AddDataField(pt.PivotFields("REPORT_BALANCE"), "Unsafe capital total", xlSum)
    pt.RowGrand = True: pt.ColumnGrand = True
    Set pt = wb.Worksheets("Base_CAP").PivotTables(1): pt.RowGrand = True: pt.ColumnGrand = True
    wb.Save: wb.Close SaveChanges:=False
    result = RepairJoinedReport(path)
    If Left$(result, 4) <> "PASS" Then Err.Raise vbObjectError + 803, , result
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    For Each ws In wb.Worksheets
        If ws.name = "Base" Or Left$(ws.name, 5) = "Case_" Then
            For Each pt In ws.PivotTables
                If pt.RowGrand Or pt.ColumnGrand Then Err.Raise vbObjectError + 804, , "Capital regression: cross-source grand totals survived repair."
                For Each pf In pt.dataFields
                    If SafeUpperText(pf.sourceName) = "REPORT_BALANCE" Then Err.Raise vbObjectError + 805, , "Capital regression: mixed capital measure survived repair."
                Next pf
            Next pt
        End If
    Next ws
    Set ws = wb.Worksheets(JOINED_SHEET)
    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
        If SafeText(ws.Cells(r, 1).Value2) = "CAP" And SafeText(ws.Cells(r, flagCol).Value2) <> "No" Then Err.Raise vbObjectError + 806, , "Capital regression: unsafe membership survived repair."
    Next r
    Set pt = wb.Worksheets("Base_CAP").PivotTables(1)
    If pt.RowGrand Or pt.ColumnGrand Then Err.Raise vbObjectError + 807, , "Capital regression: capital grand totals survived repair."
    wb.Close SaveChanges:=False
    ReconJoinedCapitalRegressionTests = "PASS: native capital components and ratios stay separate; CAP excluded from testcase populations; existing-report repair corrected unsafe layout and membership."
End Function

' Native regression over the production membership loop. No input file is read.
Public Function ReconJoinedDateRegressionTests() As String
    Dim wb As Workbook, ws As Worksheet, casesWs As Worksheet, heads As Object, contexts As Object, layout As Object
    Dim cases As Collection, spec As Variant, hdr As Variant, expected As Variant, key As String, reportDate As String
    Dim oldEnabled As Variant, oldOverride As Variant, oldIgnore As Boolean, c As Long, r As Long, flagCol As Long, n As Long
    Dim errorText As String, saved As Boolean
    On Error GoTo Failed
    Set casesWs = modScenarioBuilder_PreShock.PsCasesSheet()
    key = SafeUpperText(casesWs.Cells(PS_CASE_FIRST_ROW, 2).Value2) & "|" & SafeUpperText(casesWs.Cells(PS_CASE_FIRST_ROW, 3).Value2)
    Set contexts = CaseContexts()
    If Not contexts.Exists(key) Then Err.Raise vbObjectError + 808, , "Joined date regression requires a testcase with a reporting date."
    reportDate = CStr(contexts(key))
    If reportDate = "AMBIGUOUS" Or Len(reportDate) = 0 Then Err.Raise vbObjectError + 809, , "Joined date regression requires one reporting date."
    oldEnabled = casesWs.Cells(PS_CASE_FIRST_ROW, 1).formula
    oldOverride = casesWs.Cells(PS_CASE_FIRST_ROW, 7).formula
    oldIgnore = modScenarioBuilder_PreShock.IgnoreAsOfDate(): saved = True
    casesWs.Cells(PS_CASE_FIRST_ROW, 1).Value2 = "Yes": casesWs.Cells(PS_CASE_FIRST_ROW, 7).Value2 = "ALL"
    Set wb = Workbooks.Add: Set ws = wb.Worksheets(1): Set heads = NewMap(): Set layout = NewMap()
    hdr = Array(ROW_SOURCE_COL, "JOIN_STATUS", "AS_OF_DATE", "RECON_BANK", "BRANCH_CODE", "BASE_ELIGIBLE", "ACCOUNT_NUMBER")
    For c = 1 To 7: heads(CStr(hdr(c - 1))) = c: ws.Cells(1, c).Value2 = hdr(c - 1): Next c
    For r = 2 To 7
        ws.Cells(r, 1).Value2 = "ECL": ws.Cells(r, 2).Value2 = "ECL source row"
        ws.Cells(r, 4).Value2 = "JKB": ws.Cells(r, 5).Value2 = "1": ws.Cells(r, 7).Value2 = "A" & r
    Next r
    ws.Cells(2, 3).Value2 = CDbl(CDate(reportDate)): ws.Cells(3, 3).Value2 = CDbl(CDate(reportDate)) - 1
    ws.Cells(4, 3).Value2 = "not-a-date"
    ws.Cells(6, 2).Value2 = "Conflicting ECL dimensions - review": ws.Cells(6, 3).Value2 = "not-a-date"
    ws.Cells(7, 1).Value2 = "CAP": ws.Cells(7, 2).Value2 = "Base-level capital row"
    modScenarioBuilder_PreShock.SetAsOfMatching "Match"
    Set cases = MarkTestCases(ws, layout, heads, 6): spec = cases(1): flagCol = CLng(spec(4))
    expected = Array("Yes", "Date mismatch", "Not tested", "Not tested", "Review", "No")
    For r = 2 To 7
        If SafeText(ws.Cells(r, flagCol).Value2) <> CStr(expected(r - 2)) Then Err.Raise vbObjectError + 810, , "Joined date regression: Match changed row " & r & " status."
        n = n + 1
    Next r
    modScenarioBuilder_PreShock.SetAsOfMatching "Ignore"
    Set cases = MarkTestCases(ws, layout, heads, 6)
    expected = Array("Yes", "Yes", "Yes", "Yes", "Review", "No")
    For r = 2 To 7
        If SafeText(ws.Cells(r, flagCol).Value2) <> CStr(expected(r - 2)) Then Err.Raise vbObjectError + 811, , "Joined date regression: Ignore changed row " & r & " status."
        n = n + 1
    Next r
    heads.Remove "AS_OF_DATE": heads("UNUSED_DATE") = 3: ws.Cells(1, 3).Value2 = "UNUSED_DATE"
    modScenarioBuilder_PreShock.SetAsOfMatching "Match"
    Set cases = MarkTestCases(ws, layout, heads, 6)
    For r = 2 To 5
        If SafeText(ws.Cells(r, flagCol).Value2) <> "Not tested" Then Err.Raise vbObjectError + 812, , "Joined date regression: missing date column was accepted."
        n = n + 1
    Next r
    ReconJoinedDateRegressionTests = "PASS: " & n & " joined membership date-policy and precedence checks"
    GoTo Done
Failed:
    errorText = Err.description
Done:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If saved Then
        casesWs.Cells(PS_CASE_FIRST_ROW, 1).formula = oldEnabled: casesWs.Cells(PS_CASE_FIRST_ROW, 7).formula = oldOverride
        modScenarioBuilder_PreShock.SetAsOfMatching IIf(oldIgnore, "Ignore", "Match")
    End If
    On Error GoTo 0
    If Len(errorText) > 0 Then Err.Raise vbObjectError + 813, "ReconJoinedDateRegressionTests", errorText
End Function

Public Function ReconJoinedFinancialOwnershipRegressionTests(ByVal folder As String) As String
    Dim ecl(1 To 3, 1 To 14) As Variant, ll(1 To 3, 1 To 6) As Variant, hdr As Variant, a As Variant, field As Variant
    Dim h As Object, columns As Object, wb As Workbook, ws As Worksheet, path As String, r As Long, c As Long
    Dim nEcl As Long, nLl As Long, eclCollateral As Double, eclEad As Double, llCollateral As Double, llCashflow As Double
    Dim controlsFound As Long, sourceRows As Range, measureRows As Range
    If IsMeasureField("COLLATERAL_TYPE_CODE") Or IsMeasureField("INTEREST_RATE_CODE") Or IsMeasureField("COLLATERAL_DESCRIPTION") Then Err.Raise vbObjectError + 814, , "Financial ownership regression: descriptive dimensions were removed."
    For Each field In Array("TOTAL_COLLATERAL", "CASH_COLLATERAL_POST_HC", "PREV_EAD", "GUARANTEED_LOANS", "TOTAL_INTEREST_ADVANCE", "EFFECTIVE_PD")
        If Not IsMeasureField(CStr(field)) Then Err.Raise vbObjectError + 815, , "Financial ownership regression: financial measure classified as a dimension: " & CStr(field)
    Next field
    hdr = Array("ACCOUNT_NUMBER", "AS_OF_DATE", "BANK_CODE", "BRANCH_CODE", "STAGE_ID", "SECTOR", _
                "COLLATERAL_TYPE_CODE", "INTEREST_RATE_CODE", "TOTAL_COLLATERAL", "PREV_EAD", _
                "GUARANTEED_LOANS", "TOTAL_INTEREST_ADVANCE", "EFFECTIVE_PD", "OUTSTANDING_LCY")
    For c = 1 To 14: ecl(1, c) = hdr(c - 1): Next c
    For r = 2 To 3
        ecl(r, 1) = "A": ecl(r, 2) = CDbl(DateSerial(2025, 12, 31)): ecl(r, 3) = "JKB": ecl(r, 4) = "1"
        ecl(r, 5) = 1: ecl(r, 6) = "RETAIL": ecl(r, 7) = "CASH": ecl(r, 8) = "FIXED"
        ecl(r, 9) = (r - 1) * 1000: ecl(r, 10) = (r - 1) * 10: ecl(r, 11) = (r - 1) * 30
        ecl(r, 12) = (r - 1) * 50: ecl(r, 13) = (r - 1) / 10: ecl(r, 14) = (r - 1) * 100
    Next r
    Set h = TestHeaders(ecl): modScenarioBuilder_PreShock.RegisterTable "ECL", ecl, h, "synthetic", "ECL"
    hdr = Array("ACCOUNT_NUMBER", "AS_OF_DATE", "BANK_CODE", "BRANCH_CODE", "CASHFLOW_AMOUNT_LCY_PRE_FACTOR", "TOTAL_COLLATERAL")
    For c = 1 To 6: ll(1, c) = hdr(c - 1): Next c
    For r = 2 To 3
        ll(r, 1) = "A": ll(r, 2) = CDbl(DateSerial(2025, 12, 31)): ll(r, 3) = "JKB": ll(r, 4) = "1"
        ll(r, 5) = (r - 1) * 10: ll(r, 6) = r + 3
    Next r
    Set h = TestHeaders(ll): modScenarioBuilder_PreShock.RegisterTable "LL", ll, h, "synthetic", "LL"
    path = BuildJoinedInputTo(folder)
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(JOINED_SHEET): a = ws.UsedRange.Value2: Set columns = NewMap()
    For c = 1 To UBound(a, 2): columns(SafeText(a(1, c))) = c: Next c
    For r = 2 To UBound(a, 1)
        If SafeText(a(r, columns(ROW_SOURCE_COL))) = "ECL" Then
            nEcl = nEcl + 1: eclCollateral = eclCollateral + CDbl(a(r, columns("TOTAL_COLLATERAL")))
            eclEad = eclEad + CDbl(a(r, columns("PREV_EAD")))
        ElseIf SafeText(a(r, columns(ROW_SOURCE_COL))) = "LL" Then
            nLl = nLl + 1: llCollateral = llCollateral + CDbl(a(r, columns("TOTAL_COLLATERAL")))
            llCashflow = llCashflow + CDbl(a(r, columns("CASHFLOW_AMOUNT_LCY_PRE_FACTOR")))
            If SafeText(a(r, columns("JOIN_STATUS"))) <> "Matched identical ECL dimensions" Then Err.Raise vbObjectError + 816, , "Financial ownership regression: different ECL amounts caused a dimensional conflict."
            If CDbl(a(r, columns("ECL_MATCH_COUNT"))) <> 2 Then Err.Raise vbObjectError + 817, , "Financial ownership regression: duplicate ECL dimension membership was lost."
            If SafeText(a(r, columns("COLLATERAL_TYPE_CODE"))) <> "CASH" Or SafeText(a(r, columns("INTEREST_RATE_CODE"))) <> "FIXED" Then Err.Raise vbObjectError + 818, , "Financial ownership regression: true dimensions were not enriched."
            For Each field In Array("PREV_EAD", "GUARANTEED_LOANS", "TOTAL_INTEREST_ADVANCE", "EFFECTIVE_PD", "OUTSTANDING_LCY")
                If Len(SafeText(a(r, columns(CStr(field))))) > 0 Then Err.Raise vbObjectError + 819, , "Financial ownership regression: copied ECL measure onto liquidity fact: " & CStr(field)
            Next field
        End If
    Next r
    If nEcl <> 2 Or nLl <> 2 Then Err.Raise vbObjectError + 820, , "Financial ownership regression: original source row counts changed."
    If eclCollateral <> 3000 Or eclEad <> 30 Or llCollateral <> 11 Or llCashflow <> 30 Then Err.Raise vbObjectError + 821, , "Financial ownership regression: source-owned financial measures changed."
    Set sourceRows = ws.Range(ws.Cells(2, columns(ROW_SOURCE_COL)), ws.Cells(UBound(a, 1), columns(ROW_SOURCE_COL)))
    Set measureRows = ws.Range(ws.Cells(2, columns("TOTAL_COLLATERAL")), ws.Cells(UBound(a, 1), columns("TOTAL_COLLATERAL")))
    If Application.WorksheetFunction.SumIfs(measureRows, sourceRows, "ECL") <> 3000 Then Err.Raise vbObjectError + 822, , "Financial ownership regression: native ECL collateral total failed."
    If Application.WorksheetFunction.SumIfs(measureRows, sourceRows, "LL") <> 11 Then Err.Raise vbObjectError + 823, , "Financial ownership regression: native liquidity collateral total failed."
    Set measureRows = ws.Range(ws.Cells(2, columns("PREV_EAD")), ws.Cells(UBound(a, 1), columns("PREV_EAD")))
    If Application.WorksheetFunction.SumIfs(measureRows, sourceRows, "ECL") <> 30 Or Application.WorksheetFunction.SumIfs(measureRows, sourceRows, "LL") <> 0 Then Err.Raise vbObjectError + 824, , "Financial ownership regression: native EAD ownership total failed."
    Set ws = wb.Worksheets("Source_Controls")
    For r = 2 To ws.Cells(ws.rows.count, 1).End(xlUp).row
        If SafeText(ws.Cells(r, 1).Value2) = "ECL" And SafeText(ws.Cells(r, 2).Value2) = "OUTSTANDING_LCY" Then
            If CDbl(ws.Cells(r, 3).Value2) <> 300 Or CDbl(ws.Cells(r, 4).Value2) <> 300 Then Err.Raise vbObjectError + 825, , "Financial ownership regression: native ECL conservation control failed."
            controlsFound = controlsFound + 1
        ElseIf SafeText(ws.Cells(r, 1).Value2) = "LL" And SafeText(ws.Cells(r, 2).Value2) = "CASHFLOW_AMOUNT_LCY_PRE_FACTOR" Then
            If CDbl(ws.Cells(r, 3).Value2) <> 30 Or CDbl(ws.Cells(r, 4).Value2) <> 30 Then Err.Raise vbObjectError + 826, , "Financial ownership regression: native liquidity conservation control failed."
            controlsFound = controlsFound + 1
        End If
    Next r
    If controlsFound <> 2 Then Err.Raise vbObjectError + 827, , "Financial ownership regression: supported source-owned measure controls are missing."
    wb.Close SaveChanges:=False
    ReconJoinedFinancialOwnershipRegressionTests = "PASS: financial differences do not create dimension conflicts; ECL amounts stay on ECL; liquidity owns its cashflow and collateral; native controls reconcile."
End Function

' Full record-count fixture for the reporting projection, separate from source
' calculations. It creates 319 raw columns and 50 independent membership flags.
Public Function ReconJoinedCacheVolumeRegressionTests(ByVal folder As String) As String
    Const records As Long = 599281
    Dim wb As Workbook, ws As Worksheet, heads As Object, cases As Collection, pc As PivotCache, pt As PivotTable
    Dim names As Variant, buf() As Variant, flags() As Variant, i As Long, c As Long, r As Long, n As Long, first As Long, rowNum As Long
    Dim llRows As Long, eclRows As Long, caseRows As Long, spec As Variant, projection As Worksheet, total As Double, field As PivotField
    Dim state As Object, failure As String, result As String, fileNo As Integer, started As Double
    On Error GoTo Failed
    Set state = CaptureState(): Application.ScreenUpdating = False: Application.DisplayAlerts = False
    started = Timer
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.name = JOINED_SHEET
    Set heads = NewMap(): Set cases = New Collection
    names = Array("ROW_SOURCE", "SOURCE_ROW_NUMBER", "ROW_COUNT", "JOIN_STATUS", "ECL_MATCH_COUNT", "JOIN_KEY_FIELDS", "RECON_STAGE", "BASE_ELIGIBLE", "RECON_BANK", _
                  "ACCOUNT_NUMBER", "AS_OF_DATE", "BANK_CODE", "BRANCH_CODE", "STAGE_ID", "COA_CODE", "COA_NAME", "BUCKET_DISPLAY_NAME", "ALM_FRAMEWORK_NAME", _
                  "CASHFLOW_AMOUNT_LCY_PRE_FACTOR", "CASHFLOW_AMOUNT_LCY_POST_FACTOR", "OUTSTANDING_LCY", "ECL", "ORI_MOODY_STRESS_ECL", "OVERRIDDEN_ECL", "CALCULATED_ECL", _
                  "RWA_LCY", "OUTSTANDING_FOR_RWA", "CAP_COMPONENT_CODE", "REPORT_BALANCE")
    For c = 0 To UBound(names): heads(CStr(names(c))) = c + 1: ws.Cells(1, c + 1).Value2 = names(c): Next c
    For c = UBound(names) + 2 To 269
        heads("UNUSED_DETAIL_" & format$(c, "000")) = c: ws.Cells(1, c).Value2 = "UNUSED_DETAIL_" & format$(c, "000")
        ws.Cells(2, c).Value2 = "Full raw detail remains available"
    Next c
    For i = 1 To 50
        ws.Cells(1, 269 + i).Value2 = "TC_VOLUME_" & format$(i, "000")
        cases.Add Array("TC_VOLUME_" & format$(i, "000"), "VOLUME", "E" & i, records \ 2, 269 + i, 0, "ROW_SOURCE = 'LL'", "2025-12-31", "Calculated", "Case_" & format$(i, "000"))
    Next i
    first = 2
    Do While first <= records + 1
        n = WorksheetFunction.Min(CHUNK, records + 2 - first)
        ReDim buf(1 To n, 1 To UBound(names) + 1): ReDim flags(1 To n, 1 To 50)
        For r = 1 To n
            rowNum = first + r - 2
            If rowNum Mod 2 = 0 Then
                buf(r, 1) = "LL": llRows = llRows + 1: buf(r, 19) = 10#: buf(r, 20) = 8#: caseRows = caseRows + 1
            Else
                buf(r, 1) = "ECL": eclRows = eclRows + 1: buf(r, 21) = 20#: buf(r, 22) = 2#: buf(r, 23) = 3#: buf(r, 24) = 4#: buf(r, 25) = 5#
            End If
            buf(r, 2) = rowNum: buf(r, 3) = 1: buf(r, 4) = "Original source row": buf(r, 5) = 1: buf(r, 6) = "ACCOUNT_NUMBER|AS_OF_DATE"
            buf(r, 7) = "1": buf(r, 8) = "Yes": buf(r, 9) = "JKB": buf(r, 10) = "ACC" & rowNum
            buf(r, 11) = CDbl(DateSerial(2025, 12, 31)): buf(r, 12) = "JKB": buf(r, 13) = "1": buf(r, 14) = 1
            buf(r, 15) = "C1": buf(r, 16) = "Loans": buf(r, 17) = "0-30 days": buf(r, 18) = "LL"
            For i = 1 To 50: flags(r, i) = IIf(buf(r, 1) = "LL", "Yes", "No"): Next i
        Next r
        ws.Cells(first, 1).Resize(n, UBound(names) + 1).Value2 = buf
        ws.Cells(first, 270).Resize(n, 50).Value2 = flags
        first = first + n
        ProgressDetail "Cache volume fixture: wrote " & format$(first - 2, "#,##0") & " of " & format$(records, "#,##0") & " rows"
    Loop
    Erase buf: Erase flags
    Set pc = CreateJoinedPivotCache(wb, ws, heads, records, cases)
    Set projection = wb.Worksheets("Pivot_Source")
    If projection.Cells(1, projection.columns.count).End(xlToLeft).Column >= 110 Then Err.Raise vbObjectError + 835, , "Volume regression: unused raw fields entered the compact cache."
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = "Volume_Pivot"
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A8"), TableName:="VolumeProof")
    If pc.recordCount <> records Then Err.Raise vbObjectError + 836, , "Volume regression: cache row count changed."
    pt.PivotFields(ROW_SOURCE_COL).orientation = xlPageField
    pt.PivotFields("TC_VOLUME_050").orientation = xlPageField
    pt.AddDataField pt.PivotFields("ROW_COUNT"), "Source records", xlSum
    pt.AddDataField pt.PivotFields("CASHFLOW_AMOUNT_LCY_PRE_FACTOR"), "Own cashflow", xlSum
    pt.AddDataField pt.PivotFields("OUTSTANDING_LCY"), "Own ECL outstanding", xlSum
    pt.AddDataField pt.PivotFields("ORI_MOODY_STRESS_ECL"), "Dynamic ECL bridge measure", xlSum
    If CDbl(pt.GetPivotData("Source records").Value2) <> records Then Err.Raise vbObjectError + 837, , "Volume regression: native row total changed."
    pt.PivotFields(ROW_SOURCE_COL).CurrentPage = "LL": pt.PivotFields("TC_VOLUME_050").CurrentPage = "Yes"
    If CDbl(pt.GetPivotData("Source records").Value2) <> llRows Or CDbl(pt.GetPivotData("Own cashflow").Value2) <> llRows * 10# Then Err.Raise vbObjectError + 838, , "Volume regression: filtered liquidity scope changed."
    pt.PivotFields("TC_VOLUME_050").ClearAllFilters: pt.PivotFields(ROW_SOURCE_COL).CurrentPage = "ECL"
    If CDbl(pt.GetPivotData("Own ECL outstanding").Value2) <> eclRows * 20# Or CDbl(pt.GetPivotData("Dynamic ECL bridge measure").Value2) <> eclRows * 3# Then Err.Raise vbObjectError + 839, , "Volume regression: source-owned dynamic ECL total changed."
    result = "PASS: " & records & " records, 319 raw columns, " & projection.Cells(1, projection.columns.count).End(xlToLeft).Column & _
             " projected fields, 50 membership flags; LL rows=" & llRows & "; cashflow=" & llRows * 10# & "; ECL rows=" & eclRows & _
             "; outstanding=" & eclRows * 20# & "; dynamic ECL=" & eclRows * 3# & "; elapsed seconds=" & format$((Timer - started + 86400) Mod 86400, "0")
    fileNo = FreeFile: Open JoinPath(folder, "native_volume_results.txt") For Output As #fileNo: Print #fileNo, result: Close #fileNo
    ' Saving the large fixture is unnecessary: native cache, native scope and
    ' exact numerical assertions are recorded in the small result artifact.
    wb.Close SaveChanges:=False: Set wb = Nothing: RestoreState state
    ReconJoinedCacheVolumeRegressionTests = result
    Exit Function
Failed:
    failure = Err.description
    On Error Resume Next
    If fileNo > 0 Then Close #fileNo
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    RestoreState state
    On Error GoTo 0
    Err.Raise vbObjectError + 840, "ReconJoinedCacheVolumeRegressionTests", failure
End Function

Public Function ReconJoinedCacheRegressionTests(ByVal folder As String) As String
    Dim result As String, wb As Workbook, ws As Worksheet, i As Long, path As String, pending As String, fileNo As Integer
    Dim oldAlerts As Boolean, failure As String
    On Error GoTo Failed
    fileNo = FreeFile
    Open JoinPath(folder, "native_cache_suite_results.txt") For Output As #fileNo
    result = ReconJoinedFinancialOwnershipRegressionTests(folder): Print #fileNo, result
    ' Exercise the recovery helper using the same saved raw evidence, with all
    ' finished pivot artifacts removed from the disposable checkpoint.
    oldAlerts = Application.DisplayAlerts: Application.DisplayAlerts = False
    path = JoinPath(folder, "Joined_Input.xlsx"): pending = JoinPath(folder, "_Joined_Input_pending.xlsx")
    Set wb = Workbooks.Open(path, ReadOnly:=False, UpdateLinks:=0)
    For i = wb.Worksheets.count To 1 Step -1
        Select Case wb.Worksheets(i).name
            Case JOINED_SHEET, "Case_Coverage", "Source_Controls"
            Case Else: wb.Worksheets(i).Delete
        End Select
    Next i
    wb.SaveAs fileName:=pending, FileFormat:=xlOpenXMLWorkbook: wb.Close SaveChanges:=False: Set wb = Nothing
    path = ResumeJoinedPivotsTo(folder)
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    If CDbl(wb.Worksheets("Base_ECL").PivotTables(1).GetPivotData("OUTSTANDING_LCY ").Value2) <> 300 Then Err.Raise vbObjectError + 841, , "Resume regression: ECL native total changed."
    If CDbl(wb.Worksheets("Base_LL").PivotTables(1).GetPivotData("CASHFLOW_AMOUNT_LCY_PRE_FACTOR ").Value2) <> 30 Then Err.Raise vbObjectError + 842, , "Resume regression: liquidity native total changed."
    If wb.Worksheets(JOINED_SHEET).Cells(wb.Worksheets(JOINED_SHEET).rows.count, 1).End(xlUp).row <> 5 Then Err.Raise vbObjectError + 843, , "Resume regression: source rows changed."
    wb.Close SaveChanges:=False: Set wb = Nothing
    Application.DisplayAlerts = oldAlerts
    Print #fileNo, "PASS: resumed completed raw checkpoint; native ECL 300, liquidity 30, four source records retained."
    result = ReconJoinedCapitalRegressionTests(folder): Print #fileNo, result
    result = ReconJoinedCacheVolumeRegressionTests(folder): Print #fileNo, result
    Close #fileNo: fileNo = 0
    ReconJoinedCacheRegressionTests = "PASS: compact cache ownership, checkpoint recovery, capital separation and full record-count fixtures."
    Exit Function
Failed:
    failure = Err.description
    On Error Resume Next
    If fileNo > 0 Then Print #fileNo, "ERROR: " & failure: Close #fileNo
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    Err.Raise vbObjectError + 844, "ReconJoinedCacheRegressionTests", failure
End Function

' Small native equivalence fixture: compare a retained pre-batching joined
' report with the same source fixture, then exercise the deferred caller API.
Public Function ReconJoinedPivotBatchRegressionTests(ByVal folder As String) As String
    Dim beforeBook As Workbook, afterBook As Workbook, ws As Worksheet, other As Worksheet, pt As PivotTable
    Dim state As Object, result As String, failure As String, compared As Long, pivotCount As Long, fileNo As Integer
    Dim baseline As String, path As String, beforeRows As Long, afterRows As Long, builtHere As Boolean
    On Error GoTo Failed
    Set state = CaptureState(): Application.ScreenUpdating = False: Application.DisplayAlerts = False
    baseline = JoinPath(folder, "unbatched_baseline.xlsx")
    fileNo = FreeFile: Open JoinPath(folder, "native_batch_results.txt") For Output As #fileNo
    If Len(Dir$(baseline)) = 0 Then
        ' No retained pre-batching report in this folder, and nothing in the tool
        ' writes one. So the baseline is built here, from the same fixtures, with
        ' every pivot updated as soon as it is made: the unbatched path.
        baseline = JoinPath(folder, "unbatched_baseline_built.xlsx")
        mPivotsImmediate = True
        result = ReconJoinedFinancialOwnershipRegressionTests(folder): Print #fileNo, "baseline: " & result
        result = ReconJoinedCapitalRegressionTests(folder): Print #fileNo, "baseline: " & result
        mPivotsImmediate = False
        If Len(Dir$(baseline)) > 0 Then Kill baseline
        FileCopy JoinPath(folder, "Joined_Input.xlsx"), baseline
        builtHere = True
    End If
    result = ReconJoinedFinancialOwnershipRegressionTests(folder): Print #fileNo, result
    result = ReconJoinedCapitalRegressionTests(folder): Print #fileNo, result
    path = JoinPath(folder, "Joined_Input.xlsx")
    Set beforeBook = Workbooks.Open(baseline, UpdateLinks:=0, ReadOnly:=True)
    Set afterBook = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    beforeRows = beforeBook.Worksheets(JOINED_SHEET).Cells(beforeBook.Worksheets(JOINED_SHEET).rows.count, 1).End(xlUp).row - 1
    afterRows = afterBook.Worksheets(JOINED_SHEET).Cells(afterBook.Worksheets(JOINED_SHEET).rows.count, 1).End(xlUp).row - 1
    If beforeRows <> 7 Or afterRows <> beforeRows Then Err.Raise vbObjectError + 847, , "Batch regression: source row count changed."
    For Each ws In beforeBook.Worksheets
        If ws.PivotTables.count > 0 Then
            Set other = afterBook.Worksheets(ws.name)
            If other.PivotTables.count <> ws.PivotTables.count Then Err.Raise vbObjectError + 848, , "Batch regression: pivot count changed on " & ws.name
            For pivotCount = 1 To ws.PivotTables.count
                If PivotBatchSignature(ws.PivotTables(pivotCount)) <> PivotBatchSignature(other.PivotTables(pivotCount)) Then Err.Raise vbObjectError + 849, , "Batch regression: native values, fields, filters or layout changed on " & ws.name
                If other.PivotTables(pivotCount).ManualUpdate Then Err.Raise vbObjectError + 850, , "Batch regression: a saved report remained deferred."
                compared = compared + 1
            Next pivotCount
        End If
    Next ws
    Print #fileNo, "PASS: " & compared & " native pivots match the unbatched fixture in values, page filters, fields, subtotals and grand totals; 7 source rows retained."
    beforeBook.Close SaveChanges:=False: Set beforeBook = Nothing
    afterBook.Close SaveChanges:=False: Set afterBook = Nothing
    result = CheckDeferredPivotCaller(): Print #fileNo, result
    Close #fileNo: fileNo = 0: RestoreState state
    ReconJoinedPivotBatchRegressionTests = "PASS: unbatched/native equivalence (" & _
        IIf(builtHere, "baseline built in this run", "retained baseline") & ") and deferred update caller contract."
    Exit Function
Failed:
    failure = Err.description
    mPivotsImmediate = False
    On Error Resume Next
    If fileNo > 0 Then Print #fileNo, "ERROR: " & failure: Close #fileNo
    If Not beforeBook Is Nothing Then beforeBook.Close SaveChanges:=False
    If Not afterBook Is Nothing Then afterBook.Close SaveChanges:=False
    RestoreState state
    On Error GoTo 0
    Err.Raise vbObjectError + 851, "ReconJoinedPivotBatchRegressionTests", failure
End Function

Private Function PivotBatchSignature(ByVal pt As PivotTable) As String
    Dim pf As PivotField, a As Variant, r As Long, c As Long, i As Long, s As String
    s = CStr(pt.RowGrand) & "|" & CStr(pt.ColumnGrand)
    For Each pf In pt.PageFields
        s = s & "|PAGE:" & pf.sourceName & ":" & pf.position & ":" & CStr(pf.CurrentPage) & ":" & CStr(pf.EnableMultiplePageItems)
    Next pf
    For Each pf In pt.rowFields
        s = s & "|ROW:" & pf.sourceName & ":" & pf.position
        For i = 1 To 12: s = s & ":" & CStr(pf.Subtotals(i)): Next i
    Next pf
    For Each pf In pt.ColumnFields: s = s & "|COL:" & pf.name & ":" & pf.position: Next pf
    For Each pf In pt.dataFields: s = s & "|DATA:" & pf.sourceName & ":" & pf.name & ":" & pf.Function & ":" & pf.NumberFormat: Next pf
    a = pt.TableRange2.Value2
    If IsArray(a) Then
        s = s & "|SHAPE:" & UBound(a, 1) & "x" & UBound(a, 2)
        For r = 1 To UBound(a, 1)
            For c = 1 To UBound(a, 2): s = s & "|" & Len(SafeText(a(r, c))) & ":" & SafeText(a(r, c)): Next c
        Next r
    Else
        s = s & "|CELL:" & SafeText(a)
    End If
    PivotBatchSignature = s
End Function

Private Function CheckDeferredPivotCaller() As String
    Dim wb As Workbook, ws As Worksheet, src As Range, pc As PivotCache, immediate As PivotTable, deferred As PivotTable, refused As PivotTable
    Dim a(1 To 8, 1 To 7) As Variant, names As Variant, c As Long, r As Long, failure As String
    On Error GoTo Failed
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1): ws.name = "Facts"
    names = Array("ROW_SOURCE", "BASE_ELIGIBLE", "TC_BATCH_TEST", "RECON_STAGE", "ROW_COUNT", "AMOUNT", "BRANCH_CODE")
    For c = 1 To 7: a(1, c) = names(c - 1): Next c
    For r = 2 To 8
        a(r, 1) = "LL": a(r, 2) = "Yes": a(r, 3) = "Yes": a(r, 4) = "1": a(r, 5) = 1: a(r, 7) = "1"
    Next r
    a(2, 6) = 10: a(3, 6) = 20
    a(4, 3) = "No": a(4, 4) = "2": a(4, 6) = 30
    a(5, 1) = "ECL": a(5, 6) = 100
    a(6, 2) = "No": a(6, 6) = 999
    a(7, 4) = "2": a(7, 6) = 40
    a(8, 1) = "CAP": a(8, 3) = "No": a(8, 6) = 1000
    Set src = ws.Range("A1:G8"): src.Value2 = a
    Set pc = wb.PivotCaches.Create(xlDatabase, src.address(ReferenceStyle:=xlR1C1, External:=True))
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = "Immediate"
    Set immediate = modShared_Pivot.AddPivotFrom(pc, ws.Range("A8"), "Immediate", "RECON_STAGE", "", _
        "ROW_SOURCE|BASE_ELIGIBLE|TC_BATCH_TEST|BRANCH_CODE", "ROW_COUNT>Records>0|AMOUNT>Amount>#,##0")
    If immediate Is Nothing Then Err.Raise vbObjectError + 852, , "Batch regression: default caller returned no pivot."
    If immediate.ManualUpdate Then Err.Raise vbObjectError + 853, , "Batch regression: default caller was left deferred."
    immediate.PivotFields("ROW_SOURCE").CurrentPage = "LL"
    immediate.PivotFields("BASE_ELIGIBLE").CurrentPage = "Yes"
    immediate.PivotFields("TC_BATCH_TEST").CurrentPage = "Yes"
    modShared_Pivot.SubtotalOuterOnly immediate: ApplyJoinedPivotSafety immediate, "LL"
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count)): ws.name = "Deferred"
    Set deferred = modShared_Pivot.AddPivotFrom(pc, ws.Range("A8"), "Deferred", "RECON_STAGE", "", _
        "ROW_SOURCE|BASE_ELIGIBLE|TC_BATCH_TEST|BRANCH_CODE", "ROW_COUNT>Records>0|AMOUNT>Amount>#,##0", True)
    If deferred Is Nothing Then Err.Raise vbObjectError + 854, , "Batch regression: deferred caller returned no pivot."
    If Not deferred.ManualUpdate Then Err.Raise vbObjectError + 855, , "Batch regression: deferred caller materialized prematurely."
    deferred.PivotFields("ROW_SOURCE").CurrentPage = "LL"
    deferred.PivotFields("BASE_ELIGIBLE").CurrentPage = "Yes"
    deferred.PivotFields("TC_BATCH_TEST").CurrentPage = "Yes"
    modShared_Pivot.SubtotalOuterOnly deferred: ApplyJoinedPivotSafety deferred, "LL"
    deferred.ManualUpdate = False
    If PivotBatchSignature(immediate) <> PivotBatchSignature(deferred) Then Err.Raise vbObjectError + 856, , "Batch regression: deferred scope or values differ from the immediate caller."
    If CDbl(deferred.GetPivotData("Records").Value2) <> 3 Or CDbl(deferred.GetPivotData("Amount").Value2) <> 70 Then Err.Raise vbObjectError + 857, , "Batch regression: native filters selected the wrong records or amount."
    If CDbl(deferred.GetPivotData("Amount", "RECON_STAGE", "1").Value2) <> 30 Or CDbl(deferred.GetPivotData("Amount", "RECON_STAGE", "2").Value2) <> 40 Then Err.Raise vbObjectError + 858, , "Batch regression: stage values changed."
    Set refused = modShared_Pivot.AddPivotFrom(Nothing, ws.Range("J8"), "Refused", "", "", "", "")
    If Not refused Is Nothing Then Err.Raise vbObjectError + 859, , "Batch regression: invalid-cache failure contract changed."
    wb.Close SaveChanges:=False: Set wb = Nothing
    CheckDeferredPivotCaller = "PASS: default returns updated; deferred stays pending until scoped; both select 3 records / amount 70 (stage1 30, stage2 40); failure still returns Nothing."
    Exit Function
Failed:
    failure = Err.description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise vbObjectError + 860, "CheckDeferredPivotCaller", failure
End Function

' A row the filter selects: read from the engine's row mask when it has one for
' the active table, otherwise evaluated as before.
Private Function RowPasses(ByRef mask As Variant, ByVal node As Object, ByVal r As Long) As Boolean
    If IsArray(mask) Then
        RowPasses = (mask(r) <> 0)
    Else
        RowPasses = modScenarioBuilder_PreShock.EvalForActive(node, r)
    End If
End Function
