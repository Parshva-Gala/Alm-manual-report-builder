Option Explicit

' ============================================================================
'  Staging: the one pass that every pivot in a framework's workbook sits on.
'
'  The source file is read in blocks of forty thousand rows, only the thirteen
'  columns the pivots use are kept, and each block is written straight into the
'  staging table before the next is read. Memory stays flat on a half-million
'  row extract, and the whole workbook ends up with ONE PivotCache - so a
'  refresh is one refresh and a slicer can drive every sheet at once.
'
'  Two columns are computed here rather than pivoted on later:
'
'    LCY / FCY   whether the row's own currency is the local one. There is no
'                such column in the extract; it is the local currency compared
'                against CURRENCY_NAME, and the local currency is whichever
'                carries the most gross money. That is a guess, so it is
'                WRITTEN DOWN on the guide sheet rather than assumed silently.
'    Factor      post divided by pre. The extract has ALM_FACTOR_PCT but its
'                units are not stated anywhere - 50 and 0.5 are both defensible
'                readings of "half" - and a factor that disagrees with the two
'                amounts beside it is worse than no factor.
' ============================================================================

' What a staging pass found out, handed back to the caller.
Private mAsOf As String
Private mLocalCcy As String
Private mPreField As String
Private mPostField As String
Private mNative As Boolean
Private mRows As Double
Private mRuleNames As Object          ' rule name -> gross, for deciding sheet order
Private mCurrencies As Object         ' currency  -> gross

' Fields a Pivot config recipe names beyond the thirteen built in, staged after
' them. Kind: 0 text, 1 number, 2 date, 3 a copy of a built-in column.
Private mXCount As Long
Private mXName() As String
Private mXKind() As Long
Private mXSrc() As Long
Private mXBlank() As String
Private mXFormat() As String
Private mXFrom() As Long
Private mMissing As String

' "One sheet per" families: signature (field names joined by Chr(30)) ->
' Dictionary(combination -> gross), measured in the staging pass so the
' biggest sheets can be built first without reading the data twice.
Private mSplitCols As Object          ' signature -> Array of stage columns
Private mSplitWeights As Object       ' signature -> Dictionary

Public Function StagedRows() As Double
    StagedRows = mRows
End Function
Public Function StagedAsOf() As String
    StagedAsOf = mAsOf
End Function
Public Function LocalCurrency() As String
    LocalCurrency = mLocalCcy
End Function
Public Function AmountFieldNote() As String
    If mNative Then
        AmountFieldNote = mPreField & "  /  " & mPostField
    Else
        AmountFieldNote = mPreField & "  /  " & mPostField & "   (no native-currency field found)"
    End If
End Function
Public Function UsedNativeAmounts() As Boolean
    UsedNativeAmounts = mNative
End Function
Public Function RuleNames() As Object
    If mRuleNames Is Nothing Then Set mRuleNames = NewMap()
    Set RuleNames = mRuleNames
End Function
Public Function Currencies() As Object
    If mCurrencies Is Nothing Then Set mCurrencies = NewMap()
    Set Currencies = mCurrencies
End Function

' The staged column a field landed in: the thirteen built-ins first, in their
' fixed order, then the extras. 0 if it was not staged.
Public Function StageCol(ByVal fieldName As String) As Long
    Dim h As Variant, i As Long
    h = StageHeadings()
    For i = 0 To UBound(h)
        If StrComp(CStr(h(i)), fieldName, vbTextCompare) = 0 Then StageCol = i + 1: Exit Function
    Next i
    For i = 1 To mXCount
        If StrComp(mXName(i), fieldName, vbTextCompare) = 0 Then StageCol = C_COLS + i: Exit Function
    Next i
End Function

Public Function SplitWeightsFor(ByVal sig As String) As Object
    If Not mSplitWeights Is Nothing Then
        If mSplitWeights.Exists(sig) Then Set SplitWeightsFor = mSplitWeights(sig): Exit Function
    End If
    Set SplitWeightsFor = NewMap()
End Function

' Extra fields whose source column this file does not have. They are staged
' blank, and the guide says so.
Public Function MissingFields() As String
    MissingFields = mMissing
End Function

' ============================================================================
'  Read one framework's output into a staging table in the target workbook.
'  Returns the ListObject, or Nothing.
' ============================================================================
Public Function StageFramework(ByVal fw As String, ByVal dstWb As Workbook, ByRef errOut As String, _
                               Optional ByVal extras As Collection = Nothing, _
                               Optional ByVal splits As Collection = Nothing) As ListObject
    Dim key As String, Path As String, src As Workbook, ws As Worksheet, hdr As Long
    Dim lastR As Long, lastC As Long, hdrVals As Variant, fieldMap As Object
    Dim cols() As Long, nCols As Long, runStart() As Long, runEnd() As Long, nRuns As Long
    Dim r As Long, n As Long, i As Long, blk As Variant, buf() As Variant
    Dim stage As Worksheet, outRow As Long, wrote As Double
    Dim ix As Object, pass1 As Object

    key = "OUTPUT|" & fw
    Path = SlotFile(key)
    If Len(Path) = 0 Then errOut = "nothing loaded for " & FwLabel(fw): Exit Function

    ResetPass

    On Error GoTo Failed
    Set src = Workbooks.Open(Path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    Set ws = modPD_Files.SheetOfSlot(src, key, hdr)
    If ws Is Nothing Then errOut = "the sheet recorded for that file is not in it any more": GoTo CloseFail

    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastC = ws.Cells(hdr, ws.Columns.count).End(xlToLeft).Column
    If lastC > 300 Then lastC = 300
    If lastR <= hdr Then errOut = "the sheet has a header and no rows": GoTo CloseFail

    hdrVals = ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, lastC)).Value2
    Set fieldMap = HeaderMap(hdrVals, lastC)
    ResolveAmountFields fieldMap

    Set ix = ColumnIndex(fieldMap)
    If CLng(ix("Pre")) = 0 Then errOut = "no pre-factor amount column in that sheet": GoTo CloseFail
    PlanExtras extras, fieldMap, ix
    PlanSplits splits

    BuildWanted ix, cols, nCols
    BuildRuns cols, nCols, runStart, runEnd, nRuns

    ' --- pass one: what currency is local -----------------------------------
    '
    ' Needed BEFORE any row is written, because the LCY/FCY column depends on
    ' it. A cheap scan of one column beats staging twice.
    Set pass1 = ScanCurrencies(ws, hdr, lastR, CLng(ix("Ccy")))
    mLocalCcy = HeaviestKey(pass1)
    Set mCurrencies = pass1

    ' --- the staging table ---------------------------------------------------
    Set stage = NewStageSheet(dstWb)
    WriteStageHeader stage
    outRow = 2

    r = hdr + 1
    Do While r <= lastR
        n = lastR - r + 1
        If n > CHUNK_ROWS Then n = CHUNK_ROWS
        ReDim buf(1 To n, 1 To lastC)
        For i = 1 To nRuns
            blk = ws.Range(ws.Cells(r, runStart(i)), ws.Cells(r + n - 1, runEnd(i))).Value2
            CopyRun blk, buf, n, runStart(i), runEnd(i)
        Next i
        wrote = wrote + EmitBlock(buf, n, ix, stage, outRow)
        r = r + n
        Progress_ FwLabel(fw) & " - staged " & Fmt(wrote) & " of " & Fmt(CDbl(lastR - hdr)) & " rows", _
                  wrote / (lastR - hdr)
    Loop

    src.Close SaveChanges:=False
    Set src = Nothing
    mRows = wrote

    If wrote = 0 Then errOut = "every row in that file was empty": Exit Function
    Set StageFramework = MakeTable(stage, outRow - 1)
    modPD_Files.NoteRows key, wrote, mAsOf
    modPD_Files.NoteAmountField key, AmountFieldNote()
    Exit Function

CloseFail:
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Err.Clear
    Exit Function
Failed:
    errOut = Err.Number & " " & Err.Description
    On Error Resume Next
    If Not src Is Nothing Then src.Close SaveChanges:=False
    Err.Clear
End Function

Private Sub ResetPass()
    mAsOf = "": mLocalCcy = "": mRows = 0: mNative = False
    mPreField = F_PRE_LCY: mPostField = F_POST_LCY
    Set mRuleNames = NewMap()
    Set mCurrencies = NewMap()
    mXCount = 0
    mMissing = ""
    Set mSplitCols = NewMap()
    Set mSplitWeights = NewMap()
End Sub

' Where each extra field comes from in this file, resolved against the header
' the same way the built-ins are: case and punctuation do not matter.
Private Sub PlanExtras(ByVal extras As Collection, ByVal h As Object, ByVal ix As Object)
    Dim i As Long, f As Object, src As String
    If extras Is Nothing Then Exit Sub
    mXCount = extras.count
    If mXCount = 0 Then Exit Sub
    ReDim mXName(1 To mXCount)
    ReDim mXKind(1 To mXCount)
    ReDim mXSrc(1 To mXCount)
    ReDim mXBlank(1 To mXCount)
    ReDim mXFormat(1 To mXCount)
    ReDim mXFrom(1 To mXCount)
    For i = 1 To mXCount
        Set f = extras(i)
        mXName(i) = CStr(f("Name"))
        mXBlank(i) = CStr(f("Blank"))
        mXFormat(i) = CStr(f("Format"))
        src = UCase$(CStr(f("Source")))
        Select Case src
            Case modPD_Config.SRC_PRE: mXKind(i) = 3: mXFrom(i) = C_PRE
            Case modPD_Config.SRC_POST: mXKind(i) = 3: mXFrom(i) = C_POST
            Case modPD_Config.SRC_CCYCLASS: mXKind(i) = 3: mXFrom(i) = C_CCYCLASS
            Case modPD_Config.SRC_FACTOR: mXKind(i) = 3: mXFrom(i) = C_FACTOR
            Case Else
                Select Case CStr(f("Kind"))
                    Case "Number": mXKind(i) = 1
                    Case "Date": mXKind(i) = 2
                    Case Else: mXKind(i) = 0
                End Select
                mXSrc(i) = At(h, src)
                If mXSrc(i) = 0 Then
                    mMissing = mMissing & IIf(Len(mMissing) > 0, ", ", "") & mXName(i) & " (" & src & ")"
                Else
                    ix("X" & i) = mXSrc(i)
                End If
        End Select
    Next i
End Sub

Private Sub PlanSplits(ByVal splits As Collection)
    Dim sig As Variant, parts As Variant, cols() As Long, i As Long, ok As Boolean
    If splits Is Nothing Then Exit Sub
    For Each sig In splits
        If Not mSplitCols.Exists(CStr(sig)) Then
            parts = Split(CStr(sig), Chr$(30))
            ReDim cols(0 To UBound(parts))
            ok = True
            For i = 0 To UBound(parts)
                cols(i) = StageCol(CStr(parts(i)))
                If cols(i) = 0 Then ok = False
            Next i
            If ok Then
                mSplitCols(CStr(sig)) = cols
                Set mSplitWeights(CStr(sig)) = NewMap()
            End If
        End If
    Next sig
End Sub

Private Function HeaderMap(ByRef hdrVals As Variant, ByVal lastC As Long) As Object
    Dim d As Object, c As Long, v As String
    Set d = NewMap()
    Set HeaderMap = d
    For c = 1 To lastC
        v = NormKey(SafeText(hdrVals(1, c)))
        If Len(v) > 0 Then If Not d.Exists(v) Then d(v) = c
    Next c
End Function

' ============================================================================
'  Which amount columns this file actually has.
'
'  LCY is in every extract seen. The native-currency pair may be spelled CCY or
'  ACY or something else; each candidate is tried against the header in turn and
'  whichever exists wins. If none does, LCY is used AND SAID SO - on the Files
'  sheet, on the guide, and in the log - because a ladder labelled native
'  currency while carrying converted amounts is the kind of wrong that gets
'  submitted.
' ============================================================================
Private Sub ResolveAmountFields(ByVal h As Object)
    Dim k As Variant
    mPreField = F_PRE_LCY
    mPostField = F_POST_LCY
    mNative = False
    For Each k In PreNativeSpellings()
        If h.Exists(NormKey(CStr(k))) Then mPreField = CStr(k): mNative = True: Exit For
    Next k
    If Not mNative Then Exit Sub
    ' A pre without a post is not a pair; fall back rather than mix bases.
    mPostField = ""
    For Each k In PostNativeSpellings()
        If h.Exists(NormKey(CStr(k))) Then mPostField = CStr(k): Exit For
    Next k
    If Len(mPostField) = 0 Then
        mPreField = F_PRE_LCY
        mPostField = F_POST_LCY
        mNative = False
    End If
End Sub

Private Function ColumnIndex(ByVal h As Object) As Object
    Dim d As Object
    Set d = NewMap()
    Set ColumnIndex = d
    d("Order") = At(h, F_RULE_ORDER)
    d("Cat") = At(h, F_RULE_CAT)
    d("Rule") = At(h, F_RULE_NAME)
    d("Type") = At(h, F_TYPE)
    d("Line") = At(h, F_LINE)
    d("Sub") = At(h, F_SUBLINE)
    d("Coa") = At(h, F_COA_NAME)
    d("Ccy") = At(h, F_CURRENCY)
    d("Bucket") = At(h, F_BUCKET)
    d("Pre") = At(h, mPreField)
    d("Post") = At(h, mPostField)
    d("AsOf") = At(h, F_AS_OF)
End Function

Private Function At(ByVal h As Object, ByVal nm As String) As Long
    Dim k As String
    k = NormKey(nm)
    If h.Exists(k) Then At = CLng(h(k))
End Function

Private Sub BuildWanted(ByVal ix As Object, ByRef cols() As Long, ByRef nCols As Long)
    Dim k As Variant, c As Long
    ReDim cols(1 To ix.count + 1)
    nCols = 0
    For Each k In ix.keys
        c = CLng(ix(k))
        If c > 0 Then
            nCols = nCols + 1
            cols(nCols) = c
        End If
    Next k
End Sub

' Wanted columns grouped into contiguous runs, bridging gaps of up to three.
' On these files that turns a dozen scattered columns into three or four block
' reads per chunk instead of a dozen.
Private Sub BuildRuns(ByRef cols() As Long, ByVal nCols As Long, _
                      ByRef runStart() As Long, ByRef runEnd() As Long, ByRef nRuns As Long)
    Const GAP As Long = 3
    Dim srt() As Long, i As Long, j As Long, t As Long
    nRuns = 0
    If nCols <= 0 Then Exit Sub
    ReDim srt(1 To nCols)
    For i = 1 To nCols
        srt(i) = cols(i)
    Next i
    For i = 1 To nCols - 1
        For j = 1 To nCols - i
            If srt(j) > srt(j + 1) Then t = srt(j): srt(j) = srt(j + 1): srt(j + 1) = t
        Next j
    Next i
    ReDim runStart(1 To nCols)
    ReDim runEnd(1 To nCols)
    nRuns = 1
    runStart(1) = srt(1): runEnd(1) = srt(1)
    For i = 2 To nCols
        If srt(i) - runEnd(nRuns) <= GAP + 1 Then
            runEnd(nRuns) = srt(i)
        Else
            nRuns = nRuns + 1
            runStart(nRuns) = srt(i): runEnd(nRuns) = srt(i)
        End If
    Next i
End Sub

Private Sub CopyRun(ByRef blk As Variant, ByRef buf() As Variant, ByVal n As Long, _
                    ByVal c1 As Long, ByVal c2 As Long)
    Dim rr As Long, cc As Long
    If n = 1 And c1 = c2 Then
        buf(1, c1) = blk
    ElseIf n = 1 Then
        For cc = c1 To c2
            buf(1, cc) = blk(1, cc - c1 + 1)
        Next cc
    ElseIf c1 = c2 Then
        For rr = 1 To n
            buf(rr, c1) = blk(rr, 1)
        Next rr
    Else
        For rr = 1 To n
            For cc = c1 To c2
                buf(rr, cc) = blk(rr, cc - c1 + 1)
            Next cc
        Next rr
    End If
End Sub

' One block of source rows turned into one block of staged rows and written in a
' single assignment. Everything in here runs once per row on a half-million-row
' file, so it holds no lookup it can resolve outside the loop.
Private Function EmitBlock(ByRef buf As Variant, ByVal n As Long, ByVal ix As Object, _
                           ByVal stage As Worksheet, ByRef outRow As Long) As Double
    Dim out() As Variant, i As Long, k As Long, pre As Double, post As Double
    Dim ccy As String, rule As String, local_ As String
    Dim iOrder As Long, iCat As Long, iRule As Long, iType As Long, iLine As Long
    Dim iSub As Long, iCoa As Long, iCcy As Long, iBkt As Long, iPre As Long, iPost As Long, iAsOf As Long

    iOrder = CLng(ix("Order")): iCat = CLng(ix("Cat")): iRule = CLng(ix("Rule"))
    iType = CLng(ix("Type")): iLine = CLng(ix("Line")): iSub = CLng(ix("Sub"))
    iCoa = CLng(ix("Coa")): iCcy = CLng(ix("Ccy")): iBkt = CLng(ix("Bucket"))
    iPre = CLng(ix("Pre")): iPost = CLng(ix("Post")): iAsOf = CLng(ix("AsOf"))
    local_ = mLocalCcy

    ReDim out(1 To n, 1 To C_COLS + mXCount)
    For i = 1 To n
        pre = Amt(buf, i, iPre)
        post = Amt(buf, i, iPost)
        ccy = Txt(buf, i, iCcy)
        rule = Txt(buf, i, iRule)

        k = k + 1
        out(k, C_RULE_ORDER) = Amt(buf, i, iOrder)
        out(k, C_RULE_CAT) = Blank(Txt(buf, i, iCat), "(no category)")
        out(k, C_RULE_NAME) = Blank(rule, "(no rule)")
        ' post/pre, not the source's factor column: it agrees with the two
        ' amounts beside it by construction, whatever units that field is in.
        '
        ' Written as TEXT - "100%", "50%". The factor is a grouping label in the
        ' row area, and a pivot row field ignores NumberFormat: that applies to
        ' the field's data area, not to its item captions, so a numeric 1 was
        ' printing as "1" where it meant 100%. As text it reads right and groups
        ' right.
        If Abs(pre) > TOLERANCE Then
            out(k, C_FACTOR) = Format$(post / pre, "0%")
        Else
            out(k, C_FACTOR) = Empty
        End If
        out(k, C_TYPE) = Blank(Txt(buf, i, iType), "(no type)")
        out(k, C_LINE) = Blank(Txt(buf, i, iLine), "(no line)")
        out(k, C_SUBLINE) = Blank(Txt(buf, i, iSub), "(no subline)")
        out(k, C_COA_NAME) = Blank(Txt(buf, i, iCoa), "(no COA)")
        out(k, C_CURRENCY) = Blank(ccy, "(no currency)")
        out(k, C_CCYCLASS) = IIf(StrComp(ccy, local_, vbTextCompare) = 0, "LCY", "FCY")
        out(k, C_BUCKET) = Blank(Txt(buf, i, iBkt), "(no bucket)")
        out(k, C_PRE) = pre
        out(k, C_POST) = post

        If mXCount > 0 Then EmitExtras buf, i, out, k
        If mSplitCols.count > 0 Then Weigh out, k, pre
        If Len(rule) > 0 Then mRuleNames(rule) = SafeNum(mRuleNames(rule)) + Abs(pre)
        ' AsOfText, not Txt: the block was read with .Value2, which hands a date
        ' over as its serial number, and "45991" is not an as-of date.
        If Len(mAsOf) = 0 And iAsOf > 0 Then mAsOf = AsOfText(buf(i, iAsOf))
    Next i

    If k = 0 Then Exit Function
    stage.Range(stage.Cells(outRow, 1), stage.Cells(outRow + k - 1, C_COLS + mXCount)).Value2 = out
    outRow = outRow + k
    EmitBlock = k
End Function

' The extra fields of one row. Text keeps its blank label, a number stays a
' number, a date becomes a date whatever shape it arrived in.
Private Sub EmitExtras(ByRef buf As Variant, ByVal i As Long, ByRef out() As Variant, ByVal k As Long)
    Dim x As Long, v As Variant, t As String
    For x = 1 To mXCount
        Select Case mXKind(x)
            Case 3
                out(k, C_COLS + x) = out(k, mXFrom(x))
            Case 1
                If mXSrc(x) > 0 Then v = buf(i, mXSrc(x)) Else v = Empty
                If IsNumeric(v) And Not IsEmpty(v) Then out(k, C_COLS + x) = CDbl(v) Else out(k, C_COLS + x) = Empty
            Case 2
                If mXSrc(x) > 0 Then out(k, C_COLS + x) = DateValueOf(buf(i, mXSrc(x))) Else out(k, C_COLS + x) = Empty
            Case Else
                If mXSrc(x) > 0 Then t = Txt(buf, i, mXSrc(x)) Else t = ""
                If Len(t) = 0 Then
                    If Len(mXBlank(x)) > 0 Then out(k, C_COLS + x) = mXBlank(x) Else out(k, C_COLS + x) = Empty
                Else
                    out(k, C_COLS + x) = t
                End If
        End Select
    Next x
End Sub

Private Function DateValueOf(ByVal v As Variant) As Variant
    On Error Resume Next
    DateValueOf = Empty
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then DateValueOf = CDbl(v): Exit Function
    If IsNumeric(v) Then DateValueOf = CDbl(v): Exit Function
    If IsDate(v) Then DateValueOf = CDbl(CDate(v))
    Err.Clear
End Function

' Adds this row's gross to the combination it belongs to, for every "one
' sheet per" family being built.
Private Sub Weigh(ByRef out() As Variant, ByVal k As Long, ByVal pre As Double)
    Dim sig As Variant, cols As Variant, j As Long, key As String, d As Object
    For Each sig In mSplitCols.keys
        cols = mSplitCols(sig)
        key = CStr(out(k, cols(0)))
        For j = 1 To UBound(cols)
            key = key & Chr$(30) & CStr(out(k, cols(j)))
        Next j
        Set d = mSplitWeights(sig)
        d(key) = SafeNum(d(key)) + Abs(pre)
    Next sig
End Sub

Private Function Txt(ByRef buf As Variant, ByVal i As Long, ByVal c As Long) As String
    If c = 0 Then Exit Function
    Txt = SafeText(buf(i, c))
End Function

Private Function Amt(ByRef buf As Variant, ByVal i As Long, ByVal c As Long) As Double
    If c = 0 Then Exit Function
    Amt = SafeNum(buf(i, c))
End Function

Private Function Blank(ByVal s As String, ByVal whenEmpty As String) As String
    If Len(s) = 0 Then Blank = whenEmpty Else Blank = s
End Function

' ===================== which currency is local ==============================

Private Function ScanCurrencies(ByVal ws As Worksheet, ByVal hdr As Long, ByVal lastR As Long, _
                                ByVal col As Long) As Object
    Dim d As Object, r As Long, n As Long, blk As Variant, i As Long, v As String
    Set d = NewMap()
    Set ScanCurrencies = d
    If col = 0 Then Exit Function
    r = hdr + 1
    Do While r <= lastR
        n = lastR - r + 1
        If n > CHUNK_ROWS Then n = CHUNK_ROWS
        blk = ws.Range(ws.Cells(r, col), ws.Cells(r + n - 1, col)).Value2
        If n = 1 Then
            v = SafeText(blk)
            If Len(v) > 0 Then d(v) = SafeNum(d(v)) + 1
        Else
            For i = 1 To n
                v = SafeText(blk(i, 1))
                If Len(v) > 0 Then d(v) = SafeNum(d(v)) + 1
            Next i
        End If
        r = r + n
    Loop
End Function

Private Function HeaviestKey(ByVal d As Object) As String
    Dim k As Variant, best As Double, v As Double
    For Each k In d.keys
        v = SafeNum(d(k))
        If v > best Then best = v: HeaviestKey = CStr(k)
    Next k
End Function

' ===================== the staging sheet ====================================

Private Function NewStageSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    KillSheet SH_STAGE, wb
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    On Error Resume Next
    ws.Name = SH_STAGE
    Err.Clear
    Set NewStageSheet = ws
End Function

Private Sub WriteStageHeader(ByVal ws As Worksheet)
    Dim h As Variant, c As Long
    h = StageHeadings()
    For c = 0 To UBound(h)
        ws.Cells(1, c + 1).Value2 = h(c)
    Next c
    For c = 1 To mXCount
        ws.Cells(1, C_COLS + c).Value2 = mXName(c)
        ' Text columns are Text BEFORE the data lands, for the reason given
        ' for the factor below: a branch code "0020" must not become 20.
        Select Case mXKind(c)
            Case 0: ws.Columns(C_COLS + c).NumberFormat = "@"
            Case 2: ws.Columns(C_COLS + c).NumberFormat = IIf(Len(mXFormat(c)) > 0, mXFormat(c), "d mmm yyyy")
            Case 1: If Len(mXFormat(c)) > 0 Then ws.Columns(C_COLS + c).NumberFormat = mXFormat(c)
        End Select
    Next c
    ' The factor column is written as text - "100%", "50%" - and the column has
    ' to be Text BEFORE the first block lands. Assigning the string "100%" to a
    ' General cell makes Excel parse it back into the number 1 with a percent
    ' format, and the pivot then shows the underlying 1 as its row label.
    ws.Columns(C_FACTOR).NumberFormat = "@"
End Sub

Private Function MakeTable(ByVal ws As Worksheet, ByVal LastRow As Long) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, C_COLS + mXCount)), , xlYes)
    If lo Is Nothing Then Exit Function
    lo.Name = "tbl_data"
    lo.TableStyle = "TableStyleLight1"
    ws.Range(ws.Cells(2, C_PRE), ws.Cells(LastRow, C_POST)).NumberFormat = NUM_FMT
    ' The factor column is text ("100%"), so it carries no number format.
    ' The staging sheet is the engine room, not a deliverable. Very hidden so it
    ' cannot be unhidden by accident and edited under the pivots.
    ws.visible = xlSheetVeryHidden
    Err.Clear
    Set MakeTable = lo
End Function
