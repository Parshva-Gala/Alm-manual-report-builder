Option Explicit

' ============================================================================
'  The outputs against control reports 3 and 6.
'
'  SCOPE FIRST. A control report covering three COAs cannot disagree with six
'  hundred and sixty-four others; it simply does not carry them. Measured
'  against this bank's own sample control reports, the naive comparison produced
'  eleven thousand "breaks" of which all but sixty-six were the control report
'  not being the ledger. So the overlap is computed before anything else:
'
'    on both sides, differing   a break, ranked by gross
'    only in the control        always a break - the books carry a balance the
'                               engine classified nowhere
'    only in the output         SCOPE, not error. Counted and totalled, never
'                               printed as rows.
'    nothing on both sides      the two files are keyed differently. Said once,
'                               with an example from each, and nothing printed.
'
'  And the net is never reported without the gross. Contra pairs cancel, so a
'  net of nothing over a gross of eight billion is not agreement.
' ============================================================================

Private Const N_CTRL As Long = 1
Private Const N_FW As Long = 2
Private Const N_KEY As Long = 3
Private Const N_OUT As Long = 4
Private Const N_CTL As Long = 5
Private Const N_DIFF As Long = 6
Private Const N_VERDICT As Long = 7
Private Const N_NOTE As Long = 8
Private Const N_COLS As Long = 8

Private Const MAX_ROWS As Long = 400

Private mOut As Long

Public Sub BuildReconSheet()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_RECON)
    modPD_Theme.Dress ws, "Reconciliation", _
        "Each loaded output against control report 3 (the ledger, by COA) and control report 6 " & _
        "(the reporting balance, by account)."
    modPD_Theme.Head ws, Array("Control", "Framework", "Key", "Output", "Control", "Difference", _
                               "Verdict", "What it means"), _
                        Array(12, 18, 26, 18, 18, 18, 13, 74)
    Clear ws
    modPD_Theme.SetStatus ws, "Nothing reconciled yet.", "Idle"
End Sub

Private Sub Clear(ByVal ws As Worksheet)
    Dim lastR As Long
    lastR = LastRow(ws, N_KEY)
    If lastR >= modPD_Theme.R_FIRST Then ws.Range(ws.Rows(modPD_Theme.R_FIRST), ws.Rows(lastR)).Delete Shift:=xlUp
    mOut = modPD_Theme.R_FIRST
End Sub

Public Sub PD_Reconcile()
    Dim st As Object, ws As Worksheet, ctl3 As Object, ctl6 As Object
    Dim fw As Variant, summaries As Collection, ran As Boolean, t0 As Single

    Dim fromDesk As Boolean
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    If Not modPD_Files.AnyFrameworkLoaded() Then
        Notify "No framework output has been loaded yet - add an LCR, NSFR or maturity ladder output first.", V_CHECK
        Exit Sub
    End If
    If Not modPD_Files.SlotLoaded("CTRL3|") And Not modPD_Files.SlotLoaded("CTRL6|") Then
        Notify "Neither control report has been loaded. Add control report 3 or 6 - without one there is " & _
               "nothing to reconcile against, and that is not a clean reconciliation, it is no reconciliation.", V_CHECK
        Exit Sub
    End If

    On Error GoTo Failed
    fromDesk = modPD_Desk.DeskInFront()
    modPD_Desk.BusyOn
    Set st = CaptureState(): PD_Busy = True
    t0 = Timer
    BuildReconSheet
    Set ws = GetSheet(SH_RECON)
    Set summaries = New Collection

    If modPD_Files.SlotLoaded("CTRL3|") Then
        Step_ "reading control report 3"
        Set ctl3 = LoadControl("CTRL3|", F_COA_CODE, F_CTRL3_AMT)
        SettingSet "ctl3_keys", CStr(ctl3.count)
        LogIt IIf(ctl3.count > 0, V_OK, V_CHECK), "Recon", _
              "Control 3: " & Fmt(CDbl(ctl3.count)) & " COA(s).", ""
    End If
    If modPD_Files.SlotLoaded("CTRL6|") Then
        Step_ "reading control report 6"
        Set ctl6 = LoadControl("CTRL6|", F_ACCOUNT, F_CTRL6_AMT)
        SettingSet "ctl6_keys", CStr(ctl6.count)
        LogIt IIf(ctl6.count > 0, V_OK, V_CHECK), "Recon", _
              "Control 6: " & Fmt(CDbl(ctl6.count)) & " account(s).", ""
    End If

    ' ONE pass per framework, summing by COA and by account together.
    '
    ' It used to open the output once per control report - six full opens of
    ' 57-79MB files in one Excel process for three frameworks - and Excel ran
    ' out of memory on the fourth. Both keys come off the same read.
    Dim byCoaMap As Object, byAcctMap As Object
    For Each fw In Frameworks()
        If modPD_Files.SlotLoaded("OUTPUT|" & CStr(fw)) Then
            Step_ FwLabel(CStr(fw)) & " - reading once for both controls"
            SumBoth CStr(fw), Not (ctl3 Is Nothing), Not (ctl6 Is Nothing), byCoaMap, byAcctMap
            If Not ctl3 Is Nothing Then
                If ctl3.count > 0 Then
                    summaries.Add One("Control 3", CStr(fw), byCoaMap, ctl3)
                    ran = True
                End If
            End If
            If Not ctl6 Is Nothing Then
                If ctl6.count > 0 Then
                    summaries.Add One("Control 6", CStr(fw), byAcctMap, ctl6)
                    ran = True
                End If
            End If
            ' Let a half-million-key dictionary go before the next framework's.
            Set byCoaMap = Nothing
            Set byAcctMap = Nothing
        End If
    Next fw

    Finish ws, summaries, ran, Timer - t0
    RestoreState st: PD_Busy = False
    modPD_Desk.RefreshDesk
    modPD_Desk.BusyOff
    ' Started from the Desk, the answer is shown there - the matrix, the
    ' verdict and a toast - and the detail is one click away. Started from
    ' the sheet, the sheet is where the reader already is.
    If fromDesk Then
        modPD_Theme.GoTo_ SH_HOME
        Notify SafeText(ws.Cells(modPD_Theme.R_STATUS, 1).Value2), SettingGet("recon_level", V_OK)
    Else
        modPD_Theme.GoTo_ SH_RECON
    End If
    Exit Sub

Failed:
    Step_ "RECONCILE FAILED: " & Err.Number & " " & Err.Description
    RestoreState st: PD_Busy = False
    modPD_Desk.BusyOff
    LogIt V_BREAK, "Recon", Err.Number & " " & Err.Description, ""
    Tell "The reconciliation stopped:" & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub

' ===================== one control against one output =======================

Private Function One(ByVal ctrlName As String, ByVal fw As String, ByVal outMap As Object, _
                     ByVal ctlMap As Object) As Object
    Dim s As Object, k As Variant, key As String, o As Double, c As Double, d As Double
    Dim rowsK() As String, rowsD() As Double, n As Long, i As Long, cap As Long
    Dim keyD() As Double, payload() As Variant

    Set s = NewSummary(ctrlName, fw)
    Set One = s
    If outMap Is Nothing Or ctlMap Is Nothing Then Exit Function
    s("CtlKeys") = CDbl(ctlMap.count)
    s("OutKeys") = CDbl(outMap.count)

    ReDim rowsK(0 To ctlMap.count + 1)
    ReDim rowsD(0 To ctlMap.count + 1)

    For Each k In ctlMap.keys
        key = CStr(k)
        c = SafeNum(ctlMap(k))
        If outMap.Exists(key) Then
            o = SafeNum(outMap(key))
            s("Both") = CDbl(s("Both")) + 1
            s("CoveredGross") = CDbl(s("CoveredGross")) + Abs(o)
            d = o - c
            If Abs(d) <= TOLERANCE Then
                s("Agree") = CDbl(s("Agree")) + 1
            Else
                s("Differ") = CDbl(s("Differ")) + 1
                s("Net") = CDbl(s("Net")) + d
                s("Gross") = CDbl(s("Gross")) + Abs(d)
                If Abs(d) > Abs(CDbl(s("Worst"))) Then s("Worst") = d: s("WorstKey") = key
                rowsK(n) = key: rowsD(n) = d: n = n + 1
            End If
        ElseIf Abs(c) > TOLERANCE Then
            s("OnlyCtl") = CDbl(s("OnlyCtl")) + 1
            s("OnlyCtlAmt") = CDbl(s("OnlyCtlAmt")) + c
            If Len(CStr(s("CtlSample"))) = 0 Then s("CtlSample") = key
            rowsK(n) = key: rowsD(n) = -c: n = n + 1
        Else
            ' Nothing against nothing is agreement, and a red row worth zero is
            ' how a reconciliation acquires a hundred findings that mean nothing.
            s("ZeroCtl") = CDbl(s("ZeroCtl")) + 1
            If Len(CStr(s("CtlSample"))) = 0 Then s("CtlSample") = key
        End If
    Next k

    For Each k In outMap.keys
        key = CStr(k)
        If Len(key) > 0 Then
            If Len(CStr(s("OutSample"))) = 0 Then s("OutSample") = key
            o = SafeNum(outMap(key))
            s("OutGross") = CDbl(s("OutGross")) + Abs(o)
            If Not ctlMap.Exists(key) Then
                s("OnlyOut") = CDbl(s("OnlyOut")) + 1
                s("OnlyOutAmt") = CDbl(s("OnlyOutAmt")) + o
            End If
        End If
    Next k

    ' Coverage on GROSS. Against a netted total a framework whose assets and
    ' liabilities offset would report a negative percentage covered, which is
    ' not a thing a share of a population can be.
    If CDbl(s("OutGross")) > TOLERANCE Then _
        s("Covered") = CDbl(s("CoveredGross")) / CDbl(s("OutGross"))

    If CDbl(s("Both")) = 0 Then
        s("Unkeyed") = True
        s("Verdict") = V_BREAK
        Exit Function
    End If

    If n > 0 Then
        ReDim keyD(0 To n - 1)
        ReDim payload(0 To n - 1)
        For i = 0 To n - 1
            keyD(i) = Abs(rowsD(i)): payload(i) = rowsK(i)
        Next i
        SortDesc keyD, payload, 0, n - 1
        cap = n
        If cap > MAX_ROWS Then cap = MAX_ROWS: s("Capped") = True
        For i = 0 To cap - 1
            key = CStr(payload(i))
            o = 0: c = 0
            If outMap.Exists(key) Then o = SafeNum(outMap(key))
            If ctlMap.Exists(key) Then c = SafeNum(ctlMap(key))
            WriteRow ctrlName, fw, key, o, c, outMap.Exists(key)
        Next i
        s("Printed") = cap
        s("Differing") = n
    End If

    s("Verdict") = VerdictFor(s)
End Function

Private Function NewSummary(ByVal ctrlName As String, ByVal fw As String) As Object
    Dim s As Object
    Set s = NewMap()
    s("Control") = ctrlName: s("Framework") = fw
    s("CtlKeys") = 0#: s("OutKeys") = 0#: s("Both") = 0#
    s("Agree") = 0#: s("Differ") = 0#: s("Differing") = 0#
    s("OnlyOut") = 0#: s("OnlyCtl") = 0#: s("ZeroCtl") = 0#
    s("OnlyOutAmt") = 0#: s("OnlyCtlAmt") = 0#
    s("CoveredGross") = 0#: s("OutGross") = 0#: s("Covered") = 0#
    s("Net") = 0#: s("Gross") = 0#: s("Worst") = 0#
    s("WorstKey") = "": s("CtlSample") = "": s("OutSample") = ""
    s("Printed") = 0#: s("Capped") = False: s("Unkeyed") = False
    s("Verdict") = V_NONE
    Set NewSummary = s
End Function

' The verdict is about the keys the two files SHARE. Keys outside the overlap
' are scope, reported separately, so a narrow control cannot turn a clean
' reconciliation red.
Private Function VerdictFor(ByVal s As Object) As String
    If CDbl(s("CtlKeys")) = 0 Then VerdictFor = V_NONE: Exit Function
    If CBool(s("Unkeyed")) Then VerdictFor = V_BREAK: Exit Function
    If CDbl(s("OnlyCtl")) > 0 Then VerdictFor = V_BREAK: Exit Function
    If CDbl(s("Differ")) = 0 Then VerdictFor = V_OK: Exit Function
    If Abs(CDbl(s("Gross"))) > TOLERANCE * 1000 Then VerdictFor = V_BREAK: Exit Function
    VerdictFor = V_CHECK
End Function

Private Sub WriteRow(ByVal ctrlName As String, ByVal fw As String, ByVal key As String, _
                     ByVal o As Double, ByVal c As Double, ByVal inOut As Boolean)
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(SH_RECON)
    If ws Is Nothing Then Exit Sub
    If mOut < modPD_Theme.R_FIRST Then mOut = modPD_Theme.R_FIRST
    r = mOut
    ws.Cells(r, N_CTRL).Value2 = ctrlName
    ws.Cells(r, N_FW).Value2 = FwLabel(fw)
    ws.Cells(r, N_KEY).Value2 = key
    ws.Cells(r, N_OUT).Value2 = o
    ws.Cells(r, N_CTL).Value2 = c
    ws.Cells(r, N_DIFF).Value2 = o - c
    If inOut Then
        ws.Cells(r, N_VERDICT).Value2 = V_CHECK
        ws.Cells(r, N_NOTE).Value2 = "Both sides carry this key and the amounts differ."
    Else
        ws.Cells(r, N_VERDICT).Value2 = V_BREAK
        ws.Cells(r, N_NOTE).Value2 = "The control carries this and the output classified nothing " & _
            "against it. The whole amount is unclassified."
    End If
    mOut = mOut + 1
End Sub

' ===================== the summary at the top ===============================

Private Sub Finish(ByVal ws As Worksheet, ByVal summaries As Collection, ByVal ran As Boolean, _
                   ByVal secs As Double)
    Dim lastR As Long, i As Long, s As Object, nBreak As Long, worst As Double, msg As String

    lastR = mOut - 1
    If lastR >= modPD_Theme.R_FIRST Then
        ws.Range(ws.Cells(modPD_Theme.R_FIRST, N_OUT), ws.Cells(lastR, N_DIFF)).NumberFormat = NUM_FMT
        modPD_Theme.DressTable ws, N_COLS, lastR, N_VERDICT
        ws.Range(ws.Cells(modPD_Theme.R_FIRST, N_KEY), ws.Cells(lastR, N_KEY)).Font.Name = modPD_Theme.UI_MONO
        ws.Range(ws.Cells(modPD_Theme.R_FIRST, N_NOTE), ws.Cells(lastR, N_NOTE)).Font.Color = modPD_Theme.C_MUTED
        DiffBars ws.Range(ws.Cells(modPD_Theme.R_FIRST, N_DIFF), ws.Cells(lastR, N_DIFF))
    End If

    For i = 1 To summaries.count
        Set s = summaries(i)
        If CStr(s("Verdict")) = V_BREAK Then
            nBreak = nBreak + 1
            If CDbl(s("Gross")) > worst Then worst = CDbl(s("Gross"))
            If CDbl(s("OnlyCtlAmt")) > worst Then worst = Abs(CDbl(s("OnlyCtlAmt")))
        End If
        LogIt CStr(s("Verdict")), "Recon", Describe(s), CStr(s("Control")) & " vs " & FwLabel(CStr(s("Framework")))
    Next i

    modPD_Desk.NoteRecon summaries, ran, nBreak, worst
    If Not ran Then
        msg = "Nothing was reconciled - no control report produced any keys."
        modPD_Theme.SetStatus ws, msg, "Idle"
        SettingSet "recon_level", "Idle"
    ElseIf nBreak > 0 Then
        msg = nBreak & " of " & summaries.count & " comparison(s) broke" & _
              IIf(worst > 0, ", the largest involving " & Fmt(worst) & " LCY", "") & ".   " & _
              "Scope is on Activity for each - read it before treating a difference as an error.   " & _
              "(" & Format$(secs, "0.0") & "s)"
        modPD_Theme.SetStatus ws, msg, "Break"
        SettingSet "recon_level", V_BREAK
    Else
        msg = "Every shared key agrees, across " & summaries.count & " comparison(s).   (" & _
              Format$(secs, "0.0") & "s)"
        modPD_Theme.SetStatus ws, msg, "OK"
        SettingSet "recon_level", V_OK
    End If
End Sub

' In-cell bars on the difference, so the size of each break reads at a glance
' down the column - red either side of an axis, because a difference is
' signed and the sign matters.
Private Sub DiffBars(ByVal rng As Range)
    Dim db As Object
    On Error Resume Next
    rng.FormatConditions.Delete
    Set db = rng.FormatConditions.AddDatabar
    If db Is Nothing Then Exit Sub
    db.BarColor.Color = modPD_Theme.HX("F4A29A")
    db.BarFillType = 1                      ' xlDataBarFillSolid
    db.ShowValue = True
    db.AxisPosition = 0                     ' xlDataBarAxisAutomatic
    db.NegativeBarFormat.ColorType = 0      ' xlDataBarColor
    db.NegativeBarFormat.Color.Color = modPD_Theme.HX("F4A29A")
    db.AxisColor.Color = modPD_Theme.HX("B42318")
    Err.Clear
End Sub

Private Function Describe(ByVal s As Object) As String
    Dim t As String
    If CBool(s("Unkeyed")) Then
        Describe = "NOT THE SAME IDENTIFIER - the control is keyed like " & Chr$(34) & _
            CStr(s("CtlSample")) & Chr$(34) & " and the output like " & Chr$(34) & _
            CStr(s("OutSample")) & Chr$(34) & ". Nothing was compared; every difference " & _
            "would be an artefact of the mismatch."
        Exit Function
    End If
    t = Fmt(CDbl(s("Both"))) & " key(s) on both sides, covering " & _
        Format$(CDbl(s("Covered")), "0.00%") & " of the output's balances. "
    If CDbl(s("OnlyOut")) > 0 Then
        t = t & "The control does not carry " & Fmt(CDbl(s("OnlyOut"))) & " key(s) the output has - " & _
            "unverified rather than wrong. "
    End If
    If CDbl(s("OnlyCtl")) > 0 Then
        t = t & Fmt(CDbl(s("OnlyCtl"))) & " key(s) worth " & Fmt(CDbl(s("OnlyCtlAmt"))) & _
            " are in the control and were never classified - an error whatever the scope. "
    End If
    If CDbl(s("Differ")) = 0 Then
        t = t & "Every shared key agrees."
    Else
        t = t & Fmt(CDbl(s("Differ"))) & " shared key(s) differ, gross " & Fmt(CDbl(s("Gross"))) & _
            ", net " & Fmt(CDbl(s("Net")))
        If Abs(CDbl(s("Net"))) * 10 < Abs(CDbl(s("Gross"))) Then _
            t = t & " - the net is small because the differences offset, which is not agreement"
        t = t & "."
    End If
    If CBool(s("Capped")) Then
        t = t & "  Only the worst " & Fmt(CDbl(s("Printed"))) & " of " & _
            Fmt(CDbl(s("Differing"))) & " are listed; the figures here are over all of them."
    End If
    Describe = t
End Function

' ===================== reading the sources ==================================

' One framework's output, read ONCE, summed by COA code and by account number in
' the same pass. Streamed in blocks: nothing needs the whole file in memory.
'
' The account map is only built when a control 6 is loaded - on a half-million
' row extract it is half a million dictionary inserts to answer a question
' nobody asked.
Private Sub SumBoth(ByVal fw As String, ByVal wantCoa As Boolean, ByVal wantAcct As Boolean, _
                    ByRef coaOut As Object, ByRef acctOut As Object)
    Dim key As String, Path As String, wb As Workbook, ws As Worksheet, hdr As Long
    Dim lastR As Long, lastC As Long, hdrVals As Variant, c As Long, v As String
    Dim cCoa As Long, cAcct As Long, cAmt As Long
    Dim r As Long, n As Long, i As Long, blk As Variant, s As String, Amt As Double
    Dim c1 As Long, c2 As Long

    Set coaOut = NewMap()
    Set acctOut = NewMap()
    key = "OUTPUT|" & fw
    Path = modPD_Files.SlotFile(key)
    If Len(Path) = 0 Then Exit Sub

    On Error GoTo Done
    Set wb = Workbooks.Open(Path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    Set ws = modPD_Files.SheetOfSlot(wb, key, hdr)
    If ws Is Nothing Then GoTo CloseDone

    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastC = ws.Cells(hdr, ws.Columns.count).End(xlToLeft).Column
    hdrVals = ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, lastC)).Value2
    For c = 1 To lastC
        v = NormKey(SafeText(hdrVals(1, c)))
        If v = NormKey(F_COA_CODE) And cCoa = 0 Then cCoa = c
        If v = NormKey(F_ACCOUNT) And cAcct = 0 Then cAcct = c
        If v = NormKey(F_PRE_LCY) And cAmt = 0 Then cAmt = c
    Next c
    If cAmt = 0 Then GoTo CloseDone
    If Not wantCoa Then cCoa = 0
    If Not wantAcct Then cAcct = 0

    ' One block spanning the three columns rather than three separate reads.
    c1 = cAmt: c2 = cAmt
    If cCoa > 0 Then
        If cCoa < c1 Then c1 = cCoa
        If cCoa > c2 Then c2 = cCoa
    End If
    If cAcct > 0 Then
        If cAcct < c1 Then c1 = cAcct
        If cAcct > c2 Then c2 = cAcct
    End If

    r = hdr + 1
    Do While r <= lastR
        n = lastR - r + 1
        If n > CHUNK_ROWS Then n = CHUNK_ROWS
        If n = 1 Then
            ' A one-row, one-column range gives .Value2 a SCALAR, not an array,
            ' and the last block of a file is one row often enough to matter.
            ' Three cell reads is nothing next to getting it wrong.
            Amt = SafeNum(ws.Cells(r, cAmt).Value2)
            If cCoa > 0 Then
                s = SafeText(ws.Cells(r, cCoa).Value2)
                If Len(s) > 0 Then coaOut(s) = SafeNum(coaOut(s)) + Amt
            End If
            If cAcct > 0 Then
                s = SafeText(ws.Cells(r, cAcct).Value2)
                If Len(s) > 0 Then acctOut(s) = SafeNum(acctOut(s)) + Amt
            End If
        Else
            blk = ws.Range(ws.Cells(r, c1), ws.Cells(r + n - 1, c2)).Value2
            For i = 1 To n
                Amt = SafeNum(blk(i, cAmt - c1 + 1))
                If cCoa > 0 Then
                    s = SafeText(blk(i, cCoa - c1 + 1))
                    If Len(s) > 0 Then coaOut(s) = SafeNum(coaOut(s)) + Amt
                End If
                If cAcct > 0 Then
                    s = SafeText(blk(i, cAcct - c1 + 1))
                    If Len(s) > 0 Then acctOut(s) = SafeNum(acctOut(s)) + Amt
                End If
            Next i
        End If
        r = r + n
        Progress_ FwLabel(fw) & " - summed " & Fmt(CDbl(r - hdr - 1)) & " of " & Fmt(CDbl(lastR - hdr)) & " rows", _
                  (r - hdr - 1) / (lastR - hdr)
    Loop

CloseDone:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Err.Clear
Done:
End Sub

Private Function LoadControl(ByVal key As String, ByVal keyField As String, ByVal amtField As String) As Object
    Dim d As Object, wb As Workbook, ws As Worksheet, hdr As Long, Path As String
    Dim lastR As Long, lastC As Long, c As Long, kc As Long, ac As Long, v As String
    Dim a As Variant, r As Long, k As String

    Set d = NewMap()
    Set LoadControl = d
    Path = modPD_Files.SlotFile(key)
    If Len(Path) = 0 Then Exit Function

    On Error GoTo Done
    Set wb = Workbooks.Open(Path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    Set ws = modPD_Files.SheetOfSlot(wb, key, hdr)
    If ws Is Nothing Then GoTo CloseDone
    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastC = ws.Cells(hdr, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastC
        v = NormKey(SafeText(ws.Cells(hdr, c).Value2))
        If v = NormKey(keyField) And kc = 0 Then kc = c
        If v = NormKey(amtField) And ac = 0 Then ac = c
    Next c
    If kc = 0 Or ac = 0 Or lastR <= hdr Then GoTo CloseDone
    a = ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(lastR, lastC)).Value2
    For r = 1 To UBound(a, 1)
        k = SafeText(a(r, kc))
        If Len(k) > 0 Then d(k) = SafeNum(d(k)) + SafeNum(a(r, ac))
    Next r

CloseDone:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Err.Clear
Done:
End Function

' Quicksort, not the obvious pair of nested loops: four hundred thousand
' accounts the slow way is not a wait anybody sits through.
Public Sub SortDesc(ByRef key() As Double, ByRef payload() As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As Double, td As Double, tv As Variant
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    p = key((lo + hi) \ 2)
    Do While i <= j
        Do While key(i) > p
            i = i + 1
        Loop
        Do While key(j) < p
            j = j - 1
        Loop
        If i <= j Then
            td = key(i): key(i) = key(j): key(j) = td
            tv = payload(i): payload(i) = payload(j): payload(j) = tv
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortDesc key, payload, lo, j
    If i < hi Then SortDesc key, payload, i, hi
End Sub
