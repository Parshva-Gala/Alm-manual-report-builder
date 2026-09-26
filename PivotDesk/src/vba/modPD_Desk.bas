Option Explicit

' ============================================================================
'  The Desk.
'
'  One sheet that is the whole application: what is loaded, what can be
'  built, what reconciles, what happened, and the one thing worth doing next.
'  It replaces the console window - everything the console did, and the
'  status the sheets used to hold, is here, on one screen, live.
'
'  The Desk is DESIGNED, not drawn. Every shape on it was laid out by
'  build/desk.py and ships in the workbook; this module never adds or deletes
'  a shape. It repaints the ones that carry state - a text, a colour, a
'  switch - by name. The names are a contract with the design, and the build
'  checks that every name used here exists in the drawing.
'
'  APP VIEW. While the Desk's workbook is in front, Excel's own chrome steps
'  aside - no ribbon, no formula bar on the Desk, no sheet tabs - and the
'  workbook reads as an application. The moment another workbook comes to the
'  front, or this one closes, the chrome is put back exactly as it was. The
'  "Excel view" button turns the whole thing off for anyone who wants the
'  ribbon while they work here.
'
'  Nothing in here may break an operation. Every routine swallows its own
'  errors: a Desk that fails to repaint is a cosmetic fault, a build that
'  stops because the Desk failed to repaint is not.
' ============================================================================

' The design's canvas - 75 columns and 44 rows of 15 pt, the 1120 x 660 pt
' drawing in build/desk.py - and a cell under the black app bar where the
' cursor can sit without drawing a selection box on the design.
Private Const FIT_RANGE As String = "A1:BW44"
Private Const PARK_CELL As String = "A1"
Private Const DESK_H As Double = 660

Private Const TOAST_SECONDS As Long = 9

' The maturity-gap chart, relative to its card (build/desk.py GAP_PAD and
' GAP_CHART; the build checks the two agree). Rows are exact points; columns
' can come out a little wider or narrower with the screen's DPI, so widths
' are scaled by how wide the card actually is.
Private Const GAP_CARD_W As Double = 436
Private Const GAP_PAD_X As Double = 20
Private Const GAP_PAD_Y As Double = 36
Private Const GAP_CW As Double = 264
Private Const GAP_CH As Double = 50
Private Const GAP_BARS As Long = 12

' The tour (build/desk.py TOUR, tour_place).
Private Const TOUR_STEPS As Long = 6
Private Const TOUR_PAD As Double = 6
Private Const TOUR_GAP As Double = 14

Private mPainting As Boolean

' ===================== shapes by name =======================================

Private Function DeskSheet() As Worksheet
    Set DeskSheet = GetSheet(SH_HOME)
End Function

Private Function Shp(ByVal ws As Worksheet, ByVal nm As String) As Shape
    On Error Resume Next
    Set Shp = ws.Shapes(nm)
    Err.Clear
End Function

' Replacing the text keeps the formatting of the first character, which is
' how every live label keeps its font. An empty string would drop to the
' default font on the next write, so a label is never set to nothing.
Private Sub SetText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String)
    Dim sh As Shape
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    If Len(txt) = 0 Then txt = " "
    If sh.TextFrame2.TextRange.Text <> txt Then sh.TextFrame2.TextRange.Text = txt
    Err.Clear
End Sub

Private Sub SetTextColor(ByVal ws As Worksheet, ByVal nm As String, ByVal clr As Long)
    Dim sh As Shape
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = clr
    Err.Clear
End Sub

Private Sub SetFill(ByVal ws As Worksheet, ByVal nm As String, ByVal clr As Long, _
                    Optional ByVal transp As Single = 0)
    Dim sh As Shape
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    sh.Fill.visible = msoTrue
    sh.Fill.Solid
    sh.Fill.ForeColor.RGB = clr
    sh.Fill.Transparency = transp
    Err.Clear
End Sub

Private Sub SetLine(ByVal ws As Worksheet, ByVal nm As String, ByVal clr As Long, _
                    Optional ByVal transp As Single = 0)
    Dim sh As Shape
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    sh.Line.visible = msoTrue
    sh.Line.ForeColor.RGB = clr
    sh.Line.Transparency = transp
    Err.Clear
End Sub

Private Sub SetVisible(ByVal ws As Worksheet, ByVal nm As String, ByVal v As Boolean)
    Dim sh As Shape
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    If v Then sh.visible = msoTrue Else sh.visible = msoFalse
    Err.Clear
End Sub

' The Desk's buttons come in three states: soft (emerald, can act), ghost
' (neutral, can act) and off (cannot act, and looks it).
Private Sub PaintButton(ByVal ws As Worksheet, ByVal nm As String, ByVal kind As String)
    Select Case kind
        Case "soft"
            SetFill ws, nm, HX("001D14")
            SetLine ws, nm, HX("006141")
            SetTextColor ws, nm, HX("C2EBDA")
        Case "ghost"
            SetFill ws, nm, HX("FFFFFF"), 0.97
            SetLine ws, nm, HX("243A31")
            SetTextColor ws, nm, HX("F2F7F4")
        Case Else
            SetFill ws, nm, HX("FFFFFF"), 0.98
            SetLine ws, nm, HX("1A2922")
            SetTextColor ws, nm, HX("4A5E55")
    End Select
End Sub

' Keeps an icon beside centred button text whose length has just changed.
Private Sub CenterIcon(ByVal ws As Worksheet, ByVal btn As String, ByVal ic As String)
    Dim b As Shape, i As Shape, tw As Double, pad As Double
    Set b = Shp(ws, btn)
    Set i = Shp(ws, ic)
    If b Is Nothing Or i Is Nothing Then Exit Sub
    On Error Resume Next
    tw = b.TextFrame2.TextRange.BoundWidth
    pad = b.TextFrame2.MarginLeft
    If tw > 0 Then i.Left = b.Left + (b.Width + pad) / 2 - tw / 2 - i.Width - 5
    Err.Clear
End Sub

' ===================== what is on the desk ==================================

' Top to bottom, the order the Desk and the Files sheet list them in.
Private Function SlotKeys() As Variant
    SlotKeys = Array("OUTPUT|" & FW_LCR, "OUTPUT|" & FW_NSFR, "OUTPUT|" & FW_ML, "CTRL3|", "CTRL6|")
End Function

Private Function SlotLabel(ByVal key As String) As String
    Select Case UCase$(key)
        Case UCase$("OUTPUT|" & FW_LCR): SlotLabel = "LCR output"
        Case UCase$("OUTPUT|" & FW_NSFR): SlotLabel = "NSFR output"
        Case UCase$("OUTPUT|" & FW_ML): SlotLabel = "Maturity ladder output"
        Case "CTRL3|": SlotLabel = "Control report 3"
        Case "CTRL6|": SlotLabel = "Control report 6"
        Case Else: SlotLabel = key
    End Select
End Function

Private Function SlotCell(ByVal key As String, ByVal col As Long) As String
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Function
    r = modPD_Files.SlotRow(key)
    If r = 0 Then Exit Function
    SlotCell = SafeText(ws.Cells(r, col).Value2)
    Err.Clear
End Function

' EMPTY, MISSING, FORCED (placed by hand against what its columns say) or
' LOADED.
Private Function SlotState(ByVal key As String) As String
    Dim p As String, found As Boolean
    p = modPD_Files.SlotFile(key)
    If Len(p) = 0 Then SlotState = "EMPTY": Exit Function
    On Error Resume Next
    found = (Len(Dir$(p)) > 0)
    Err.Clear
    On Error GoTo 0
    If Not found Then SlotState = "MISSING": Exit Function
    If InStr(1, SlotCell(key, modPD_Files.S_NOTE), "YOU CHOSE THIS FILE", vbBinaryCompare) > 0 Then
        SlotState = "FORCED"
    Else
        SlotState = "LOADED"
    End If
End Function

Private Function IsIn(ByVal state As String) As Boolean
    IsIn = (state = "LOADED" Or state = "FORCED")
End Function

Private Function SlotMeta(ByVal key As String, ByVal state As String) As String
    Dim leaf As String, rowsTxt As String
    Select Case state
        Case "EMPTY"
            SlotMeta = "Not added  " & ChrW(183) & "  click to choose"
        Case "MISSING"
            SlotMeta = "Moved or renamed  " & ChrW(183) & "  click to find it"
        Case Else
            leaf = MidTrim(FileLeaf(modPD_Files.SlotFile(key)), 30)
            rowsTxt = SlotCell(key, modPD_Files.S_ROWS)
            If state = "FORCED" Then
                SlotMeta = leaf & "  " & ChrW(183) & "  placed by hand - check"
            ElseIf Len(rowsTxt) > 0 And IsNumeric(rowsTxt) Then
                SlotMeta = leaf & "  " & ChrW(183) & "  " & Fmt(CDbl(rowsTxt)) & " rows"
            Else
                SlotMeta = leaf
            End If
    End Select
End Function

Private Function SlotMetaColor(ByVal state As String) As Long
    Select Case state
        Case "EMPTY": SlotMetaColor = HX("4A5E55")
        Case "MISSING": SlotMetaColor = HX("FF6B5E")
        Case "FORCED": SlotMetaColor = HX("F2B544")
        Case Else: SlotMetaColor = HX("7E9388")
    End Select
End Function

Private Function DotColor(ByVal state As String) As Long
    Select Case state
        Case "LOADED": DotColor = HX("2FC48D")
        Case "MISSING": DotColor = HX("FF6B5E")
        Case "FORCED": DotColor = HX("F2B544")
        Case Else: DotColor = HX("3A4B43")
    End Select
End Function

' Which outputs are on the desk, as one string - a change here means the last
' build or reconciliation no longer describes what is loaded.
Private Function OutputsSig() As String
    Dim fw As Variant, s As String
    For Each fw In Frameworks()
        s = s & modPD_Files.SlotFile("OUTPUT|" & CStr(fw)) & "|"
    Next fw
    OutputsSig = s
End Function

Public Function ReconSig() As String
    ReconSig = OutputsSig() & modPD_Files.SlotFile("CTRL3|") & "|" & modPD_Files.SlotFile("CTRL6|")
End Function

Public Function BuildSig() As String
    BuildSig = OutputsSig()
End Function

Private Function FwIndexLabel(ByVal fw As String) As String
    If StrComp(fw, FW_ML, vbTextCompare) = 0 Then FwIndexLabel = "the maturity ladder" Else FwIndexLabel = FwLabel(fw)
End Function

' "LCR", "LCR and NSFR", "LCR, NSFR and the maturity ladder"
Private Function JoinAnd(ByVal items As Collection) As String
    Dim i As Long, s As String
    For i = 1 To items.count
        If i = 1 Then
            s = CStr(items(i))
        ElseIf i = items.count Then
            s = s & " and " & CStr(items(i))
        Else
            s = s & ", " & CStr(items(i))
        End If
    Next i
    JoinAnd = s
End Function

Private Function Plural(ByVal n As Long, ByVal one As String, ByVal many As String) As String
    If n = 1 Then Plural = one Else Plural = many
End Function

' ===================== the next thing to do =================================
'
' One action is always the obvious one, and the hero says why in a sentence.
' The same function answers "what does the big button do" and "what does the
' hero say", so the two can never disagree.

Private Function NextAction(ByRef label As String, ByRef lede As String) As String
    Dim keys As Variant, i As Long, st As String, nFiles As Long, nFw As Long, nCtl As Long
    Dim missingKey As String, ready As New Collection, fw As Variant, nSel As Long
    Dim ctl As String, built As Boolean, reconciled As Boolean, nBreak As Long, nCmp As Long
    Dim worst As Double, folder As String
    On Error Resume Next      ' a question it cannot answer must not stop the Desk

    keys = SlotKeys()
    For i = 0 To UBound(keys)
        st = SlotState(CStr(keys(i)))
        If IsIn(st) Then
            nFiles = nFiles + 1
            If i <= 2 Then nFw = nFw + 1 Else nCtl = nCtl + 1
        ElseIf st = "MISSING" And Len(missingKey) = 0 Then
            missingKey = CStr(keys(i))
        End If
    Next i
    For Each fw In Frameworks()
        If IsIn(SlotState("OUTPUT|" & CStr(fw))) Then
            ready.Add FwIndexLabel(CStr(fw))
            If SettingGet("sel_" & CStr(fw), "1") = "1" Then nSel = nSel + 1
        End If
    Next fw
    If IsIn(SlotState("CTRL3|")) And IsIn(SlotState("CTRL6|")) Then
        ctl = "control reports 3 and 6 are"
    ElseIf IsIn(SlotState("CTRL3|")) Then
        ctl = "control report 3 is"
    ElseIf IsIn(SlotState("CTRL6|")) Then
        ctl = "control report 6 is"
    End If
    built = (nFw > 0 And SettingGet("build_sig") = BuildSig())
    reconciled = (SettingGet("recon_sig") = ReconSig() And SettingGet("recon_ran") = "1")
    nBreak = CLng(Val(SettingGet("recon_breaks", "0")))
    nCmp = CLng(Val(SettingGet("recon_n", "0")))
    worst = Val(SettingGet("recon_worst", "0"))
    folder = SettingGet("last_out_folder")

    If nFiles = 0 And Len(missingKey) = 0 Then
        label = "Scan a folder"
        lede = "Nothing is on the desk yet. Point PivotDesk at the folder holding your framework outputs " & _
               "and control reports 3 and 6 - it works out which file is which."
        NextAction = "SCAN"
    ElseIf Len(missingKey) > 0 Then
        label = "Find " & SlotLabel(missingKey)
        If Len(label) > 24 Then label = "Find the missing file"
        lede = SlotLabel(missingKey) & " has moved or been renamed since it was added. Find it again to " & _
               "carry on - nothing else on the desk has changed."
        NextAction = "SLOT:" & missingKey
    ElseIf nFw = 0 Then
        label = "Add an output"
        lede = "Only control reports are on the desk so far. Add an LCR, NSFR or maturity ladder output " & _
               "to build pivots and reconcile."
        NextAction = "PICK"
    ElseIf Not built And modPD_Config.Engine() = "recipes" And Val(SettingGet("config_bad", "0")) > 0 Then
        label = "Fix the pivot config"
        lede = "Some pivots on the Pivot config sheet will not build - the reason is written beside each. " & _
               "Fix them or switch them off, then build."
        NextAction = "GO_CONFIG"
    ElseIf Not built Then
        If nSel > 0 Then
            label = "Build " & nSel & Plural(nSel, " workbook", " workbooks")
            NextAction = "BUILD"
        Else
            label = "Choose frameworks"
            NextAction = "NONE"
        End If
        lede = Words(nFiles) & " of five files are on the desk. " & UCase$(Left$(JoinAnd(ready), 1)) & _
               Mid$(JoinAnd(ready), 2) & Plural(ready.count, " is", " are") & " ready to pivot"
        If Len(ctl) = 0 Then
            lede = lede & " - add control report 3 or 6 to reconcile as well."
        Else
            lede = lede & ", and " & ctl & " there to reconcile against."
        End If
    ElseIf Len(ctl) = 0 Then
        label = "Add a control report"
        lede = "The pivots are built. Add control report 3 or 6 and PivotDesk will reconcile the outputs " & _
               "against the ledger."
        NextAction = "PICK"
    ElseIf Not reconciled Then
        label = "Reconcile now"
        lede = "The pivots are built and " & ctl & " on the desk, so the outputs can be checked against " & _
               "the ledger now."
        NextAction = "RECON"
    ElseIf nBreak > 0 Then
        label = "Open the breaks"
        lede = "The last reconciliation found " & nBreak & Plural(nBreak, " break", " breaks") & " across " & _
               nCmp & Plural(nCmp, " comparison", " comparisons")
        If worst > 0 Then lede = lede & " - the largest involves " & Compact(worst) & " LCY"
        lede = lede & ". Scope is on Activity; read it before calling a difference an error."
        NextAction = "GO_RECON"
    Else
        label = "Open the workbooks"
        lede = "Everything on the desk is built and every shared key reconciles."
        If Len(folder) > 0 Then lede = lede & " The workbooks are in " & MidTrim(folder, 60) & "."
        NextAction = "OPEN_FOLDER"
    End If
End Function

Public Sub PD_NextAction()
    Dim label As String, lede As String, act As String
    If PD_Busy Then Exit Sub
    PressFx
    ToastHide
    act = NextAction(label, lede)
    Select Case True
        Case act = "SCAN": modPD_Files.PD_LoadFolder
        Case act = "PICK": modPD_Files.PD_LoadFiles
        Case Left$(act, 5) = "SLOT:": modPD_Files.UseFileFor Mid$(act, 6)
        Case act = "BUILD": PD_BuildSelected
        Case act = "RECON": modPD_Recon.PD_Reconcile
        Case act = "GO_RECON": modPD_Theme.PD_GoRecon
        Case act = "OPEN_FOLDER": PD_OpenOutputFolder
        Case act = "GO_CONFIG": modPD_Theme.PD_GoConfig
        Case Else: Toast "Switch on at least one framework under Build pivots.", "CHECK"
    End Select
End Sub

' ===================== painting =============================================

Public Sub RefreshDesk()
    Dim ws As Worksheet, su As Boolean
    If mPainting Then Exit Sub
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    mPainting = True
    On Error Resume Next
    su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ws.Unprotect

    PaintHero ws
    PaintFiles ws
    PaintBuild ws
    PaintRecon ws
    PaintActivity ws
    PaintGap ws
    SetVisible ws, "pdx_busy", False
    SetVisible ws, "pdx_macros", False
    SetText ws, "pdx_btn_view", IIf(AppView(), "Excel view", "App view")
    SetText ws, "pdx_foot", TOOL_NAME & " " & TOOL_VERSION & "  " & ChrW(183) & _
        "  Every workbook it writes is live PivotTables on one cache " & ChrW(8212) & _
        " drag a field, add a slicer, double-click a total."

    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Application.ScreenUpdating = su
    mPainting = False
    Err.Clear
End Sub

Private Sub PaintHero(ByVal ws As Worksheet)
    Dim label As String, lede As String, h As Long, g As String
    Dim keys As Variant, i As Long, nFiles As Long, nFw As Long, nCtl As Long, st As String
    Dim ran As Boolean, nBreak As Long, nCmp As Long, asOf As String, fw As Variant
    On Error Resume Next      ' one label that will not paint must not stop the rest

    h = Hour(Now)
    If h < 12 Then
        g = "Good morning."
    ElseIf h < 17 Then
        g = "Good afternoon."
    Else
        g = "Good evening."
    End If
    SetText ws, "pdx_hero_greet", g
    SetText ws, "pdx_hero_date", UCase$(Format$(Date, "dddd")) & "  " & ChrW(183) & "  " & _
                                 UCase$(Format$(Date, "d mmmm yyyy"))
    NextAction label, lede
    SetText ws, "pdx_hero_lede", lede
    SetText ws, "pdx_cta", label
    CenterIcon ws, "pdx_cta", "pdx_cta_ic"

    keys = SlotKeys()
    For i = 0 To UBound(keys)
        st = SlotState(CStr(keys(i)))
        If IsIn(st) Then
            nFiles = nFiles + 1
            If i <= 2 Then nFw = nFw + 1 Else nCtl = nCtl + 1
        End If
    Next i
    Kpi ws, "files", nFiles, 5, "of 5 on the desk"
    Kpi ws, "fw", nFw, 3, "of 3 ready to pivot"
    Kpi ws, "ctl", nCtl, 2, "of 2 control reports"

    ran = (SettingGet("recon_ran") = "1")
    nBreak = CLng(Val(SettingGet("recon_breaks", "0")))
    nCmp = CLng(Val(SettingGet("recon_n", "0")))
    If Not ran Then
        SetText ws, "pdx_kpi_recon_value", ChrW(8212)
        SetTextColor ws, "pdx_kpi_recon_value", HX("F2F7F4")
        SetText ws, "pdx_kpi_recon_sub", "not run yet"
        SetFill ws, "pdx_kpi_recon_seg1", HX("FFFFFF"), 0.9
    ElseIf nBreak > 0 Then
        SetText ws, "pdx_kpi_recon_value", CStr(nBreak)
        SetTextColor ws, "pdx_kpi_recon_value", HX("FF6B5E")
        SetText ws, "pdx_kpi_recon_sub", Plural(nBreak, "break", "breaks") & " in " & nCmp & _
                                         Plural(nCmp, " comparison", " comparisons")
        SetFill ws, "pdx_kpi_recon_seg1", HX("FF6B5E")
    Else
        SetText ws, "pdx_kpi_recon_value", "0"
        SetTextColor ws, "pdx_kpi_recon_value", HX("2FC48D")
        SetText ws, "pdx_kpi_recon_sub", "breaks across " & nCmp & Plural(nCmp, " comparison", " comparisons")
        SetFill ws, "pdx_kpi_recon_seg1", HX("2FC48D")
    End If

    ' The as-of date of the data, off the first output that knows it.
    For Each fw In Frameworks()
        If Len(asOf) = 0 Then asOf = SlotCell("OUTPUT|" & CStr(fw), modPD_Files.S_ASOF)
    Next fw
    If Len(asOf) > 0 Then
        SetText ws, "pdx_asof", "DATA AS OF " & UCase$(asOf)
    ElseIf nFiles > 0 Then
        SetText ws, "pdx_asof", nFiles & Plural(nFiles, " FILE", " FILES") & " ON THE DESK"
    Else
        SetText ws, "pdx_asof", "NO DATA LOADED"
    End If
End Sub

Private Sub Kpi(ByVal ws As Worksheet, ByVal key As String, ByVal n As Long, ByVal total As Long, _
                ByVal subText As String)
    Dim j As Long
    On Error Resume Next      ' one label that will not paint must not stop the rest
    SetText ws, "pdx_kpi_" & key & "_value", CStr(n)
    SetText ws, "pdx_kpi_" & key & "_sub", subText
    For j = 1 To total
        If j <= n Then
            SetFill ws, "pdx_kpi_" & key & "_seg" & j, HX("16B07F")
        Else
            SetFill ws, "pdx_kpi_" & key & "_seg" & j, HX("FFFFFF"), 0.9
        End If
    Next j
End Sub

Private Sub PaintFiles(ByVal ws As Worksheet)
    Dim keys As Variant, i As Long, st As String, n As Long
    On Error Resume Next      ' one label that will not paint must not stop the rest
    keys = SlotKeys()
    For i = 0 To UBound(keys)
        st = SlotState(CStr(keys(i)))
        If IsIn(st) Then n = n + 1
        SetFill ws, "pdx_slot" & (i + 1) & "_dot", DotColor(st)
        SetText ws, "pdx_slot" & (i + 1) & "_meta", SlotMeta(CStr(keys(i)), st)
        SetTextColor ws, "pdx_slot" & (i + 1) & "_meta", SlotMetaColor(st)
    Next i
    SetText ws, "pdx_card1_count", n & " OF 5 IN PLACE"
End Sub

Private Sub PaintBuild(ByVal ws As Worksheet)
    Dim fws As Variant, j As Long, fw As String, st As String, sel As Boolean, nm As String
    Dim nReady As Long, nSel As Long, rowsTxt As String, asOf As String, meta As String
    Dim track As Shape, knob As Shape, lastWhen As String, lastN As String, folder As String
    On Error Resume Next      ' one label that will not paint must not stop the rest

    fws = Frameworks()
    For j = 0 To UBound(fws)
        fw = CStr(fws(j))
        nm = "pdx_fw" & (j + 1)
        st = SlotState("OUTPUT|" & fw)
        sel = (SettingGet("sel_" & fw, "1") = "1")
        If IsIn(st) Then
            nReady = nReady + 1
            If sel Then nSel = nSel + 1
            rowsTxt = SlotCell("OUTPUT|" & fw, modPD_Files.S_ROWS)
            asOf = SlotCell("OUTPUT|" & fw, modPD_Files.S_ASOF)
            If Len(rowsTxt) > 0 And IsNumeric(rowsTxt) Then
                meta = Fmt(CDbl(rowsTxt)) & " rows"
            Else
                meta = MidTrim(FileLeaf(modPD_Files.SlotFile("OUTPUT|" & fw)), 34)
            End If
            If Len(asOf) > 0 Then meta = meta & "  " & ChrW(183) & "  as of " & asOf
            SetText ws, nm & "_meta", meta
            SetTextColor ws, nm & "_label", HX("F2F7F4")
            SetTextColor ws, nm & "_meta", HX("7E9388")
            If sel Then
                SetFill ws, nm & "_row", HX("001D14")
                SetLine ws, nm & "_row", HX("004A32")
                SetFill ws, nm & "_track", HX("16B07F")
            Else
                SetFill ws, nm & "_row", HX("111F19")
                SetLine ws, nm & "_row", HX("1A2922")
                SetFill ws, nm & "_track", HX("2A3B33")
            End If
            SetFill ws, nm & "_knob", HX("FFFFFF")
        Else
            sel = False
            If st = "MISSING" Then
                SetText ws, nm & "_meta", "The file has moved " & ChrW(183) & " click to find it"
            Else
                SetText ws, nm & "_meta", "Add " & IIf(j = 2, "a maturity ladder", "an " & FwLabel(fw)) & _
                                          " output first " & ChrW(183) & " click to choose"
            End If
            SetTextColor ws, nm & "_label", HX("4A5E55")
            SetTextColor ws, nm & "_meta", HX("4A5E55")
            SetFill ws, nm & "_row", HX("FFFFFF"), 1
            SetLine ws, nm & "_row", HX("1A2922"), 0.3
            SetFill ws, nm & "_track", HX("17221D")
            SetFill ws, nm & "_knob", HX("34443D")
        End If
        Set track = Shp(ws, nm & "_track")
        Set knob = Shp(ws, nm & "_knob")
        If Not track Is Nothing And Not knob Is Nothing Then
            On Error Resume Next
            If sel Then
                knob.Left = track.Left + track.Width - knob.Width - 2
            Else
                knob.Left = track.Left + 2
            End If
            Err.Clear
        End If
    Next j

    If nReady = 0 Then
        SetText ws, "pdx_card2_count", "NOTHING TO BUILD YET"
    Else
        SetText ws, "pdx_card2_count", nReady & " OF 3 READY"
    End If
    If modPD_Config.Engine() = "classic" Then
        SetText ws, "pdx_c2_config", "1.0 layout  " & ChrW(8594)
    ElseIf Val(SettingGet("config_bad", "0")) > 0 Then
        SetText ws, "pdx_c2_config", SettingGet("config_bad") & " to fix  " & ChrW(8594)
    Else
        SetText ws, "pdx_c2_config", SettingGet("config_on", "4") & " pivots  " & ChrW(8594)
    End If
    If nSel > 0 Then
        SetText ws, "pdx_c2_build", "Build " & nSel & Plural(nSel, " workbook", " workbooks")
        PaintButton ws, "pdx_c2_build", "soft"
    ElseIf nReady > 0 Then
        SetText ws, "pdx_c2_build", "Switch one on"
        PaintButton ws, "pdx_c2_build", "off"
    Else
        SetText ws, "pdx_c2_build", "Build pivots"
        PaintButton ws, "pdx_c2_build", "off"
    End If

    lastWhen = SettingGet("last_build_when")
    lastN = SettingGet("last_build_n")
    folder = SettingGet("last_out_folder")
    If Len(lastWhen) > 0 Then
        SetText ws, "pdx_c2_last", "Last built " & lastWhen & "  " & ChrW(183) & "  " & lastN & _
            Plural(CLng(Val(lastN)), " workbook", " workbooks")
    Else
        SetText ws, "pdx_c2_last", "Nothing built yet."
    End If
    If Len(folder) > 0 Then
        PaintButton ws, "pdx_c2_open", "ghost"
    Else
        PaintButton ws, "pdx_c2_open", "off"
    End If
End Sub

Private Sub PaintRecon(ByVal ws As Worksheet)
    Dim st3 As String, st6 As String, nFw As Long, fw As Variant, ran As Boolean
    Dim r As Long, c As Long, v As String, ctlKey As Variant, fws As Variant, nm As String
    Dim nBreak As Long, nCmp As Long, lvl As String
    On Error Resume Next      ' one label that will not paint must not stop the rest

    st3 = SlotState("CTRL3|")
    st6 = SlotState("CTRL6|")
    For Each fw In Frameworks()
        If IsIn(SlotState("OUTPUT|" & CStr(fw))) Then nFw = nFw + 1
    Next fw

    SetFill ws, "pdx_ctl1_dot", DotColor(st3)
    SetFill ws, "pdx_ctl2_dot", DotColor(st6)
    SetText ws, "pdx_ctl1_meta", "by COA  " & ChrW(183) & "  " & CtlMeta(st3, SettingGet("ctl3_keys"))
    SetText ws, "pdx_ctl2_meta", "by account  " & ChrW(183) & "  " & CtlMeta(st6, SettingGet("ctl6_keys"))
    SetTextColor ws, "pdx_ctl1_meta", SlotMetaColor(st3)
    SetTextColor ws, "pdx_ctl2_meta", SlotMetaColor(st6)

    Select Case True
        Case nFw = 0 And Not IsIn(st3) And Not IsIn(st6)
            SetText ws, "pdx_card3_count", "NEEDS AN OUTPUT AND A CONTROL"
        Case nFw = 0
            SetText ws, "pdx_card3_count", "NEEDS AN OUTPUT"
        Case Not IsIn(st3) And Not IsIn(st6)
            SetText ws, "pdx_card3_count", "NEEDS A CONTROL REPORT"
        Case Else
            SetText ws, "pdx_card3_count", (Abs(IsIn(st3)) + Abs(IsIn(st6))) & " OF 2 CONTROLS IN PLACE"
    End Select

    ran = (SettingGet("recon_ran") = "1")
    If ran Then
        SetText ws, "pdx_c3_when", "LAST RUN  " & ChrW(183) & "  " & UCase$(SettingGet("recon_when"))
    Else
        SetText ws, "pdx_c3_when", "NOT RUN YET"
    End If

    fws = Frameworks()
    r = 0
    For Each ctlKey In Array("CR3", "CR6")
        r = r + 1
        For c = 0 To UBound(fws)
            nm = "pdx_m_" & r & (c + 1)
            v = SettingGet("m_" & CStr(ctlKey) & "_" & CStr(fws(c)))
            PaintCell ws, nm, v
        Next c
    Next ctlKey

    nBreak = CLng(Val(SettingGet("recon_breaks", "0")))
    nCmp = CLng(Val(SettingGet("recon_n", "0")))
    If Not ran Then
        lvl = "IDLE"
        SetText ws, "pdx_c3_vtitle", "NOT RUN"
        SetText ws, "pdx_c3_vvalue", ChrW(8212)
    ElseIf nBreak > 0 Then
        lvl = "BREAK"
        SetText ws, "pdx_c3_vtitle", UCase$(Plural(nBreak, "break", "breaks"))
        SetText ws, "pdx_c3_vvalue", nBreak & " of " & nCmp
    Else
        lvl = "OK"
        SetText ws, "pdx_c3_vtitle", "ALL CLEAR"
        SetText ws, "pdx_c3_vvalue", nCmp & " of " & nCmp
    End If
    SetFill ws, "pdx_c3_vbox", DarkBg(lvl)
    SetLine ws, "pdx_c3_vbox", LevelDark(lvl), 0.55
    SetTextColor ws, "pdx_c3_vtitle", LevelDark(lvl)

    If nFw > 0 And (IsIn(st3) Or IsIn(st6)) Then
        PaintButton ws, "pdx_c3_run", "soft"
    Else
        PaintButton ws, "pdx_c3_run", "off"
    End If
End Sub

Private Function CtlMeta(ByVal state As String, ByVal keys As String) As String
    Select Case state
        Case "EMPTY": CtlMeta = "not added"
        Case "MISSING": CtlMeta = "file moved"
        Case "FORCED": CtlMeta = "placed by hand"
        Case Else
            If Len(keys) > 0 Then CtlMeta = Fmt(Val(keys)) & " keys" Else CtlMeta = "loaded"
    End Select
End Function

Private Function DarkBg(ByVal lvl As String) As Long
    Select Case UCase$(lvl)
        Case "OK": DarkBg = HX("0D2A20")
        Case "BREAK": DarkBg = HX("2E1614")
        Case "CHECK": DarkBg = HX("2A2310")
        Case Else: DarkBg = HX("121D19")
    End Select
End Function

' One cell of the verdict matrix: a control against a framework.
Private Sub PaintCell(ByVal ws As Worksheet, ByVal nm As String, ByVal verdict As String)
    Dim lvl As String, word As String
    On Error Resume Next      ' one label that will not paint must not stop the rest
    Select Case UCase$(verdict)
        Case "OK": lvl = "OK": word = "OK"
        Case "BREAK": lvl = "BREAK": word = "Break"
        Case "CHECK": lvl = "CHECK": word = "Check"
        Case Else: lvl = "IDLE": word = ChrW(8212)
    End Select
    SetFill ws, nm, DarkBg(lvl)
    If lvl = "IDLE" Then
        SetLine ws, nm, HX("243A31")
        SetTextColor ws, nm, HX("4A5E55")
    Else
        SetLine ws, nm, LevelDark(lvl), 0.6
        SetTextColor ws, nm, LevelDark(lvl)
    End If
    SetText ws, nm, word
End Sub

Private Sub PaintActivity(ByVal ws As Worksheet)
    Dim lg As Worksheet, a As Long, r As Long, n As Long, lvl As String, msg As String
    On Error Resume Next      ' one label that will not paint must not stop the rest
    Set lg = GetSheet(SH_LOG)
    For a = 1 To 4
        r = modPD_Theme.R_FIRST + a - 1
        msg = ""
        If Not lg Is Nothing Then msg = SafeText(lg.Cells(r, 4).Value2)
        If Len(msg) > 0 Then
            n = n + 1
            lvl = SafeText(lg.Cells(r, 2).Value2)
            If Len(msg) > 180 Then msg = Left$(msg, 177) & ChrW(8230)
            SetText ws, "pdx_act" & a & "_when", SafeText(lg.Cells(r, 1).Value2)
            SetText ws, "pdx_act" & a & "_lvl", UCase$(lvl)
            SetTextColor ws, "pdx_act" & a & "_lvl", LevelDark(lvl)
            SetFill ws, "pdx_act" & a & "_lvl", DarkBg(lvl)
            SetLine ws, "pdx_act" & a & "_lvl", LevelDark(lvl), 0.65
            SetText ws, "pdx_act" & a & "_stage", SafeText(lg.Cells(r, 3).Value2)
            SetText ws, "pdx_act" & a & "_msg", msg
        End If
        SetVisible ws, "pdx_act" & a & "_when", Len(msg) > 0
        SetVisible ws, "pdx_act" & a & "_lvl", Len(msg) > 0
        SetVisible ws, "pdx_act" & a & "_stage", Len(msg) > 0
        SetVisible ws, "pdx_act" & a & "_msg", Len(msg) > 0
    Next a
    SetVisible ws, "pdx_act_empty", n = 0
    SetVisible ws, "pdx_act_art", n = 0
End Sub

' ===================== the maturity gap =====================================
'
' The net pre-factor balance in each maturity bucket of the framework last
' built, shortest tenor first: bars above the zero line in emerald, below it
' in grey, "(no bucket)" dim at the end. The numbers are what staging measured
' (modPD_Stage.NoteGap), so the card costs nothing to paint.

Private Sub PaintGap(ByVal ws As Worksheet)
    Dim fw As String, raw As String, parts As Variant, n As Long, i As Long, f As Variant, bits As Variant
    Dim lbl() As String, v() As Double, kind() As String, nShow As Long, tot As Variant
    Dim card As Shape, sx As Double, cx As Double, top As Double, cw As Double, ch As Double
    Dim gp As Double, bw As Double, x0 As Double, pos As Double, neg As Double, span As Double
    Dim zero As Double, h As Double, bar As Shape, lab As Shape, nm As String, nData As Long
    On Error Resume Next      ' one label that will not paint must not stop the rest

    fw = SettingGet("gap_fw")
    If Len(fw) > 0 Then raw = SettingGet("gap_" & fw)
    If Len(raw) = 0 Then
        For Each f In Frameworks()
            raw = SettingGet("gap_" & CStr(f))
            If Len(raw) > 0 Then fw = CStr(f): Exit For
        Next f
    End If
    For Each f In Frameworks()
        If Len(SettingGet("gap_" & CStr(f))) > 0 Then nData = nData + 1
    Next f

    Set card = Shp(ws, "pdx_gap")
    If card Is Nothing Then Exit Sub
    sx = card.Width / GAP_CARD_W
    If sx <= 0 Then sx = 1
    cx = card.Left + GAP_PAD_X * sx
    top = card.Top + GAP_PAD_Y
    cw = GAP_CW * sx
    ch = GAP_CH

    If Len(raw) = 0 Then
        ' Nothing built yet: the ghost of a gap under a pill that says so.
        nShow = 9
        ReDim v(1 To nShow)
        bits = Array(0.62, 0.38, 0.22, -0.12, -0.3, -0.46, -0.6, -0.36, 0.28)
        For i = 1 To nShow
            v(i) = CDbl(bits(i - 1))
        Next i
    Else
        parts = Split(raw, ";")
        n = UBound(parts) + 1
        nShow = n
        If nShow > GAP_BARS Then nShow = GAP_BARS
        ReDim lbl(1 To nShow)
        ReDim v(1 To nShow)
        ReDim kind(1 To nShow)
        For i = 1 To n
            bits = Split(CStr(parts(i - 1)), "|")
            If UBound(bits) >= 3 Then
                If i < GAP_BARS Or n = GAP_BARS Then
                    lbl(i) = CStr(bits(0)): v(i) = Val(bits(1)): kind(i) = CStr(bits(3))
                Else
                    ' More buckets than bars: the longest tenors share the last.
                    lbl(GAP_BARS) = "MORE": v(GAP_BARS) = v(GAP_BARS) + Val(bits(1)): kind(GAP_BARS) = ""
                End If
            End If
        Next i
    End If

    ' The same arithmetic as build/desk.py gap_bars.
    If nShow > 8 Then gp = 5 * sx Else gp = 8 * sx
    bw = (cw - (nShow - 1) * gp) / nShow
    If bw > 26 * sx Then bw = 26 * sx
    x0 = cx + (cw - (nShow * bw + (nShow - 1) * gp)) / 2
    For i = 1 To nShow
        If v(i) > pos Then pos = v(i)
        If -v(i) > neg Then neg = -v(i)
    Next i
    span = pos + neg
    If span <= 0 Then span = 1
    zero = top + ch * pos / span

    For i = 1 To GAP_BARS
        nm = "pdx_gap_bar" & i
        Set bar = Shp(ws, nm)
        Set lab = Shp(ws, "pdx_gap_lbl" & i)
        If i <= nShow And Not bar Is Nothing Then
            h = ch * Abs(v(i)) / span
            If h < 1 Then h = 1
            bar.Left = x0 + (i - 1) * (bw + gp)
            bar.Width = bw
            bar.Height = h
            If v(i) >= 0 Then bar.Top = zero - h Else bar.Top = zero
            If Len(raw) = 0 Then
                SetFill ws, nm, HX("FFFFFF"), 0.95
            ElseIf kind(i) = "none" Then
                SetFill ws, nm, HX("4A5E55")
            ElseIf v(i) >= 0 Then
                SetFill ws, nm, HX("16B07F")
            Else
                SetFill ws, nm, HX("7E9388")
            End If
            SetVisible ws, nm, True
            If Not lab Is Nothing Then
                lab.Left = bar.Left - gp / 2 - 4
                lab.Width = bw + gp + 8
                lab.Top = top + ch + 3
                If Len(raw) > 0 Then SetText ws, "pdx_gap_lbl" & i, lbl(i)
            End If
            SetVisible ws, "pdx_gap_lbl" & i, Len(raw) > 0
        Else
            SetVisible ws, nm, False
            SetVisible ws, "pdx_gap_lbl" & i, False
        End If
    Next i
    Set bar = Shp(ws, "pdx_gap_zero")
    If Not bar Is Nothing Then bar.Top = zero - bar.Height / 2

    SetVisible ws, "pdx_gap_empty", Len(raw) = 0
    SetVisible ws, "pdx_gap_fw", Len(raw) > 0
    If Len(raw) = 0 Then
        SetText ws, "pdx_gap_unit", "NET PRE-FACTOR  " & ChrW(183) & "  BY BUCKET"
        SetText ws, "pdx_gap_v1", ChrW(8212)
        SetText ws, "pdx_gap_v2", ChrW(8212)
        SetText ws, "pdx_gap_v3", ChrW(8212)
        Exit Sub
    End If

    SetText ws, "pdx_gap_fw", GapChip(fw) & IIf(nData > 1, "  " & ChrW(8250), "")
    tot = Split(SettingGet("gap_" & fw & "_tot") & "||||", "|")
    SetText ws, "pdx_gap_unit", "NET PRE-FACTOR  " & ChrW(183) & "  LCY  " & ChrW(183) & "  " & _
                                n & Plural(n, " BUCKET", " BUCKETS")
    SetText ws, "pdx_gap_v1", Replace(Compact(Val(tot(0))), "-", ChrW(8722))
    SetText ws, "pdx_gap_v2", Compact(Val(tot(1)))
    If Val(tot(1)) > 0 Then
        SetText ws, "pdx_gap_v3", Format$(Val(tot(2)) / Val(tot(1)), "0.0%")
    Else
        SetText ws, "pdx_gap_v3", ChrW(8212)
    End If
End Sub

Private Function GapChip(ByVal fw As String) As String
    If StrComp(fw, FW_ML, vbTextCompare) = 0 Then GapChip = "LADDER" Else GapChip = UCase$(FwLabel(fw))
End Function

' The chip on the card: the next framework that has been built.
Public Sub PD_GapNext()
    Dim fws As Variant, cur As String, i As Long, k As Long, f As String, ws As Worksheet
    If PD_Busy Then Exit Sub
    On Error Resume Next
    fws = Frameworks()
    cur = SettingGet("gap_fw")
    For i = 0 To UBound(fws)
        If StrComp(CStr(fws(i)), cur, vbTextCompare) = 0 Then k = i
    Next i
    For i = 1 To UBound(fws) + 1
        f = CStr(fws((k + i) Mod (UBound(fws) + 1)))
        If Len(SettingGet("gap_" & f)) > 0 Then
            SettingSet "gap_fw", f
            Exit For
        End If
    Next i
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    ws.Unprotect
    PaintGap ws
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Err.Clear
End Sub

' ===================== results reported by the engine =======================

' Called by the reconciliation when it finishes: what the matrix and the
' verdict tile show, kept so they survive closing the workbook.
Public Sub NoteRecon(ByVal summaries As Collection, ByVal ran As Boolean, _
                     ByVal nBreak As Long, ByVal worst As Double)
    Dim i As Long, s As Object, ctl As String
    On Error Resume Next
    SettingClear "m_"
    For i = 1 To summaries.count
        Set s = summaries(i)
        If InStr(1, CStr(s("Control")), "3", vbBinaryCompare) > 0 Then ctl = "CR3" Else ctl = "CR6"
        SettingSet "m_" & ctl & "_" & CStr(s("Framework")), CStr(s("Verdict"))
    Next i
    SettingSet "recon_ran", IIf(ran, "1", "0")
    SettingSet "recon_n", CStr(summaries.count)
    SettingSet "recon_breaks", CStr(nBreak)
    SettingSet "recon_worst", Format$(worst, "0")
    SettingSet "recon_when", Format$(Now, "dd mmm hh:nn")
    SettingSet "recon_sig", ReconSig()
    Err.Clear
End Sub

Public Sub NoteBuild(ByVal folder As String, ByVal n As Long)
    On Error Resume Next
    SettingSet "last_out_folder", folder
    SettingSet "last_build_when", Format$(Now, "dd mmm hh:nn")
    SettingSet "last_build_n", CStr(n)
    If n > 0 Then SettingSet "build_sig", BuildSig()
    Err.Clear
End Sub

' ===================== actions on the Desk ==================================

Private Function CallerName() As String
    On Error Resume Next
    CallerName = CStr(Application.Caller)
    Err.Clear
End Function

' Any row in Add files, and either control tile, opens the file picker for
' exactly that slot.
Public Sub PD_SlotClick()
    Dim nm As String, keys As Variant, key As String
    If PD_Busy Then Exit Sub
    ToastHide
    nm = CallerName()
    keys = SlotKeys()
    If nm Like "pdx_slot#_*" Then
        key = CStr(keys(CLng(Mid$(nm, 9, 1)) - 1))
    ElseIf nm = "pdx_ctl1_hit" Then
        key = "CTRL3|"
    ElseIf nm = "pdx_ctl2_hit" Then
        key = "CTRL6|"
    End If
    If Len(key) > 0 Then modPD_Files.UseFileFor key
End Sub

' A framework switch. A framework with nothing loaded behind it cannot be
' switched on, so its row opens the picker for its output instead.
Public Sub PD_FwToggle()
    Dim nm As String, i As Long, fws As Variant, fw As String
    If PD_Busy Then Exit Sub
    ToastHide
    nm = CallerName()
    If Not nm Like "pdx_fw#_*" Then Exit Sub
    i = CLng(Mid$(nm, 7, 1))
    fws = Frameworks()
    If i < 1 Or i > UBound(fws) + 1 Then Exit Sub
    fw = CStr(fws(i - 1))
    If Not IsIn(SlotState("OUTPUT|" & fw)) Then
        modPD_Files.UseFileFor "OUTPUT|" & fw
        Exit Sub
    End If
    If SettingGet("sel_" & fw, "1") = "1" Then
        SettingSet "sel_" & fw, "0"
        SlideKnob i, False
    Else
        SettingSet "sel_" & fw, "1"
        SlideKnob i, True
    End If
    RefreshDesk
End Sub

' The switch slides rather than jumps: four frames over about a tenth of a
' second, easing out, with the track taking its new colour half way.
Private Sub SlideKnob(ByVal i As Long, ByVal toOn As Boolean)
    Dim ws As Worksheet, track As Shape, knob As Shape, x0 As Double, x1 As Double, f As Long, e As Double
    On Error Resume Next
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    Set track = Shp(ws, "pdx_fw" & i & "_track")
    Set knob = Shp(ws, "pdx_fw" & i & "_knob")
    If track Is Nothing Or knob Is Nothing Then Exit Sub
    x0 = knob.Left
    If toOn Then x1 = track.Left + track.Width - knob.Width - 2 Else x1 = track.Left + 2
    ws.Unprotect
    Application.ScreenUpdating = True
    For f = 1 To 4
        e = 1 - (1 - f / 4) ^ 3
        knob.Left = x0 + (x1 - x0) * e
        If f = 2 Then
            If toOn Then track.Fill.ForeColor.RGB = HX("16B07F") Else track.Fill.ForeColor.RGB = HX("2A3B33")
        End If
        Pause 0.025
    Next f
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Err.Clear
End Sub

' A short wait that keeps the screen painting.
Private Sub Pause(ByVal secs As Single)
    Dim t0 As Single
    DoEvents
    t0 = Timer
    Do While Timer - t0 < secs And Timer >= t0
        DoEvents
    Loop
End Sub

Public Sub PD_BuildSelected()
    Dim fw As Variant, s As String
    If PD_Busy Then Exit Sub
    PressFx
    ToastHide
    For Each fw In Frameworks()
        If IsIn(SlotState("OUTPUT|" & CStr(fw))) And SettingGet("sel_" & CStr(fw), "1") = "1" Then
            If Len(s) > 0 Then s = s & ","
            s = s & CStr(fw)
        End If
    Next fw
    If Len(s) = 0 Then
        If modPD_Files.AnyFrameworkLoaded() Then
            Toast "Switch on at least one framework to build.", "CHECK"
        Else
            Toast "Add an LCR, NSFR or maturity ladder output first " & ChrW(8212) & " there is nothing to build yet.", "CHECK"
        End If
        Exit Sub
    End If
    modPD_Run.BuildPivotsFor s
End Sub

Public Sub PD_OpenOutputFolder()
    Dim f As String, there As Boolean
    PressFx
    ToastHide
    f = SettingGet("last_out_folder")
    If Len(f) = 0 Then
        Toast "Nothing has been built yet, so there is no folder to open.", "CHECK"
        Exit Sub
    End If
    On Error Resume Next
    there = (Len(Dir$(f, vbDirectory)) > 0)
    Err.Clear
    If Not there Then
        Toast "The last output folder is not there any more: " & f, "CHECK"
        Exit Sub
    End If
    Shell "explorer.exe """ & f & """", vbNormalFocus
    Err.Clear
End Sub

' Kept so a button or shortcut made for 1.0 still lands somewhere sensible.
Public Sub PD_Console()
    modPD_Theme.GoTo_ SH_HOME
End Sub

' The PivotDesk tab on the ribbon (build/ribbon.py), for Excel view. Every
' button calls this, and it dispatches on the button's id. Declared As Object
' rather than IRibbonControl, so it needs no reference to the Office library.
Public Sub PD_RibbonClick(control As Object)
    Dim id As String
    On Error Resume Next
    id = CStr(control.id)
    On Error GoTo 0
    Select Case id
        Case "pdDesk": modPD_Theme.PD_GoHome
        Case "pdFiles": modPD_Theme.PD_GoFiles
        Case "pdConfig": modPD_Theme.PD_GoConfig
        Case "pdRecon": modPD_Theme.PD_GoRecon
        Case "pdLog": modPD_Theme.PD_GoLog
        Case "pdScan": modPD_Files.PD_LoadFolder
        Case "pdPick": modPD_Files.PD_LoadFiles
        Case "pdBuild": PD_BuildSelected
        Case "pdReconcile": modPD_Recon.PD_Reconcile
        Case "pdApp": PD_ToggleAppView
        Case "pdTour": PD_TourStart
    End Select
End Sub

' ===================== toast ================================================

' A one-line report on the Desk's bar that clears itself. It replaces the
' message box after every action: the result is on screen, and nothing has
' to be clicked away.
Public Sub Toast(ByVal msg As String, Optional ByVal level As String = "OK")
    Dim ws As Worksheet, clr As Long, dotSh As Shape, t As Date
    On Error Resume Next
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    msg = Replace(msg, vbCrLf & vbCrLf, "  ")
    msg = Replace(msg, vbCrLf, " ")
    msg = Replace(msg, vbLf, " ")
    Do While InStr(msg, "   ") > 0
        msg = Replace(msg, "   ", "  ")
    Loop
    msg = Trim$(msg)
    If Len(msg) > 190 Then msg = Left$(msg, 187) & ChrW(8230)
    clr = LevelDark(level)
    ws.Unprotect
    SetText ws, "pdx_toast", msg
    SetLine ws, "pdx_toast", clr, 0.45
    SetFill ws, "pdx_toast_dot", clr
    Set dotSh = Shp(ws, "pdx_toast_dot")
    If Not dotSh Is Nothing Then dotSh.Glow.Color.RGB = clr
    If Application.ScreenUpdating And DeskInFront() Then
        FadeIn ws, "pdx_toast", 0.45
    Else
        SetVisible ws, "pdx_toast", True
    End If
    SetVisible ws, "pdx_toast_dot", True
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True

    ' Cleared by a timer. The time is kept on the settings sheet, not only in
    ' a variable, so closing the workbook can cancel it even after a code
    ' reset - a timer left pending re-opens a closed workbook to run.
    CancelToastTimer
    t = Now + TimeSerial(0, 0, TOAST_SECONDS)
    SettingSet "toast_at", Format$(t, "yyyy-mm-dd hh:nn:ss")
    Application.OnTime t, "PD_ToastAutoHide"
    Err.Clear
End Sub

' Comes up over three frames instead of appearing: fill, outline and text
' from clear to their own opacity.
Private Sub FadeIn(ByVal ws As Worksheet, ByVal nm As String, ByVal lineTransp As Single)
    Dim sh As Shape, f As Long, e As Single
    On Error Resume Next
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    sh.Fill.Transparency = 1
    sh.Line.Transparency = 1
    sh.TextFrame2.TextRange.Font.Fill.Transparency = 1
    sh.visible = msoTrue
    For f = 1 To 3
        e = f / 3
        sh.Fill.Transparency = 1 - e
        sh.Line.Transparency = 1 - e * (1 - lineTransp)
        sh.TextFrame2.TextRange.Font.Fill.Transparency = 1 - e
        Pause 0.03
    Next f
    sh.Fill.Transparency = 0
    sh.Line.Transparency = lineTransp
    sh.TextFrame2.TextRange.Font.Fill.Transparency = 0
    Err.Clear
End Sub

Public Sub PD_ToastAutoHide()
    Dim due As String
    On Error Resume Next
    due = SettingGet("toast_at")
    ' A newer toast moved the time on; this timer belongs to the old one.
    If Len(due) > 0 Then
        If Now < CDate(due) - TimeSerial(0, 0, 1) Then Exit Sub
    End If
    SettingSet "toast_at", ""
    ToastHide
    Err.Clear
End Sub

Public Sub PD_ToastHide()
    CancelToastTimer
    ToastHide
End Sub

Private Sub ToastHide()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    If Shp(ws, "pdx_toast") Is Nothing Then Exit Sub
    If ws.Shapes("pdx_toast").visible = msoFalse Then Exit Sub
    ws.Unprotect
    SetVisible ws, "pdx_toast", False
    SetVisible ws, "pdx_toast_dot", False
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Err.Clear
End Sub

Private Sub CancelToastTimer()
    Dim due As String
    On Error Resume Next
    due = SettingGet("toast_at")
    If Len(due) > 0 Then Application.OnTime CDate(due), "PD_ToastAutoHide", , False
    SettingSet "toast_at", ""
    Err.Clear
End Sub

' ===================== the tour =============================================
'
' Six steps, each lighting one part of the desk: four veils darken everything
' around it, a ring marks it, and a card beside it says what it is for. The
' placement is build/desk.py's tour_place, so the preview and Excel agree, and
' the build checks every title here is the one the design shows.

Private Sub TourStep(ByVal n As Long, ByRef targets As String, ByRef title As String, ByRef body As String)
    Select Case n
        Case 1
            targets = "pdx_hero_greet,pdx_hero_lede,pdx_cta,pdx_cta2"
            title = "The next step, always"
            body = "PivotDesk reads the desk and puts the next sensible step on this button. " & _
                   "The sentence above it says why."
        Case 2
            targets = "pdx_card1"
            title = "Put the files on the desk"
            body = "Scan a folder or pick files. Each is recognised by its columns, not its name. " & _
                   "Click a row to choose the file for that slot."
        Case 3
            targets = "pdx_card2"
            title = "Build what you need"
            body = "Switch frameworks on or off, then build. Every sheet comes from a row on Pivot config, " & _
                   "and every row can be changed."
        Case 4
            targets = "pdx_card3"
            title = "Breaks before anyone asks"
            body = "Reconcile the outputs against control reports 3 and 6. Each cell of the matrix is one " & _
                   "control against one framework."
        Case 5
            targets = "pdx_gap"
            title = "The shape of the book"
            body = "After a build, the net balance in each maturity bucket is drawn here, shortest tenor " & _
                   "first. The chip switches framework."
        Case Else
            targets = "pdx_nav"
            title = "Everything is one click away"
            body = "The bar goes to every sheet, as do Ctrl+Shift+D, F, P, R and A. Excel view brings the " & _
                   "ribbon back. F1 plays this tour again."
    End Select
End Sub

Private Function TourCardParts() As Variant
    TourCardParts = Array("pdx_tour_card", "pdx_tour_step", "pdx_tour_pip1", "pdx_tour_pip2", "pdx_tour_pip3", _
                          "pdx_tour_pip4", "pdx_tour_pip5", "pdx_tour_pip6", "pdx_tour_title", "pdx_tour_body", _
                          "pdx_tour_skip", "pdx_tour_back", "pdx_tour_next")
End Function

Private Function TourVeils() As Variant
    TourVeils = Array("pdx_tour_dim_t", "pdx_tour_dim_b", "pdx_tour_dim_l", "pdx_tour_dim_r", "pdx_tour_ring")
End Function

Public Sub PD_TourStart()
    If PD_Busy Then Exit Sub
    TourShow 1
End Sub

Public Sub PD_TourNext()
    Dim n As Long
    If PD_Busy Then Exit Sub
    n = CLng(Val(SettingGet("tour_step", "0"))) + 1
    If n > TOUR_STEPS Then
        PD_TourEnd
    Else
        TourShow n
    End If
End Sub

Public Sub PD_TourBack()
    Dim n As Long
    If PD_Busy Then Exit Sub
    n = CLng(Val(SettingGet("tour_step", "1")))
    If n > 1 Then TourShow n - 1
End Sub

Public Sub PD_TourEnd()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DeskSheet()
    If Not ws Is Nothing Then
        ws.Unprotect
        TourHide ws
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    SettingSet "toured", "1"
    Err.Clear
End Sub

' Hides the tour. The caller has the sheet unprotected.
Private Sub TourHide(ByVal ws As Worksheet)
    Dim nm As Variant
    On Error Resume Next
    If Len(SettingGet("tour_step")) = 0 Then Exit Sub
    For Each nm In TourVeils()
        SetVisible ws, CStr(nm), False
    Next nm
    For Each nm In TourCardParts()
        SetVisible ws, CStr(nm), False
    Next nm
    SettingSet "tour_step", ""
    Err.Clear
End Sub

Private Sub TourShow(ByVal n As Long)
    Dim ws As Worksheet, targets As String, title As String, body As String, nm As Variant
    Dim sh As Shape, card As Shape, found As Boolean, i As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, wd As Double
    Dim rx As Double, ry As Double, rw As Double, rh As Double, kx As Double, ky As Double
    Dim dx As Double, dy As Double
    On Error Resume Next
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    If Not DeskInFront() Then modPD_Theme.GoTo_ SH_HOME
    CancelToastTimer
    ToastHide
    TourStep n, targets, title, body

    For Each nm In Split(targets, ",")
        Set sh = Shp(ws, CStr(nm))
        If Not sh Is Nothing Then
            If Not found Then
                x0 = sh.Left: y0 = sh.Top: x1 = sh.Left + sh.Width: y1 = sh.Top + sh.Height
                found = True
            Else
                If sh.Left < x0 Then x0 = sh.Left
                If sh.Top < y0 Then y0 = sh.Top
                If sh.Left + sh.Width > x1 Then x1 = sh.Left + sh.Width
                If sh.Top + sh.Height > y1 Then y1 = sh.Top + sh.Height
            End If
        End If
    Next nm
    Set card = Shp(ws, "pdx_tour_card")
    If Not found Or card Is Nothing Then Exit Sub

    Set sh = Shp(ws, "pdx_bar")
    If sh Is Nothing Then wd = 1120 Else wd = sh.Width
    rx = x0 - TOUR_PAD: ry = y0 - TOUR_PAD
    rw = x1 - x0 + 2 * TOUR_PAD: rh = y1 - y0 + 2 * TOUR_PAD
    If ry + rh < 60 Then
        kx = Clamp(rx, 12, wd - card.Width - 12): ky = ry + rh + TOUR_GAP
    ElseIf rx + rw + TOUR_GAP + card.Width <= wd - 12 Then
        kx = rx + rw + TOUR_GAP: ky = Clamp(ry, 60, DESK_H - card.Height - 12)
    ElseIf rx - TOUR_GAP - card.Width >= 12 Then
        kx = rx - TOUR_GAP - card.Width: ky = Clamp(ry, 60, DESK_H - card.Height - 12)
    ElseIf ry + rh + TOUR_GAP + card.Height <= DESK_H - 12 Then
        kx = Clamp(rx, 12, wd - card.Width - 12): ky = ry + rh + TOUR_GAP
    Else
        kx = Clamp(rx, 12, wd - card.Width - 12): ky = ry - TOUR_GAP - card.Height
    End If

    ws.Unprotect
    ' Past the canvas on the right and below, for a window wider than the design.
    Place ws, "pdx_tour_dim_t", 0, 0, wd + 800, Clamp(ry, 0.75, ry)
    Place ws, "pdx_tour_dim_b", 0, ry + rh, wd + 800, DESK_H + 800 - ry - rh
    Place ws, "pdx_tour_dim_l", 0, ry, Clamp(rx, 0.75, rx), rh
    Place ws, "pdx_tour_dim_r", rx + rw, ry, wd + 800 - rx - rw, rh
    Place ws, "pdx_tour_ring", rx, ry, rw, rh
    dx = kx - card.Left
    dy = ky - card.Top
    For Each nm In TourCardParts()
        Set sh = Shp(ws, CStr(nm))
        If Not sh Is Nothing Then
            sh.Left = sh.Left + dx
            sh.Top = sh.Top + dy
        End If
    Next nm

    SetText ws, "pdx_tour_step", "STEP " & n & " OF " & TOUR_STEPS
    SetText ws, "pdx_tour_title", title
    SetText ws, "pdx_tour_body", body
    For i = 1 To TOUR_STEPS
        If i = n Then
            SetFill ws, "pdx_tour_pip" & i, HX("16B07F")
        ElseIf i < n Then
            SetFill ws, "pdx_tour_pip" & i, HX("006141")
        Else
            SetFill ws, "pdx_tour_pip" & i, HX("243A31")
        End If
    Next i
    If n > 1 Then
        PaintButton ws, "pdx_tour_back", "ghost"
    Else
        PaintButton ws, "pdx_tour_back", "off"
    End If
    If n < TOUR_STEPS Then
        SetText ws, "pdx_tour_next", "Next"
    Else
        SetText ws, "pdx_tour_next", "Done"
    End If
    For Each nm In TourVeils()
        SetVisible ws, CStr(nm), True
    Next nm
    For Each nm In TourCardParts()
        SetVisible ws, CStr(nm), True
    Next nm
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    SettingSet "tour_step", CStr(n)
    Err.Clear
End Sub

Private Sub Place(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, _
                  ByVal w As Double, ByVal h As Double)
    Dim sh As Shape
    Set sh = Shp(ws, nm)
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    If w < 0.75 Then w = 0.75
    If h < 0.75 Then h = 0.75
    sh.Left = x: sh.Top = y: sh.Width = w: sh.Height = h
    Err.Clear
End Sub

Private Function Clamp(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v > hi Then v = hi
    If v < lo Then v = lo
    Clamp = v
End Function

' ===================== busy =================================================

' Long work runs with the screen frozen. The overlay is painted first, so the
' frozen screen says "working" rather than looking like nothing happened.
Public Sub BusyOn()
    Dim ws As Worksheet
    On Error Resume Next
    Application.Cursor = xlWait
    If Not DeskInFront() Then Exit Sub
    Set ws = DeskSheet()
    ws.Unprotect
    TourHide ws
    SetVisible ws, "pdx_busy", True
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Application.ScreenUpdating = True
    DoEvents
    Err.Clear
End Sub

Public Sub BusyOff()
    Dim ws As Worksheet
    On Error Resume Next
    Application.Cursor = xlDefault
    Set ws = DeskSheet()
    If ws Is Nothing Then Exit Sub
    ws.Unprotect
    SetVisible ws, "pdx_busy", False
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Err.Clear
End Sub

' ===================== app view =============================================

Public Function DeskInFront() As Boolean
    On Error Resume Next
    DeskInFront = False
    If ActiveWorkbook Is Nothing Then Exit Function
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Function
    DeskInFront = (StrComp(ActiveSheet.Name, SH_HOME, vbTextCompare) = 0)
    Err.Clear
End Function

Public Function AppView() As Boolean
    AppView = (SettingGet("view", "app") <> "excel")
End Function

' Open, and every time this workbook's window comes back to the front.
Public Sub AppEnter()
    On Error Resume Next
    If PD_Busy Then Exit Sub
    RegisterKeys True
    ApplyChrome
    If DeskInFront() Then FitDesk
    Err.Clear
End Sub

' Another workbook came to the front, or this one is closing: Excel is put
' back exactly as the person left it.
Public Sub AppLeave(ByVal closing As Boolean)
    On Error Resume Next
    RestoreChrome
    RegisterKeys False
    If closing Then CancelToastTimer
    Err.Clear
End Sub

Public Sub ApplyChrome()
    Dim onDesk As Boolean, wn As Window
    On Error Resume Next
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    Set wn = ThisWorkbook.Windows(1)
    onDesk = DeskInFront()
    If AppView() Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        Application.DisplayFormulaBar = Not onDesk
        wn.DisplayWorkbookTabs = False
        wn.DisplayHeadings = False
    Else
        RestoreChrome
        wn.DisplayWorkbookTabs = True
        wn.DisplayHeadings = Not onDesk
    End If
    wn.DisplayGridlines = False
    wn.Caption = TOOL_NAME & "  " & ChrW(183) & "  " & BANK_NAME
    Err.Clear
End Sub

Public Sub RestoreChrome()
    On Error Resume Next
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    Err.Clear
End Sub

Public Sub PD_ToggleAppView()
    Dim ws As Worksheet
    On Error Resume Next
    PressFx
    If AppView() Then
        SettingSet "view", "excel"
    Else
        SettingSet "view", "app"
    End If
    ApplyChrome
    Set ws = DeskSheet()
    If Not ws Is Nothing Then
        ws.Unprotect
        SetText ws, "pdx_btn_view", IIf(AppView(), "Excel view", "App view")
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    FitDesk
    Err.Clear
End Sub

' Every sheet of the desk gets the same window treatment when it comes up.
Public Sub SheetChrome(ByVal sh As Object)
    On Error Resume Next
    If PD_Busy Then Exit Sub
    ApplyChrome
    If StrComp(sh.Name, SH_HOME, vbTextCompare) = 0 Then
        FitDesk
    Else
        FreezeHeader sh
    End If
    Err.Clear
End Sub

' The table header stays in view however far down the reader scrolls.
Private Sub FreezeHeader(ByVal sh As Object)
    Dim wn As Window
    On Error Resume Next
    Set wn = ActiveWindow
    If wn.FreezePanes Then Exit Sub
    wn.ScrollRow = 1
    wn.ScrollColumn = 1
    ' Pivot config is wide: its On and Pivot columns stay put as well.
    If StrComp(sh.Name, SH_CONFIG, vbTextCompare) = 0 Then
        sh.Cells(modPD_Theme.R_FIRST, 3).Select
    Else
        sh.Cells(modPD_Theme.R_FIRST, 1).Select
    End If
    wn.FreezePanes = True
    sh.Cells(modPD_Theme.R_FIRST, 1).Select
    Err.Clear
End Sub

' The one piece of motion a shape can carry: pressed, it sinks a point and
' comes back, so a click is felt before the work it starts. Only buttons -
' not the invisible rows or the switches, which answer by changing.
Public Sub PressFx()
    Dim nm As String, sh As Shape, ws As Worksheet, l As Double, t As Double, w As Double, h As Double
    Dim t0 As Single, prot As Boolean
    On Error Resume Next
    nm = CStr(Application.Caller)
    If Err.Number <> 0 Or Len(nm) = 0 Then Err.Clear: Exit Sub
    If InStr(nm, "_hit") > 0 Or InStr(nm, "_fw") > 0 Then Exit Sub
    If Left$(nm, 4) <> "pdx_" And Left$(nm, 4) <> "pdr_" Then Exit Sub
    Set ws = ActiveSheet
    Set sh = ws.Shapes(nm)
    If sh Is Nothing Then Exit Sub
    If Right$(nm, 3) = "_ic" Then Set sh = ws.Shapes(Left$(nm, Len(nm) - 3))
    If sh Is Nothing Then Exit Sub
    prot = ws.ProtectDrawingObjects
    If prot Then ws.Unprotect
    l = sh.Left: t = sh.Top: w = sh.Width: h = sh.Height
    Application.ScreenUpdating = True
    sh.Left = l + 1: sh.Top = t + 1: sh.Width = w - 2: sh.Height = h - 2
    DoEvents
    t0 = Timer
    Do While Timer - t0 < 0.08 And Timer >= t0
        DoEvents
    Loop
    sh.Left = l: sh.Top = t: sh.Width = w: sh.Height = h
    DoEvents
    If prot Then ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Err.Clear
End Sub

' The design fills the window whatever its size: zoom to the canvas, clamp
' so text stays legible, lock scrolling to it, and park the cursor under the
' app bar where no selection box can show.
Public Sub FitDesk()
    Dim ws As Worksheet, wn As Window
    On Error Resume Next
    If Not DeskInFront() Then Exit Sub
    Set ws = DeskSheet()
    Set wn = ActiveWindow
    ws.ScrollArea = ""
    wn.DisplayGridlines = False
    wn.DisplayHeadings = False
    ws.Range(FIT_RANGE).Select
    wn.Zoom = True
    ' Below this the smallest labels stop being readable; a window that small
    ' scrolls instead.
    If wn.Zoom < 60 Then wn.Zoom = 60
    If wn.Zoom > 200 Then wn.Zoom = 200
    wn.ScrollRow = 1
    wn.ScrollColumn = 1
    ws.Range(PARK_CELL).Select
    ws.ScrollArea = FIT_RANGE
    Err.Clear
End Sub

Public Sub FitIfDesk()
    If DeskInFront() Then FitDesk
End Sub

' Ctrl+Shift+D / F / R / A / P, and F1 for the tour, only while this
' workbook is in front.
Private Sub RegisterKeys(ByVal onOff As Boolean)
    On Error Resume Next
    If onOff Then
        Application.OnKey "^+d", "PD_GoHome"
        Application.OnKey "^+f", "PD_GoFiles"
        Application.OnKey "^+r", "PD_GoRecon"
        Application.OnKey "^+a", "PD_GoLog"
        Application.OnKey "^+p", "PD_GoConfig"
        Application.OnKey "{F1}", "PD_TourStart"
    Else
        Application.OnKey "^+d"
        Application.OnKey "^+f"
        Application.OnKey "^+r"
        Application.OnKey "^+a"
        Application.OnKey "^+p"
        Application.OnKey "{F1}"
    End If
    Err.Clear
End Sub

' ===================== open and save ========================================

Public Sub Opened()
    On Error Resume Next
    RefreshDesk
    modPD_Theme.GoTo_ SH_HOME
    AppEnter
    ' Once per version. The first time ever, the tour; after that, a line on
    ' what is new.
    If SettingGet("welcome") = "1" Then
        SettingSet "welcome", ""
        If SettingGet("toured") <> "1" And DeskInFront() Then
            TourShow 1
        Else
            Toast "New in " & TOOL_VERSION & ": the maturity gap of each build on the Desk, buckets in tenor " & _
                  "order in every pivot, and a tour of the desk on F1.", V_OK
        End If
    End If
    Err.Clear
End Sub

' A copy opened with macros off shows the banner and its sheet tabs; the
' live desk shows neither. So the file is saved in the macros-off state and
' put straight back.
Public Sub BeforeSave()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DeskSheet()
    If Not ws Is Nothing Then
        ws.Unprotect
        SetVisible ws, "pdx_macros", True
        SetVisible ws, "pdx_busy", False
        SetVisible ws, "pdx_toast", False
        SetVisible ws, "pdx_toast_dot", False
        TourHide ws
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    ThisWorkbook.Windows(1).DisplayWorkbookTabs = True
    Err.Clear
End Sub

Public Sub AfterSave()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DeskSheet()
    If Not ws Is Nothing Then
        ws.Unprotect
        SetVisible ws, "pdx_macros", False
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    End If
    ApplyChrome
    ' Putting the banner away again is not a change worth a "save?" prompt.
    ThisWorkbook.Saved = True
    Err.Clear
End Sub
