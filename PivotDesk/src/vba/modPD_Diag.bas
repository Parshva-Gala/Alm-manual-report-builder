Option Explicit

' ============================================================================
'  A headless run, for the build harness.
'
'  The same code path the console takes, with the window and the message boxes
'  removed. Nothing here is a second implementation of the run - if it were, it
'  would pass while the thing the user presses fails.
' ============================================================================

Private mLog As String

' args: "<sourceFolder>|<outputFolder>[|<frameworks>]"
Public Sub PD_DiagRun(ByVal logPath As String, ByVal args As String)
    Dim parts As Variant, srcFolder As String, outFolder As String, fws As String
    Dim t0 As Single, n As Long

    PD_Quiet = True
    mLog = logPath
    OpenLog
    Set PD_Trace = New PDTrace
    PD_Trace.Path = logPath

    parts = Split(args & "||", "|")
    srcFolder = Trim$(CStr(parts(0)))
    outFolder = Trim$(CStr(parts(1)))
    fws = Trim$(CStr(parts(2)))

    Wr TOOL_NAME & " " & TOOL_VERSION & " headless run"
    Wr "source: " & srcFolder
    Wr "output: " & outFolder
    Wr String$(74, "-")

    On Error GoTo Failed
    t0 = Timer
    modPD_Run.PD_Setup
    Wr "setup " & Format$(Timer - t0, "0.0") & "s"

    t0 = Timer
    Wr "load: starting"
    modPD_Files.ClearRefused
    n = modPD_Files.ScanFolder(srcFolder, 0)
    modPD_Files.RefreshStatuses
    Wr "load: " & n & " file(s) placed, " & Format$(Timer - t0, "0.0") & "s"
    DumpFiles
    DumpConsoleState

    If Len(fws) = 0 Then fws = LoadedFrameworks()
    If Len(fws) > 0 Then
        t0 = Timer
        Wr "pivots: " & fws
        BuildQuiet fws, outFolder
        Wr "pivots: " & Format$(Timer - t0, "0.0") & "s"
    Else
        Wr "pivots: no framework output was loaded"
    End If

    t0 = Timer
    Wr "reconcile: starting"
    modPD_Recon.PD_Reconcile
    Wr "reconcile: " & Format$(Timer - t0, "0.0") & "s"
    DumpRecon
    DumpLog

    Wr "DONE"
    Exit Sub

Failed:
    Wr "FAILED: " & Err.Number & " " & Err.Description
End Sub

Private Function LoadedFrameworks() As String
    Dim fw As Variant, s As String
    For Each fw In Frameworks()
        If modPD_Files.SlotLoaded("OUTPUT|" & CStr(fw)) Then
            If Len(s) > 0 Then s = s & ","
            s = s & CStr(fw)
        End If
    Next fw
    LoadedFrameworks = s
End Function

' The same build the console runs, with the folder picker taken out.
Private Sub BuildQuiet(ByVal fws As String, ByVal outFolder As String)
    Dim parts As Variant, i As Long, fw As String, errOut As String, Path As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outFolder) Then fso.CreateFolder outFolder
    parts = Split(fws, ",")
    For i = 0 To UBound(parts)
        fw = Trim$(CStr(parts(i)))
        If Len(fw) > 0 Then
            errOut = ""
            Path = modPD_Build.BuildFramework(fw, outFolder, errOut)
            If Len(Path) > 0 Then
                Wr "   " & PadR(FwLabel(fw), 20) & "-> " & FileLeaf(Path) & _
                   "   (" & Fmt(modPD_Stage.StagedRows()) & " rows, local " & _
                   modPD_Stage.LocalCurrency() & ", " & modPD_Stage.AmountFieldNote() & ")"
                DumpSheets Path
            Else
                Wr "   " & PadR(FwLabel(fw), 20) & "NOT BUILT: " & errOut
            End If
        End If
    Next i
End Sub

Private Sub DumpSheets(ByVal Path As String)
    Dim wb As Workbook, ws As Worksheet, n As Long, pv As Long, names As String
    On Error Resume Next
    Set wb = Workbooks.Open(Path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    If wb Is Nothing Then Exit Sub
    For Each ws In wb.Worksheets
        n = n + 1
        pv = pv + ws.PivotTables.count
        If Len(names) < 150 Then names = names & IIf(Len(names) > 0, " | ", "") & ws.Name
    Next ws
    Wr "        " & n & " sheet(s), " & pv & " pivot(s):  " & names & " ..."
    wb.Close SaveChanges:=False
    Err.Clear
End Sub

Private Sub DumpFiles()
    Dim ws As Worksheet, k As Variant, r As Long
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Sub
    Wr "files:"
    r = modPD_Theme.R_FIRST
    For Each k In modPD_Files.Slots.keys
        Wr "   " & PadR(SafeText(ws.Cells(r, modPD_Files.S_WHAT).Value2), 26) & _
           PadR(SafeText(ws.Cells(r, modPD_Files.S_STATUS).Value2), 10) & _
           PadR(FileLeaf(SafeText(ws.Cells(r, modPD_Files.S_FILE).Value2)), 40) & _
           SafeText(ws.Cells(r, modPD_Files.S_NOTE).Value2)
        r = r + 1
    Next k
End Sub

' What the console would be handed, written and read back the way the console
' reads it. This is the check that was missing: the desk and the window each
' worked on their own, and only the round trip showed they disagreed.
Private Sub DumpConsoleState()
    Dim p As String, st As Object, raw As String, lines_ As Variant, i As Long
    On Error Resume Next
    modPD_Console.WriteState
    p = modPD_Console.StatePath()
    Wr "console state (" & p & "):"
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile p
    raw = st.ReadText
    st.Close
    If Err.Number <> 0 Then
        Wr "   COULD NOT READ IT BACK: " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Sub
    End If
    If Len(raw) > 0 Then
        If AscW(Left$(raw, 1)) = 65279 Then raw = Mid$(raw, 2)
    End If
    lines_ = Split(Replace(raw, vbCr, ""), vbLf)
    For i = 0 To UBound(lines_)
        If InStr(CStr(lines_(i)), "=") > 1 Then Wr "   " & CStr(lines_(i))
    Next i
    Err.Clear
End Sub

Private Sub DumpRecon()
    Dim ws As Worksheet, lastR As Long, r As Long, n As Long
    Set ws = GetSheet(SH_RECON)
    If ws Is Nothing Then Exit Sub
    Wr "reconciliation status: " & SafeText(ws.Cells(modPD_Theme.R_STATUS, 1).Value2)
    lastR = LastRow(ws, 3)
    If lastR < modPD_Theme.R_FIRST Then Wr "   no differing keys": Exit Sub
    For r = modPD_Theme.R_FIRST To lastR
        n = n + 1
        If n > 12 Then Wr "   ... and " & (lastR - r + 1) & " more": Exit For
        Wr "   " & PadR(SafeText(ws.Cells(r, 1).Value2), 11) & _
           PadR(SafeText(ws.Cells(r, 2).Value2), 18) & _
           PadR(SafeText(ws.Cells(r, 3).Value2), 24) & _
           PadL(Fmt(SafeNum(ws.Cells(r, 4).Value2)), 18) & _
           PadL(Fmt(SafeNum(ws.Cells(r, 5).Value2)), 18) & _
           PadL(Fmt(SafeNum(ws.Cells(r, 6).Value2)), 18) & "  " & _
           SafeText(ws.Cells(r, 7).Value2)
    Next r
End Sub

Private Sub DumpLog()
    Dim ws As Worksheet, r As Long, lastR As Long
    Set ws = GetSheet(SH_LOG)
    If ws Is Nothing Then Exit Sub
    lastR = LastRow(ws, 4)
    Wr "activity:"
    For r = modPD_Theme.R_FIRST To lastR
        If r > modPD_Theme.R_FIRST + 24 Then Exit For
        Wr "   " & PadR(SafeText(ws.Cells(r, 2).Value2), 8) & _
           PadR(SafeText(ws.Cells(r, 3).Value2), 10) & _
           SafeText(ws.Cells(r, 4).Value2) & _
           IIf(Len(SafeText(ws.Cells(r, 5).Value2)) > 0, "   [" & SafeText(ws.Cells(r, 5).Value2) & "]", "")
    Next r
End Sub

' ===================== the log ==============================================

Private Sub OpenLog()
    Dim fh As Integer
    On Error Resume Next
    fh = FreeFile
    Open mLog For Output As #fh
    Close #fh
    Err.Clear
End Sub

Private Sub Wr(ByVal s As String)
    Dim fh As Integer
    On Error Resume Next
    If Len(mLog) = 0 Then Exit Sub
    fh = FreeFile
    Open mLog For Append As #fh
    Print #fh, Format$(Now, "hh:nn:ss") & "  " & s
    Close #fh
    Err.Clear
End Sub

Private Function PadR(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then PadR = Left$(s, n - 1) & " " Else PadR = s & Space$(n - Len(s))
End Function

Private Function PadL(ByVal s As String, ByVal n As Long) As String
    If Len(s) >= n Then PadL = Right$(s, n) Else PadL = Space$(n - Len(s)) & s
End Function
