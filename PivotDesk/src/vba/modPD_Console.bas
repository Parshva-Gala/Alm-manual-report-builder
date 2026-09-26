Option Explicit

' ============================================================================
'  The console window.
'
'  It is an ordinary .hta, kept as a file in source and injected into a very
'  hidden sheet at build time, so the tool ships as ONE workbook and the page
'  is still editable as a page. Excel cannot draw what that window draws -
'  there is no hover state on a Shape, no transition, no blur - and the console
'  is the surface a person actually touches.
'
'  It talks back through a file, not through COM. VBA writes what it knows to
'  state.txt, runs the window and WAITS for it to close, then reads cmd.txt and
'  acts. No callback into a running Excel, nothing to leak if the window is
'  killed, and the whole exchange is two files a person can read.
' ============================================================================

Private Const SRC_SHEET As String = "_Console_Src"

Private Function WorkDir() As String
    Dim p As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = fso.GetSpecialFolder(2).Path & "\pivotdesk"
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    WorkDir = p
End Function

Public Sub PD_Console()
    Dim cmd As String
    If PD_Busy Then Exit Sub
    cmd = ShowConsole()
    If Len(cmd) = 0 Then Exit Sub
    Dispatch cmd
End Sub

' Writes the state, runs the window, waits, returns whatever it chose.
Public Function ShowConsole() As String
    Dim hta As String, cmdFile As String, sh As Object, fso As Object

    hta = Extract()
    If Len(hta) = 0 Then
        Tell "The console page is missing from this workbook." & vbCrLf & vbCrLf & _
             "Use the buttons on the Desk sheet instead - they do the same three things.", vbExclamation
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    cmdFile = WorkDir() & "\cmd.txt"
    If fso.FileExists(cmdFile) Then fso.DeleteFile cmdFile, True
    WriteState

    On Error GoTo Failed
    Set sh = CreateObject("WScript.Shell")
    ' waitOnReturn: the console is modal by intention. Two of these open at once
    ' and the second would act on a desk the first had already changed.
    sh.Run "mshta.exe """ & hta & """", 1, True
    On Error GoTo 0

    If Not fso.FileExists(cmdFile) Then Exit Function
    ShowConsole = Trim$(ReadAll(cmdFile))
    Exit Function

Failed:
    LogIt V_BREAK, "Console", Err.Description, ""
    Tell "The console could not be opened:" & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
         "The buttons on the Desk sheet do the same three things.", vbExclamation
End Function

' What the window needs to know: which frameworks have data, how many files are
' loaded, the bank's name, and where the logo is if there is one.
'
' Public so the headless harness can write it and read it back. The console
' once showed an empty desk because this file and its reader disagreed about
' the encoding, and nothing on either side said so.
Public Sub WriteState()
    Dim s As String, fw As Variant, n As Long, k As Variant, logo As String

    For Each k In modPD_Files.Slots.keys
        If modPD_Files.SlotLoaded(CStr(k)) Then n = n + 1
    Next k

    s = "BANK=" & BANK_NAME & vbCrLf
    s = s & "NFILES=" & n & vbCrLf
    s = s & "CTRL3=" & IIf(modPD_Files.SlotLoaded("CTRL3|"), "1", "0") & vbCrLf
    s = s & "CTRL6=" & IIf(modPD_Files.SlotLoaded("CTRL6|"), "1", "0") & vbCrLf

    Dim fws As String
    For Each fw In Frameworks()
        If Len(fws) > 0 Then fws = fws & ","
        fws = fws & CStr(fw) & ":" & FwLabel(CStr(fw)) & ":" & _
              IIf(modPD_Files.SlotLoaded("OUTPUT|" & CStr(fw)), "1", "0")
    Next fw
    s = s & "FRAMEWORKS=" & fws & vbCrLf

    logo = LogoPath()
    If Len(logo) > 0 Then s = s & "LOGO=" & logo & vbCrLf

    If Not WriteAll(WorkDir() & "\state.txt", s) Then
        LogIt V_CHECK, "Console", "Could not write the console's state file. The console will " & _
              "say so rather than showing an empty desk.", WorkDir() & "\state.txt"
    End If
End Sub

' Where the state file is, so the harness can read back what was written.
Public Function StatePath() As String
    StatePath = WorkDir() & "\state.txt"
End Function

' The bank's mark, if it has been put next to the workbook. Nothing is embedded:
' a logo dropped into the folder is a logo the tool picks up on next open, and
' the console falls back to a plain gold monogram when there is none.
Public Function LogoPath() As String
    Dim fso As Object, base As String, cand As Variant, p As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = ThisWorkbook.Path
    If Len(base) = 0 Then Exit Function
    If Not IsLocalPath(base) Then Exit Function
    For Each cand In Array("logo.png", "logo.jpg", "assets\logo.png", "assets\logo.jpg", _
                           "midbank.png", "assets\midbank.png")
        p = PathJoin(base, CStr(cand))
        If fso.FileExists(p) Then LogoPath = p: Exit Function
    Next cand
End Function

' ===================== acting on what it chose ==============================

Public Sub Dispatch(ByVal cmd As String)
    Dim verb As String, args As String, i As Long
    i = InStr(cmd, "|")
    If i > 0 Then
        verb = Left$(cmd, i - 1)
        args = Mid$(cmd, i + 1)
    Else
        verb = cmd
    End If

    Select Case UCase$(Trim$(verb))
        Case "UPLOAD_FOLDER": modPD_Files.PD_LoadFolder
        Case "UPLOAD_FILES": modPD_Files.PD_LoadFiles
        Case "PIVOTS": modPD_Run.BuildPivotsFor args
        Case "RECONCILE": modPD_Recon.PD_Reconcile
        Case Else
            If Len(Trim$(verb)) > 0 Then LogIt V_CHECK, "Console", "Unknown command: " & verb, ""
    End Select
End Sub

' ===================== the page itself ======================================

' Written out beside the workbook's own temp area each time rather than cached:
' a stale page from a previous version is a bug nobody thinks to look for.
Private Function Extract() As String
    Dim ws As Worksheet, n As Long, i As Long, s As String, p As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SRC_SHEET)
    If ws Is Nothing Then Exit Function
    n = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If n < 5 Then Exit Function
    Dim a As Variant
    a = ws.Range(ws.Cells(1, 1), ws.Cells(n, 1)).Value2
    For i = 1 To n
        s = s & SafeTextRaw(a(i, 1)) & vbCrLf
    Next i
    p = WorkDir() & "\console.hta"
    WriteAll p, s
    Extract = p
    Err.Clear
End Function

' Like SafeText but without the trimming - leading spaces are indentation in a
' stylesheet, and a page re-indented by the reader is a page that still works
' but stops being readable.
Private Function SafeTextRaw(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeTextRaw = CStr(v)
    Err.Clear
End Function

' UTF-8, and the console reads it as UTF-8. Both sides say so in their own
' comments because the one time they disagreed, everything still "worked" -
' the file was written, the file was read, and the console reported an empty
' desk over a workbook with five files on it.
'
' ADODB rather than FileSystemObject: the page is UTF-8 and holds characters a
' locale-encoded write turns into question marks.
Private Function WriteAll(ByVal Path As String, ByVal body As String) As Boolean
    Dim st As Object
    On Error GoTo Failed
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText body
    st.SaveToFile Path, 2
    st.Close
    WriteAll = True
    Exit Function
Failed:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
    Err.Clear
End Function

Private Function ReadAll(ByVal Path As String) As String
    Dim st As Object
    On Error Resume Next
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile Path
    ReadAll = st.ReadText
    st.Close
    ' A BOM survives the round trip and would make "PIVOTS" not equal "PIVOTS".
    If Len(ReadAll) > 0 Then
        If AscW(Left$(ReadAll, 1)) = 65279 Then ReadAll = Mid$(ReadAll, 2)
    End If
    Err.Clear
End Function
