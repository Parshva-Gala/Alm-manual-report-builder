Attribute VB_Name = "modShared_Progress"
Option Explicit

' ============================================================================
'  Saying what is happening, and how much longer.
'
'  Some of this work takes minutes - a 150 MB extract is not read quickly - and
'  a person watching a frozen Excel has no way to tell a long job from a hung
'  one. Excel's status bar is the obvious place to say so and the wrong one: it
'  is easy to miss, it carries no sense of progress, and while VBA is busy Excel
'  frequently will not repaint it.
'
'  So the workbook writes a small status file and a separate window polls it.
'  Different process, so it keeps drawing however hard Excel is working.
'
'  The estimate is deliberately plain arithmetic - elapsed time per unit of work
'  done, extrapolated - and it is phrased as "about", because that is what it is.
' ============================================================================

Private mPath As String
Private mTitle As String
Private mSub As String
Private mStep As String
Private mDetail As String
Private mLog As String
Private mTotal As Double
Private mDone As Double
Private mStart As Double
Private mActive As Boolean
Private mLastWrite As Double

' ===================== the four calls ======================================

' totalWork is in whatever unit suits the job - files, rows, steps. It only has
' to be consistent with what is passed to ProgressStep.
Public Sub ProgressStart(ByVal title As String, ByVal subtitle As String, ByVal totalWork As Double)
    Dim fso As Object, folder As String, token As String, htaPath As String, stream As Object, html As String
    On Error GoTo Failed
    mTitle = title: mSub = subtitle
    mTotal = totalWork: mDone = 0
    mStep = "Starting...": mDetail = "": mLog = ""
    mStart = Timer
    mLastWrite = 0
    mActive = True

    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetSpecialFolder(2)
    token = format$(Now, "yyyymmddhhnnss")
    CleanOldProgress fso, folder
    mPath = fso.BuildPath(folder, "JKB_Progress_" & token & ".txt")
    WriteStatus "running", True

    ' No window during a headless run: it would be a window nobody is there to see.
    If Not Application.Visible Then Exit Sub

    html = ProgressHtml(fso.GetFileName(mPath))
    If Len(html) = 0 Then Exit Sub
    htaPath = fso.BuildPath(folder, "JKB_ProgressWin_" & token & ".hta")
    Set stream = fso.CreateTextFile(htaPath, True, True)
    stream.Write html
    stream.Close
    On Error Resume Next
    Shell "mshta.exe """ & htaPath & """", vbNormalFocus
    Err.Clear
    On Error GoTo 0
    Exit Sub
Failed:
    mActive = False
End Sub

' Names the step now beginning, and banks the work the previous one finished.
Public Sub ProgressStep(ByVal label As String, ByVal workJustDone As Double)
    If Not mActive Then Exit Sub
    If Len(mStep) > 0 And mStep <> "Starting..." And workJustDone > 0 Then
        If Len(mLog) > 0 Then mLog = mLog & "|"
        mLog = mLog & mStep
    End If
    mDone = mDone + workJustDone
    mStep = label
    mDetail = ""
    WriteStatus "running", True
End Sub

' A line under the step, for the part that moves while one step runs. Throttled,
' because writing a file on every thousandth row would itself cost real time.
Public Sub ProgressDetail(ByVal text As String, Optional ByVal workSoFar As Double = -1)
    If Not mActive Then Exit Sub
    mDetail = text
    If workSoFar >= 0 Then mDone = workSoFar
    WriteStatus "running", False
End Sub

Public Sub ProgressDone(ByVal summary As String)
    If Not mActive Then Exit Sub
    If Len(mStep) > 0 And mStep <> "Starting..." Then
        If Len(mLog) > 0 Then mLog = mLog & "|"
        mLog = mLog & mStep
    End If
    mDone = mTotal
    mStep = summary
    mDetail = ""
    WriteStatus "done", True
    mActive = False
End Sub

Public Sub ProgressFailed(ByVal reason As String)
    If Not mActive Then Exit Sub
    mStep = reason
    WriteStatus "failed", True
    mActive = False
End Sub

Public Function ProgressIsActive() As Boolean
    ProgressIsActive = mActive
End Function

' ===================== the status file ======================================

Private Sub WriteStatus(ByVal state As String, ByVal force As Boolean)
    Dim fso As Object, stream As Object, pct As Double, el As Double
    If Len(mPath) = 0 Then Exit Sub
    el = Timer - mStart
    If el < 0 Then el = el + 86400          ' rolled past midnight
    ' Two writes a second is plenty for something a person is reading.
    If Not force Then
        If el - mLastWrite < 0.5 Then Exit Sub
    End If
    mLastWrite = el

    If mTotal > 0 Then pct = mDone / mTotal
    If pct > 1 Then pct = 1
    If pct < 0 Then pct = 0

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(mPath, True, True)
    stream.WriteLine "title=" & OneLine(mTitle)
    stream.WriteLine "sub=" & OneLine(mSub)
    stream.WriteLine "state=" & state
    stream.WriteLine "pct=" & Replace(format$(pct, "0.0000"), ",", ".")
    stream.WriteLine "step=" & OneLine(mStep)
    stream.WriteLine "detail=" & OneLine(mDetail)
    stream.WriteLine "elapsed=" & Friendly(el) & " elapsed"
    stream.WriteLine "eta=" & Eta(pct, el)
    stream.WriteLine "log=" & OneLine(mLog)
    stream.Close
    Err.Clear
    On Error GoTo 0
    DoEvents
End Sub

' Plain arithmetic, honestly phrased. Below 5% done the estimate is worthless and
' saying so is better than inventing one.
Private Function Eta(ByVal pct As Double, ByVal elapsed As Double) As String
    Dim remain As Double
    If pct <= 0.05 Then Eta = "working out how long this will take": Exit Function
    If pct >= 0.999 Then Eta = "almost done": Exit Function
    remain = elapsed * (1 - pct) / pct
    Eta = "about " & Friendly(remain) & " left"
End Function

Private Function Friendly(ByVal secs As Double) As String
    Dim m As Long
    If secs < 45 Then
        Friendly = CLng(secs) & "s"
    ElseIf secs < 90 Then
        Friendly = "a minute"
    ElseIf secs < 3600 Then
        m = CLng(secs / 60)
        Friendly = m & " minute" & IIf(m = 1, "", "s")
    Else
        Friendly = format$(secs / 3600, "0.0") & " hours"
    End If
End Function

Private Function OneLine(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    OneLine = s
End Function

Private Function ProgressHtml(ByVal statusFile As String) As String
    Dim ws As Worksheet, lastR As Long, a As Variant, r As Long, sb As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_Progress_Src")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastR < 1 Then Exit Function
    a = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 1)).Value2
    If Not IsArray(a) Then
        sb = CStr(a)
    Else
        For r = 1 To UBound(a, 1)
            sb = sb & CStr(a(r, 1)) & vbCrLf
        Next r
    End If
    ProgressHtml = Replace(sb, "__STATUS_FILE__", Replace(StatusFullPath(statusFile), "\", "\\"))
End Function

Private Function StatusFullPath(ByVal fileName As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    StatusFullPath = fso.BuildPath(fso.GetSpecialFolder(2), fileName)
End Function

Private Sub CleanOldProgress(ByVal fso As Object, ByVal folder As String)
    Dim f As Object
    On Error Resume Next
    For Each f In fso.GetFolder(folder).files
        If Left$(f.name, 13) = "JKB_Progress_" Or Left$(f.name, 16) = "JKB_ProgressWin_" Then
            If DateDiff("n", f.DateLastModified, Now) > 30 Then f.Delete True
        End If
    Next f
    Err.Clear
End Sub
