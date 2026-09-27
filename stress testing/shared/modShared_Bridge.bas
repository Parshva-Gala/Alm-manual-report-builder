Attribute VB_Name = "modShared_Bridge"
Option Explicit

' ============================================================================
'  Talking to the windows, now that VBScript is gone.
'
'  The console, the condition builder and the file loader are HTAs, and they
'  reached back into Excel the only way an HTA could:
'
'      GetObject("", "Excel.Application").Run "SomeMacro", arg
'
'  On Windows 11 that no longer works. VBScript has been deprecated and is
'  disabled by default, so a <script language="VBScript"> block does not merely
'  fail at the call - it never PARSES, and every function declared in it is
'  undefined. And MSHTML's JScript, which still runs, has ActiveXObject but no
'  GetObject at all, so it cannot attach to a running instance either. Both
'  routes into Excel are closed.
'
'  What still works in an HTA is the file system. So the window and the workbook
'  talk through three files in the temporary folder:
'
'      JKB_Bridge_<token>.req    the window asks: macro, then its arguments
'      JKB_Bridge_<token>.res    the workbook answers
'      JKB_Bridge_<token>.done   the window has closed; stop listening
'
'  The macro that opened the window then PUMPS: it waits, runs whatever is
'  asked, writes the answer back, and returns when the window closes. Excel is
'  busy for as long as the window is open, which is exactly what it was before -
'  the window was always modal in practice.
'
'  Arguments are joined with Chr(2) and their newlines held as Chr(1), because
'  neither character can occur in a file path, a macro name or a filter
'  expression, and that makes the format unambiguous without any escaping.
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub SleepMs Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#Else
    Private Declare Sub SleepMs Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#End If

Private mToken As String

' ===================== the token and its file names =========================

Public Function BridgeNewToken() As String
    mToken = format$(Now, "yyyymmddhhnnss") & Right$("000" & CLng(Rnd() * 999), 3)
    BridgeCleanOld
    BridgeNewToken = mToken
End Function

Public Function BridgeFile(ByVal token As String, ByVal ext As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    BridgeFile = fso.BuildPath(fso.GetSpecialFolder(2), "JKB_Bridge_" & token & "." & ext)
End Function

' The path the window needs, with its backslashes doubled for JScript.
Public Function BridgeStemForJs(ByVal token As String) As String
    BridgeStemForJs = Replace(Left$(BridgeFile(token, "x"), Len(BridgeFile(token, "x")) - 2), "\", "\\")
End Function

' ===================== the pump =============================================
'
' Runs while the window is open. Returns the last answer it gave, or a reason.
'
' allowed is a pipe-separated list of the macro names this window is permitted
' to run. A window can only ask for what the macro that opened it said it could
' ask for - the request file is written by a page, and a page is not something
' to hand an unrestricted Run to.
Public Function BridgePump(ByVal token As String, ByVal allowed As String, _
                           Optional ByVal timeoutSec As Long = 1800) As String
    Dim reqPath As String, resPath As String, donePath As String
    Dim t0 As Double, body As String, parts As Variant
    Dim macro As String, res As String, i As Long, last As String

    reqPath = BridgeFile(token, "req")
    resPath = BridgeFile(token, "res")
    donePath = BridgeFile(token, "done")
    t0 = Timer

    Do
        If Len(Dir$(donePath)) > 0 Then
            last = ReadAllText(donePath)
            KillQuiet donePath
            BridgePump = IIf(Len(last) > 0, last, "closed")
            Exit Function
        End If

        If Len(Dir$(reqPath)) > 0 Then
            body = ReadAllText(reqPath)
            KillQuiet reqPath
            parts = Split(body, Chr$(2))
            macro = Trim$(CStr(parts(0)))
            For i = LBound(parts) To UBound(parts)
                parts(i) = Replace(CStr(parts(i)), Chr$(1), vbLf)
            Next i
            If Not IsAllowed(macro, allowed) Then
                res = "ERR:" & macro & " is not one of the actions this window may run."
            Else
                res = RunRequested(macro, parts)
            End If
            WriteAllText resPath, res
            last = res
            ' The clock restarts after real work: a person who spent ten minutes
            ' choosing has not stalled, and should not be timed out for thinking.
            t0 = Timer
        End If

        DoEvents
        SleepMs 90
        If Timer - t0 < 0 Then t0 = Timer          ' rolled past midnight
    Loop While Timer - t0 < timeoutSec

    BridgePump = "timeout"
End Function

Private Function IsAllowed(ByVal macro As String, ByVal allowed As String) As Boolean
    Dim p As Variant
    For Each p In Split(allowed, "|")
        If StrComp(Trim$(CStr(p)), macro, vbTextCompare) = 0 Then IsAllowed = True: Exit Function
    Next p
End Function

' parts(0) is the macro; anything after it is an argument.
Private Function RunRequested(ByVal macro As String, ByRef parts As Variant) As String
    Dim n As Long, v As Variant
    On Error GoTo Failed
    n = UBound(parts) - LBound(parts)
    Select Case n
        Case 0: v = Application.Run(macro)
        Case 1: v = Application.Run(macro, parts(LBound(parts) + 1))
        Case 2: v = Application.Run(macro, parts(LBound(parts) + 1), parts(LBound(parts) + 2))
        Case Else: v = Application.Run(macro, parts(LBound(parts) + 1), parts(LBound(parts) + 2), parts(LBound(parts) + 3))
    End Select
    If IsEmpty(v) Then RunRequested = "" Else RunRequested = CStr(v)
    Exit Function
Failed:
    RunRequested = "ERR:" & Err.description
End Function

' ===================== files ================================================

Private Function ReadAllText(ByVal path As String) As String
    Dim fso As Object, st As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Unicode, matching what the window writes. A filter expression can carry
    ' anything the bank's data carries.
    Set st = fso.OpenTextFile(path, 1, False, -1)
    If Not st Is Nothing Then
        If Not st.AtEndOfStream Then ReadAllText = st.ReadAll
        st.Close
    End If
    Err.Clear
End Function

Private Sub WriteAllText(ByVal path As String, ByVal s As String)
    Dim fso As Object, st As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set st = fso.CreateTextFile(path, True, True)
    st.Write s
    st.Close
    Err.Clear
End Sub

Private Sub KillQuiet(ByVal path As String)
    On Error Resume Next
    Kill path
    Err.Clear
End Sub

Public Sub BridgeCleanUp(ByVal token As String)
    KillQuiet BridgeFile(token, "req")
    KillQuiet BridgeFile(token, "res")
    KillQuiet BridgeFile(token, "done")
End Sub

Private Sub BridgeCleanOld()
    Dim fso As Object, f As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(fso.GetSpecialFolder(2)).files
        If Left$(f.name, 11) = "JKB_Bridge_" Then
            If DateDiff("h", f.DateLastModified, Now) > 6 Then f.Delete True
        End If
    Next f
    Err.Clear
End Sub

' ===================== the window's half ====================================
'
' The JScript the pages use to talk to the pump. Held here, as one string, so
' there is one implementation of the protocol rather than one per page - the two
' halves of a protocol drifting apart is the whole reason this file exists.
Public Function BridgeJs(ByVal token As String) As String
    Dim s As String
    s = s & "var JKB_BRIDGE_STEM = """ & BridgeStemForJs(token) & """;" & vbLf
    s = s & "function jkbFso(){ return new ActiveXObject(""Scripting.FileSystemObject""); }" & vbLf
    s = s & "function jkbWrite(ext, text){" & vbLf
    s = s & "  var f = jkbFso().CreateTextFile(JKB_BRIDGE_STEM + ""."" + ext, true, true);" & vbLf
    s = s & "  f.Write(text); f.Close();" & vbLf
    s = s & "}" & vbLf
    s = s & "function jkbReadOnce(ext){" & vbLf
    s = s & "  var fso = jkbFso(), p = JKB_BRIDGE_STEM + ""."" + ext, t = """";" & vbLf
    s = s & "  if (!fso.FileExists(p)) return null;" & vbLf
    s = s & "  try { var st = fso.OpenTextFile(p, 1, false, -1); if (!st.AtEndOfStream) t = st.ReadAll(); st.Close(); } catch(e){ return null; }" & vbLf
    s = s & "  try { fso.DeleteFile(p); } catch(e){}" & vbLf
    s = s & "  return t;" & vbLf
    s = s & "}" & vbLf
    ' JScript in a page cannot block, so a round trip is a callback rather than a
    ' return value. Every caller is written that way.
    s = s & "function jkbCall(macro, args, onDone){" & vbLf
    s = s & "  var i, parts = [macro];" & vbLf
    s = s & "  for (i = 0; i < args.length; i++) parts.push(String(args[i]).replace(/\r/g, """").replace(/\n/g, """"));" & vbLf
    s = s & "  try { jkbWrite(""req"", parts.join("""")); }" & vbLf
    s = s & "  catch(e){ if (onDone) onDone(""ERR:the request could not be written - "" + e.message); return; }" & vbLf
    s = s & "  var tries = 0;" & vbLf
    s = s & "  var h = window.setInterval(function(){" & vbLf
    s = s & "    tries++;" & vbLf
    s = s & "    var r = jkbReadOnce(""res"");" & vbLf
    s = s & "    if (r !== null){ window.clearInterval(h); if (onDone) onDone(r); return; }" & vbLf
    ' Twenty minutes. A folder of extracts is genuinely that slow, and a window
    ' that gives up while Excel is still working is worse than one that waits.
    s = s & "    if (tries > 12000){ window.clearInterval(h); if (onDone) onDone(""ERR:Excel did not answer. It may still be busy - check the workbook.""); }" & vbLf
    s = s & "  }, 100);" & vbLf
    s = s & "}" & vbLf
    s = s & "function jkbDone(msg){ try { jkbWrite(""done"", msg || ""closed""); } catch(e){} }" & vbLf
    s = s & "window.onunload = function(){ jkbDone(""closed""); };" & vbLf
    BridgeJs = s
End Function
