Attribute VB_Name = "modShared_Console"
Option Explicit

' ============================================================================
'  ONE console, built from ONE idea.
'
'  Both tools do the same thing:
'
'      take a number someone reported, rebuild it from the raw data,
'      show the difference and the rows behind it.
'
'  Base and pre-shock, portfolio segmentation, bucket placement, weighting
'  factors, control reports, leakage - every one of them is that shape. So there
'  is no per-check display code here, and there is none in either tool.
'
'  The trick that makes it possible is that every check sheet in both workbooks
'  already has the SAME layout:
'
'      row 3  title        row 6  status
'      row 4  about        row 7  header      row 8+ data
'
'  So one reader turns any sheet into a console view. Adding a check to either
'  tool means writing a sheet, which the tool already does; the console picks it
'  up with no change here at all.
'
'  This file is built into BOTH workbooks from shared\. One console to design,
'  one to improve, and the two tools cannot drift apart where a person looks.
' ============================================================================

Private Const CON_TITLE_ROW As Long = 3
Private Const CON_ABOUT_ROW As Long = 4
Private Const CON_STATUS_ROW As Long = 6
Private Const CON_HDR_ROW As Long = 7
Private Const CON_FIRST_ROW As Long = 8
Private Const CON_MAX_ROWS As Long = 4000
Private Const CON_MAX_COLS As Long = 24

' ===================== the one public entry point ===========================
'
' spec is a tiny declaration, because a launcher should be a declaration:
'   tool      what to call it
'   version   shown under the name
'   asOf      the reporting date, if there is one
'   sheets    "Group/SheetName" per item - the group heads the rail
'   tiles     built by the caller with Tile()
'   flow      built by the caller with Hop()
Public Function OpenConsole(ByVal toolName As String, ByVal version As String, ByVal asOf As String, _
                            ByVal sheets As Variant, ByVal tiles As Collection, ByVal flow As Collection, _
                            ByVal actions As Collection, Optional ByVal accent As String = "#173763", _
                            Optional ByVal accentDark As String = "#12243B") As String
    Dim fso As Object, folder As String, token As String, jsPath As String, htaPath As String
    Dim stream As Object, html As String, payload As String

    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetSpecialFolder(2)
    token = format$(Now, "yyyymmddhhnnss")
    CleanOld fso, folder
    jsPath = fso.BuildPath(folder, "JKB_Console_Payload_" & token & ".js")
    htaPath = fso.BuildPath(folder, "JKB_Console_" & token & ".hta")

    payload = BuildPayload(toolName, version, asOf, sheets, tiles, flow, actions, accent, accentDark)
    Set stream = fso.CreateTextFile(jsPath, True, True)
    stream.Write payload
    stream.Close
    Set stream = Nothing

    Dim bridge As String, allowed As String, a As Variant
    bridge = modShared_Bridge.BridgeNewToken()
    ' The window may run the actions it was given and nothing else. The request
    ' file is written by a page, and a page is not something to hand an
    ' unrestricted Application.Run.
    If Not actions Is Nothing Then
        For Each a In actions
            allowed = allowed & IIf(Len(allowed) > 0, "|", "") & CStr(a("macro"))
        Next a
    End If

    html = ConsoleHtml(toolName, fso.GetFileName(jsPath))
    html = Replace(html, "__BRIDGE_JS__", modShared_Bridge.BridgeJs(bridge))
    Set stream = fso.CreateTextFile(htaPath, True, True)
    stream.Write html
    stream.Close
    Set stream = Nothing

    ' A headless run builds the files and stops there. Popping a window from a
    ' build script would be a window nobody is there to close.
    If Application.Visible Then
        On Error Resume Next
        Shell "mshta.exe """ & htaPath & """", vbNormalFocus
        If Err.Number <> 0 Then
            ' mshta is blocked in some environments. The browser still renders the
            ' console; only the write-back into Excel is lost.
            Err.Clear
            modShared_Bridge.BridgeCleanUp bridge
            ThisWorkbook.FollowHyperlink address:=htaPath, NewWindow:=True
            On Error GoTo 0
            OpenConsole = htaPath & "|" & jsPath
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        ' Excel listens for as long as the console is open. It has to: an HTA can
        ' no longer reach back into a running Excel, so the workbook does the
        ' waiting and the window does the asking. Closing the console releases it.
        modShared_Bridge.BridgePump bridge, allowed
        modShared_Bridge.BridgeCleanUp bridge
    End If
    OpenConsole = htaPath & "|" & jsPath
    Exit Function

Failed:
    MsgBox "The review console could not be opened." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, toolName
End Function

' ===================== builders the launchers use ===========================

Public Function Tile(ByVal label As String, ByVal value As Variant, ByVal format As String, _
                     ByVal verdict As String, ByVal Note As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("label") = label: d("value") = value: d("format") = format
    d("verdict") = verdict: d("note") = Note
    Set Tile = d
End Function

Public Function Hop(ByVal fromName As String, ByVal toName As String, ByVal amountIn As Double, _
                    ByVal amountOut As Double, ByVal dropped As Double, ByVal verdict As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("from") = fromName: d("to") = toName
    d("amountIn") = amountIn: d("amountOut") = amountOut
    d("dropped") = dropped: d("verdict") = verdict
    Set Hop = d
End Function

Public Function Act(ByVal label As String, ByVal macro As String, ByVal primary As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("label") = label: d("macro") = macro: d("primary") = primary
    Set Act = d
End Function

' ===================== the payload ==========================================

Private Function BuildPayload(ByVal toolName As String, ByVal version As String, ByVal asOf As String, _
                              ByVal sheets As Variant, ByVal tiles As Collection, ByVal flow As Collection, _
                              ByVal actions As Collection, ByVal accent As String, _
                              ByVal accentDark As String) As String
    Dim sb As String, i As Long, first As Boolean, item As Variant, parts As Variant

    sb = "window.JKB_PAYLOAD={"
    sb = sb & """tool"":" & j(toolName)
    sb = sb & ",""version"":" & j(version)
    sb = sb & ",""asOf"":" & j(asOf)
    sb = sb & ",""ranAt"":" & j(format$(Now, "dd-mmm-yyyy HH:mm"))
    sb = sb & ",""accent"":" & j(accent) & ",""accentDark"":" & j(accentDark)

    sb = sb & ",""tiles"":["
    first = True
    If Not tiles Is Nothing Then
        For Each item In tiles
            If Not first Then sb = sb & ","
            first = False
            sb = sb & "{""label"":" & j(CStr(item("label"))) & ",""value"":" & Num(item("value")) & _
                      ",""format"":" & j(CStr(item("format"))) & ",""verdict"":" & j(CStr(item("verdict"))) & _
                      ",""note"":" & j(CStr(item("note"))) & "}"
        Next item
    End If
    sb = sb & "]"

    sb = sb & ",""flow"":["
    first = True
    If Not flow Is Nothing Then
        For Each item In flow
            If Not first Then sb = sb & ","
            first = False
            sb = sb & "{""from"":" & j(CStr(item("from"))) & ",""to"":" & j(CStr(item("to"))) & _
                      ",""amountIn"":" & Num(item("amountIn")) & ",""amountOut"":" & Num(item("amountOut")) & _
                      ",""dropped"":" & Num(item("dropped")) & ",""verdict"":" & j(CStr(item("verdict"))) & "}"
        Next item
    End If
    sb = sb & "]"

    sb = sb & ",""actions"":["
    first = True
    If Not actions Is Nothing Then
        For Each item In actions
            If Not first Then sb = sb & ","
            first = False
            sb = sb & "{""label"":" & j(CStr(item("label"))) & ",""macro"":" & j(CStr(item("macro"))) & _
                      ",""primary"":" & LCase$(CStr(CBool(item("primary")))) & "}"
        Next item
    End If
    sb = sb & "]"

    sb = sb & ",""sections"":["
    first = True
    For i = LBound(sheets) To UBound(sheets)
        parts = Split(CStr(sheets(i)) & "/", "/")
        Dim one As String
        one = SectionJson(CStr(parts(1)), CStr(parts(0)), (i = LBound(sheets)))
        If Len(one) > 0 Then
            If Not first Then sb = sb & ","
            first = False
            sb = sb & one
        End If
    Next i
    sb = sb & "]};"
    BuildPayload = sb
End Function

' One sheet becomes one section. No knowledge of what the sheet is about.
Private Function SectionJson(ByVal sheetName As String, ByVal groupName As String, ByVal isFirst As Boolean) As String
    Dim ws As Worksheet, lastR As Long, lastC As Long, r As Long, c As Long
    Dim sb As String, types() As String, heads() As String, n As Long, first As Boolean
    Dim verdict As String, v As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastC = 0
    For c = 1 To CON_MAX_COLS
        If Len(Trim$(CStr(ws.Cells(CON_HDR_ROW, c).Value2 & ""))) > 0 Then lastC = c
    Next c
    If lastC = 0 Then Exit Function

    ReDim heads(1 To lastC)
    ReDim types(1 To lastC)
    For c = 1 To lastC
        heads(c) = Trim$(CStr(ws.Cells(CON_HDR_ROW, c).Value2 & ""))
    Next c

    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastR > CON_FIRST_ROW + CON_MAX_ROWS Then lastR = CON_FIRST_ROW + CON_MAX_ROWS
    If lastR < CON_FIRST_ROW Then lastR = CON_FIRST_ROW - 1

    ' Column types are read off the data, not declared. A column whose cells are
    ' numbers is a number; a column headed Verdict or Status is a verdict. That is
    ' the whole type system, and it has never needed to be more than that.
    For c = 1 To lastC
        types(c) = ColumnType(ws, c, lastR, heads(c))
    Next c

    ' A sheet is "one of ours" if it carries the convention, which shows in row 4:
    ' every sheet the tools build says there what it is for. A sheet that does not -
    ' an older log, say - is still shown, but titled by its own name rather than by
    ' whatever happens to sit in row 3.
    Dim title As String, about As String
    about = CellText(ws, CON_ABOUT_ROW, 1, "")
    If Len(about) > 0 Then
        title = CellText(ws, CON_TITLE_ROW, 1, Replace(sheetName, "_", " "))
    Else
        title = Replace(sheetName, "_", " ")
    End If

    sb = "{""id"":" & j(SafeId(sheetName)) & ",""group"":" & j(groupName) & _
         ",""title"":" & j(title) & _
         ",""about"":" & j(about) & _
         ",""status"":" & j(CellText(ws, CON_STATUS_ROW, 1, "")) & _
         ",""overview"":" & LCase$(CStr(isFirst))

    sb = sb & ",""verdict"":" & j(SectionVerdict(ws, lastR, lastC, types))

    sb = sb & ",""columns"":["
    For c = 1 To lastC
        If c > 1 Then sb = sb & ","
        sb = sb & "{""label"":" & j(heads(c)) & ",""type"":" & j(types(c)) & "}"
    Next c
    sb = sb & "]"

    sb = sb & ",""rows"":["
    first = True
    For r = CON_FIRST_ROW To lastR
        If RowHasContent(ws, r, lastC) Then
            If Not first Then sb = sb & ","
            first = False
            sb = sb & "["
            For c = 1 To lastC
                If c > 1 Then sb = sb & ","
                v = ws.Cells(r, c).Value2
                If types(c) = "money" Or types(c) = "count" Or types(c) = "pct" Then
                    sb = sb & Num(v)
                Else
                    sb = sb & j(CellText(ws, r, c, ""))
                End If
            Next c
            sb = sb & "]"
            n = n + 1
        End If
    Next r
    sb = sb & "]}"
    SectionJson = sb
End Function

Private Function ColumnType(ByVal ws As Worksheet, ByVal c As Long, ByVal lastR As Long, ByVal head As String) As String
    Dim r As Long, seen As Long, nums As Long, v As Variant, h As String
    h = UCase$(head)
    If InStr(h, "VERDICT") > 0 Or h = "STATUS" Or h = "ON" Then ColumnType = "verdict": Exit Function
    If InStr(h, "SHARE") > 0 Or InStr(h, "%") > 0 Then ColumnType = "pct": Exit Function
    For r = CON_FIRST_ROW To lastR
        v = ws.Cells(r, c).Value2
        If Not IsEmpty(v) Then
            If Len(Trim$(CStr(v) & "")) > 0 Then
                seen = seen + 1
                If IsNumeric(v) Then nums = nums + 1
            End If
        End If
        If seen >= 40 Then Exit For
    Next r
    If seen > 0 Then
        If nums / seen >= 0.8 Then
            If InStr(h, "ROW") > 0 Or InStr(h, "COUNT") > 0 Or InStr(h, "ORDER") > 0 Or InStr(h, "RUNG") > 0 Then
                ColumnType = "count"
            Else
                ColumnType = "money"
            End If
            Exit Function
        End If
    End If
    ColumnType = "text"
End Function

' The worst verdict on the sheet is the sheet's verdict, which is what a status
' dot in the rail should mean.
Private Function SectionVerdict(ByVal ws As Worksheet, ByVal lastR As Long, ByVal lastC As Long, ByRef types() As String) As String
    Dim r As Long, c As Long, t As String, worst As Long, rank As Long
    For c = 1 To lastC
        If types(c) = "verdict" Then
            For r = CON_FIRST_ROW To lastR
                t = UCase$(Trim$(CStr(ws.Cells(r, c).Value2 & "")))
                rank = 0
                If t Like "BREAK*" Or t Like "FAIL*" Or t Like "MISMATCH*" Or t Like "MISSING*" Then rank = 3
                If rank = 0 Then
                    If t Like "CHECK*" Or t Like "PARTIAL*" Or t Like "WARN*" Or t Like "UNMAPPED*" Or t Like "COVERAGE*" Then rank = 2
                End If
                If rank = 0 And Len(t) > 0 Then rank = 1
                If rank > worst Then worst = rank
                If worst = 3 Then Exit For
            Next r
        End If
        If worst = 3 Then Exit For
    Next c
    Select Case worst
        Case 3: SectionVerdict = "Break"
        Case 2: SectionVerdict = "Check"
        Case 1: SectionVerdict = "Clean"
        Case Else: SectionVerdict = ""
    End Select
End Function

Private Function RowHasContent(ByVal ws As Worksheet, ByVal r As Long, ByVal lastC As Long) As Boolean
    Dim c As Long
    For c = 1 To lastC
        If Len(Trim$(CStr(ws.Cells(r, c).Value2 & ""))) > 0 Then RowHasContent = True: Exit Function
    Next c
End Function

Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal fallback As String) As String
    Dim s As String
    On Error Resume Next
    s = Trim$(CStr(ws.Cells(r, c).text))
    If Len(s) = 0 Then s = Trim$(CStr(ws.Cells(r, c).Value2 & ""))
    Err.Clear
    On Error GoTo 0
    If Len(s) = 0 Then s = fallback
    CellText = s
End Function

Private Function SafeId(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        Else
            out = out & "_"
        End If
    Next i
    SafeId = LCase$(out)
End Function

' ===================== json =================================================

Private Function j(ByVal s As String) As String
    Dim i As Long, n As Long, out As String
    For i = 1 To Len(s)
        n = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case n
            Case 34: out = out & "\"""
            Case 92: out = out & "\\"
            Case 0 To 31, 38, 60, 62, 8232, 8233: out = out & "\u" & Right$("0000" & Hex$(n), 4)
            Case Else: out = out & ChrW$(n)
        End Select
    Next i
    j = """" & out & """"
End Function

Private Function Num(ByVal v As Variant) As String
    On Error GoTo Zero
    If IsEmpty(v) Then Num = "null": Exit Function
    If IsError(v) Then Num = "null": Exit Function
    If Not IsNumeric(v) Then Num = "null": Exit Function
    Num = Replace(CStr(CDbl(v)), ",", ".")
    If InStr(Num, "E") > 0 Or InStr(Num, "e") > 0 Then Num = format$(CDbl(v), "0.############")
    Exit Function
Zero:
    Num = "null"
End Function

' ===================== the html =============================================

' The console itself lives in shared\console.hta and is injected into a very
' hidden sheet at build time, exactly as the condition builder is.
Private Function ConsoleHtml(ByVal toolName As String, ByVal payloadFile As String) As String
    Dim ws As Worksheet, lastR As Long, a As Variant, r As Long, sb As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_Console_Src")
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 900, , "The console source is missing from this workbook (sheet _Console_Src). Rebuild the tool."
    End If
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    a = ws.Range(ws.Cells(1, 1), ws.Cells(lastR, 1)).Value2
    If Not IsArray(a) Then
        sb = CStr(a)
    Else
        For r = 1 To UBound(a, 1)
            sb = sb & CStr(a(r, 1)) & vbCrLf
        Next r
    End If
    sb = Replace(sb, "__PAYLOAD_FILE__", payloadFile)
    sb = Replace(sb, "__TITLE__", toolName)
    ConsoleHtml = sb
End Function

Private Sub CleanOld(ByVal fso As Object, ByVal folder As String)
    Dim f As Object
    On Error Resume Next
    For Each f In fso.GetFolder(folder).files
        If Left$(f.name, 12) = "JKB_Console_" Then
            If DateDiff("h", f.DateLastModified, Now) > 6 Then f.Delete True
        End If
    Next f
    Err.Clear
End Sub
