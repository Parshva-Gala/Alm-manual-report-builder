Option Explicit

' ============================================================================
'  Workbooks: how many files a framework is built into.
'
'  One row per framework. Left blank, a framework is one workbook, as it
'  always was. Name a field under "One workbook per" and the build makes one
'  workbook for each of its values - the Maturity ladder, one per currency -
'  and inside each, the sheets Pivot config asks for: a ladder book per
'  currency holds one sheet per rule, as the LCR and NSFR books do.
'
'  The output is read once to find the values, and held open while each
'  workbook stages only its own rows. Every total, tile and sheet in a
'  currency's book is that currency's.
' ============================================================================

Private Const B_FW As Long = 1
Private Const B_PER As Long = 2
Private Const B_FILE As Long = 3
Private Const B_ONLY As Long = 4
Private Const B_MAX As Long = 5
Private Const B_DESC As Long = 6
Private Const B_CHECK As Long = 7
Private Const B_WHY As Long = 8
Private Const B_LAST As Long = 8

Public Const MAX_BOOKS_DEFAULT As Long = 40

Public Function BookHeads() As Variant
    BookHeads = Array("Framework", "One workbook per", "File name", "Only these", "Max workbooks", _
                      "What you get", "Check", "What to fix")
End Function

Private Function BookWidths() As Variant
    BookWidths = Array(18, 22, 26, 30, 14, 64, 10, 60)
End Function

Private Function Frameworks() As Variant
    Frameworks = Array(FW_LCR, FW_NSFR, FW_ML)
End Function

' ===================== the sheet ============================================

Public Sub BuildBooksSheet(Optional ByVal withDefaults As Boolean = False)
    Dim ws As Worksheet, fresh As Boolean, old As Collection
    Set ws = GetSheet(SH_BOOKS)
    fresh = ws Is Nothing
    Set ws = EnsureSheet(SH_BOOKS)
    If Not fresh And Not withDefaults Then
        If Not modPD_Config.HeadersMatch(ws, BookHeads()) Then Set old = modPD_Config.Remember(ws)
    End If
    If fresh Or withDefaults Or Not old Is Nothing Then
        ws.Cells.Clear
        ws.Cells.Validation.Delete
    End If
    modPD_Theme.Dress ws, "Workbooks", _
        "How many files each framework is built into. Blank: one workbook. Name a field, and each of its " & _
        "values gets a workbook of its own - with every sheet Pivot config asks for inside it.", _
        "REPORTS  " & ChrW(183) & "  WORKBOOKS"
    modPD_Theme.Head ws, BookHeads(), BookWidths()
    If Not old Is Nothing Then modPD_Config.PutBack ws, old, BookHeads()
    If fresh Or withDefaults Or Len(SafeText(ws.Cells(modPD_Theme.R_FIRST, B_FW).Value2)) = 0 Then
        WriteDefaultBooks ws
    End If
    DressBooks ws
    BookHints ws
    modPD_Theme.PrintReady ws, B_LAST
    CheckBooks True
End Sub

' LCR and NSFR as one workbook each; the ladder one per currency.
Private Sub WriteDefaultBooks(ByVal ws As Worksheet)
    Dim r As Long
    r = modPD_Theme.R_FIRST
    ws.Range(ws.Cells(r, 1), ws.Cells(r + 2, B_LAST)).NumberFormat = "@"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, B_DESC)).Value2 = _
        Array(FwLabel(FW_LCR), "", "{fw}", "", "", "One workbook: every rule, local and foreign side by side.")
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, B_DESC)).Value2 = _
        Array(FwLabel(FW_NSFR), "", "{fw}", "", "", "One workbook: every rule, local and foreign side by side.")
    ws.Range(ws.Cells(r + 2, 1), ws.Cells(r + 2, B_DESC)).Value2 = _
        Array(FwLabel(FW_ML), H_CURRENCY, "{fw} - {part}", "", CStr(MAX_BOOKS_DEFAULT), _
              "One workbook per currency, biggest first - and in each, one sheet per rule.")
End Sub

Private Sub DressBooks(ByVal ws As Worksheet)
    Dim lastR As Long
    On Error Resume Next
    lastR = modPD_Theme.R_FIRST + UBound(Frameworks())
    modPD_Theme.DressTable ws, B_LAST, lastR
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, 1), ws.Cells(lastR, B_LAST)).NumberFormat = "@"
    With ws.Range(ws.Cells(modPD_Theme.R_FIRST, B_FW), ws.Cells(lastR, B_FW)).Font
        .Name = modPD_Theme.UI_SEMI
        .Color = modPD_Theme.C_TEXT
    End With
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, B_PER), ws.Cells(lastR, B_ONLY)).Font.Name = modPD_Theme.UI_FONT
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, B_DESC), ws.Cells(lastR, B_DESC)).Font.Color = modPD_Theme.C_TEXT_3
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, B_WHY), ws.Cells(lastR, B_WHY)).Font.Color = modPD_Theme.C_TEXT_2
    Err.Clear
End Sub

Private Sub BookHints(ByVal ws As Worksheet)
    Dim r1 As Long, r2 As Long
    r1 = modPD_Theme.R_FIRST
    r2 = r1 + UBound(Frameworks())
    modPD_Config.Hint ws, r1, r2, B_FW, "Framework", "One row per framework - the rows are fixed."
    modPD_Config.ListRule ws, r1, r2, B_PER, H_CURRENCY & "," & H_CCYCLASS & "," & H_RULE_CAT & "," & H_TYPE, _
        "One workbook per", "Blank: one workbook. Or a text field from Pivot fields - each of its values gets a " & _
        "workbook of its own. e.g. Currency.", False
    modPD_Config.Hint ws, r1, r2, B_FILE, "File name", "{fw} is the framework and {part} the value - e.g. " & _
        "{fw} - {part} makes Maturity Ladder - USD.xlsx. With One workbook per, {part} must be in it."
    modPD_Config.Hint ws, r1, r2, B_ONLY, "Only these", "Optional. The values to build, separated by | - " & _
        "e.g. EGP | USD | EUR. Blank builds every value, biggest first."
    modPD_Config.Hint ws, r1, r2, B_MAX, "Max workbooks", "The most workbooks to make. Blank means " & _
        MAX_BOOKS_DEFAULT & "; the smallest values past it are left out, and Activity says which."
    modPD_Config.Hint ws, r1, r2, B_DESC, "What you get", "A note for whoever reads this sheet next."
End Sub

' ===================== reading ==============================================

' fw -> Dictionary(Per, File, Only, Max, Problem, Row)
Public Function PlanFor(ByVal fw As String) As Object
    Dim ws As Worksheet, r As Long, p As Object, fl As Object
    Set ws = GetSheet(SH_BOOKS)
    Set fl = modPD_Config.Fields()
    If Not ws Is Nothing Then
        For r = modPD_Theme.R_FIRST To modPD_Theme.R_FIRST + 10
            If StrComp(NormFw(SafeText(ws.Cells(r, B_FW).Value2)), fw, vbTextCompare) = 0 Then
                Set PlanFor = ParsePlan(ws, r, fw, fl)
                Exit Function
            End If
        Next r
    End If
    ' No sheet, or no row for it: what the defaults say.
    Set p = NewMap()
    p("Row") = 0
    If StrComp(fw, FW_ML, vbTextCompare) = 0 Then
        p("Per") = H_CURRENCY
        p("File") = "{fw} - {part}"
    Else
        p("Per") = ""
        p("File") = "{fw}"
    End If
    Set p("Only") = New Collection
    p("Max") = MAX_BOOKS_DEFAULT
    p("Problem") = ""
    Set PlanFor = p
End Function

Private Function ParsePlan(ByVal ws As Worksheet, ByVal r As Long, ByVal fw As String, ByVal fl As Object) As Object
    Dim p As Object, s As String
    Set p = NewMap()
    Set ParsePlan = p
    p("Row") = r
    p("Per") = SafeText(ws.Cells(r, B_PER).Value2)
    p("File") = SafeText(ws.Cells(r, B_FILE).Value2)
    If Len(p("File")) = 0 Then
        If Len(p("Per")) > 0 Then p("File") = "{fw} - {part}" Else p("File") = "{fw}"
    End If
    Set p("Only") = modPD_Recipe.SplitList(SafeText(ws.Cells(r, B_ONLY).Value2), "|")
    s = SafeText(ws.Cells(r, B_MAX).Value2)
    p("Problem") = ""
    If Len(s) = 0 Then
        p("Max") = MAX_BOOKS_DEFAULT
    ElseIf IsNumeric(s) Then
        p("Max") = CLng(Val(s))
        If CLng(p("Max")) < 1 Or CLng(p("Max")) > 500 Then p("Problem") = "Max workbooks must be between 1 and 500."
    Else
        p("Max") = MAX_BOOKS_DEFAULT
        p("Problem") = "Max workbooks must be a number."
    End If
    If Len(p("Problem")) > 0 Then Exit Function
    If Len(p("Per")) > 0 Then
        If Not fl.Exists(CStr(p("Per"))) Then
            p("Problem") = Chr$(34) & p("Per") & Chr$(34) & " is not on the Pivot fields sheet."
        ElseIf fl(CStr(p("Per")))("Kind") <> "Text" Then
            p("Problem") = "One workbook per takes a text field - " & Chr$(34) & p("Per") & Chr$(34) & " is " & _
                           LCase$(fl(CStr(p("Per")))("Kind")) & "."
        ElseIf InStr(1, p("File"), "{part}", vbTextCompare) = 0 Then
            p("Problem") = "Put {part} in the file name - without it every workbook would be saved over the last."
        ElseIf modPD_Config.Engine() <> "recipes" Then
            p("Problem") = "The 1.0 layout builds one workbook per framework - switch Pivot config back on to split."
        End If
    End If
    If Len(p("Problem")) = 0 And Len(p("Per")) = 0 And p("Only").count > 0 Then
        p("Problem") = "Only these picks values of One workbook per, which is blank."
    End If
End Function

Private Function NormFw(ByVal s As String) As String
    Select Case UCase$(Replace(Replace(Trim$(s), " ", ""), "_", ""))
        Case "LCR": NormFw = FW_LCR
        Case "NSFR": NormFw = FW_NSFR
        Case "MATURITYLADDER", "LADDER", "ML": NormFw = FW_ML
        Case Else: NormFw = s
    End Select
End Function

' The file a workbook is saved as, without its extension.
Public Function FileBase(ByVal plan As Object, ByVal fw As String, ByVal part As String) As String
    Dim s As String
    s = Replace(CStr(plan("File")), "{fw}", FwLabel(fw), , , vbTextCompare)
    s = Replace(s, "{part}", part, , , vbTextCompare)
    ' "{fw} - {part}" with no part: no dangling separator.
    s = Trim$(s)
    Do While Len(s) > 0 And InStr(" -_.", Right$(s, 1)) > 0
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(Trim$(s)) = 0 Then s = FwLabel(fw)
    FileBase = modPD_Build.SafeFileName(Trim$(s))
End Function

' Whether a value gets a workbook: not a blank, and among Only these when
' the row lists some.
Public Function Wanted(ByVal plan As Object, ByVal part As String) As Boolean
    Dim fl As Object, blank As String
    If Len(part) = 0 Then Exit Function
    Set fl = modPD_Config.Fields()
    If fl.Exists(CStr(plan("Per"))) Then blank = CStr(fl(CStr(plan("Per")))("Blank"))
    If Len(blank) > 0 And StrComp(part, blank, vbTextCompare) = 0 Then Exit Function
    If plan("Only").count > 0 Then
        Wanted = modPD_Recipe.InCollection(plan("Only"), part)
    Else
        Wanted = True
    End If
End Function

' ===================== Check ================================================

' The status-bar line for a row just edited - read, never written, so Excel's
' undo survives the edit.
Public Function BookRowSays(ByVal r As Long) As String
    Dim ws As Worksheet, fw As String, p As Object
    Set ws = GetSheet(SH_BOOKS)
    If ws Is Nothing Then Exit Function
    fw = NormFw(SafeText(ws.Cells(r, B_FW).Value2))
    If Len(fw) = 0 Then Exit Function
    Set p = PlanFor(fw)
    If CLng(p("Row")) <> r Then Exit Function
    If Len(p("Problem")) > 0 Then
        BookRowSays = FwLabel(fw) & " will not build: " & p("Problem")
    ElseIf Len(p("Per")) > 0 Then
        BookRowSays = FwLabel(fw) & ": one workbook per " & p("Per") & ", saved as " & _
                  FileBase(p, fw, "[" & p("Per") & "]") & ".xlsx"
    Else
        BookRowSays = FwLabel(fw) & ": one workbook, saved as " & FileBase(p, fw, "") & ".xlsx"
    End If
End Function

' Writes OK / Break beside each row and returns how many will not build.
Public Function CheckBooks(Optional ByVal quiet As Boolean = False) As Long
    Dim ws As Worksheet, fw As Variant, p As Object, r As Long, nBad As Long, nSplit As Long
    Set ws = GetSheet(SH_BOOKS)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    For Each fw In Frameworks()
        Set p = PlanFor(CStr(fw))
        r = CLng(p("Row"))
        If r > 0 Then
            ws.Range(ws.Cells(r, B_CHECK), ws.Cells(r, B_WHY)).ClearContents
            If Len(p("Problem")) > 0 Then
                nBad = nBad + 1
                ws.Cells(r, B_CHECK).Value2 = V_BREAK
                ws.Cells(r, B_WHY).Value2 = p("Problem")
            Else
                ws.Cells(r, B_CHECK).Value2 = V_OK
                If Len(p("Per")) > 0 Then nSplit = nSplit + 1
            End If
            modPD_Theme.PaintVerdict ws.Cells(r, B_CHECK)
        End If
    Next fw
    If nBad > 0 Then
        modPD_Theme.SetStatus ws, nBad & " row(s) will not build - the reason is beside each, under What to fix.", _
            "Break"
    ElseIf nSplit = 0 Then
        modPD_Theme.SetStatus ws, "Every framework is built as one workbook.", "OK"
    Else
        modPD_Theme.SetStatus ws, nSplit & " framework(s) split into one workbook per value; the rest one " & _
            "workbook each.", "OK"
    End If
    If Not quiet Then
        If nBad = 0 Then
            Notify "Every row on Workbooks is ready.", V_OK
        Else
            Notify nBad & " row(s) on Workbooks will not build - see What to fix beside each.", V_BREAK
        End If
    End If
    CheckBooks = nBad
    Err.Clear
End Function

Public Sub PD_BooksCheck()
    If PD_Busy Then Exit Sub
    modPD_Desk.PressFx
    CheckBooks
End Sub

Public Sub PD_BooksDefaults()
    If PD_Busy Then Exit Sub
    If MsgBox("Put Workbooks back to its defaults?" & vbCrLf & vbCrLf & _
              "LCR and NSFR one workbook each; the Maturity ladder one per currency.", _
              vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
    Application.ScreenUpdating = False
    BuildBooksSheet True
    modPD_Theme.Rail GetSheet(SH_BOOKS)
    Application.ScreenUpdating = True
    CheckBooks
End Sub
