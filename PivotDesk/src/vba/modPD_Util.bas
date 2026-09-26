Option Explicit

' ============================================================================
'  The small shared things, in one place so no module carries a private copy -
'  which is how two copies come to disagree.
' ============================================================================

' The log keeps this many entries; older ones fall off the bottom. A log
' nobody can scroll to the end of is not read.
Private Const LOG_KEEP As Long = 2000

Public Function NewMap() As Object
    Set NewMap = CreateObject("Scripting.Dictionary")
    NewMap.CompareMode = 1          ' TextCompare: keys differing only in case are one key
End Function

Public Function SafeText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    SafeText = Trim$(CStr(v))
    ' The exports write a literal "NULL" for an empty cell, and carrying that
    ' through puts the word NULL on a pivot row label.
    If StrComp(SafeText, "NULL", vbTextCompare) = 0 Then SafeText = ""
    Err.Clear
End Function

Public Function SafeUpper(ByVal v As Variant) As String
    SafeUpper = UCase$(SafeText(v))
End Function

Public Function SafeNum(ByVal v As Variant) As Double
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsNumeric(v) Then SafeNum = CDbl(v)
    Err.Clear
End Function

' A heading reduced to the one spelling everything compares against: upper case,
' every run of anything that is not a letter or digit becoming one underscore.
Public Function NormKey(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, lastU As Boolean
    s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch: lastU = False
        ElseIf Not lastU And Len(out) > 0 Then
            out = out & "_": lastU = True
        End If
    Next i
    If Right$(out, 1) = "_" Then out = Left$(out, Len(out) - 1)
    NormKey = out
End Function

Public Function Fmt(ByVal v As Double) As String
    Fmt = Format$(v, "#,##0")
End Function

Public Function FileLeaf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then FileLeaf = Mid$(p, i + 1) Else FileLeaf = p
End Function

Public Function FolderOf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then FolderOf = Left$(p, i - 1)
End Function

' The as-of date as a person writes it, whatever shape the cell handed over:
' a serial number (what .Value2 gives for a date), a Date, or text already.
Public Function AsOfText(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then AsOfText = Format$(v, "d mmm yyyy"): Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 20000 And CDbl(v) < 80000 Then
            AsOfText = Format$(CDate(CDbl(v)), "d mmm yyyy")
            Exit Function
        End If
    End If
    AsOfText = SafeText(v)
    Err.Clear
End Function

Public Function PathJoin(ByVal folder As String, ByVal leaf As String) As String
    If Right$(folder, 1) = "\" Then PathJoin = folder & leaf Else PathJoin = folder & "\" & leaf
End Function

' OneDrive hands out https:// paths for synced folders, and Workbooks.Open on
' one of those fails with "cannot access the file" rather than anything useful.
Public Function IsLocalPath(ByVal p As String) As Boolean
    If Len(p) < 3 Then Exit Function
    If LCase$(Left$(p, 4)) = "http" Then Exit Function
    IsLocalPath = True
End Function

Public Function GetSheet(ByVal nm As String, Optional ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set GetSheet = wb.Worksheets(nm)
    Err.Clear
End Function

Public Function EnsureSheet(ByVal nm As String, Optional ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = GetSheet(nm, wb)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        On Error Resume Next
        ws.Name = nm
        Err.Clear
    End If
    Set EnsureSheet = ws
End Function

Public Sub KillSheet(ByVal nm As String, Optional ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    Set ws = GetSheet(nm, wb)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Err.Clear
End Sub

Public Function LastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    On Error Resume Next
    LastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row
    Err.Clear
End Function

' A sheet name Excel will accept: 31 characters, none of the seven it forbids.
Public Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, b As Variant, o As String
    o = s
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each b In bad
        o = Replace(o, CStr(b), " ")
    Next b
    Do While InStr(o, "  ") > 0
        o = Replace(o, "  ", " ")
    Loop
    o = Trim$(o)
    If Len(o) > 31 Then o = Trim$(Left$(o, 31))
    If Len(o) = 0 Then o = "sheet"
    SafeSheetName = o
End Function

' A name Excel will accept AND that is free in this workbook, remembering what
' it gave out so two callers cannot be handed the same one.
Public Function FreeSheetName(ByVal wanted As String, ByVal wb As Workbook) As String
    Dim stem As String, nm As String, i As Long
    stem = SafeSheetName(wanted)
    nm = stem
    i = 1
    Do While Not GetSheet(nm, wb) Is Nothing
        i = i + 1
        nm = SafeSheetName(Left$(stem, 28 - Len(CStr(i))) & " " & i)
        If i > 500 Then Exit Do
    Loop
    FreeSheetName = nm
End Function

' ===================== the activity log =====================================

Public Sub LogIt(ByVal level As String, ByVal stage As String, ByVal msg As String, _
                 Optional ByVal ctx As String = "")
    Dim ws As Worksheet, r As Long, lastR As Long
    On Error Resume Next
    Set ws = GetSheet(SH_LOG)
    If ws Is Nothing Then Exit Sub
    ' Newest first: the thing that just happened is the thing being looked for.
    '
    ' Formatted from BELOW. The default takes the row above - which is the
    ' black table header - and every entry came out looking like a heading.
    ws.Rows(modPD_Theme.R_FIRST).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    r = modPD_Theme.R_FIRST
    ws.Cells(r, 1).Value2 = Format$(Now, "dd mmm  hh:nn:ss")
    ws.Cells(r, 2).Value2 = level
    ws.Cells(r, 3).Value2 = stage
    ws.Cells(r, 4).Value2 = msg
    ws.Cells(r, 5).Value2 = ctx
    modPD_Theme.DressLogRow ws, r
    lastR = LastRow(ws, 4)
    If lastR > modPD_Theme.R_FIRST + LOG_KEEP Then
        ws.Range(ws.Rows(modPD_Theme.R_FIRST + LOG_KEEP), ws.Rows(lastR)).Delete
    End If
    Err.Clear
End Sub

' ===================== settings =============================================
'
' What the Desk remembers between sessions - which frameworks are switched on,
' where the last build went, what the last reconciliation found - lives on a
' very hidden sheet as plain key / value rows. A sheet rather than document
' properties or the registry: it travels with the workbook, a person can read
' it, and nothing about it needs permission.

Private Function SettingsSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = GetSheet(SH_SETTINGS)
    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = SH_SETTINGS
        ws.Cells(1, 1).Value2 = "Key"
        ws.Cells(1, 2).Value2 = "Value"
        ws.visible = xlSheetVeryHidden
        Err.Clear
        On Error GoTo 0
    End If
    Set SettingsSheet = ws
End Function

Private Function SettingRow(ByVal ws As Worksheet, ByVal key As String) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For r = 2 To lastR
        If StrComp(SafeText(ws.Cells(r, 1).Value2), key, vbTextCompare) = 0 Then SettingRow = r: Exit Function
    Next r
End Function

Public Function SettingGet(ByVal key As String, Optional ByVal dflt As String = "") As String
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    SettingGet = dflt
    Set ws = SettingsSheet()
    If ws Is Nothing Then Exit Function
    r = SettingRow(ws, key)
    If r > 0 Then SettingGet = CStr(ws.Cells(r, 2).Value2)
    Err.Clear
End Function

Public Sub SettingSet(ByVal key As String, ByVal value As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SettingsSheet()
    If ws Is Nothing Then Exit Sub
    r = SettingRow(ws, key)
    If r = 0 Then r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value2 = key
    ' Text, always: a folder called 2025 or a verdict called "1 of 2" must come
    ' back exactly as it went in, not as a number or a date.
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value2 = value
    Err.Clear
End Sub

' Every key starting with a prefix, gone - how a reset forgets results
' without forgetting preferences.
Public Sub SettingClear(ByVal prefix As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = SettingsSheet()
    If ws Is Nothing Then Exit Sub
    For r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row To 2 Step -1
        If StrComp(Left$(SafeText(ws.Cells(r, 1).Value2), Len(prefix)), prefix, vbTextCompare) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r
    Err.Clear
End Sub

' Small numbers as words, the way a sentence reads them.
Public Function Words(ByVal n As Long) As String
    Select Case n
        Case 0: Words = "No"
        Case 1: Words = "One"
        Case 2: Words = "Two"
        Case 3: Words = "Three"
        Case 4: Words = "Four"
        Case 5: Words = "Five"
        Case Else: Words = CStr(n)
    End Select
End Function

' 412806 -> "412.8 k", 30077375161 -> "30.1 bn". For tiles, never tables.
Public Function Compact(ByVal v As Double) As String
    Dim a As Double
    a = Abs(v)
    Select Case True
        Case a >= 1000000000#: Compact = Format$(v / 1000000000#, "0.0") & " bn"
        Case a >= 1000000#: Compact = Format$(v / 1000000#, "0.0") & " m"
        Case a >= 10000#: Compact = Format$(v / 1000#, "0.0") & " k"
        Case Else: Compact = Format$(v, "#,##0")
    End Select
End Function

' A long file name shortened in the middle, so the start and the extension -
' the two parts a person recognises - both survive.
Public Function MidTrim(ByVal s As String, ByVal maxLen As Long) As String
    Dim keepR As Long, keepL As Long
    If Len(s) <= maxLen Or maxLen < 8 Then MidTrim = s: Exit Function
    keepR = 9
    keepL = maxLen - keepR - 1
    MidTrim = Left$(s, keepL) & ChrW(8230) & Right$(s, keepR)
End Function

' ===================== application state ====================================

Public Function CaptureState() As Object
    Dim s As Object
    Set s = NewMap()
    On Error Resume Next
    s("Calc") = Application.Calculation
    s("Screen") = Application.ScreenUpdating
    s("Events") = Application.EnableEvents
    s("Alerts") = Application.DisplayAlerts
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Err.Clear
    Set CaptureState = s
End Function

Public Sub RestoreState(ByVal s As Object)
    On Error Resume Next
    If s Is Nothing Then
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.StatusBar = False
        Exit Sub
    End If
    Application.Calculation = s("Calc")
    Application.ScreenUpdating = CBool(s("Screen"))
    Application.EnableEvents = CBool(s("Events"))
    Application.DisplayAlerts = CBool(s("Alerts"))
    Application.StatusBar = False
    Err.Clear
End Sub

' ===================== tidy labels ==========================================
'
' The outputs name lines the way the ledger codes them:
'
'     1.07.00.MBGL.1360.LOANS TO CUSTOMERS
'     COA_MBGL.7220.PERFORMANCE
'     4.01.00.MBGL.6510.CONTINGENT LIABILITIES& COMMITMENTS
'
' A field's Labels setting on Pivot fields says what to do with them:
'
'   0  As is
'   1  Drop codes                 LOANS TO CUSTOMERS
'   2  Drop codes, title case     Loans to Customers
'   3  Title case                 (the code, if any, stays)
'
' Title case keeps what should stay in capitals - acronyms (ECL, FVTOCI, LCY
' and any listed on Pivot fields), single letters (T.Bills), anything with a
' digit - and leaves alone any word that already has lower case in it.
Public Function TidyLabel(ByVal s As String, ByVal mode As Long, Optional ByVal keep As String = "") As String
    Dim t As String
    TidyLabel = s
    If mode <= 0 Or Len(s) = 0 Then Exit Function
    t = Trim$(s)
    If mode = 1 Or mode = 2 Then t = DropCode(t)
    t = FixAmpersand(t)
    If mode >= 2 Then t = TitleWords(t, keep)
    TidyLabel = t
End Function

' The code in front: dotted pieces with no spaces, the last of which to carry
' a digit ends the code. At least two such pieces, and never the last piece -
' so "2.5% RESERVE" and "T.Bills" are left whole.
Private Function DropCode(ByVal t As String) As String
    Dim parts As Variant, i As Long, k As Long, j As Long, out As String, u As Long
    DropCode = t
    ' 3686_LTL UNSECURED LEASING: digits and an underscore in front.
    u = InStr(t, "_")
    If u > 2 And u < Len(t) Then
        If IsAllDigits(Left$(t, u - 1)) Then DropCode = Trim$(Mid$(t, u + 1)): Exit Function
    End If
    If InStr(t, ".") = 0 Then Exit Function
    parts = Split(t, ".")
    k = -1
    For i = 0 To UBound(parts) - 1
        If Not IsCodePiece(CStr(parts(i))) Then Exit For
        If HasDigit(CStr(parts(i))) Then k = i
    Next i
    If k < 1 Then Exit Function
    For j = k + 1 To UBound(parts)
        If j > k + 1 Then out = out & "."
        out = out & CStr(parts(j))
    Next j
    out = Trim$(out)
    If Len(out) > 0 Then DropCode = out
End Function

Private Function IsCodePiece(ByVal p As String) As Boolean
    Dim i As Long, ch As String
    If Len(p) = 0 Or Len(p) > 12 Then Exit Function
    For i = 1 To Len(p)
        ch = Mid$(p, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or ch = "_" Or ch = "-") Then Exit Function
    Next i
    IsCodePiece = True
End Function

Private Function IsAllDigits(ByVal p As String) As Boolean
    Dim i As Long, ch As String
    If Len(p) = 0 Then Exit Function
    For i = 1 To Len(p)
        ch = Mid$(p, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

' LTL, STL, CRM: a word with no vowel is an abbreviation, kept in capitals.
Private Function NoVowel(ByVal u As String) As Boolean
    Dim i As Long
    For i = 1 To Len(u)
        If InStr("AEIOUY", Mid$(u, i, 1)) > 0 Then Exit Function
    Next i
    NoVowel = (Len(u) > 1)
End Function

Private Function HasDigit(ByVal p As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(p)
        ch = Mid$(p, i, 1)
        If ch >= "0" And ch <= "9" Then HasDigit = True: Exit Function
    Next i
End Function

' "LIABILITIES& COMMITMENTS" -> "LIABILITIES & COMMITMENTS"; "P&L" is left as
' it is. Runs of spaces become one.
Private Function FixAmpersand(ByVal t As String) As String
    Dim i As Long, ch As String, out As String, lft As String, rgt As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = "&" Then
            If i > 1 Then lft = Mid$(t, i - 1, 1) Else lft = " "
            If i < Len(t) Then rgt = Mid$(t, i + 1, 1) Else rgt = " "
            If (lft = " ") <> (rgt = " ") Then ch = " & "
        End If
        out = out & ch
    Next i
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    FixAmpersand = Trim$(out)
End Function

Private Function TitleWords(ByVal t As String, ByVal keep As String) As String
    Dim words As Variant, i As Long, out As String, w As String
    words = Split(t, " ")
    For i = 0 To UBound(words)
        w = CStr(words(i))
        If Len(w) > 0 Then w = TitleWord(w, i = 0, UBound(words) = 0, keep)
        If i > 0 Then out = out & " "
        out = out & w
    Next i
    TitleWords = out
End Function

' One word: pieces between - / . ' ( ) & each cased on their own.
Private Function TitleWord(ByVal w As String, ByVal first As Boolean, ByVal only As Boolean, _
                           ByVal keep As String) As String
    Dim i As Long, ch As String, piece As String, out As String, afterApos As Boolean, nPieces As Long
    If w <> UCase$(w) Then TitleWord = w: Exit Function          ' already has lower case
    If HasDigit(w) Then TitleWord = w: Exit Function
    For i = 1 To Len(w) + 1
        If i <= Len(w) Then ch = Mid$(w, i, 1) Else ch = ""
        If ch = "-" Or ch = "/" Or ch = "." Or ch = "'" Or ch = "(" Or ch = ")" Or ch = "&" Or ch = "" Then
            If Len(piece) > 0 Then
                nPieces = nPieces + 1
                out = out & CasePiece(piece, first And nPieces = 1, afterApos, _
                                      Len(w) = Len(piece) And Not only, keep)
            End If
            afterApos = (ch = "'")
            out = out & ch
            piece = ""
        Else
            piece = piece & ch
        End If
    Next i
    TitleWord = out
End Function

Private Function CasePiece(ByVal p As String, ByVal first As Boolean, ByVal afterApos As Boolean, _
                           ByVal wholeWord As Boolean, ByVal keep As String) As String
    Dim u As String
    u = UCase$(p)
    If afterApos Then CasePiece = LCase$(p): Exit Function          ' LC'S -> LC's
    If KeptCaps(u, keep) Or NoVowel(u) Then CasePiece = u: Exit Function
    If Len(u) = 1 And Not (wholeWord And u = "A") Then CasePiece = u: Exit Function
    If wholeWord And Not first And IsSmallWord(u) Then CasePiece = LCase$(p): Exit Function
    CasePiece = UCase$(Left$(p, 1)) & LCase$(Mid$(p, 2))
End Function

Private Function IsSmallWord(ByVal u As String) As Boolean
    IsSmallWord = (InStr(1, " A AN AND AS AT BY FOR FROM IN INTO OF ON OR PER THE TO VIA VS WITH ", _
                         " " & u & " ", vbBinaryCompare) > 0)
End Function

' Acronyms kept in capitals: the ones every ALM output uses, and any listed
' under "Words kept in capitals" on Pivot fields.
Public Function KeptCaps(ByVal u As String, Optional ByVal extra As String = "") As Boolean
    Const BUILT_IN As String = " ALM ASF AC AED ATM BHD BV CASA CBE CD CDS CHF CNY COA CR DR EAD ECL EGP EMI EU EUR " & _
        "FCY FVOCI FVTOCI FVTPL FX GBP GL HQLA ID IFRS INT IRS JOD JPY KSA KWD KYC LC LCR LCS LCY LG LGD LGS MBGL MM " & _
        "NBE NII NIM NPA NPL NSFR OCI OD ODS OMR OVD PD POS QAR ROA ROE RSF SAR SME SMES SPV TB TD TDS UAE UK US USD VAT "
    If InStr(1, BUILT_IN, " " & u & " ", vbBinaryCompare) > 0 Then KeptCaps = True: Exit Function
    If Len(extra) > 0 Then
        KeptCaps = (InStr(1, " " & UCase$(Replace(Replace(extra, ",", " "), ";", " ")) & " ", " " & u & " ", _
                          vbBinaryCompare) > 0)
    End If
End Function
