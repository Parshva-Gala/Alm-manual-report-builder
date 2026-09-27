Attribute VB_Name = "modScenarioBuilder_Core"
Option Explicit

' The ECL pre-shock subsystem now lives in modScenarioBuilder_PreShock, which owns the
' single ECL_Pre_Shock worksheet and the loaded ECL extract.

' Formula Studio V2 / manual-formula natural logic translator.
' Supports natural IF ... THEN ... ELSE (including ELSE IF nesting) and infix AND / OR.
' It deliberately leaves row tokens symbolic; ConvertManualFormula resolves them afterwards.
Public Function NormalizeNaturalFormulaLogic(ByVal formula As String, ByVal context As String) As String
    Dim t As String
    t = Trim$(formula)
    If Len(t) = 0 Then Exit Function
    t = LogicTranslateNaturalIf(t, context)
    t = LogicNormalizeBooleanExpr(t, context)
    NormalizeNaturalFormulaLogic = t
End Function

Public Function TranslateManualFormulaSyntax(ByVal formula As String, ByVal context As String) As String
    TranslateManualFormulaSyntax = NormalizeNaturalFormulaLogic(NormalizeConfigFormulaText(formula), context)
End Function

Private Function LogicTranslateNaturalIf(ByVal s As String, ByVal context As String) As String
    Dim t As String, pThen As Long, pElse As Long, cond As String, yesPart As String, noPart As String
    t = Trim$(s)
    If Len(t) < 3 Or UCase$(Left$(t, 3)) <> "IF " Then LogicTranslateNaturalIf = t: Exit Function
    pThen = LogicFindTopKeyword(t, "THEN", 3)
    If pThen = 0 Then Err.Raise vbObjectError + 580, , "Natural IF requires THEN: " & context
    pElse = LogicFindNaturalElse(t, pThen + 4)
    If pElse = 0 Then Err.Raise vbObjectError + 580, , "Natural IF requires ELSE: " & context
    cond = Trim$(Mid$(t, 3, pThen - 3))
    yesPart = Trim$(Mid$(t, pThen + 4, pElse - (pThen + 4)))
    noPart = Trim$(Mid$(t, pElse + 4))
    If Len(cond) = 0 Or Len(yesPart) = 0 Or Len(noPart) = 0 Then Err.Raise vbObjectError + 580, , "Natural IF has an empty condition/branch: " & context
    yesPart = LogicTranslateNaturalIf(yesPart, context)
    noPart = LogicTranslateNaturalIf(noPart, context)
    LogicTranslateNaturalIf = "IF(" & LogicNormalizeBooleanExpr(cond, context) & "," & LogicNormalizeBooleanExpr(yesPart, context) & "," & LogicNormalizeBooleanExpr(noPart, context) & ")"
End Function

Private Function LogicFindTopKeyword(ByVal s As String, ByVal keyword As String, ByVal startPos As Long) As Long
    Dim i As Long, depth As Long, quoted As Boolean, ch As String, klen As Long
    klen = Len(keyword)
    For i = IIf(startPos < 1, 1, startPos) To Len(s) - klen + 1
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If quoted And Mid$(s, i + 1, 1) = """" Then i = i + 1 Else quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If depth = 0 And LogicKeywordAt(s, i, keyword) Then LogicFindTopKeyword = i: Exit Function
        End If
    Next i
End Function

Private Function LogicFindNaturalElse(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long, depth As Long, quoted As Boolean, nestedIf As Long, ch As String, j As Long
    For i = IIf(startPos < 1, 1, startPos) To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If quoted And Mid$(s, i + 1, 1) = """" Then i = i + 1 Else quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If depth = 0 Then
                If LogicKeywordAt(s, i, "IF") Then
                    j = i + 2
                    Do While j <= Len(s)
                        If Not LogicWhitespace(Mid$(s, j, 1)) Then Exit Do
                        j = j + 1
                    Loop
                    ' IF(...) is already Excel-style. Natural IF has an expression after IF.
                    If j <= Len(s) Then
                        If Mid$(s, j, 1) <> "(" Then nestedIf = nestedIf + 1
                    End If
                ElseIf LogicKeywordAt(s, i, "ELSE") Then
                    If nestedIf = 0 Then LogicFindNaturalElse = i: Exit Function
                    nestedIf = nestedIf - 1
                End If
            End If
        End If
    Next i
End Function

Private Function LogicWhitespace(ByVal ch As String) As Boolean
    LogicWhitespace = (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf)
End Function

Private Function LogicKeywordAt(ByVal s As String, ByVal pos As Long, ByVal keyword As String) As Boolean
    Dim beforeCh As String, afterCh As String, klen As Long
    klen = Len(keyword)
    If pos < 1 Or pos + klen - 1 > Len(s) Then Exit Function
    If StrComp(Mid$(s, pos, klen), keyword, vbTextCompare) <> 0 Then Exit Function
    If pos > 1 Then beforeCh = Mid$(s, pos - 1, 1)
    If pos + klen <= Len(s) Then afterCh = Mid$(s, pos + klen, 1)
    If Len(beforeCh) > 0 And beforeCh Like "[A-Za-z0-9_]" Then Exit Function
    If Len(afterCh) > 0 And afterCh Like "[A-Za-z0-9_]" Then Exit Function
    LogicKeywordAt = True
End Function

Private Function LogicNormalizeBooleanExpr(ByVal s As String, ByVal context As String) As String
    Dim parts As Collection
    s = Trim$(s)
    If Len(s) = 0 Then LogicNormalizeBooleanExpr = s: Exit Function
    If Len(s) >= 3 And UCase$(Left$(s, 3)) = "IF " Then s = LogicTranslateNaturalIf(s, context)
    Set parts = LogicSplitTopLogical(s, "OR")
    If parts.count > 1 Then LogicNormalizeBooleanExpr = "OR(" & LogicJoinNormalized(parts, context) & ")": Exit Function
    Set parts = LogicSplitTopLogical(s, "AND")
    If parts.count > 1 Then LogicNormalizeBooleanExpr = "AND(" & LogicJoinNormalized(parts, context) & ")": Exit Function
    LogicNormalizeBooleanExpr = LogicNormalizeParenGroups(s, context)
End Function

Private Function LogicSplitTopLogical(ByVal s As String, ByVal keyword As String) As Collection
    Dim parts As New Collection, i As Long, startPos As Long, depth As Long, quoted As Boolean, ch As String, beforeCh As String, afterCh As String, klen As Long
    startPos = 1: klen = Len(keyword)
    For i = 1 To Len(s) - klen + 1
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If quoted And Mid$(s, i + 1, 1) = """" Then i = i + 1 Else quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If depth = 0 And LogicKeywordAt(s, i, keyword) Then
                beforeCh = "": afterCh = ""
                If i > 1 Then beforeCh = Mid$(s, i - 1, 1)
                If i + klen <= Len(s) Then afterCh = Mid$(s, i + klen, 1)
                If LogicWhitespace(beforeCh) And LogicWhitespace(afterCh) Then
                    parts.Add Trim$(Mid$(s, startPos, i - startPos))
                    startPos = i + klen
                    i = i + klen - 1
                End If
            End If
        End If
    Next i
    If parts.count > 0 Then parts.Add Trim$(Mid$(s, startPos)) Else parts.Add Trim$(s)
    Set LogicSplitTopLogical = parts
End Function

Private Function LogicJoinNormalized(ByVal parts As Collection, ByVal context As String) As String
    Dim i As Long, out As String
    For i = 1 To parts.count
        If Len(out) > 0 Then out = out & ","
        out = out & LogicNormalizeBooleanExpr(CStr(parts(i)), context)
    Next i
    LogicJoinNormalized = out
End Function

Private Function LogicNormalizeParenGroups(ByVal s As String, ByVal context As String) As String
    Dim out As String, i As Long, closePos As Long, inner As String, fn As String, parts As Collection, ch As String, quoted As Boolean
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            out = out & ch: i = i + 1: quoted = True
            Do While i <= Len(s) And quoted
                ch = Mid$(s, i, 1): out = out & ch
                If ch = """" Then
                    If i < Len(s) And Mid$(s, i + 1, 1) = """" Then out = out & """": i = i + 1 Else quoted = False
                End If
                i = i + 1
            Loop
        ElseIf ch = "(" Then
            closePos = LogicMatchingParen(s, i)
            If closePos = 0 Then Err.Raise vbObjectError + 581, , "Unmatched parentheses: " & context
            inner = Mid$(s, i + 1, closePos - i - 1)
            fn = LogicIdentifierBefore(s, i)
            If Len(fn) > 0 Then
                Set parts = LogicSplitArguments(inner)
                out = out & "(" & LogicJoinNormalized(parts, context) & ")"
            Else
                out = out & "(" & LogicNormalizeBooleanExpr(inner, context) & ")"
            End If
            i = closePos + 1
        Else
            out = out & ch: i = i + 1
        End If
    Loop
    LogicNormalizeParenGroups = out
End Function

' Returns the function name immediately preceding an opening parenthesis, or "" when the
' parenthesis is a plain grouping bracket.
'
' VBA evaluates both sides of And/Or - it does not short-circuit - so a loop written as
' "Do While i >= 1 And LogicWhitespace(Mid$(s, i, 1))" still evaluates Mid$(s, 0, 1) on the
' final iteration and raises run-time error 5. Every scan that can run off the left edge of
' the string therefore has to test the bound in its own statement.
Private Function LogicIdentifierBefore(ByVal s As String, ByVal openPos As Long) As String
    Dim i As Long, e As Long
    i = openPos - 1
    Do While i >= 1
        If Not LogicWhitespace(Mid$(s, i, 1)) Then Exit Do
        i = i - 1
    Loop
    e = i
    Do While i >= 1
        If Not (Mid$(s, i, 1) Like "[A-Za-z0-9_]") Then Exit Do
        i = i - 1
    Loop
    If e >= i + 1 Then LogicIdentifierBefore = Mid$(s, i + 1, e - i)
End Function

Public Function LogicMatchingParen(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long, depth As Long, quoted As Boolean, ch As String
    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If quoted And Mid$(s, i + 1, 1) = """" Then i = i + 1 Else quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then
                depth = depth - 1
                If depth = 0 Then LogicMatchingParen = i: Exit Function
            End If
        End If
    Next i
End Function

Public Function LogicSplitArguments(ByVal s As String) As Collection
    Dim parts As New Collection, startPos As Long, i As Long, depth As Long, quoted As Boolean, ch As String
    startPos = 1
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If quoted And Mid$(s, i + 1, 1) = """" Then i = i + 1 Else quoted = Not quoted
        ElseIf Not quoted Then
            If ch = "(" Then depth = depth + 1
            If ch = ")" Then depth = depth - 1
            If ch = "," And depth = 0 Then parts.Add Trim$(Mid$(s, startPos, i - startPos)): startPos = i + 1
        End If
    Next i
    parts.Add Trim$(Mid$(s, startPos))
    Set LogicSplitArguments = parts
End Function



Public Function NewMap() As Object
    Set NewMap = CreateObject("Scripting.Dictionary")
    NewMap.CompareMode = vbTextCompare
End Function

Public Function KeyPart(ByVal s As String) As String
    KeyPart = CStr(Len(s)) & ":" & UCase$(s)
End Function

Public Function NormalHeader(ByVal v As Variant) As String
    Dim s As String
    s = Replace(SafeUpperText(v), ChrW(160), " ")
    s = Replace(s, "-", "_"): s = Replace(s, "/", "_"): s = Replace(s, ".", "_")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    s = Replace(s, " ", "_")
    Do While InStr(s, "__") > 0: s = Replace(s, "__", "_"): Loop
    NormalHeader = s
End Function

' Safe compatibility resolver for older testcase exports. Exact names always win.
' Aliases are deliberately one-way: only latest PRE_* config fields fall back to
' the known legacy pre-shock names, avoiding unsafe interpretation in reverse.
Public Function ResolveSourceFieldKey(ByVal headers As Object, ByVal requested As Variant) As String
    Dim f As String, alt As String
    f = NormalHeader(requested)
    If Len(f) = 0 Then Exit Function
    If headers.Exists(f) Then ResolveSourceFieldKey = f: Exit Function
    Select Case f
        Case "PRE_LCR_HQLA": alt = "LCR_HQLA"
        Case "PRE_LCR_INFLOW": alt = "LCR_INFLOW"
        Case "PRE_LCR_OUTFLOW": alt = "LCR_OUTFLOW"
        Case "PRE_LEG_LIQ_TOTAL_ASSETS": alt = "LEG_LIQ_TOTAL_ASSETS"
        Case "PRE_LEG_LIQ_TOTAL_LIAB": alt = "LEG_LIQ_TOTAL_LIAB"
    End Select
    If Len(alt) > 0 Then
        If headers.Exists(alt) Then ResolveSourceFieldKey = alt
    End If
End Function

Private Function CanonicalHeader(ByVal v As Variant) As String
    Dim s As String: s = NormalHeader(v)
    Select Case s
        Case "AOD", "ASOFDATE", "ASOF_DATE", "AS_OFDATE", "AS_OF_DT", "REPORT_DATE", "REPORTING_DATE": s = "AS_OF_DATE"
        Case "ENTITYID", "ENTITY_IDENTIFIER": s = "ENTITY_ID"
        Case "ENTITYCODE": s = "ENTITY_CODE"
        Case "RUNID": s = "RUN_ID"
        Case "TEST_CASE_CODE", "TESTCASE_CODE", "SCENARIO_TESTCASE_CODE": s = "SCENARIO_TEST_CASE_CODE"
        Case "TEST_CASE_NAME", "TESTCASE_NAME", "SCENARIO_TESTCASE_NAME": s = "SCENARIO_TEST_CASE_NAME"
        Case "TEST_ELEMENT_CODE", "ELEMENT_CODE", "SCENARIOELEMENT_CODE": s = "SCENARIO_ELEMENT_CODE"
        Case "TEST_ELEMENT_NAME", "ELEMENT_NAME", "SCENARIOELEMENT_NAME": s = "SCENARIO_ELEMENT_NAME"
        Case "ELEMENTTYPE": s = "ELEMENT_TYPE"
        Case "SEVERITY", "SEVERITY_NAME", "SEVERITYCODE": s = "SEVERITY_CODE"
        Case "AMOUNTCHANGE": s = "AMOUNT_CHANGE"
        Case "FILTER_CONDITION", "FILTER_CONDITIONS", "SCENARIO_FILTER", "TEST_ELEMENT_FILTER", "ELEMENT_FILTER": s = "SCENARIO_ELEMENT_FORMULA_NAME"
    End Select
    CanonicalHeader = s
End Function

Private Function HeaderSummary(ByVal h As Object) As String
    Dim k As Variant, n As Long
    For Each k In h.keys
        n = n + 1
        If n <= 18 Then
            If Len(HeaderSummary) > 0 Then HeaderSummary = HeaderSummary & ", "
            HeaderSummary = HeaderSummary & CStr(k)
        End If
    Next k
    If h.count > 18 Then HeaderSummary = HeaderSummary & ", ... (" & h.count & " columns)"
End Function

Public Function DetectHeaderRow(ByVal ws As Worksheet) As Long
    Dim a As Variant, r As Long, c As Long, h As Object, found As Long, foundRows As String
    a = ws.Range("A1:IV50").Value2
    For r = 1 To UBound(a, 1)
        Set h = NewMap()
        For c = 1 To UBound(a, 2)
            If Len(SafeText(a(r, c))) > 0 Then h(CanonicalHeader(a(r, c))) = c
        Next c
        If h.Exists("AS_OF_DATE") And h.Exists("SCENARIO_ELEMENT_CODE") And h.Exists("SEVERITY_CODE") And h.Exists("ELEMENT_TYPE") Then
            found = found + 1: DetectHeaderRow = r
            If Len(foundRows) > 0 Then foundRows = foundRows & ", "
            foundRows = foundRows & CStr(r)
        End If
    Next r
    If found > 1 Then Err.Raise vbObjectError + 510, , "More than one possible source header row was found in '" & ws.name & "' (rows " & foundRows & "). Keep one header row only."
End Function

Public Sub ReadSource(ByVal ws As Worksheet, ByRef a As Variant, ByRef h As Object, ByRef headerRow As Long)
    Dim lastCell As Range, lastRow As Long, lastCol As Long, c As Long, k As String, req As Variant, missing As String, originals As Object
    headerRow = DetectHeaderRow(ws)
    If headerRow = 0 Then Err.Raise vbObjectError + 511, , "No usable Scenario Element Output header was found in '" & ws.name & "'. The header may be anywhere in the first 50 rows. Minimum identifying fields are AS_OF_DATE (or AOD), SCENARIO_ELEMENT_CODE, SEVERITY_CODE and ELEMENT_TYPE."
    Set lastCell = ws.Cells.Find(what:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If lastCell Is Nothing Then Err.Raise vbObjectError + 512, , "The selected sheet is empty."
    lastRow = lastCell.row
    lastCol = ws.Cells(headerRow, ws.columns.count).End(xlToLeft).Column
    If lastRow <= headerRow Then Err.Raise vbObjectError + 512, , "The selected source contains a header but no data rows. Sheet: " & ws.name & ", header row: " & headerRow & "."
    a = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol)).Value2
    Set h = NewMap(): Set originals = NewMap()
    For c = 1 To UBound(a, 2)
        k = CanonicalHeader(a(1, c))
        If Len(k) > 0 Then
            If h.Exists(k) Then Err.Raise vbObjectError + 513, , "Two source columns resolve to '" & k & "': '" & originals(k) & "' and '" & SafeText(a(1, c)) & "'. Keep only one to avoid ambiguous mapping."
            h(k) = c: originals(k) = SafeText(a(1, c))
        End If
    Next c
    For Each req In Array("AS_OF_DATE", "ENTITY_ID", "ENTITY_CODE", "RUN_ID", "SCENARIO_TEST_CASE_CODE", "SCENARIO_ELEMENT_CODE", "ELEMENT_TYPE", "SEVERITY_CODE", "AMOUNT_CHANGE")
        If Not h.Exists(CStr(req)) Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & CStr(req)
        End If
    Next req
    If Len(missing) > 0 Then
        Err.Raise vbObjectError + 514, , "Upload cannot be reconciled safely because required fields are missing: " & missing & "." & vbCrLf & _
            "AOD: AS_OF_DATE (AOD/report date aliases accepted). Entity: ENTITY_ID + ENTITY_CODE. Test case: SCENARIO_TEST_CASE_CODE. Test element: SCENARIO_ELEMENT_CODE + ELEMENT_TYPE. Severity: SEVERITY_CODE. Shock value: AMOUNT_CHANGE. Audit: RUN_ID." & vbCrLf & _
            "Sheet: " & ws.name & ", header row: " & headerRow & ". Recognised columns: " & HeaderSummary(h)
    End If
    ' Names are descriptive, not identity keys. If absent, safely use their codes.
    If Not h.Exists("SCENARIO_TEST_CASE_NAME") Then h("SCENARIO_TEST_CASE_NAME") = h("SCENARIO_TEST_CASE_CODE")
    If Not h.Exists("SCENARIO_ELEMENT_NAME") Then h("SCENARIO_ELEMENT_NAME") = h("SCENARIO_ELEMENT_CODE")
End Sub

Private Function RowContext(ByRef a As Variant, ByVal h As Object, ByVal r As Long, ByVal headerRow As Long) As String
    Dim s As String
    s = "Source row " & CStr(r + headerRow - 1)
    If h.Exists("AS_OF_DATE") Then s = s & " | AOD=" & SafeText(a(r, h("AS_OF_DATE")))
    If h.Exists("ENTITY_CODE") Then s = s & " | Entity=" & SafeText(a(r, h("ENTITY_CODE")))
    If h.Exists("ENTITY_ID") Then s = s & " (ID " & SafeText(a(r, h("ENTITY_ID"))) & ")"
    If h.Exists("SCENARIO_TEST_CASE_CODE") Then s = s & " | Testcase=" & SafeText(a(r, h("SCENARIO_TEST_CASE_CODE")))
    If h.Exists("SCENARIO_ELEMENT_CODE") Then s = s & " | Test element=" & SafeText(a(r, h("SCENARIO_ELEMENT_CODE")))
    If h.Exists("SEVERITY_CODE") Then s = s & " | Severity=" & SafeText(a(r, h("SEVERITY_CODE")))
    RowContext = s
End Function

Private Function DateKeyAt(ByRef a As Variant, ByVal h As Object, ByVal r As Long, ByVal headerRow As Long, ByVal date1904 As Boolean) As String
    Dim n As Long, d As String
    On Error GoTo BadDate
    DateKeyAt = DateKey(a(r, h("AS_OF_DATE")), date1904)
    Exit Function
BadDate:
    n = Err.Number: d = Err.description
    Err.Raise n, , d & " " & RowContext(a, h, r, headerRow)
End Function

Public Function DateKey(ByVal v As Variant, Optional ByVal date1904 As Boolean = False) As String
    Dim s As String, p As Variant, y As Long, m As Long, d As Long, dt As Date, serial As Double
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Err.Raise vbObjectError + 515, , "AS_OF_DATE is missing or invalid."
    If VarType(v) = vbDate Then
        dt = dateValue(v)
    ElseIf IsNumeric(v) Then
        serial = Fix(CDbl(v))
        If date1904 Then serial = serial + 1462
        If serial < 61 Or serial > 2958465 Then Err.Raise vbObjectError + 516, , "AS_OF_DATE is outside the supported Excel date range."
        dt = CDate(serial)
    Else
        s = Replace(Trim$(CStr(v)), "/", "-")
        If s Like "####-##-##*" Then
            y = val(Left$(s, 4)): m = val(Mid$(s, 6, 2)): d = val(Mid$(s, 9, 2))
        Else
            s = Replace(s, "/", "-"): s = Split(s, " ")(0): p = Split(s, "-")
            If UBound(p) <> 2 Then Err.Raise vbObjectError + 517, , "Use an Excel date or yyyy-mm-dd for AS_OF_DATE: " & s
            y = val(p(2))
            If Not IsNumeric(p(1)) Then
                If Len(p(1)) < 3 Then Err.Raise vbObjectError + 517, , "Invalid month in AS_OF_DATE: " & s
                m = (InStr(1, "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC", UCase$(Left$(p(1), 3)), vbBinaryCompare) + 2) \ 3
                d = val(p(0))
            ElseIf UCase$(CStr(SettingValue("Text date order", "DMY"))) = "MDY" Then
                m = val(p(0)): d = val(p(1))
            Else
                d = val(p(0)): m = val(p(1))
            End If
            If y < 100 Then y = y + 2000
        End If
        If y < 1900 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Err.Raise vbObjectError + 518, , "Invalid AS_OF_DATE: " & s
        dt = DateSerial(y, m, d)
        If Year(dt) <> y Or Month(dt) <> m Or day(dt) <> d Then Err.Raise vbObjectError + 519, , "Invalid AS_OF_DATE: " & s
    End If
    DateKey = format$(dt, "yyyy-mm-dd")
End Function

Private Function RowsEqual(ByRef a As Variant, ByVal r1 As Long, ByVal r2 As Long) As Boolean
    Dim c As Long
    For c = 1 To UBound(a, 2)
        If VarType(a(r1, c)) <> VarType(a(r2, c)) Then Exit Function
        If CStr(a(r1, c)) <> CStr(a(r2, c)) Then Exit Function
    Next c
    RowsEqual = True
End Function

Private Function IsIdentityColumn(ByVal col As Long, ByVal h As Object) As Boolean
    Dim k As Variant
    For Each k In Array("AS_OF_DATE", "ENTITY_ID", "ENTITY_CODE", "RUN_ID", "SCENARIO_TEST_CASE_CODE", "SCENARIO_TEST_CASE_NAME", "SCENARIO_ELEMENT_CODE", "SCENARIO_ELEMENT_NAME", "ELEMENT_TYPE", "SEVERITY_CODE", "SEVERITY_ID")
        If h.Exists(CStr(k)) Then
            If CLng(h(CStr(k))) = col Then IsIdentityColumn = True: Exit Function
        End If
    Next k
End Function

Public Function IndexSource(ByRef a As Variant, ByVal h As Object, ByVal headerRow As Long, ByVal date1904 As Boolean, ByRef stats As Object) As Object
    Dim result As Object, bases As Object, dates As Object, entities As Object, cases As Object, categories As Object
    Dim tc As Object, els As Object, el As Object, sevRows As Object, catMap As Object
    Dim r As Long, c As Long, dt As String, eid As String, ec As String, ent As String, pair As String, tk As String, ek As String
    Dim testCode As String, sev As String, cat As String, et As String, k As Variant, n As Long, isBlank As Boolean, meaningful As Boolean
    Set result = NewMap(): Set bases = NewMap(): Set dates = NewMap(): Set entities = NewMap(): Set cases = NewMap(): Set categories = NewMap(): Set catMap = CategoryMap()
    Set stats = NewMap(): stats("SourceRows") = UBound(a, 1) - 1: stats("Duplicates") = 0: stats("EmptySeverityRows") = 0
    For r = 2 To UBound(a, 1)
        isBlank = True
        For c = 1 To UBound(a, 2)
            If IsError(a(r, c)) Then Err.Raise vbObjectError + 520, , "Excel error in source column " & c & ". " & RowContext(a, h, r, headerRow)
            If Len(SafeText(a(r, c))) > 0 Then isBlank = False
        Next c
        If isBlank Then GoTo NextSourceRow
        dt = DateKeyAt(a, h, r, headerRow, date1904)
        eid = SafeText(a(r, h("ENTITY_ID"))): ec = SafeText(a(r, h("ENTITY_CODE")))
        If Len(eid) = 0 Or Len(ec) = 0 Then Err.Raise vbObjectError + 521, , "ENTITY_ID or ENTITY_CODE is blank. " & RowContext(a, h, r, headerRow)
        ent = KeyPart(eid) & KeyPart(ec): pair = dt & "|" & ent
        testCode = SafeText(a(r, h("SCENARIO_TEST_CASE_CODE"))): ek = SafeUpperText(a(r, h("SCENARIO_ELEMENT_CODE"))): sev = SafeUpperText(a(r, h("SEVERITY_CODE")))
        If Len(testCode) = 0 And Len(ek) = 0 And (sev = "" Or sev = "BASE" Or sev = "BASELINE" Or sev = "NORMAL" Or sev = "0") Then
            If bases.Exists(pair) Then
                If Not RowsEqual(a, CLng(bases(pair)), r) Then Err.Raise vbObjectError + 522, , "Conflicting base rows for AOD " & dt & " / entity " & ec & ": source rows " & bases(pair) + headerRow - 1 & " and " & r + headerRow - 1 & ". Keep one base row per AOD/entity. " & RowContext(a, h, r, headerRow)
                stats("Duplicates") = stats("Duplicates") + 1
            Else
                bases(pair) = r
            End If
            GoTo NextSourceRow
        End If
        If Len(testCode) = 0 Or Len(ek) = 0 Then Err.Raise vbObjectError + 523, , "A scenario row is missing SCENARIO_TEST_CASE_CODE or SCENARIO_ELEMENT_CODE. " & RowContext(a, h, r, headerRow)
        meaningful = False
        For c = 1 To UBound(a, 2)
            If Not IsIdentityColumn(c, h) Then
                If Not IsEmpty(a(r, c)) Then
                    If Len(CStr(a(r, c))) > 0 Then meaningful = True: Exit For
                End If
            End If
        Next c
        If Not meaningful Then stats("EmptySeverityRows") = stats("EmptySeverityRows") + 1: GoTo NextSourceRow
        If sev <> "MODERATE" And sev <> "MEDIUM" And sev <> "SEVERE" Then Err.Raise vbObjectError + 524, , "Unrecognised severity '" & sev & "'. Expected MODERATE, MEDIUM or SEVERE. " & RowContext(a, h, r, headerRow)
        et = SafeText(a(r, h("ELEMENT_TYPE")))
        If Len(et) = 0 Then Err.Raise vbObjectError + 525, , "ELEMENT_TYPE is blank. " & RowContext(a, h, r, headerRow)
        cat = CategoryFor(testCode, catMap): tk = pair & "|" & KeyPart(testCode)
        If Not result.Exists(tk) Then
            Set tc = NewMap(): Set els = NewMap()
            tc.Add "TestCaseCode", testCode: tc.Add "TestCaseName", SafeText(a(r, h("SCENARIO_TEST_CASE_NAME")))
            tc.Add "Date", dt: tc.Add "EntityKey", ent: tc.Add "EntityCode", ec: tc.Add "EntityID", eid: tc.Add "Pair", pair: tc.Add "Category", cat: tc.Add "Elements", els
            result.Add tk, tc
        End If
        Set tc = result(tk): Set els = tc("Elements")
        If Not els.Exists(ek) Then
            Set el = NewMap(): Set sevRows = NewMap()
            el.Add "ScenarioCode", SafeText(a(r, h("SCENARIO_ELEMENT_CODE")))
            el.Add "ScenarioName", SafeText(a(r, h("SCENARIO_ELEMENT_NAME")))
            el.Add "ElementType", et: el.Add "SeverityRows", sevRows
            If h.Exists("SCENARIO_ELEMENT_FORMULA_NAME") Then el.Add "FilterExpression", SafeText(a(r, h("SCENARIO_ELEMENT_FORMULA_NAME"))) Else el.Add "FilterExpression", ""
            els.Add ek, el
        End If
        Set el = els(ek): Set sevRows = el("SeverityRows")
        If StrComp(CStr(el("ElementType")), et, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 526, , "Conflicting ELEMENT_TYPE values for the same testcase/test element. " & RowContext(a, h, r, headerRow)
        If h.Exists("SCENARIO_ELEMENT_FORMULA_NAME") Then
            If Len(SafeText(a(r, h("SCENARIO_ELEMENT_FORMULA_NAME")))) > 0 Then
                If Len(CStr(el("FilterExpression"))) = 0 Then
                    el("FilterExpression") = SafeText(a(r, h("SCENARIO_ELEMENT_FORMULA_NAME")))
                ElseIf StrComp(Trim$(CStr(el("FilterExpression"))), Trim$(SafeText(a(r, h("SCENARIO_ELEMENT_FORMULA_NAME")))), vbTextCompare) <> 0 Then
                    Err.Raise vbObjectError + 534, , "Conflicting filter conditions for the same test element. " & RowContext(a, h, r, headerRow)
                End If
            End If
        End If
        If sevRows.Exists(sev) Then
            n = CLng(sevRows(sev))
            If Not RowsEqual(a, n, r) Then Err.Raise vbObjectError + 527, , "Conflicting duplicate scenario row. Existing source row " & n + headerRow - 1 & "; duplicate " & RowContext(a, h, r, headerRow) & ". Exact duplicates are accepted; conflicting values are not."
            stats("Duplicates") = stats("Duplicates") + 1
        Else
            sevRows(sev) = r
        End If
        dates(dt) = dt: entities(ent) = ec & " (ID " & eid & ")": cases(testCode) = tc("TestCaseName"): categories(cat) = cat
NextSourceRow:
    Next r
    For Each k In result.keys
        Set tc = result(k)
        If Not bases.Exists(tc("Pair")) Then Err.Raise vbObjectError + 528, , "Missing base row for AOD " & tc("Date") & " / entity " & tc("EntityCode") & " (ID " & tc("EntityID") & "), first affected testcase " & tc("TestCaseCode") & ". Add one base row for this AOD/entity with blank testcase + test element and severity blank/BASE/BASELINE/NORMAL/0."
        tc.Add "BaseRow", CLng(bases(tc("Pair")))
    Next k
    If result.count = 0 Then Err.Raise vbObjectError + 529, , "No populated scenario severities were found."
    stats.Add "Dates", dates: stats.Add "Entities", entities: stats.Add "Cases", cases: stats.Add "Categories", categories: stats("BaseRows") = bases.count
    Set IndexSource = result
End Function

Private Sub ValidateUploadFields(ByVal idx As Object, ByVal h As Object)
    Dim rulesByType As Object, tc As Object, el As Object, rules As Collection, ri As Object, tk As Variant, ek As Variant, t As String, f As String, checked As Object
    Set rulesByType = LoadRulesByElementType(ThisWorkbook): Set checked = NewMap()
    For Each tk In idx.keys
        Set tc = idx(tk)
        For Each ek In tc("Elements").keys
            Set el = tc("Elements")(ek): t = SafeUpperText(el("ElementType"))
            If Not rulesByType.Exists(t) Then Err.Raise vbObjectError + 532, , "No enabled config rules match ELEMENT_TYPE '" & el("ElementType") & "'. AOD=" & tc("Date") & " | Entity=" & tc("EntityCode") & " (ID " & tc("EntityID") & ") | Testcase=" & tc("TestCaseCode") & " | Test element=" & el("ScenarioCode")
            If Not checked.Exists(t) Then
                Set rules = rulesByType(t)
                For Each ri In rules
                    f = ResolveSourceFieldKey(h, ri("SourceFieldName"))
                    If Len(NormalHeader(ri("SourceFieldName"))) > 0 And Len(f) = 0 Then
                        Err.Raise vbObjectError + 533, , "Source field '" & ri("SourceFieldName") & "' is required by the enabled config for ELEMENT_TYPE '" & el("ElementType") & "' but is missing from this upload." & vbCrLf & _
                            "AOD=" & tc("Date") & " | Entity=" & tc("EntityCode") & " (ID " & tc("EntityID") & ") | Testcase=" & tc("TestCaseCode") & " | Test element=" & el("ScenarioCode")
                    End If
                Next ri
                checked(t) = True
            End If
        Next ek
    Next tk
End Sub

Public Sub UploadScenarioOutput()
    Dim chosen As Variant, srcWb As Workbook, openWb As Workbook, candidate As Worksheet, preferred As Worksheet, ws As Worksheet, stage As Worksheet, oldSource As Worksheet, backup As Worksheet
    Dim a As Variant, h As Object, idx As Object, stats As Object, hr As Long, count As Long, state As Object, openedHere As Boolean, sourceName As String
    Dim er As String, committed As Boolean, startedCommit As Boolean, oldArea As String, oldDashboard As Variant, preferredCount As Long, candidates As String
    If JKB_Busy Then Exit Sub
    chosen = Application.GetOpenFilename("Excel workbooks (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", , "Upload testcase output")
    If VarType(chosen) = vbBoolean Then Exit Sub
    If StrComp(CStr(chosen), ThisWorkbook.FullName, vbTextCompare) = 0 Then UiNotice "Load system output", "That file is this workbook.", , "Choose the test case output file produced by the stress testing system.": Exit Sub
    On Error GoTo Failed
    Set state = CaptureState(): JKB_Busy = True
    Application.EnableEvents = False: Application.ScreenUpdating = False: Application.DisplayAlerts = False
    Application.DefaultSheetDirection = xlLTR
    Application.AutomationSecurity = 3
    For Each openWb In Application.Workbooks
        If StrComp(openWb.FullName, CStr(chosen), vbTextCompare) = 0 Then Set srcWb = openWb: Exit For
    Next openWb
    If srcWb Is Nothing Then
        Set srcWb = Workbooks.Open(fileName:=CStr(chosen), UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
        openedHere = True
    End If
    For Each ws In srcWb.Worksheets
        If DetectHeaderRow(ws) > 0 Then
            count = count + 1: Set candidate = ws
            If Len(candidates) > 0 Then candidates = candidates & ", "
            candidates = candidates & ws.name
            If Replace(Replace(Replace(UCase$(ws.name), " ", ""), "_", ""), "-", "") = "SCENARIOELEMENTOUTPUT" Then
                preferredCount = preferredCount + 1: Set preferred = ws
            End If
        End If
    Next ws
    If count = 0 Then Err.Raise vbObjectError + 530, , "No usable testcase-output sheet was found. The sheet name is not fixed. Put the source headers anywhere in the first 50 rows; required identifiers include AS_OF_DATE/AOD, SCENARIO_ELEMENT_CODE, SEVERITY_CODE and ELEMENT_TYPE. Extra columns and extra sheets are allowed."
    If preferredCount = 1 Then Set candidate = preferred
    If count > 1 And preferredCount <> 1 Then Err.Raise vbObjectError + 531, , "Several sheets look like testcase output: " & candidates & ". Rename only the intended source sheet to 'Scenario Element Output' so the tool can choose it without guessing."
    sourceName = candidate.name
    ReadSource candidate, a, h, hr
    Set idx = IndexSource(a, h, hr, srcWb.date1904, stats)
    ValidateUploadFields idx, h
    ' Canonical dates are numeric Excel dates in the master workbook's date system.
    Dim r As Long, dk As String
    For r = 2 To UBound(a, 1)
        If Len(SafeText(a(r, h("AS_OF_DATE")))) > 0 Then
            dk = DateKey(a(r, h("AS_OF_DATE")), srcWb.date1904)
            a(r, h("AS_OF_DATE")) = CDbl(DateSerial(val(Left$(dk, 4)), val(Mid$(dk, 6, 2)), val(Right$(dk, 2))))
            If ThisWorkbook.date1904 Then a(r, h("AS_OF_DATE")) = a(r, h("AS_OF_DATE")) - 1462
        End If
    Next r
    Set oldSource = ThisWorkbook.Worksheets(SHEET_SOURCE)
    Sheet1.EnsureRibbon oldSource
    Set stage = ThisWorkbook.Worksheets.Add(After:=oldSource)
    stage.name = UniqueSheetName(ThisWorkbook, "__Upload_Staging")
    stage.Range("A1").Value2 = "JKB"
    WriteLiteralValues stage.Cells(3, 1).Resize(UBound(a, 1), UBound(a, 2)), a
    stage.columns(CLng(h("AS_OF_DATE"))).NumberFormat = "dd-mmm-yyyy"
    stage.rows(3).Font.Bold = True: stage.rows(3).WrapText = True: stage.rows(3).RowHeight = 42
    stage.Range(stage.Cells(3, 1), stage.Cells(3, UBound(a, 2))).Interior.Color = RGB(52, 64, 84)
    stage.Range(stage.Cells(3, 1), stage.Cells(3, UBound(a, 2))).Font.Color = vbWhite
    stage.UsedRange.columns.ColumnWidth = 18
    stage.Range(stage.Cells(3, 1), stage.Cells(UBound(a, 1) + 2, UBound(a, 2))).AutoFilter
    Set backup = ThisWorkbook.Worksheets.Add(After:=stage)
    backup.name = UniqueSheetName(ThisWorkbook, "__Upload_Backup")
    oldArea = oldSource.UsedRange.address
    oldSource.UsedRange.Copy Destination:=backup.Range(oldArea)
    oldDashboard = ThisWorkbook.Worksheets(SHEET_DASHBOARD).Range("F18:F25").Value2
    ' Keep the original source sheet's identity so formulas, names and navigation remain valid.
    startedCommit = True
    oldSource.AutoFilterMode = False
    oldSource.UsedRange.Clear
    stage.UsedRange.Copy Destination:=oldSource.Range(stage.UsedRange.address)
    oldSource.UsedRange.columns.ColumnWidth = 18
    oldSource.rows(3).RowHeight = 42
    oldSource.Range(oldSource.Cells(3, 1), oldSource.Cells(UBound(a, 1) + 2, UBound(a, 2))).AutoFilter
    Sheet1.EnsureRibbon oldSource
    With ThisWorkbook.Worksheets(SHEET_DASHBOARD)
        .Range("F18").Value2 = CStr(chosen): .Range("F19").Value2 = sourceName
        .Range("F20").value = Now: .Range("F20").NumberFormat = "dd-mmm-yyyy hh:mm:ss"
        .Range("F21").Value2 = stats("SourceRows"): .Range("F22").Value2 = stats("Cases").count
        .Range("F23").Value2 = stats("Dates").count: .Range("F24").Value2 = stats("Entities").count
        .Range("F25").Value2 = "Ready. Choose scenarios, dates and entities."
    End With
    EnsurePreShockSheet
    SyncPreShockTestCases idx
    committed = True
    stage.Delete: Set stage = Nothing
    backup.Delete: Set backup = Nothing
    ClearRunSelection
    If openedHere Then srcWb.Close SaveChanges:=False
    RestoreState state: JKB_Busy = False
    OpenReconWorkbench   ' Home, not the hidden Dashboard.
    UiNotice "Load system output", stats("SourceRows") & " rows loaded from " & sourceName & ".", , "Check scenarios and test cases (step 2) on the home screen."
    Exit Sub
Failed:
    er = Err.description
    On Error Resume Next
    If Not committed Then
        If startedCommit Then
            oldSource.UsedRange.Clear
            backup.Range(oldArea).Copy Destination:=oldSource.Range(oldArea)
            Dim dr As Long
            For dr = 1 To 8
                ThisWorkbook.Worksheets(SHEET_DASHBOARD).Cells(dr + 17, 6).Value2 = oldDashboard(dr, 1)
            Next dr
            Sheet1.EnsureRibbon oldSource
        End If
    End If
    If Not stage Is Nothing Then stage.Delete
    If Not backup Is Nothing Then backup.Delete
    If openedHere Then srcWb.Close SaveChanges:=False
    RestoreState state: JKB_Busy = False
    UiProblem "Load system output", "The system output could not be loaded.", er
End Sub
