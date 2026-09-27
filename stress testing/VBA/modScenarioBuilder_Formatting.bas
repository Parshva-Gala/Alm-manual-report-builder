Attribute VB_Name = "modScenarioBuilder_Formatting"
Option Explicit
Private mCompiled As Object
Private mRowMapSeq As Long
Private mGuard As Boolean
Private mGuardChecked As Boolean

Public Sub ResetFormulaCache()
    Set mCompiled = NewMap()
    mGuardChecked = False
End Sub

' Driven by "Guard formula errors" on Config_MasterFormatting; defaults to Yes.
' Read once per run because it is consulted for every formula.
Private Function GuardDivisionErrors() As Boolean
    If Not mGuardChecked Then
        mGuard = (SafeUpperText(SettingValue("Guard formula errors", "Yes")) <> "NO")
        mGuardChecked = True
    End If
    GuardDivisionErrors = mGuard
End Function

Public Function ManualFormulaText(ByVal rule As Object, ByVal severity As Long) As String
    Dim keys As Variant
    keys = Array("ManualFormulaModerate", "ManualFormulaMedium", "ManualFormulaSevere")
    ManualFormulaText = NormalizeConfigFormulaText(rule(keys(severity)))
    If Len(ManualFormulaText) = 0 Then ManualFormulaText = NormalizeConfigFormulaText(rule("ManualFormulaDefault"))
End Function

' derivedRows: metric -> row on the output workbook's Derived_Values sheet, for the element
' being written. Supplied only during generation; when absent, a `derived.X` token falls back
' to the system-output cell so a formula never fails to compile.
Public Function ConvertManualFormula(ByVal formula As String, ByVal rowMap As Object, ByVal severity As Long, _
                                     ByVal context As String, Optional ByVal derivedRows As Object) As String
    Dim s As String, cacheKey As String, hasDerived As Boolean
    s = NormalizeConfigFormulaText(formula)
    If Len(s) = 0 Then Exit Function
    If mCompiled Is Nothing Then Set mCompiled = NewMap()
    ' A derived token resolves against THIS element's slice of the derived-values sheet, so
    ' the result is not reusable across elements. Those formulas skip the cache outright
    ' rather than risk one element's numbers being reused on another's sheet.
    hasDerived = (RegexExecute("@?\b(derived|derivedbase|base_derived)\.", s).count > 0)
    If Not hasDerived Then
        ' The compiled result is a set of absolute cell references, so it is only valid for the
        ' row map it was resolved against. The same element type is laid out differently during
        ' validation and generation (section headers, detail-group ordering), so the row map
        ' MUST take part in the cache key - otherwise a formula validated against one layout is
        ' silently reused in another and points at the wrong metrics.
        cacheKey = RowMapFingerprint(rowMap) & "|" & KeyPart(context) & CStr(severity) & ":" & CStr(Len(s)) & ":" & s
        If mCompiled.Exists(cacheKey) Then ConvertManualFormula = mCompiled(cacheKey): Exit Function
    End If
    If InStr(1, s, "#REF!", vbTextCompare) > 0 Then
        WarnOnce context & "#REF", "Legacy #REF! placeholder: formula remains blank.", context
        Exit Function
    End If
    s = NormalizeNaturalFormulaLogic(s, context)
    s = ReplaceIfElseName(s)
    ' Derived tokens are resolved first, so a metric name that also exists as an output row
    ' cannot be captured by the generic token pass below.
    If hasDerived Then s = ReplaceDerivedTokens(s, rowMap, severity, context, derivedRows)
    s = ReplaceTokens(s, "\[([^\]]+)\]", 0, rowMap, severity, context)
    s = ReplaceTokens(s, "@([A-Za-z_]+)\(([^)]+)\)", 1, rowMap, severity, context)
    s = ReplaceTokens(s, "@?\b(base|system|systemoutput|manual|man|difference|differences|diff)\.([A-Za-z0-9_]+)\b", 1, rowMap, severity, context)
    s = ExpandHelpers(s, context)
    ' A reconciliation sheet showing #DIV/0! teaches the reviewer nothing: a ratio whose
    ' denominator is zero is simply not defined for this upload. Division guards are added
    ' automatically so the cell reads blank instead, while genuine configuration faults are
    ' still surfaced by the validator before generation starts.
    If GuardDivisionErrors() Then s = "IFERROR(" & s & ",NA())"
    ConvertManualFormula = "=" & s
    If Not hasDerived Then mCompiled(cacheKey) = ConvertManualFormula
End Function

' `derived.X`  -> the independently rebuilt PRE-SHOCK figure for metric X
' `derivedbase.X` (alias `base_derived.X`) -> the independently rebuilt BASE figure
'
' These are the numbers this tool derives from the uploaded extracts. They are written to a
' Derived_Values sheet inside the output workbook and referenced from there, so the reviewer
' can see the figure, the source it came from, the filter that produced it and how many rows
' matched - instead of a value pasted into a cell with no provenance.
'
' Without a derived figure the token degrades to the system-output cell, which is what the
' configuration used to hold outright. That keeps every formula valid when an extract has not
' been uploaded, while a run WITH the extract loaded stops depending on the system's own
' number for its own check.
Private Function ReplaceDerivedTokens(ByVal s As String, ByVal rm As Object, ByVal severity As Long, _
                                      ByVal context As String, ByVal derivedRows As Object) As String
    Dim matches As Object, m As Object, i As Long, kind As String, label As String, value As String, rowIx As Long
    Set matches = RegexExecute("@?\b(derivedbase|base_derived|derived)\.([A-Za-z0-9_]+)\b", s)
    For i = matches.count - 1 To 0 Step -1
        Set m = matches(i)
        If Not InsideString(s, m.FirstIndex + 1) Then
            kind = UCase$(Trim$(m.SubMatches(0)))
            label = UCase$(Trim$(m.SubMatches(1)))
            rowIx = 0
            If Not derivedRows Is Nothing Then
                If derivedRows.Exists(label) Then rowIx = CLng(derivedRows(label))
            End If
            If rowIx > 0 Then
                If kind = "DERIVED" Then
                    value = "'" & DERIVED_SHEET & "'!$E$" & CStr(rowIx)
                Else
                    value = "'" & DERIVED_SHEET & "'!$D$" & CStr(rowIx)
                End If
                value = "IF(ISNUMBER(" & value & ")," & value & ",NA())"
            Else
                WarnOnce context & "|derived-missing|" & label, _
                    "Independent source value is unavailable. The check remains unavailable until the source is loaded and recalculated.", context
                value = "NA()"
            End If
            s = Left$(s, m.FirstIndex) & value & Mid$(s, m.FirstIndex + m.Length + 1)
        End If
    Next i
    ReplaceDerivedTokens = s
End Function

' Tags each row map with an identity the compiled-formula cache can key on.
'
' A content hash would risk a collision silently binding a formula to the wrong layout, so
' every map object instead gets its own sequential tag, stored under a key no configuration
' label can produce. Two structurally identical maps simply recompile - a cost, never a
' correctness risk. The counter is deliberately never reset: an old map keeps the tag it was
' given, so reusing a number after a cache reset could alias two live maps.
Private Function RowMapFingerprint(ByVal rowMap As Object) As String
    Const TOKEN_KEY As String = "*<JKB row map token>*"
    If rowMap Is Nothing Then RowMapFingerprint = "none": Exit Function
    If Not rowMap.Exists(TOKEN_KEY) Then
        mRowMapSeq = mRowMapSeq + 1
        rowMap(TOKEN_KEY) = "rm" & CStr(mRowMapSeq)
    End If
    RowMapFingerprint = CStr(rowMap(TOKEN_KEY))
End Function

Private Function ReplaceIfElseName(ByVal s As String) As String
    Dim matches As Object, m As Object, i As Long
    Set matches = RegexExecute("\bIFELSE\s*\(", s)
    For i = matches.count - 1 To 0 Step -1
        Set m = matches(i)
        If Not InsideString(s, m.FirstIndex + 1) Then s = Left$(s, m.FirstIndex) & "IF(" & Mid$(s, m.FirstIndex + m.Length + 1)
    Next i
    ReplaceIfElseName = s
End Function

Private Function ReplaceTokens(ByVal s As String, ByVal pattern As String, ByVal kind As Long, ByVal rm As Object, ByVal severity As Long, ByVal context As String) As String
    Dim matches As Object, m As Object, i As Long, token As String, p As Long, sec As String, label As String, value As String
    Set matches = RegexExecute(pattern, s)
    For i = matches.count - 1 To 0 Step -1
        Set m = matches(i)
        If Not InsideString(s, m.FirstIndex + 1) Then
            sec = "": label = ""
            If kind = 0 Then
                token = CStr(m.SubMatches(0)): p = InStr(token, ":")
                If p > 0 Then sec = Trim$(Left$(token, p - 1)): label = Trim$(Mid$(token, p + 1)) Else label = Trim$(token)
            Else
                sec = Trim$(m.SubMatches(0)): label = Trim$(m.SubMatches(1))
            End If
            If rm.Exists(UCase$(label)) Then
                value = "$" & colLetter(SectionColumn(sec, severity)) & "$" & CStr(rm(UCase$(label)))
            Else
                ' Preserve the supplied engine's zero fallback; make it explicit in the log.
                WarnOnce context & "|" & sec & "|" & label, "Configured row is absent or disabled. Independent result is unavailable.", context & " / " & sec & "." & label
                value = "NA()"
            End If
            s = Left$(s, m.FirstIndex) & value & Mid$(s, m.FirstIndex + m.Length + 1)
        End If
    Next i
    ReplaceTokens = s
End Function

Public Sub WarnOnce(ByVal key As String, ByVal message As String, ByVal context As String)
    Static seen As Object
    If key = "__RESET__" Then Set seen = NewMap(): Exit Sub
    If seen Is Nothing Then Set seen = NewMap()
    If seen.Exists(key) Then Exit Sub
    seen(key) = True: LogIssue LOG_LEVEL_WARN, "Configuration", message, context
End Sub

Private Function InsideString(ByVal s As String, ByVal position As Long) As Boolean
    Dim i As Long, quoted As Boolean
    i = 1
    Do While i < position
        If Mid$(s, i, 1) = """" Then
            If quoted And Mid$(s, i + 1, 1) = """" Then i = i + 1 Else quoted = Not quoted
        End If
        i = i + 1
    Loop
    InsideString = quoted
End Function

Private Function ExpandHelpers(ByVal s As String, ByVal context As String) As String
    Dim matches As Object, m As Object, p As Long, closePos As Long, inner As String, parts As Collection, replacement As String, n As Long
    Do
        Set matches = RegexExecute("\b(IFNULL|ISNULL)\s*\(", s)
        Set m = Nothing
        For n = matches.count - 1 To 0 Step -1
            If Not InsideString(s, matches(n).FirstIndex + 1) Then Set m = matches(n): Exit For
        Next n
        If m Is Nothing Then Exit Do
        p = m.FirstIndex + m.Length
        closePos = LogicMatchingParen(s, p)
        If closePos = 0 Then Err.Raise vbObjectError + 581, , "Unmatched parentheses: " & context
        inner = Mid$(s, p + 1, closePos - p - 1)
        Set parts = LogicSplitArguments(inner)
        If UCase$(m.SubMatches(0)) = "IFNULL" Then
            If parts.count <> 2 Then Err.Raise vbObjectError + 582, , "IFNULL requires exactly two arguments: " & context
            replacement = "IF(" & NullTest(CStr(parts(1))) & "," & parts(2) & "," & parts(1) & ")"
        Else
            If parts.count <> 1 Then Err.Raise vbObjectError + 583, , "ISNULL requires one argument: " & context
            replacement = NullTest(CStr(parts(1)))
        End If
        s = Left$(s, m.FirstIndex) & replacement & Mid$(s, closePos + 1)
        If Len(s) > 8191 Then Err.Raise vbObjectError + 584, , "Formula exceeds Excel's length limit: " & context
    Loop
    ExpandHelpers = s
End Function
Private Function NullTest(ByVal s As String) As String
    NullTest = "IFERROR(LEN(" & Trim$(s) & "&"""")=0,TRUE)"
End Function
