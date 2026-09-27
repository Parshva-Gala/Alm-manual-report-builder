Attribute VB_Name = "modScenarioBuilder_Config"
Option Explicit

Public Function LoadRulesByElementType(ByVal wb As Workbook) As Object
    Dim lo As ListObject, h As Object, result As Object, ri As Object, rules As Collection, a As Variant, hdr As Variant
    Dim props As Variant, cols As Variant, r As Long, c As Long, k As String, value As Variant, i As Long
    Set lo = wb.Worksheets(SHEET_RULES).ListObjects(TABLE_RULES)
    Set result = NewMap(): Set h = NewMap()
    hdr = lo.HeaderRowRange.Value2
    For c = 1 To UBound(hdr, 2): h(NormalHeader(hdr(1, c))) = c: Next c
    props = Split("ElementType|OutputOrder|IsEnabled|ShowManual|ShowDifference|OutputRowLabel|SourceFieldName|ValueFormatKey|DifferenceFormatKey|ManualFormulaDefault|ManualFormulaModerate|ManualFormulaMedium|ManualFormulaSevere|LabelStyleKey|BaseStyleKey|SystemStyleKey|ManualStyleKey|DifferenceStyleKey|HideRowOnOpen|Notes", "|")
    cols = Split("ELEMENT_TYPE|OUTPUT_ORDER|IS_ENABLED|SHOW_MANUAL|SHOW_DIFFERENCE|OUTPUT_ROW_LABEL|SOURCE_FIELD_NAME|VALUE_FORMAT_KEY|DIFFERENCE_FORMAT_KEY|MANUAL_FORMULA_DEFAULT|MANUAL_FORMULA_MODERATE|MANUAL_FORMULA_MEDIUM|MANUAL_FORMULA_SEVERE|LABEL_STYLE_KEY|BASE_STYLE_KEY|SYSTEM_STYLE_KEY|MANUAL_STYLE_KEY|DIFFERENCE_STYLE_KEY|HIDE_ROW_ON_OPEN|NOTES", "|")
    For i = 0 To UBound(cols)
        If Not h.Exists(cols(i)) Then Err.Raise vbObjectError + 550, , "Missing configuration column: " & cols(i)
    Next i
    If lo.DataBodyRange Is Nothing Then Err.Raise vbObjectError + 551, , "The configuration table has no rules."
    a = lo.DataBodyRange.Value2
    For r = 1 To UBound(a, 1)
        If Not ParseBool(a(r, h("IS_ENABLED"))) Then GoTo NextRule
        k = SafeUpperText(a(r, h("ELEMENT_TYPE")))
        If Len(k) = 0 Then Err.Raise vbObjectError + 552, , "An enabled configuration row has no ELEMENT_TYPE."
        Set ri = NewMap()
        For i = 0 To UBound(cols)
            value = a(r, h(cols(i)))
            Select Case props(i)
                Case "IsEnabled", "ShowManual", "ShowDifference", "HideRowOnOpen": value = ParseBool(value)
                Case "OutputOrder"
                    If Not IsNumeric(value) Then Err.Raise vbObjectError + 553, , "Invalid OUTPUT_ORDER in configuration row " & r & "."
                    value = CLng(value)
                Case Else: value = SafeText(value)
            End Select
            ri.Add props(i), value
        Next i
        If Not result.Exists(k) Then Set rules = New Collection: result.Add k, rules
        Set rules = result(k)
        For i = 1 To rules.count
            If CLng(ri("OutputOrder")) < CLng(rules(i)("OutputOrder")) Then rules.Add ri, Before:=i: GoTo NextRule
        Next i
        rules.Add ri
NextRule:
    Next r
    Set LoadRulesByElementType = result
End Function

' RuleRowMap (a flat, ungrouped label -> row map) was removed: it disagreed with the layout
' the generator actually writes. modScenarioBuilder_ExternalRun.BuildOutputRowMap is now the
' single source of truth for output row positions.

' The selection validator lives in modScenarioBuilder_ExternalRun, next to the generator
' it guards, so that it reports every configuration problem in one pass.
