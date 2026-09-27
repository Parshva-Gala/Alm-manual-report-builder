Attribute VB_Name = "modUI_Theme"
Option Explicit

' ============================================================================
'  JKB design system v2: the one palette and the three button styles.
'
'  The workbook, the review console, the filter builder and the progress
'  window all draw from these tokens. Before v2 the VBA carried about 160
'  slightly different colours; every one of them now maps to a token below.
'
'  Colour carries meaning and nothing else:
'    brand navy        structure, navigation and the one primary action
'    green/amber/red   verdicts only (reconciled / review / break)
'    greys             everything else
'
'  Sheet1 (Dashboard) keeps its own RGB literals on purpose: a generated
'  review workbook carries a copy of Sheet1 and nothing else, so it cannot
'  reference this module. Its literals use exactly these values.
'
'  VBA colours are BGR, so each constant is written &HBBGGRR&.
' ============================================================================

Public Const UI_FONT As String = "Aptos"

' Neutrals
Public Const UI_WHITE As Long = &HFFFFFF        ' #FFFFFF surface
Public Const UI_SUBTLE As Long = &HFBFAF9       ' #F9FAFB raised rows
Public Const UI_CANVAS As Long = &HFAF7F5       ' #F5F7FA page background
Public Const UI_FILL As Long = &HF7F4F2         ' #F2F4F7 quiet fill
Public Const UI_LINE As Long = &HF0ECEA         ' #EAECF0 hairline
Public Const UI_LINE_2 As Long = &HDDD5D0       ' #D0D5DD control border
Public Const UI_FAINT As Long = &HB3A298        ' #98A2B3 placeholder
Public Const UI_MUTED As Long = &H857066        ' #667085 secondary text
Public Const UI_TEXT_2 As Long = &H675447       ' #475467 labels
Public Const UI_TEXT As Long = &H544034         ' #344054 body text
Public Const UI_INK As Long = &H281810          ' #101828 headings

' Brand
Public Const UI_BRAND As Long = &H893F19        ' #193F89 primary
Public Const UI_BRAND_DK As Long = &H6B3012     ' #12306B primary hover / dark
Public Const UI_NAVY As Long = &H40220E         ' #0E2240 header bands
Public Const UI_NAVY_TEXT As Long = &HD8C4B7    ' #B7C4D8 text on navy
Public Const UI_BRAND_TINT As Long = &HFBF3EE   ' #EEF3FB selected
Public Const UI_BRAND_TINT_2 As Long = &HF5E2D6 ' #D6E2F5 selected border

' Verdicts: text, background, border
Public Const UI_OK As Long = &H477606&          ' #067647
Public Const UI_OK_BG As Long = &HF3FDEC        ' #ECFDF3
Public Const UI_OK_LINE As Long = &HC6EFAB      ' #ABEFC6
Public Const UI_WARN As Long = &H847B5          ' #B54708
Public Const UI_WARN_BG As Long = &HEBFAFF      ' #FFFAEB
Public Const UI_WARN_LINE As Long = &H89DFFE    ' #FEDF89
Public Const UI_BAD As Long = &H1823B4          ' #B42318
Public Const UI_BAD_BG As Long = &HF2F3FE       ' #FEF3F2
Public Const UI_BAD_LINE As Long = &HCACDFE     ' #FECDCA

' Heatmap and chart fills, and the saturated dots used for status markers
Public Const UI_OK_FILL As Long = &HE6FADC      ' #DCFAE6
Public Const UI_WARN_FILL As Long = &HC7F0FE    ' #FEF0C7
Public Const UI_BAD_FILL As Long = &HE2E4FE     ' #FEE4E2
Public Const UI_OK_DOT As Long = &H6AB217       ' #17B26A
Public Const UI_WARN_DOT As Long = &H990F7      ' #F79009
Public Const UI_BAD_DOT As Long = &H3844F0      ' #F04438
Public Const UI_BAD_BAR As Long = &H9BA2FD      ' #FDA29B
Public Const UI_UP_BAR As Long = &HFFAD84       ' #84ADFF
Public Const UI_NAVY_MUTED As Long = &HC2A38F   ' #8FA3C2
Public Const UI_OK_SOFT As Long = &HF9FEF6      ' #F6FEF9
Public Const UI_WARN_SOFT As Long = &HF5FCFF    ' #FFFCF5
Public Const UI_BAD_SOFT As Long = &HFAFBFF     ' #FFFBFA

' Button kinds for UiButton
Public Const UI_BTN_PRIMARY As String = "primary"
Public Const UI_BTN_SECONDARY As String = "secondary"
Public Const UI_BTN_QUIET As String = "quiet"

' A button on a sheet, fitted inside a cell range.
'
'   primary     filled brand navy, white text. One per screen at most.
'   secondary   white, grey border, navy text. The normal action.
'   quiet       no fill, no border, muted text. Links and minor tools.
'
' Bound to the workbook that owns the sheet, so a copy of the workbook keeps
' working under any file name. proc may be qualified ("Sheet1.GoConfig").
Public Sub UiButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal proc As String, _
                    ByVal rg As Range, Optional ByVal kind As String = "secondary", Optional ByVal fontSize As Double = 10.5)
    Dim sh As Shape
    On Error Resume Next: ws.Shapes(nm).Delete: On Error GoTo 0
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, rg.Left + 3, rg.Top + 3, rg.Width - 6, rg.Height - 6)
    sh.name = nm
    sh.OnAction = "'" & Replace(ws.Parent.name, "'", "''") & "'!" & proc
    sh.Placement = xlMove: sh.Shadow.Visible = msoFalse
    On Error Resume Next: sh.Adjustments.item(1) = 0.16: On Error GoTo 0
    Select Case kind
        Case UI_BTN_PRIMARY
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = UI_BRAND
            sh.line.Visible = msoTrue: sh.line.ForeColor.RGB = UI_BRAND
        Case UI_BTN_QUIET
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = UI_FILL
            sh.line.Visible = msoFalse
        Case Else
            sh.fill.Visible = msoTrue: sh.fill.ForeColor.RGB = UI_WHITE
            sh.line.Visible = msoTrue: sh.line.ForeColor.RGB = UI_LINE_2
    End Select
    sh.line.Weight = 0.75
    With sh.TextFrame2
        .textRange.text = caption
        .textRange.Font.name = UI_FONT: .textRange.Font.Size = fontSize: .textRange.Font.Bold = msoTrue
        Select Case kind
            Case UI_BTN_PRIMARY: .textRange.Font.fill.ForeColor.RGB = UI_WHITE
            Case UI_BTN_QUIET: .textRange.Font.fill.ForeColor.RGB = UI_TEXT
            Case Else: .textRange.Font.fill.ForeColor.RGB = UI_BRAND
        End Select
        .textRange.ParagraphFormat.Alignment = msoAlignCenter: .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 1: .MarginBottom = 1
        .WordWrap = msoFalse
    End With
End Sub

' A white card with a hairline border, the one container shape on every sheet.
Public Sub UiCard(ByVal rg As Range, Optional ByVal fillColor As Long = UI_WHITE)
    Dim e As Variant
    rg.Interior.Color = fillColor
    For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With rg.Borders(e)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = UI_LINE
        End With
    Next e
End Sub

' A small uppercase section heading.
Public Sub UiEyebrow(ByVal c As Range, ByVal text As String)
    c.Value2 = UCase$(text)
    c.Font.name = UI_FONT: c.Font.Size = 9: c.Font.Bold = True: c.Font.Color = UI_MUTED
    c.VerticalAlignment = xlBottom
End Sub

' Verdict colours for a status cell, applied as conditional formats so the
' cell recolours itself whenever the run writes a new status into it.
Public Sub UiStatusFormats(ByVal rg As Range)
    Dim fc As FormatCondition
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Reconciled", TextOperator:=xlBeginsWith): fc.Font.Color = UI_OK
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Review", TextOperator:=xlBeginsWith): fc.Font.Color = UI_WARN
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Failed", TextOperator:=xlBeginsWith): fc.Font.Color = UI_BAD
    Set fc = rg.FormatConditions.Add(xlTextString, String:="Running", TextOperator:=xlBeginsWith): fc.Font.Color = UI_BRAND
End Sub

' ------------------------------------------------------------ report books ---
' The last step before any generated report workbook is saved, so every file the
' tool hands out opens the same way: the house font, document properties filled
' in, each sheet scrolled to A1 with the first sheet in front, and a print layout
' (landscape, one page wide, title and page number in the footer) so a printed or
' PDF'd report reads as a report rather than a screen dump.
' Nothing here changes a value, a formula or a pivot.
Public Sub UiFinishReportBook(ByVal wb As Workbook, ByVal reportTitle As String, Optional ByVal subject As String = "", _
                              Optional ByVal resetView As Boolean = True)
    Dim ws As Worksheet, first As Worksheet, printComm As Boolean, oldActive As Workbook, fixedPage As Boolean
    On Error Resume Next
    Set oldActive = ActiveWorkbook
    With wb.Styles("Normal").Font
        .name = UI_FONT: .Size = 10
    End With
    wb.BuiltinDocumentProperties("Title").Value = reportTitle
    wb.BuiltinDocumentProperties("Subject").Value = IIf(Len(subject) > 0, subject, "Stress testing reconciliation")
    wb.BuiltinDocumentProperties("Company").Value = "JKB Risk & Finance"
    wb.BuiltinDocumentProperties("Comments").Value = "Generated by the JKB Stress Testing tool on " & format$(Now, "dd-mmm-yyyy hh:nn")

    printComm = Application.PrintCommunication
    Application.PrintCommunication = False
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If first Is Nothing Then Set first = ws
            ' A sheet laid out as one printed page (the pack's cover and summary)
            ' carries the sheet name JKB_FixedPage and keeps its own page setup.
            fixedPage = False
            fixedPage = (Len(ws.Names("JKB_FixedPage").name) > 0)
            Err.Clear
            If Not fixedPage Then
            With ws.PageSetup
                .Orientation = xlLandscape
                ' One page wide only while that stays legible; a wide data sheet
                ' prints across pages at 100% instead of shrinking to nothing.
                If ws.UsedRange.Columns.count <= 30 Then
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                End If
                .LeftMargin = Application.CentimetersToPoints(1.2)
                .RightMargin = Application.CentimetersToPoints(1.2)
                .TopMargin = Application.CentimetersToPoints(1.5)
                .BottomMargin = Application.CentimetersToPoints(1.5)
                .LeftFooter = "&8" & Replace(reportTitle, "&", "&&") & "  |  &A"
                .CenterFooter = "&8Page &P of &N"
                .RightFooter = "&8Printed &D"
            End With
            End If
            Err.Clear
        End If
    Next ws
    Application.PrintCommunication = printComm

    ' Open on A1 everywhere. Needs the book's own window, which a book made by
    ' Workbooks.Add has even while screen updating is off. Generated output books
    ' set their own view (ApplyOutputWindowSettings), so they skip this.
    If Not resetView Then Exit Sub
    wb.Activate
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Activate
            ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
            ws.Range("A1").Select
            Err.Clear
        End If
    Next ws
    If Not first Is Nothing Then first.Activate
    If Not oldActive Is Nothing Then If Not oldActive Is wb Then oldActive.Activate
    Err.Clear
End Sub

' ---------------------------------------------------------------- messages ---
' Every popup in the tool's own modules goes through these three, so they all
' read the same way:
'
'   headline     one sentence: what happened, or what is being asked
'   detail       the facts behind it (counts, a file path, the reason)
'   next step    what to do now, when there is something to do
'
' The window title names the task ("Load inputs", "Run tests"), never "Error".
' Sheet1, Sheet2, ThisWorkbook and shared\*.bas keep plain MsgBox calls: Sheet1
' is copied into generated workbooks and shared\ is also compiled into ALM_V10,
' and neither can see this module. Their wording follows the same pattern.

Public Sub UiNotice(ByVal title As String, ByVal headline As String, Optional ByVal detail As String = "", _
                    Optional ByVal nextStep As String = "", Optional ByVal warn As Boolean = False)
    MsgBox UiMessageText(headline, detail, nextStep, ""), IIf(warn, vbExclamation, vbInformation), UiTitle(title)
End Sub

Public Sub UiProblem(ByVal title As String, ByVal headline As String, Optional ByVal reason As String = "", _
                     Optional ByVal nextStep As String = "")
    MsgBox UiMessageText(headline, "", nextStep, reason), vbExclamation, UiTitle(title)
End Sub

' True when the person chooses OK (or Yes for a destructive question, where No is
' the default button so Enter never deletes anything).
Public Function UiAsk(ByVal title As String, ByVal question As String, Optional ByVal detail As String = "", _
                      Optional ByVal destructive As Boolean = False) As Boolean
    Dim text As String
    text = UiMessageText(question, detail, "", "")
    If destructive Then
        UiAsk = (MsgBox(text, vbYesNo + vbExclamation + vbDefaultButton2, UiTitle(title)) = vbYes)
    Else
        UiAsk = (MsgBox(text, vbOKCancel + vbQuestion, UiTitle(title)) = vbOK)
    End If
End Function

Private Function UiMessageText(ByVal headline As String, ByVal detail As String, ByVal nextStep As String, _
                               ByVal reason As String) As String
    Dim t As String
    t = Trim$(headline)
    If Len(Trim$(detail)) > 0 Then t = t & vbCrLf & vbCrLf & Trim$(detail)
    If Len(Trim$(reason)) > 0 Then t = t & vbCrLf & vbCrLf & "Reason: " & UiCleanReason(reason)
    If Len(Trim$(nextStep)) > 0 Then t = t & vbCrLf & vbCrLf & "Next: " & Trim$(nextStep)
    UiMessageText = t
End Function

' Raw error text often arrives with a leading "ERROR:" or "ERR:" and no full stop.
Private Function UiCleanReason(ByVal s As String) As String
    s = Trim$(s)
    If UCase$(Left$(s, 6)) = "ERROR:" Then s = Trim$(Mid$(s, 7))
    If UCase$(Left$(s, 4)) = "ERR:" Then s = Trim$(Mid$(s, 5))
    If Len(s) > 0 Then
        If InStr(".!?", Right$(s, 1)) = 0 Then s = s & "."
        s = UCase$(Left$(s, 1)) & Mid$(s, 2)
    End If
    UiCleanReason = s
End Function

Private Function UiTitle(ByVal title As String) As String
    If Len(Trim$(title)) = 0 Then UiTitle = "JKB Stress Testing" Else UiTitle = "JKB Stress Testing - " & Trim$(title)
End Function

' One entry point for loading a single extract type, so the sheets can offer
' one "Load inputs" button instead of a row of per-type loaders.
Public Sub UiLoadOneSource()
    Dim choice As String, procs As Variant
    procs = Array("", "UploadECLSource", "UploadCAPRWASource", "UploadLLSource", "UploadLCRSource", "UploadCAPSource", "UploadNSFRSource", "UploadInputFolder")
    choice = InputBox("Load just one input. Type its number:" & vbCrLf & vbCrLf & _
                      "   1   ECL output" & vbCrLf & _
                      "   2   CAPRWA output" & vbCrLf & _
                      "   3   Legal liquidity" & vbCrLf & _
                      "   4   LCR" & vbCrLf & _
                      "   5   Capital component" & vbCrLf & _
                      "   6   NSFR" & vbCrLf & _
                      "   7   Every file in a folder", "Load one input", "1")
    choice = Trim$(choice)
    If Len(choice) = 0 Then Exit Sub
    If Not IsNumeric(choice) Then Exit Sub
    If CLng(choice) < 1 Or CLng(choice) > 7 Then
        UiNotice "Load one input", "Type a number from 1 to 7."
        Exit Sub
    End If
    Application.Run "'" & Replace(ThisWorkbook.name, "'", "''") & "'!" & CStr(procs(CLng(choice)))
End Sub

' Opens Config_ValueSources: where each base and pre-shock value is linked to the
' input files, with its per-row switch. Sets it up first if it is not there yet.
Public Sub GoValueSources()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(VS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        JKB_ApplyValueSourceSetup True
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(VS_SHEET)
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Sub
    ws.Visible = xlSheetVisible
    ws.Activate
    DressWorkSheet ws
    ws.Range("A1").Select
End Sub

' Run once by Rebuild.ps1 after the modules are imported: the one-time config
' setups for grouping and for the base and pre-shock value sources.
Public Sub JKB_ApplyReleaseSetup()
    JKB_ApplyGroupingSetup True, True
    JKB_ApplyValueSourceSetup True
    JKB_ApplySimplifySetup True
End Sub
