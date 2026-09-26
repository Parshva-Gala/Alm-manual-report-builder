Option Explicit

' ============================================================================
'  The look, in one module.
'
'  MIDBANK'S OWN COLOURS, not an invented corporate palette. Sampled off the
'  bank's logo mark: jet black, white, and the emerald the word MID is set in.
'  Those three and nothing else - no navy, no gold, no second accent competing
'  with the first.
'
'    black     #000000   every masthead, every header row, the primary button
'    emerald   #00905F   the rule under the masthead, the key line, the accent
'    white     #FFFFFF   the canvas the numbers sit on
'
'  BLACK CHROME ON A WHITE CANVAS. The chrome is heavy and unbroken so the
'  sheet reads as the bank's, and the data area is left white because a balance
'  is what the sheet is for and nobody reads one off a dark field.
'
'  Two rules hold the whole thing together:
'
'    1  every figure uses ONE number format, from modPD_Const.NUM_FMT
'    2  every sheet has the same skeleton, so a sheet you have not opened reads
'       the way the last one did
'
'  To re-skin the tool, change C_INK and C_BRAND. Nothing else names a colour.
' ============================================================================

' --- the skeleton -----------------------------------------------------------
Public Const R_BAR As Long = 1        ' the button rail
Public Const R_TITLE As Long = 2
Public Const R_ABOUT As Long = 3
Public Const R_STATUS As Long = 5
Public Const R_HDR As Long = 7
Public Const R_FIRST As Long = 8

Public Const UI_FONT As String = "Aptos"
Public Const UI_MONO As String = "Consolas"

' --- the palette, off the logo ----------------------------------------------
'
' The black is the logo's own field and the emerald is the colour MID is set
' in - both measured from the mark rather than chosen to look like a bank.
Public Function C_INK() As Long
    C_INK = RGB(0, 0, 0)
End Function
Public Function C_SURFACE() As Long
    C_SURFACE = RGB(12, 18, 15)          ' black with the faintest green cast
End Function
Public Function C_ELEV() As Long
    C_ELEV = RGB(22, 34, 28)
End Function
Public Function C_BRAND() As Long
    C_BRAND = RGB(0, 144, 96)            ' #009060, the emerald of the mark
End Function
Public Function C_BRAND_DEEP() As Long
    C_BRAND_DEEP = RGB(0, 90, 60)
End Function
Public Function C_BRAND_SOFT() As Long
    C_BRAND_SOFT = RGB(214, 244, 231)    ' pale mint, for type reversed on black
End Function
Public Function C_PAPER() As Long
    C_PAPER = RGB(255, 255, 255)
End Function
Public Function C_MIST() As Long
    C_MIST = RGB(245, 248, 252)
End Function
Public Function C_HAIR() As Long
    C_HAIR = RGB(214, 222, 235)
End Function
Public Function C_MUTED() As Long
    C_MUTED = RGB(120, 136, 161)
End Function
Public Function C_OK_TX() As Long
    C_OK_TX = RGB(22, 101, 70)
End Function
Public Function C_OK_BG() As Long
    C_OK_BG = RGB(226, 245, 236)
End Function
Public Function C_WARN_TX() As Long
    C_WARN_TX = RGB(133, 96, 12)
End Function
Public Function C_WARN_BG() As Long
    C_WARN_BG = RGB(253, 244, 214)
End Function
Public Function C_BAD_TX() As Long
    C_BAD_TX = RGB(159, 45, 45)
End Function
Public Function C_BAD_BG() As Long
    C_BAD_BG = RGB(252, 233, 232)
End Function
Public Function C_IDLE_TX() As Long
    C_IDLE_TX = RGB(110, 126, 150)
End Function
Public Function C_IDLE_BG() As Long
    C_IDLE_BG = RGB(238, 242, 248)
End Function

' ===================== dressing a sheet =====================================

Public Sub Dress(ByVal ws As Worksheet, ByVal title As String, ByVal about As String)
    On Error Resume Next
    With ws.Cells
        .Font.Name = UI_FONT
        .Font.Size = 10
        .Interior.Color = C_PAPER
    End With

    ' The masthead: a solid ink band with the title reversed out of it, and a
    ' gold hairline under. One band, on every sheet, in the same two rows.
    With ws.Range(ws.Cells(R_TITLE, 1), ws.Cells(R_ABOUT, 14))
        .Interior.Color = C_INK
    End With
    With ws.Cells(R_TITLE, 1)
        .Value2 = title
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = C_PAPER
        .IndentLevel = 1
    End With
    With ws.Cells(R_ABOUT, 1)
        .Value2 = about
        .Font.Size = 9
        .Font.Color = RGB(150, 176, 163)     ' grey with the brand's green in it
        .IndentLevel = 1
    End With
    ' An emerald rule under the masthead and an emerald bar down its leading
    ' edge - the logo's own mark sits to the left of the word, and the sheet
    ' repeats that shape rather than borrowing a generic accent stripe.
    With ws.Range(ws.Cells(R_ABOUT, 1), ws.Cells(R_ABOUT, 14)).Borders(xlEdgeBottom)
        .Color = C_BRAND
        .Weight = xlMedium
    End With
    With ws.Range(ws.Cells(R_TITLE, 1), ws.Cells(R_ABOUT, 1)).Borders(xlEdgeLeft)
        .Color = C_BRAND
        .Weight = xlThick
    End With

    ws.Rows(R_BAR).RowHeight = 30
    ws.Rows(R_TITLE).RowHeight = 26
    ws.Rows(R_ABOUT).RowHeight = 15
    ws.Rows(4).RowHeight = 7
    ws.Rows(6).RowHeight = 7
    ws.Tab.Color = C_INK
    ws.DisplayPageBreaks = False
    Err.Clear
End Sub

Public Sub Head(ByVal ws As Worksheet, ByRef headings As Variant, ByRef widths As Variant)
    Dim c As Long
    On Error Resume Next
    For c = 0 To UBound(headings)
        ws.Cells(R_HDR, c + 1).Value2 = headings(c)
        If c <= UBound(widths) Then ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    With ws.Range(ws.Cells(R_HDR, 1), ws.Cells(R_HDR, UBound(headings) + 1))
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = C_BRAND_SOFT
        .Interior.Color = C_INK
        .WrapText = True
        .VerticalAlignment = xlBottom
        .IndentLevel = 1
        .Borders(xlEdgeBottom).Color = C_BRAND
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(R_HDR).RowHeight = 30
    Err.Clear
End Sub

Public Sub SetStatus(ByVal ws As Worksheet, ByVal msg As String, ByVal level As String)
    On Error Resume Next
    With ws.Range(ws.Cells(R_STATUS, 1), ws.Cells(R_STATUS, 14))
        .Interior.Color = LevelBg(level)
        .Borders(xlEdgeLeft).Color = LevelTx(level)
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
    With ws.Cells(R_STATUS, 1)
        .Value2 = msg
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = LevelTx(level)
        .IndentLevel = 1
    End With
    ws.Rows(R_STATUS).RowHeight = 22
    Err.Clear
End Sub

Private Function LevelTx(ByVal level As String) As Long
    Select Case UCase$(level)
        Case "OK", "READY": LevelTx = C_OK_TX
        Case "BREAK", "BAD": LevelTx = C_BAD_TX
        Case "CHECK", "WARN": LevelTx = C_WARN_TX
        Case Else: LevelTx = C_IDLE_TX
    End Select
End Function

Private Function LevelBg(ByVal level As String) As Long
    Select Case UCase$(level)
        Case "OK", "READY": LevelBg = C_OK_BG
        Case "BREAK", "BAD": LevelBg = C_BAD_BG
        Case "CHECK", "WARN": LevelBg = C_WARN_BG
        Case Else: LevelBg = C_IDLE_BG
    End Select
End Function

Public Sub PaintVerdict(ByVal cell As Range)
    Dim v As String
    On Error Resume Next
    v = UCase$(SafeText(cell.Value2))
    cell.HorizontalAlignment = xlCenter
    cell.Font.Bold = True
    cell.Font.Size = 9
    Select Case v
        Case "OK", "READY", "AGREES", "LOADED"
            cell.Font.Color = C_OK_TX: cell.Interior.Color = C_OK_BG
        Case "BREAK", "MISSING", "FAILED"
            cell.Font.Color = C_BAD_TX: cell.Interior.Color = C_BAD_BG
        Case "CHECK", "WARN", "PARTIAL"
            cell.Font.Color = C_WARN_TX: cell.Interior.Color = C_WARN_BG
        Case Else
            cell.Font.Color = C_IDLE_TX: cell.Interior.Color = C_IDLE_BG
    End Select
    Err.Clear
End Sub

Public Sub PaintVerdictColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal LastRow As Long)
    Dim r As Long
    If LastRow < R_FIRST Then Exit Sub
    For r = R_FIRST To LastRow
        PaintVerdict ws.Cells(r, col)
    Next r
End Sub

' Banding, hairlines and a filter - applied once per sheet at the end rather
' than per row, which is the difference between a second and a minute.
Public Sub DressTable(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal LastRow As Long)
    Dim rng As Range
    On Error Resume Next
    If LastRow < R_FIRST Then Exit Sub
    Set rng = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(LastRow, lastCol))
    rng.Borders(xlInsideHorizontal).Color = C_HAIR
    rng.Borders(xlEdgeBottom).Color = C_HAIR
    rng.VerticalAlignment = xlTop
    ws.Range(ws.Cells(R_HDR, 1), ws.Cells(LastRow, lastCol)).AutoFilter
    Err.Clear
End Sub

' ===================== the button rail ======================================
'
' Excel shapes have no hover state, so "animation" here is what a shape CAN do
' and do well: a soft shadow, a gold key line on the primary action, and a
' pressed look on the one you are standing on. The console window is where the
' real motion lives - see console.hta.

' Clears only the family it is asked for.
'
' It used to clear every "pd_" shape, and the rail is laid down AFTER the desk's
' cards - so installing the rail deleted the three card buttons that are the
' whole point of the Desk sheet. The rail owns "pdr_", the cards own "pd_card_",
' and neither can now remove the other.
Public Sub ClearButtons(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long, sh As Shape
    On Error Resume Next
    For i = ws.Shapes.count To 1 Step -1
        Set sh = ws.Shapes(i)
        If Left$(sh.Name, Len(prefix)) = prefix Then sh.Delete
    Next i
    Err.Clear
End Sub

' kind: 0 quiet, 1 primary (gold), 2 current (filled ink)
Public Function Button(ByVal ws As Worksheet, ByVal caption As String, ByVal proc As String, _
                       ByVal x As Double, ByVal w As Double, Optional ByVal kind As Long = 0) As Double
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, 5, w, 21)
    If sh Is Nothing Then Button = x: Exit Function
    sh.Name = "pdr_" & CLng(x) & "_" & CLng(Rnd * 100000)
    sh.Placement = xlFreeFloating
    sh.Adjustments(1) = 0.22
    sh.Line.visible = msoFalse
    With sh.Shadow
        .visible = msoTrue
        .style = msoShadowStyleOuterShadow
        .Blur = 4
        .Transparency = 0.82
        .Size = 100
        .OffsetX = 0
        .OffsetY = 1
        .ForeColor.RGB = C_INK
    End With
    Select Case kind
        Case 1
            sh.Fill.ForeColor.RGB = C_BRAND          ' the action you came for
        Case 2
            sh.Fill.ForeColor.RGB = C_INK            ' the sheet you are standing on
        Case Else
            sh.Fill.ForeColor.RGB = RGB(238, 243, 240)
    End Select
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 3: .MarginRight = 3
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Name = UI_FONT
        Select Case kind
            Case 1: .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Case 2: .TextRange.Font.Fill.ForeColor.RGB = C_BRAND_SOFT
            Case Else: .TextRange.Font.Fill.ForeColor.RGB = C_INK
        End Select
    End With
    If Len(proc) > 0 Then sh.OnAction = proc
    Err.Clear
    Button = x + w + 5
End Function

' A thin gold rule used as a separator inside the rail.
Public Function Divider(ByVal ws As Worksheet, ByVal x As Double) As Double
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x + 3, 9, 1.5, 13)
    If Not sh Is Nothing Then
        sh.Name = "pdr_div_" & CLng(x)
        sh.Placement = xlFreeFloating
        sh.Line.visible = msoFalse
        sh.Fill.ForeColor.RGB = C_HAIR
        sh.Shadow.visible = msoFalse
    End If
    Err.Clear
    Divider = x + 12
End Function

' The rail every desk sheet carries. Where you are does not change how you
' get out.
Public Sub Rail(ByVal ws As Worksheet)
    Dim x As Double
    On Error Resume Next
    ClearButtons ws, "pdr_"
    x = 5
    x = Button(ws, "Open the console", "PD_Console", x, 116, 1)
    x = Divider(ws, x)
    x = Button(ws, "Desk", "PD_GoHome", x, 46, IIf(ws.Name = SH_HOME, 2, 0))
    x = Button(ws, "Files", "PD_GoFiles", x, 46, IIf(ws.Name = SH_SOURCES, 2, 0))
    x = Button(ws, "Reconciliation", "PD_GoRecon", x, 84, IIf(ws.Name = SH_RECON, 2, 0))
    x = Button(ws, "Activity", "PD_GoLog", x, 60, IIf(ws.Name = SH_LOG, 2, 0))
    ws.Rows(R_BAR).RowHeight = 30
    Err.Clear
End Sub

Public Sub RailEverywhere()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.visible = xlSheetVisible Then Rail ws
    Next ws
End Sub

' ===================== navigation ===========================================

Public Sub GoTo_(ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetSheet(nm)
    If ws Is Nothing Then Exit Sub
    ws.visible = xlSheetVisible
    ws.Activate
    ws.Range("A1").Select
    Err.Clear
End Sub

Public Sub PD_GoHome()
    GoTo_ SH_HOME
End Sub
Public Sub PD_GoFiles()
    GoTo_ SH_SOURCES
End Sub
Public Sub PD_GoRecon()
    GoTo_ SH_RECON
End Sub
Public Sub PD_GoLog()
    GoTo_ SH_LOG
End Sub
