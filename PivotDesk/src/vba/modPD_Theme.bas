Option Explicit

' ============================================================================
'  The look, in one module.
'
'  MIDBANK'S OWN COLOURS, not an invented corporate palette. Sampled off the
'  bank's logo mark: jet black, white, and the emerald the word MID is set in.
'  Those three and their tones - no navy, no gold, no second accent competing
'  with the first. Status colours exist for status and are never decoration.
'
'    black     #000000   the app bar, every masthead, every table head
'    emerald   #009060   the rule under the masthead, the active tab, the accent
'    white     #FFFFFF   the canvas the numbers sit on
'
'  ONE ROOM. Since 3.0 every sheet is the Desk's: the dark canvas, the black
'  bar, emerald for what can be acted on. Figures are set in near-white on
'  near-black rows at AA contrast or better (docs/CONTRAST.md), and every
'  sheet prints in black and white.
'
'  The values here mirror build/design_tokens.py, which draws the Desk. The
'  build checks the two agree, so the sheets and the Desk cannot drift.
'
'  Two rules hold the whole thing together:
'
'    1  every figure uses ONE number format, from modPD_Const.NUM_FMT
'    2  every sheet has the same skeleton, so a sheet you have not opened reads
'       the way the last one did
' ============================================================================

' --- the skeleton -----------------------------------------------------------
Public Const R_BAR As Long = 1        ' the app bar: navigation and actions
Public Const R_TITLE As Long = 2
Public Const R_ABOUT As Long = 3
Public Const R_STATUS As Long = 5
Public Const R_HDR As Long = 7
Public Const R_FIRST As Long = 8

' How far right the chrome is painted. Wide enough to reach the edge of a
' large monitor at a normal zoom on any of the sheets.
Public Const DRESS_COLS As Long = 40

' Segoe UI ships with every Windows since Vista. Nothing here waits on a
' cloud font arriving.
Public Const UI_FONT As String = "Segoe UI"
Public Const UI_SEMI As String = "Segoe UI Semibold"
Public Const UI_LIGHT As String = "Segoe UI Light"
Public Const UI_MONO As String = "Consolas"

' The Avati mark, width over height (build/brand.py trims it to 785 x 205).
Public Const LOGO_RATIO As Double = 3.829

' ===================== the palette ==========================================

' "009060" -> the Long Excel wants. Hex strings keep the palette readable and
' identical to the design tokens.
Public Function HX(ByVal h As String) As Long
    HX = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

' --- the field ---
Public Function C_INK() As Long
    C_INK = HX("000000")
End Function
Public Function C_CANVAS() As Long
    C_CANVAS = HX("060B09")
End Function
Public Function C_SURFACE() As Long
    C_SURFACE = HX("0C1512")
End Function
Public Function C_ELEV() As Long
    C_ELEV = HX("111F19")
End Function
Public Function C_HAIR_DARK() As Long
    C_HAIR_DARK = HX("1A2922")
End Function

' --- emerald ---
Public Function C_BRAND() As Long
    C_BRAND = HX("009060")            ' the emerald of the mark
End Function
Public Function C_BRAND_BRIGHT() As Long
    C_BRAND_BRIGHT = HX("16B07F")
End Function
Public Function C_BRAND_DEEP() As Long
    C_BRAND_DEEP = HX("006141")
End Function
Public Function C_BRAND_900() As Long
    C_BRAND_900 = HX("003323")
End Function
Public Function C_BRAND_950() As Long
    C_BRAND_950 = HX("001D14")
End Function
Public Function C_BRAND_SOFT() As Long
    C_BRAND_SOFT = HX("C2EBDA")       ' pale mint, for type reversed on black
End Function
Public Function C_BRAND_TINT() As Long
    C_BRAND_TINT = HX("E6F6EF")
End Function

' --- type on dark ---
Public Function C_TX1() As Long
    C_TX1 = HX("F2F7F4")
End Function
Public Function C_TX2() As Long
    C_TX2 = HX("A9BDB3")
End Function
Public Function C_TX3() As Long
    C_TX3 = HX("7E9388")
End Function
Public Function C_TX4() As Long
    C_TX4 = HX("4A5E55")
End Function

' --- the light canvas ---
Public Function C_PAPER() As Long
    C_PAPER = HX("FFFFFF")
End Function
Public Function C_MIST() As Long
    C_MIST = HX("F4F7F5")
End Function
Public Function C_HAIR() As Long
    C_HAIR = HX("E2E9E5")
End Function
Public Function C_MUTED() As Long
    C_MUTED = HX("5F7068")
End Function
Public Function C_BODY() As Long
    C_BODY = HX("18241F")
End Function

' --- status, on the light canvas ---
Public Function C_OK_TX() As Long
    C_OK_TX = HX("0B6B47")
End Function
Public Function C_OK_BG() As Long
    C_OK_BG = HX("DDF5EA")
End Function
Public Function C_WARN_TX() As Long
    C_WARN_TX = HX("8A5A00")
End Function
Public Function C_WARN_BG() As Long
    C_WARN_BG = HX("FFF1CF")
End Function
Public Function C_BAD_TX() As Long
    C_BAD_TX = HX("B42318")
End Function
Public Function C_BAD_BG() As Long
    C_BAD_BG = HX("FDE5E2")
End Function
Public Function C_IDLE_TX() As Long
    C_IDLE_TX = HX("5E6F67")
End Function
Public Function C_IDLE_BG() As Long
    C_IDLE_BG = HX("EDF2EF")
End Function

' --- status, on dark (the Desk) ---
Public Function C_OK_DK() As Long
    C_OK_DK = HX("2FC48D")
End Function
Public Function C_WARN_DK() As Long
    C_WARN_DK = HX("F2B544")
End Function
Public Function C_BAD_DK() As Long
    C_BAD_DK = HX("FF6B5E")
End Function
Public Function C_IDLE_DK() As Long
    C_IDLE_DK = HX("7E9388")
End Function

' --- the sheets, dark since 3.0: the Desk's room on every sheet ---
Public Function C_SHEET() As Long
    C_SHEET = HX("060B09")            ' the canvas behind everything
End Function
Public Function C_ROW() As Long
    C_ROW = HX("0C1512")              ' a table row
End Function
Public Function C_ROW_ALT() As Long
    C_ROW_ALT = HX("111F19")          ' every other row
End Function
Public Function C_LINE() As Long
    C_LINE = HX("1A2922")             ' the hairline between rows
End Function
Public Function C_LINE_2() As Long
    C_LINE_2 = HX("243A31")           ' a stronger edge
End Function
Public Function C_BAR_WELL() As Long
    C_BAR_WELL = HX("0B1310")         ' the well a group of pills sits in
End Function
Public Function C_TEXT() As Long
    C_TEXT = HX("F2F7F4")
End Function
Public Function C_TEXT_2() As Long
    C_TEXT_2 = HX("A9BDB3")
End Function
Public Function C_TEXT_3() As Long
    C_TEXT_3 = HX("7E9388")
End Function
Public Function C_LINK() As Long
    C_LINK = HX("4FC79C")             ' links and the small line over a title
End Function
Public Function C_TOTAL() As Long
    C_TOTAL = HX("001D14")            ' total rows
End Function
Public Function C_OK_BG_DK() As Long
    C_OK_BG_DK = HX("0D2A20")
End Function
Public Function C_WARN_BG_DK() As Long
    C_WARN_BG_DK = HX("2A2310")
End Function
Public Function C_BAD_BG_DK() As Long
    C_BAD_BG_DK = HX("2E1614")
End Function
Public Function C_IDLE_BG_DK() As Long
    C_IDLE_BG_DK = HX("121D19")
End Function
' The Avati mark's blue, for the word when the picture cannot be placed.
Public Function C_AVATI() As Long
    C_AVATI = HX("1C8CCB")
End Function

' A leading dot on a status cell, drawn by the number format so the cell's
' VALUE stays the plain word ("Loaded", "Break") that other code reads back.
Public Function DotFormat() As String
    DotFormat = """" & ChrW(9679) & "  ""@"
End Function

' ===================== dressing a sheet =====================================
'
' Every sheet, the tool's and the ones it builds, is the Desk's room: the
' dark canvas, a black app bar with the Avati mark, the title set light and
' large with a small emerald line over it, then the sheet's own tools, its
' status, and its table on dark rows. Rows 1 to 7 are the chrome and stay
' frozen; the table starts at row 8.
'
'   1  app bar          4  toolbar (sub-tabs and actions)
'   2  title             5  status
'   3  what it is for    6  spacer (Pivot config: the column bands)
'                        7  table header

Public Sub Dress(ByVal ws As Worksheet, ByVal title As String, ByVal about As String, _
                 Optional ByVal overline As String = "")
    On Error Resume Next
    With ws.Cells
        .Font.Name = UI_FONT
        .Font.Size = 9.5
        .Font.Color = C_TEXT
        .Interior.Color = C_SHEET
    End With
    With ws.Range(ws.Cells(R_BAR, 1), ws.Cells(R_BAR, DRESS_COLS))
        .Interior.Color = C_INK
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = C_BRAND_DEEP
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ws.Rows(R_BAR).RowHeight = 44
    ws.Rows(R_TITLE).RowHeight = 44
    ws.Rows(R_ABOUT).RowHeight = 20
    ws.Rows(4).RowHeight = 34
    ws.Rows(6).RowHeight = 8
    If Len(overline) = 0 Then overline = "ALM DESK  " & ChrW(183) & "  " & UCase$(BANK_NAME)
    TitleBlock ws, overline, title, about
    ws.Tab.Color = C_INK
    ws.DisplayPageBreaks = False
    Err.Clear
End Sub

' Every cell of a workbook starts on the dark canvas in near-white type, and
' links are emerald without an underline. Used for this workbook and for every
' workbook a build writes.
Public Sub DarkNormal(ByVal wb As Workbook)
    On Error Resume Next
    With wb.Styles("Normal")
        .IncludePatterns = True
        .Interior.Color = C_SHEET
        .Font.Color = C_TEXT
    End With
    With wb.Styles("Hyperlink")
        .Font.Color = C_LINK
        .Font.Underline = xlUnderlineStyleNone
    End With
    With wb.Styles("Followed Hyperlink")
        .Font.Color = C_LINK
        .Font.Underline = xlUnderlineStyleNone
    End With
    Err.Clear
End Sub

' The title as the Desk sets its hero: a small line in emerald capitals, the
' title under it in Segoe UI Light, what the sheet is for beneath. Shapes, not
' cells, so the title never wraps in whatever width column A happens to be.
Public Sub TitleBlock(ByVal ws As Worksheet, ByVal overline As String, ByVal title As String, _
                      ByVal about As String)
    Dim sh As Shape, t As Double
    On Error Resume Next
    ClearButtons ws, "pdt_"
    t = ws.Rows(R_TITLE).Top
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 14, t + 5, 1100, 38)
    If Not sh Is Nothing Then
        sh.Name = "pdt_title"
        Plain sh
        With sh.TextFrame2
            .VerticalAnchor = msoAnchorBottom
            .TextRange.Text = overline & vbCr & title
            With .TextRange.Paragraphs(1).Font
                .Name = UI_SEMI
                .Size = 7.5
                .Spacing = 1.4
                .Fill.ForeColor.RGB = C_LINK
            End With
            With .TextRange.Paragraphs(2).Font
                .Name = UI_LIGHT
                .Size = 19
                .Spacing = 0
                .Fill.ForeColor.RGB = C_TEXT
            End With
        End With
        sh.AlternativeText = title
    End If
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 14, ws.Rows(R_ABOUT).Top + 1, 1100, 18)
    If Not sh Is Nothing Then
        sh.Name = "pdt_about"
        Plain sh
        With sh.TextFrame2
            .VerticalAnchor = msoAnchorTop
            .TextRange.Text = IIf(Len(about) > 0, about, " ")
            .TextRange.Font.Name = UI_FONT
            .TextRange.Font.Size = 9
            .TextRange.Font.Fill.ForeColor.RGB = C_TEXT_2
        End With
    End If
    Err.Clear
End Sub

' A shape that is only its text: no fill, no outline, no margins, not
' moved or sized by the cells under it, and never selectable by accident.
Public Sub Plain(ByVal sh As Shape)
    On Error Resume Next
    sh.Fill.visible = msoFalse
    sh.Line.visible = msoFalse
    sh.Shadow.visible = msoFalse
    sh.Placement = xlFreeFloating
    With sh.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
    End With
    sh.Locked = True
    Err.Clear
End Sub

Public Sub Head(ByVal ws As Worksheet, ByRef headings As Variant, ByRef widths As Variant, _
                Optional ByVal hdrRow As Long = 0)
    Dim c As Long
    On Error Resume Next
    If hdrRow = 0 Then hdrRow = R_HDR
    For c = 0 To UBound(headings)
        ws.Cells(hdrRow, c + 1).Value2 = headings(c)
        If c <= UBound(widths) Then ws.Columns(c + 1).ColumnWidth = widths(c)
    Next c
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(hdrRow, UBound(headings) + 1))
        .Font.Name = UI_SEMI
        .Font.Size = 8.5
        .Font.Color = C_BRAND_SOFT
        .Interior.Color = C_INK
        .WrapText = False
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = C_BRAND
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = C_LINE
        .Borders(xlInsideVertical).Weight = xlThin
    End With
    ws.Rows(hdrRow).RowHeight = 28
    Err.Clear
End Sub

' The sentence under the toolbar: a tinted band, a thick accent down its
' leading edge, a dot, and the words - so a status is never colour alone.
Public Sub SetStatus(ByVal ws As Worksheet, ByVal msg As String, ByVal level As String)
    On Error Resume Next
    With ws.Range(ws.Cells(R_STATUS, 1), ws.Cells(R_STATUS, 14))
        .Interior.Color = LevelBg(level)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = LevelTx(level)
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
    With ws.Cells(R_STATUS, 1)
        .Value2 = msg
        .NumberFormat = DotFormat()
        .Font.Name = UI_SEMI
        .Font.Size = 9
        .Font.Color = LevelTx(level)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    ws.Rows(R_STATUS).RowHeight = 26
    Err.Clear
End Sub

' Verdict colours. Every sheet is dark since 3.0, so these are the Desk's.
Public Function LevelTx(ByVal level As String) As Long
    LevelTx = LevelDark(level)
End Function

Public Function LevelBg(ByVal level As String) As Long
    Select Case UCase$(level)
        Case "OK", "READY", "AGREES", "LOADED": LevelBg = C_OK_BG_DK
        Case "BREAK", "BAD", "MISSING", "FAILED": LevelBg = C_BAD_BG_DK
        Case "CHECK", "WARN", "PARTIAL", "FORCED": LevelBg = C_WARN_BG_DK
        Case Else: LevelBg = C_IDLE_BG_DK
    End Select
End Function

Public Function LevelDark(ByVal level As String) As Long
    Select Case UCase$(level)
        Case "OK", "READY", "AGREES", "LOADED": LevelDark = C_OK_DK
        Case "BREAK", "BAD", "MISSING", "FAILED": LevelDark = C_BAD_DK
        Case "CHECK", "WARN", "PARTIAL", "FORCED": LevelDark = C_WARN_DK
        Case Else: LevelDark = C_IDLE_DK
    End Select
End Function

' A status cell as a pill: tinted, centred, a dot in front of the word.
Public Sub PaintVerdict(ByVal cell As Range)
    Dim v As String
    On Error Resume Next
    v = SafeText(cell.Value2)
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
    cell.Font.Name = UI_SEMI
    cell.Font.Size = 8.5
    If Len(v) = 0 Then
        cell.Font.Color = C_TEXT_3
    Else
        cell.Font.Color = LevelTx(v)
        cell.Interior.Color = LevelBg(v)
    End If
    cell.NumberFormat = DotFormat()
    Err.Clear
End Sub

Public Sub PaintVerdictColumn(ByVal ws As Worksheet, ByVal col As Long, ByVal LastRow As Long)
    Dim r As Long
    If LastRow < R_FIRST Then Exit Sub
    For r = R_FIRST To LastRow
        PaintVerdict ws.Cells(r, col)
    Next r
End Sub

' Dark rows, a hairline between them, one row height and a filter - applied
' once per sheet at the end rather than per row. The verdict column is
' repainted last so the banding cannot wash its colour out.
Public Sub DressTable(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal LastRow As Long, _
                      Optional ByVal verdictCol As Long = 0, Optional ByVal firstRow As Long = 0)
    Dim rng As Range, r As Long
    On Error Resume Next
    If firstRow = 0 Then firstRow = R_FIRST
    If LastRow < firstRow Then Exit Sub
    Set rng = ws.Range(ws.Cells(firstRow, 1), ws.Cells(LastRow, lastCol))
    rng.Interior.Color = C_ROW
    For r = firstRow + 1 To LastRow Step 2
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = C_ROW_ALT
    Next r
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = C_LINE
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = C_LINE_2
        .Weight = xlThin
    End With
    rng.VerticalAlignment = xlCenter
    rng.IndentLevel = 1
    rng.Font.Name = UI_FONT
    rng.Font.Size = 9.5
    rng.Font.Color = C_TEXT
    ws.Range(ws.Rows(firstRow), ws.Rows(LastRow)).RowHeight = 22
    If verdictCol > 0 Then PaintVerdictColumn ws, verdictCol, LastRow
    ' AutoFilter on a range TOGGLES. Applied on every refresh it switched the
    ' filter off every other time; now it is only ever turned on.
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(firstRow - 1, 1), ws.Cells(LastRow, lastCol)).AutoFilter
    Err.Clear
End Sub

' One new Activity row, dressed in place - the log is written a line at a
' time, newest first, so there is no end-of-run pass to do it.
Public Sub DressLogRow(ByVal ws As Worksheet, ByVal r As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Interior.Color = C_ROW
        .Font.Name = UI_FONT
        .Font.Size = 9.5
        .Font.Color = C_TEXT
        .Font.Bold = False
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .WrapText = False
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = C_LINE
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Cells(r, 1)
        .Font.Name = UI_MONO
        .Font.Size = 9
        .Font.Color = C_TEXT_3
    End With
    With ws.Cells(r, 3)
        .Font.Name = UI_SEMI
        .Font.Color = C_TEXT
    End With
    ws.Cells(r, 5).Font.Color = C_TEXT_3
    ws.Rows(r).RowHeight = 22
    PaintVerdict ws.Cells(r, 2)
    Err.Clear
End Sub

' Print so it can be handed to someone: landscape, one page wide, the header
' row on every page, a footer that says what and when - and in black and
' white, so the dark screen does not become a page of toner.
Public Sub PrintReady(ByVal ws As Worksheet, ByVal lastCol As Long)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .BlackAndWhite = True
        .PrintTitleRows = ws.Rows(R_HDR).Address
        .PrintArea = ws.Range(ws.Cells(R_STATUS, 1), ws.Cells(ws.Rows.count, lastCol)).Address
        .CenterHeader = "&""Segoe UI Semibold,Regular""&11" & ws.Name
        .LeftFooter = "&8" & TOOL_NAME & " ALM Desk " & TOOL_VERSION & "  " & ChrW(183) & "  " & BANK_NAME
        .RightFooter = "&8Page &P of &N  " & ChrW(183) & "  &D"
        .CenterFooter = ""
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub

' ===================== the logo =============================================
'
' A macro cannot read the images packed inside its own workbook, so the build
' leaves the Avati mark on the settings sheet as base64. The first time it is
' wanted it is written to the temp folder once, and placed from there - on
' every sheet's bar here, and in every workbook a build writes.
Public Function LogoFile() As String
    Dim f As String, b64 As String, doc As Object, node As Object, bytes() As Byte, fh As Integer
    On Error GoTo Nope
    f = Environ$("TEMP") & "\avati-logo-" & Replace(TOOL_VERSION, ".", "-") & ".png"
    If Len(Dir$(f)) > 0 Then LogoFile = f: Exit Function
    b64 = SettingGet("logo_b64")
    If Len(b64) = 0 Then Exit Function
    Set doc = CreateObject("MSXML2.DOMDocument")
    Set node = doc.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = b64
    bytes = node.nodeTypedValue
    fh = FreeFile
    Open f For Binary Access Write As #fh
    Put #fh, , bytes
    Close #fh
    LogoFile = f
    Exit Function
Nope:
    LogoFile = ""
End Function

' The mark at (x, y), this many points tall. If the picture cannot be had,
' the word AVATI in the mark's blue stands in, so no bar is ever left bare.
Public Function PlaceLogo(ByVal ws As Worksheet, ByVal nm As String, ByVal x As Double, ByVal y As Double, _
                          ByVal h As Double) As Shape
    Dim f As String, sh As Shape
    On Error Resume Next
    f = LogoFile()
    If Len(f) > 0 Then Set sh = ws.Shapes.AddPicture(f, msoFalse, msoTrue, x, y, h * LOGO_RATIO, h)
    If sh Is Nothing Then
        Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y - 2, h * LOGO_RATIO + 8, h + 4)
        If sh Is Nothing Then Exit Function
        Plain sh
        With sh.TextFrame2.TextRange
            .Text = "AVATI"
            .Font.Name = UI_SEMI
            .Font.Size = h * 0.95
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = C_AVATI
        End With
    End If
    sh.Name = nm
    sh.Placement = xlFreeFloating
    sh.AlternativeText = "Avati"
    Set PlaceLogo = sh
    Err.Clear
End Function

' ===================== the app bar and the toolbar ==========================
'
' Excel shapes have no hover state, so what a shape CAN do it does well: a
' pill group for where you can go with the sheet you are on filled in, and a
' toolbar under the title for what this sheet can do. The Desk is designed
' in the build and carries its own.

' Clears only the family it is asked for: the bar and toolbar own "pdr_", the
' title block "pdt_", the Desk's own design "pdx_", and none can remove another.
Public Sub ClearButtons(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long, sh As Shape
    On Error Resume Next
    For i = ws.Shapes.count To 1 Step -1
        Set sh = ws.Shapes(i)
        If Left$(sh.Name, Len(prefix)) = prefix Then sh.Delete
    Next i
    Err.Clear
End Sub

' A rounded pill. kind: 0 quiet (another sheet), 1 action (emerald), 2 current
' (this sheet), 3 container (the group behind a set of pills).
Public Function Pill(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, ByVal proc As String, _
                     ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                     ByVal kind As Long) As Shape
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    If sh Is Nothing Then Exit Function
    sh.Name = nm
    sh.Placement = xlFreeFloating
    sh.Adjustments(1) = 0.5
    sh.Shadow.visible = msoFalse
    sh.Line.Weight = 0.75
    Select Case kind
        Case 1
            sh.Fill.ForeColor.RGB = C_BRAND_950
            sh.Line.ForeColor.RGB = C_BRAND_DEEP
        Case 2
            sh.Fill.ForeColor.RGB = C_BRAND_900
            sh.Line.ForeColor.RGB = C_BRAND_DEEP
        Case 3
            sh.Fill.ForeColor.RGB = C_BAR_WELL
            sh.Line.ForeColor.RGB = C_LINE
        Case Else
            ' Invisible but still there: a fully transparent fill keeps the
            ' whole pill clickable, where no fill would leave only the letters.
            sh.Fill.ForeColor.RGB = C_INK
            sh.Fill.Transparency = 1
            sh.Line.visible = msoFalse
    End Select
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 4: .MarginRight = 4
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Text = IIf(Len(caption) > 0, caption, " ")
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoFalse
        .TextRange.Font.Name = UI_SEMI
        Select Case kind
            Case 1: .TextRange.Font.Fill.ForeColor.RGB = C_BRAND_SOFT
            Case 2: .TextRange.Font.Fill.ForeColor.RGB = C_TEXT
            Case Else: .TextRange.Font.Fill.ForeColor.RGB = C_TEXT_2
        End Select
    End With
    sh.AlternativeText = IIf(Len(caption) > 0, caption, nm)
    If Len(proc) > 0 Then sh.OnAction = proc
    Set Pill = sh
    Err.Clear
End Function

' Kept for the older calls: a pill on the app bar, returning where the next goes.
Public Function Button(ByVal ws As Worksheet, ByVal caption As String, ByVal proc As String, _
                       ByVal x As Double, ByVal w As Double, Optional ByVal kind As Long = 0) As Double
    Pill ws, "pdr_" & CLng(x) & "_" & CLng(Rnd * 100000), caption, proc, x, (ws.Rows(R_BAR).Height - 22) / 2, w, 22, kind
    Button = x + w + 4
End Function

' A hairline separating one group from the next.
Public Function Divider(ByVal ws As Worksheet, ByVal x As Double, Optional ByVal row As Long = 0) As Double
    Dim sh As Shape, top As Double, h As Double
    On Error Resume Next
    If row = 0 Then row = R_BAR
    h = 16
    top = ws.Rows(row).Top + (ws.Rows(row).Height - h) / 2
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x + 6, top, 0.75, h)
    If Not sh Is Nothing Then
        sh.Name = "pdr_div_" & row & "_" & CLng(x)
        sh.Placement = xlFreeFloating
        sh.Line.visible = msoFalse
        sh.Fill.ForeColor.RGB = C_LINE_2
        sh.Shadow.visible = msoFalse
    End If
    Err.Clear
    Divider = x + 18
End Function

Private Function BarText(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal x As Double, _
                       ByVal y As Double, ByVal size As Single, ByVal clr As Long, ByVal font As String, _
                       Optional ByVal spacing As Single = 0) As Shape
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, 200, size + 5)
    If sh Is Nothing Then Exit Function
    sh.Name = nm
    Plain sh
    With sh.TextFrame2.TextRange
        .Text = txt
        .Font.Name = font
        .Font.Size = size
        .Font.Spacing = spacing
        .Font.Fill.ForeColor.RGB = clr
    End With
    Set BarText = sh
    Err.Clear
End Function

' The bar every sheet carries: the Avati mark, what this is, and where you
' can go - the same on every sheet, with the one you are on filled in.
Public Sub Rail(ByVal ws As Worksheet)
    Dim x As Double, y As Double, items As Variant, it As Variant, cur As String, w As Double
    On Error Resume Next
    If StrComp(ws.Name, SH_HOME, vbTextCompare) = 0 Then Exit Sub
    ClearButtons ws, "pdr_"
    ws.Rows(R_BAR).RowHeight = 44
    PlaceLogo ws, "pdr_logo", 16, (44 - 16) / 2, 16
    ws.Shapes("pdr_logo").OnAction = "PD_GoHome"
    x = 16 + 16 * LOGO_RATIO
    x = Divider(ws, x + 4)
    BarText ws, "pdr_what", "ALM DESK", x, 11, 7.5, C_TEXT, UI_SEMI, 1.6
    BarText ws, "pdr_whose", UCase$(Replace(BANK_NAME, "  ", "  " & ChrW(183) & "  ")), x, 24, 6, C_TEXT_3, UI_SEMI, 1.4

    cur = SectionOf(ws.Name)
    items = Array(Array("Desk", "PD_GoHome", 52, "desk"), Array("Files", "PD_GoFiles", 52, "files"), _
                  Array("Reports", "PD_GoConfig", 70, "reports"), _
                  Array("Reconciliation", "PD_GoRecon", 102, "recon"), Array("Activity", "PD_GoLog", 66, "log"))
    x = x + 104
    w = 3
    For Each it In items
        w = w + CDbl(it(2)) + 2
    Next it
    y = (44 - 26) / 2
    Pill ws, "pdr_nav", "", "", x, y, w + 1, 26, 3
    x = x + 3
    For Each it In items
        Pill ws, "pdr_nav_" & CStr(it(3)), CStr(it(0)), CStr(it(1)), x, y + 2, CDbl(it(2)), 22, _
             IIf(cur = CStr(it(3)), 2, 0)
        x = x + CDbl(it(2)) + 2
    Next it
    Toolbar ws
    Err.Clear
End Sub

' Which part of the bar a sheet belongs to.
Private Function SectionOf(ByVal nm As String) As String
    Select Case True
        Case StrComp(nm, SH_SOURCES, vbTextCompare) = 0: SectionOf = "files"
        Case StrComp(nm, SH_RECON, vbTextCompare) = 0: SectionOf = "recon"
        Case StrComp(nm, SH_LOG, vbTextCompare) = 0: SectionOf = "log"
        Case IsReportsSheet(nm): SectionOf = "reports"
        Case Else: SectionOf = ""
    End Select
End Function

Public Function IsReportsSheet(ByVal nm As String) As Boolean
    Dim s As Variant
    For Each s In ReportsSheets()
        If StrComp(nm, CStr(s(0)), vbTextCompare) = 0 Then IsReportsSheet = True: Exit Function
    Next s
End Function

' The Reports section: every sheet that shapes what a build makes, as the
' tabs across its toolbar.
Public Function ReportsSheets() As Variant
    ReportsSheets = Array(Array(SH_CONFIG, "Pivots", "PD_GoConfig"), Array(SH_FIELDS, "Fields", "PD_GoFields"))
End Function

' Row 4: for the Reports sheets, their tabs; then what this sheet can do.
Public Sub Toolbar(ByVal ws As Worksheet)
    Dim x As Double, y As Double, tabs As Variant, t As Variant, w As Double, tw As Double
    On Error Resume Next
    ws.Rows(4).RowHeight = 34
    y = ws.Rows(4).Top + (34 - 24) / 2
    x = 14
    If IsReportsSheet(ws.Name) Then
        tabs = ReportsSheets()
        tw = 3
        For Each t In tabs
            tw = tw + 74 + 2
        Next t
        Pill ws, "pdr_tabs", "", "", x, y - 1, tw + 1, 26, 3
        x = x + 3
        For Each t In tabs
            Pill ws, "pdr_tab_" & Replace(LCase$(CStr(t(1))), " ", "_"), CStr(t(1)), CStr(t(2)), x, y + 1, 74, 22, _
                 IIf(StrComp(ws.Name, CStr(t(0)), vbTextCompare) = 0, 2, 0)
            x = x + 76
        Next t
        x = Divider(ws, x + 4, 4)
    End If
    Select Case ws.Name
        Case SH_SOURCES
            x = Action(ws, "Scan a folder", "PD_LoadFolder", x, y, 96)
            x = Action(ws, "Pick files", "PD_LoadFiles", x, y, 76)
            x = Action(ws, "Use a file for this row", "PD_UseFileHere", x, y, 142)
            x = Action(ws, "Clear this row", "PD_ClearRow", x, y, 98)
        Case SH_CONFIG
            x = Action(ws, "Add a pivot", "PD_ConfigAdd", x, y, 84)
            x = Action(ws, "Check", "PD_ConfigCheck", x, y, 62)
            x = Action(ws, "Restore defaults", "PD_ConfigDefaults", x, y, 112)
            If modPD_Config.Engine() = "recipes" Then
                x = Action(ws, "Use the 1.0 layout", "PD_ConfigEngine", x, y, 122)
            Else
                x = Action(ws, "Build from this sheet", "PD_ConfigEngine", x, y, 136)
            End If
        Case SH_FIELDS
            x = Action(ws, "Check", "PD_ConfigCheck", x, y, 62)
        Case SH_RECON
            x = Action(ws, "Reconcile now", "PD_Reconcile", x, y, 102)
        Case SH_LOG
            x = Action(ws, "Clear activity", "PD_ClearLog", x, y, 100)
    End Select
    Err.Clear
End Sub

Private Function Action(ByVal ws As Worksheet, ByVal caption As String, ByVal proc As String, _
                        ByVal x As Double, ByVal y As Double, ByVal w As Double) As Double
    Pill ws, "pdr_act_" & proc, caption, proc, x, y, w, 24, 1
    Action = x + w + 6
End Function

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
    If StrComp(nm, SH_HOME, vbTextCompare) = 0 Then
        modPD_Desk.FitDesk
    Else
        ws.Range("A1").Select
    End If
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
Public Sub PD_GoConfig()
    EnsureConfig
    GoTo_ SH_CONFIG
End Sub
Public Sub PD_GoFields()
    EnsureConfig
    GoTo_ SH_FIELDS
End Sub

Private Sub EnsureConfig()
    On Error Resume Next
    If GetSheet(SH_CONFIG) Is Nothing Or GetSheet(SH_FIELDS) Is Nothing Then
        Application.ScreenUpdating = False
        modPD_Config.BuildConfigSheet
        modPD_Config.CheckAll True
        Rail GetSheet(SH_CONFIG)
        Rail GetSheet(SH_FIELDS)
        Application.ScreenUpdating = True
    End If
    Err.Clear
End Sub
