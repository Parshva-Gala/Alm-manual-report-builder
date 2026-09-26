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
'  DARK CHROME, LIGHT DATA. The chrome is black and unbroken so every sheet
'  reads as the bank's and as one application; the data area stays white
'  because a balance is what the sheet is for and nobody reads one off a dark
'  field. The Desk is the one dark room - it holds no balances.
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

' A leading dot on a status cell, drawn by the number format so the cell's
' VALUE stays the plain word ("Loaded", "Break") that other code reads back.
Public Function DotFormat() As String
    DotFormat = """" & ChrW(9679) & "  ""@"
End Function

' ===================== dressing a sheet =====================================

Public Sub Dress(ByVal ws As Worksheet, ByVal title As String, ByVal about As String)
    Dim band As Range
    On Error Resume Next
    With ws.Cells
        .Font.Name = UI_FONT
        .Font.Size = 9.5
        .Font.Color = C_BODY
        .Interior.Color = C_PAPER
    End With

    ' The chrome: the app bar and the masthead as one black block, the title
    ' reversed out of it, an emerald rule underneath and an emerald bar down
    ' its leading edge - the logo's mark sits to the left of the word, and the
    ' sheet repeats that shape rather than borrowing a generic accent stripe.
    Set band = ws.Range(ws.Cells(R_BAR, 1), ws.Cells(R_ABOUT, DRESS_COLS))
    band.Interior.Color = C_INK
    band.Borders(xlInsideHorizontal).LineStyle = xlNone

    With ws.Cells(R_TITLE, 1)
        .Value2 = title
        .Font.Name = UI_SEMI
        .Font.Size = 17
        .Font.Color = C_TX1
        .IndentLevel = 1
        .VerticalAlignment = xlBottom
        .WrapText = False
    End With
    With ws.Cells(R_ABOUT, 1)
        .Value2 = about
        .Font.Name = UI_FONT
        .Font.Size = 9
        .Font.Color = C_TX2
        .IndentLevel = 1
        .VerticalAlignment = xlTop
        .WrapText = False
    End With
    With ws.Range(ws.Cells(R_ABOUT, 1), ws.Cells(R_ABOUT, DRESS_COLS)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = C_BRAND
        .Weight = xlThick
    End With
    With ws.Range(ws.Cells(R_TITLE, 1), ws.Cells(R_ABOUT, 1)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = C_BRAND
        .Weight = xlThick
    End With

    ws.Rows(R_BAR).RowHeight = 34
    ws.Rows(R_TITLE).RowHeight = 30
    ws.Rows(R_ABOUT).RowHeight = 21
    ws.Rows(4).RowHeight = 10
    ws.Rows(6).RowHeight = 10
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
    End With
    ws.Rows(R_HDR).RowHeight = 26
    Err.Clear
End Sub

' The sentence under the masthead: a tinted field, a thick accent down the
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
        .Font.Size = 9.5
        .Font.Color = LevelTx(level)
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    ws.Rows(R_STATUS).RowHeight = 26
    Err.Clear
End Sub

Public Function LevelTx(ByVal level As String) As Long
    Select Case UCase$(level)
        Case "OK", "READY", "AGREES", "LOADED": LevelTx = C_OK_TX
        Case "BREAK", "BAD", "MISSING", "FAILED": LevelTx = C_BAD_TX
        Case "CHECK", "WARN", "PARTIAL": LevelTx = C_WARN_TX
        Case Else: LevelTx = C_IDLE_TX
    End Select
End Function

Public Function LevelBg(ByVal level As String) As Long
    Select Case UCase$(level)
        Case "OK", "READY", "AGREES", "LOADED": LevelBg = C_OK_BG
        Case "BREAK", "BAD", "MISSING", "FAILED": LevelBg = C_BAD_BG
        Case "CHECK", "WARN", "PARTIAL": LevelBg = C_WARN_BG
        Case Else: LevelBg = C_IDLE_BG
    End Select
End Function

' The same verdicts on the dark Desk.
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
    cell.Font.Color = LevelTx(v)
    cell.Interior.Color = LevelBg(v)
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

' Hairlines, quiet banding, one row height and a filter - applied once per
' sheet at the end rather than per row, which is the difference between a
' second and a minute. The verdict column is repainted last so the banding
' cannot wash its colour out.
Public Sub DressTable(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal LastRow As Long, _
                      Optional ByVal verdictCol As Long = 0)
    Dim rng As Range, r As Long
    On Error Resume Next
    If LastRow < R_FIRST Then Exit Sub
    Set rng = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(LastRow, lastCol))
    rng.Interior.Color = C_PAPER
    For r = R_FIRST + 1 To LastRow Step 2
        ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = C_MIST
    Next r
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = C_HAIR
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = C_HAIR
        .Weight = xlThin
    End With
    rng.VerticalAlignment = xlCenter
    rng.IndentLevel = 1
    rng.Font.Name = UI_FONT
    rng.Font.Size = 9.5
    rng.Font.Color = C_BODY
    ws.Range(ws.Rows(R_FIRST), ws.Rows(LastRow)).RowHeight = 21
    If verdictCol > 0 Then PaintVerdictColumn ws, verdictCol, LastRow
    ' AutoFilter on a range TOGGLES. Applied on every refresh it switched the
    ' filter off every other time; now it is only ever turned on.
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(R_HDR, 1), ws.Cells(LastRow, lastCol)).AutoFilter
    Err.Clear
End Sub

' One new Activity row, dressed in place - the log is written a line at a
' time, newest first, so there is no end-of-run pass to do it.
Public Sub DressLogRow(ByVal ws As Worksheet, ByVal r As Long)
    On Error Resume Next
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Interior.Color = C_PAPER
        .Font.Name = UI_FONT
        .Font.Size = 9.5
        .Font.Color = C_BODY
        .Font.Bold = False
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .WrapText = False
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = C_HAIR
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    With ws.Cells(r, 1)
        .Font.Name = UI_MONO
        .Font.Size = 9
        .Font.Color = C_MUTED
    End With
    With ws.Cells(r, 3)
        .Font.Name = UI_SEMI
        .Font.Color = C_BODY
    End With
    ws.Cells(r, 5).Font.Color = C_MUTED
    ws.Rows(r).RowHeight = 21
    PaintVerdict ws.Cells(r, 2)
    Err.Clear
End Sub

' Print so it can be handed to someone: landscape, one page wide, the header
' row on every page, and a footer that says what and when.
Public Sub PrintReady(ByVal ws As Worksheet, ByVal lastCol As Long)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ws.Rows(R_HDR).Address
        .PrintArea = ws.Range(ws.Cells(R_TITLE, 1), ws.Cells(ws.Rows.count, lastCol)).Address
        .LeftFooter = "&8" & TOOL_NAME & " " & TOOL_VERSION & "  " & ChrW(183) & "  &A"
        .RightFooter = "&8Page &P of &N  " & ChrW(183) & "  &D"
        .CenterFooter = ""
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub

' ===================== the app bar ==========================================
'
' Excel shapes have no hover state, so what a shape CAN do it does well: a
' filled pill for the sheet you are on, quiet text for the others, and an
' emerald-tinted button for each thing this sheet can do.

' Clears only the family it is asked for.
'
' It used to clear every "pd_" shape, and the rail is laid down AFTER the desk's
' cards - so installing the rail deleted the three card buttons that are the
' whole point of the Desk sheet. The rail owns "pdr_", the Desk's own design
' owns "pdx_", and neither can now remove the other.
Public Sub ClearButtons(ByVal ws As Worksheet, ByVal prefix As String)
    Dim i As Long, sh As Shape
    On Error Resume Next
    For i = ws.Shapes.count To 1 Step -1
        Set sh = ws.Shapes(i)
        If Left$(sh.Name, Len(prefix)) = prefix Then sh.Delete
    Next i
    Err.Clear
End Sub

' kind: 0 quiet (another sheet), 1 action (emerald), 2 current (this sheet)
Public Function Button(ByVal ws As Worksheet, ByVal caption As String, ByVal proc As String, _
                       ByVal x As Double, ByVal w As Double, Optional ByVal kind As Long = 0) As Double
    Dim sh As Shape, yTop As Double
    On Error Resume Next
    yTop = (ws.Rows(R_BAR).Height - 22) / 2
    If yTop < 2 Then yTop = 2
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, yTop, w, 22)
    If sh Is Nothing Then Button = x: Exit Function
    sh.Name = "pdr_" & CLng(x) & "_" & CLng(Rnd * 100000)
    sh.Placement = xlFreeFloating
    sh.Adjustments(1) = 0.5
    sh.Shadow.visible = msoFalse
    Select Case kind
        Case 1
            sh.Fill.ForeColor.RGB = C_BRAND_950
            sh.Fill.Transparency = 0
            sh.Line.visible = msoTrue
            sh.Line.ForeColor.RGB = C_BRAND_DEEP
            sh.Line.Weight = 0.75
        Case 2
            sh.Fill.ForeColor.RGB = C_BRAND_900
            sh.Fill.Transparency = 0
            sh.Line.visible = msoTrue
            sh.Line.ForeColor.RGB = C_BRAND_DEEP
            sh.Line.Weight = 0.75
        Case Else
            ' Invisible but still there: a fully transparent fill keeps the
            ' whole pill clickable, where no fill at all would leave only the
            ' letters to hit.
            sh.Fill.ForeColor.RGB = C_INK
            sh.Fill.Transparency = 1
            sh.Line.visible = msoFalse
    End Select
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 4: .MarginRight = 4
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Text = caption
        .TextRange.Font.Size = 8.5
        .TextRange.Font.Bold = msoFalse
        .TextRange.Font.Name = UI_SEMI
        Select Case kind
            Case 1: .TextRange.Font.Fill.ForeColor.RGB = C_BRAND_SOFT
            Case 2: .TextRange.Font.Fill.ForeColor.RGB = C_TX1
            Case Else: .TextRange.Font.Fill.ForeColor.RGB = C_TX2
        End Select
    End With
    sh.AlternativeText = caption
    If Len(proc) > 0 Then sh.OnAction = proc
    Err.Clear
    Button = x + w + 4
End Function

' A hairline separating where-to-go from what-to-do.
Public Function Divider(ByVal ws As Worksheet, ByVal x As Double) As Double
    Dim sh As Shape, yTop As Double
    On Error Resume Next
    yTop = (ws.Rows(R_BAR).Height - 14) / 2
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x + 6, yTop, 0.75, 14)
    If Not sh Is Nothing Then
        sh.Name = "pdr_div_" & CLng(x)
        sh.Placement = xlFreeFloating
        sh.Line.visible = msoFalse
        sh.Fill.ForeColor.RGB = C_HAIR_DARK
        sh.Shadow.visible = msoFalse
    End If
    Err.Clear
    Divider = x + 16
End Function

' The wordmark at the head of the bar: "Pivot" in white, "Desk" in emerald.
Private Function Wordmark(ByVal ws As Worksheet, ByVal x As Double) As Double
    Dim sh As Shape, yTop As Double
    On Error Resume Next
    yTop = (ws.Rows(R_BAR).Height - 22) / 2
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, yTop, 74, 22)
    If sh Is Nothing Then Wordmark = x: Exit Function
    sh.Name = "pdr_mark"
    sh.Placement = xlFreeFloating
    sh.Fill.visible = msoFalse
    sh.Line.visible = msoFalse
    With sh.TextFrame2
        .MarginTop = 0: .MarginBottom = 0: .MarginLeft = 0: .MarginRight = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.Text = TOOL_NAME
        .TextRange.Font.Name = UI_SEMI
        .TextRange.Font.Size = 11
        .TextRange.Font.Fill.ForeColor.RGB = C_TX1
        .TextRange.Characters(6, 4).Font.Fill.ForeColor.RGB = HX("4FC79C")
    End With
    sh.OnAction = "PD_GoHome"
    sh.AlternativeText = TOOL_NAME
    Err.Clear
    Wordmark = x + 78
End Function

' The bar every desk sheet carries. Where you are does not change how you
' get out, and what the sheet can do sits right beside it.
Public Sub Rail(ByVal ws As Worksheet)
    Dim x As Double
    On Error Resume Next
    ' The Desk is designed, not drawn: it carries its own bar.
    If StrComp(ws.Name, SH_HOME, vbTextCompare) = 0 Then Exit Sub
    ClearButtons ws, "pdr_"
    x = 12
    x = Wordmark(ws, x)
    x = Divider(ws, x - 8)
    x = Button(ws, "Desk", "PD_GoHome", x, 50, 0)
    x = Button(ws, "Files", "PD_GoFiles", x, 50, IIf(ws.Name = SH_SOURCES, 2, 0))
    x = Button(ws, "Pivot config", "PD_GoConfig", x, 86, IIf(ws.Name = SH_CONFIG Or ws.Name = SH_FIELDS, 2, 0))
    x = Button(ws, "Reconciliation", "PD_GoRecon", x, 96, IIf(ws.Name = SH_RECON, 2, 0))
    x = Button(ws, "Activity", "PD_GoLog", x, 64, IIf(ws.Name = SH_LOG, 2, 0))
    x = Divider(ws, x)
    Select Case ws.Name
        Case SH_SOURCES
            x = Button(ws, "Scan a folder", "PD_LoadFolder", x, 92, 1)
            x = Button(ws, "Pick files", "PD_LoadFiles", x, 72, 1)
            x = Button(ws, "Use a file for this row", "PD_UseFileHere", x, 138, 1)
            x = Button(ws, "Clear this row", "PD_ClearRow", x, 94, 1)
        Case SH_CONFIG
            x = Button(ws, "Add a pivot", "PD_ConfigAdd", x, 82, 1)
            x = Button(ws, "Check", "PD_ConfigCheck", x, 60, 1)
            x = Button(ws, "Fields", "PD_GoFields", x, 56, 1)
            x = Button(ws, "Restore defaults", "PD_ConfigDefaults", x, 108, 1)
            If modPD_Config.Engine() = "recipes" Then
                x = Button(ws, "Use the 1.0 layout", "PD_ConfigEngine", x, 118, 1)
            Else
                x = Button(ws, "Build from this sheet", "PD_ConfigEngine", x, 132, 1)
            End If
        Case SH_FIELDS
            x = Button(ws, "Back to recipes", "PD_GoConfig", x, 104, 1)
            x = Button(ws, "Check", "PD_ConfigCheck", x, 60, 1)
        Case SH_RECON
            x = Button(ws, "Reconcile now", "PD_Reconcile", x, 98, 1)
        Case SH_LOG
            x = Button(ws, "Clear activity", "PD_ClearLog", x, 96, 1)
    End Select
    ws.Rows(R_BAR).RowHeight = 34
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
