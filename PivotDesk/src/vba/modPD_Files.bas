Option Explicit

' ============================================================================
'  What you uploaded, and what it turned out to be.
'
'  A file is identified by ITS COLUMNS, never by its name. The exports in this
'  bank are demonstrably misnamed - a file called "LCR_Portfolio Segmentation"
'  holds the NSFR set - so a tool that trusts a filename files things under the
'  wrong framework and then disagrees with itself with nothing saying why.
'
'  Nothing is copied into the desk. A slot holds a PATH; the pass opens the
'  file, reads it in blocks, keeps only the columns being staged, and lets each
'  block go. Memory stays flat whatever the file size.
' ============================================================================

' Files sheet columns
Public Const S_WHAT As Long = 1
Public Const S_STATUS As Long = 2
Public Const S_FILE As Long = 3
Public Const S_SHEET As Long = 4
Public Const S_HDRROW As Long = 5
Public Const S_ROWS As Long = 6
Public Const S_ASOF As Long = 7
Public Const S_AMOUNT As Long = 8
Public Const S_NOTE As Long = 9
Public Const S_COLS As Long = 9

Private mSlots As Object
Private mRefused As Collection

Public Function Slots() As Object
    If mSlots Is Nothing Then BuildSlots
    Set Slots = mSlots
End Function

Private Sub BuildSlots()
    Dim fw As Variant
    Set mSlots = NewMap()
    For Each fw In Frameworks()
        AddSlot "OUTPUT|" & CStr(fw), FwLabel(CStr(fw)) & " output"
    Next fw
    AddSlot "CTRL3|", "Control report 3  (by COA)"
    AddSlot "CTRL6|", "Control report 6  (by account)"
End Sub

Private Sub AddSlot(ByVal key As String, ByVal what As String)
    Dim x As Object
    Set x = NewMap()
    x("Key") = key: x("What") = what
    Set mSlots(key) = x
End Sub

Public Function SlotRow(ByVal key As String) As Long
    Dim i As Long, k As Variant
    For Each k In Slots.keys
        If StrComp(CStr(k), key, vbTextCompare) = 0 Then SlotRow = modPD_Theme.R_FIRST + i: Exit Function
        i = i + 1
    Next k
End Function

Public Function KeyForRow(ByVal r As Long) As String
    Dim i As Long, k As Variant
    For Each k In Slots.keys
        If modPD_Theme.R_FIRST + i = r Then KeyForRow = CStr(k): Exit Function
        i = i + 1
    Next k
End Function

Public Function SlotFile(ByVal key As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Function
    r = SlotRow(key)
    If r = 0 Then Exit Function
    SlotFile = SafeText(ws.Cells(r, S_FILE).Value2)
End Function

Public Function SlotLoaded(ByVal key As String) As Boolean
    Dim p As String
    p = SlotFile(key)
    If Len(p) = 0 Then Exit Function
    SlotLoaded = (Len(Dir$(p)) > 0)
End Function

Public Function AnyFrameworkLoaded() As Boolean
    Dim fw As Variant
    For Each fw In Frameworks()
        If SlotLoaded("OUTPUT|" & CStr(fw)) Then AnyFrameworkLoaded = True: Exit Function
    Next fw
End Function

' ===================== the Files sheet ======================================

Public Sub BuildFilesSheet()
    Dim ws As Worksheet, k As Variant, r As Long
    Set ws = EnsureSheet(SH_SOURCES)
    modPD_Theme.Dress ws, "Files", _
        "Every file is identified by the columns it carries, not by its name. Use the console to add them, " & _
        "or click a row here and press " & Chr$(34) & "Use a file for this row" & Chr$(34) & " to place one by hand."
    modPD_Theme.Head ws, Array("What", "Status", "File", "Sheet", "Header row", "Rows", "As of", "Amount field", "Note"), _
                        Array(30, 12, 54, 20, 11, 12, 13, 34, 56)
    r = modPD_Theme.R_FIRST
    For Each k In Slots.keys
        ws.Cells(r, S_WHAT).Value2 = CStr(Slots(k)("What"))
        r = r + 1
    Next k
    RefreshStatuses
End Sub

Public Sub RefreshStatuses()
    Dim ws As Worksheet, k As Variant, r As Long, p As String, n As Long, total As Long
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Sub
    r = modPD_Theme.R_FIRST
    For Each k In Slots.keys
        total = total + 1
        p = SafeText(ws.Cells(r, S_FILE).Value2)
        If Len(p) = 0 Then
            ws.Cells(r, S_STATUS).Value2 = "Empty"
        ElseIf Len(Dir$(p)) = 0 Then
            ws.Cells(r, S_STATUS).Value2 = "Missing"
            ws.Cells(r, S_NOTE).Value2 = "The file is no longer at that path."
        Else
            ws.Cells(r, S_STATUS).Value2 = "Loaded"
            n = n + 1
        End If
        modPD_Theme.PaintVerdict ws.Cells(r, S_STATUS)
        r = r + 1
    Next k
    ws.Range(ws.Cells(modPD_Theme.R_FIRST, S_ROWS), ws.Cells(r - 1, S_ROWS)).NumberFormat = NUM_FMT
    modPD_Theme.DressTable ws, S_COLS, r - 1
    modPD_Theme.SetStatus ws, n & " of " & total & " file(s) loaded." & _
        IIf(AnyFrameworkLoaded(), "  Ready to build pivots.", "  Load at least one framework output to build pivots."), _
        IIf(n = 0, "Idle", IIf(AnyFrameworkLoaded(), "OK", "Check"))
End Sub

' ===================== loading ==============================================

Public Function Refused() As Collection
    If mRefused Is Nothing Then Set mRefused = New Collection
    Set Refused = mRefused
End Function

Public Sub ClearRefused()
    Set mRefused = New Collection
End Sub

Public Sub PD_LoadFolder()
    Dim fld As String, st As Object, n As Long
    If PD_Busy Then Exit Sub
    On Error GoTo Failed
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Choose the folder holding the outputs and control reports"
        If .Show <> -1 Then Exit Sub
        fld = .SelectedItems(1)
    End With
    Set st = CaptureState(): PD_Busy = True
    ClearRefused
    n = ScanFolder(fld, 0)
    RefreshStatuses
    RestoreState st: PD_Busy = False
    LogIt V_OK, "Load", n & " file(s) placed.", fld
    Tell n & " file(s) placed." & vbCrLf & vbCrLf & _
         "Anything not recognised is on Activity with the reason.", vbInformation
    Exit Sub
Failed:
    RestoreState st: PD_Busy = False
    LogIt V_BREAK, "Load", Err.Description, fld
    Tell "The folder could not be read:" & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub PD_LoadFiles()
    Dim st As Object, i As Long, n As Long, tried As Long, p As String, missed As String
    If PD_Busy Then Exit Sub
    On Error GoTo Failed
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Choose the file or files to load"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx; *.xlsm; *.xlsb"
        If .Show <> -1 Then Exit Sub
        Set st = CaptureState(): PD_Busy = True
        For i = 1 To .SelectedItems.count
            p = CStr(.SelectedItems(i))
            tried = tried + 1
            Step_ "reading " & FileLeaf(p)
            If AssignFile(p) Then n = n + 1 Else missed = missed & vbCrLf & "   " & FileLeaf(p)
        Next i
    End With
    RefreshStatuses
    RestoreState st: PD_Busy = False
    LogIt V_OK, "Load", n & " of " & tried & " file(s) placed.", ""
    Tell n & " of " & tried & " file(s) placed." & _
         IIf(Len(missed) > 0, vbCrLf & vbCrLf & "Not placed:" & missed & vbCrLf & vbCrLf & _
             "Activity says why for each.", ""), vbInformation
    Exit Sub
Failed:
    RestoreState st: PD_Busy = False
    LogIt V_BREAK, "Load", Err.Description, ""
    Tell "The files could not be read:" & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub

' This file, in this slot, whatever its contents look like. The contents are
' still read and still reported - a file forced into a slot it does not match is
' worth knowing about - but the reading does not overrule you.
Public Sub PD_UseFileHere()
    Dim ws As Worksheet, r As Long, key As String, what As String, st As Object, p As String
    If PD_Busy Then Exit Sub
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Sub
    If StrComp(ActiveSheet.Name, SH_SOURCES, vbTextCompare) <> 0 Then
        modPD_Theme.GoTo_ SH_SOURCES
        Tell "Click the row you want to fill first, then press this again.", vbInformation
        Exit Sub
    End If
    r = ActiveCell.Row
    key = KeyForRow(r)
    If Len(key) = 0 Then
        Tell "That is not one of the file rows." & vbCrLf & vbCrLf & _
             "Click a row between " & modPD_Theme.R_FIRST & " and " & _
             (modPD_Theme.R_FIRST + Slots.count - 1) & " and press this again.", vbInformation
        Exit Sub
    End If
    what = SafeText(ws.Cells(r, S_WHAT).Value2)

    On Error GoTo Failed
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Choose the file to use as: " & what
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel workbooks", "*.xlsx; *.xlsm; *.xlsb"
        If .Show <> -1 Then Exit Sub
        p = CStr(.SelectedItems(1))
    End With
    Set st = CaptureState(): PD_Busy = True
    If AssignFile(p, key) Then
        RefreshStatuses
        RestoreState st: PD_Busy = False
        Tell FileLeaf(p) & " is now the " & what & "." & vbCrLf & vbCrLf & _
             "The Note column says what its contents looked like, which is worth a glance if it " & _
             "disagrees with you.", vbInformation
    Else
        RefreshStatuses
        RestoreState st: PD_Busy = False
        Tell FileLeaf(p) & " could not be used - Activity says why.", vbExclamation
    End If
    Exit Sub
Failed:
    RestoreState st: PD_Busy = False
    LogIt V_BREAK, "Load", Err.Description, FileLeaf(p)
    Tell "The file could not be read:" & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub PD_ClearRow()
    Dim ws As Worksheet, r As Long, key As String, had As String
    If PD_Busy Then Exit Sub
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Sub
    If StrComp(ActiveSheet.Name, SH_SOURCES, vbTextCompare) <> 0 Then modPD_Theme.GoTo_ SH_SOURCES: Exit Sub
    r = ActiveCell.Row
    key = KeyForRow(r)
    If Len(key) = 0 Then Exit Sub
    had = SafeText(ws.Cells(r, S_FILE).Value2)
    If Len(had) = 0 Then Exit Sub
    If MsgBox("Stop using " & FileLeaf(had) & "?" & vbCrLf & vbCrLf & "The file itself is not touched.", _
              vbQuestion + vbYesNo, TOOL_NAME) <> vbYes Then Exit Sub
    ws.Range(ws.Cells(r, S_FILE), ws.Cells(r, S_COLS)).ClearContents
    RefreshStatuses
    LogIt V_OK, "Load", "Cleared the " & Replace(key, "|", " ") & " slot.", FileLeaf(had)
End Sub

Public Function ScanFolder(ByVal fld As String, ByVal depth As Long) As Long
    Dim fso As Object, f As Object, sub_ As Object, n As Long, ext As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fld) Then Exit Function
    For Each f In fso.GetFolder(fld).Files
        ext = LCase$(fso.GetExtensionName(f.Name))
        If (ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb") And Left$(f.Name, 1) <> "~" Then
            Step_ "reading " & f.Name
            If AssignFile(CStr(f.Path)) Then n = n + 1
        End If
    Next f
    If depth < 2 Then
        For Each sub_ In fso.GetFolder(fld).SubFolders
            n = n + ScanFolder(CStr(sub_.Path), depth + 1)
        Next sub_
    End If
    ScanFolder = n
End Function

Public Function AssignFile(ByVal Path As String, Optional ByVal forceKey As String = "") As Boolean
    Dim wb As Workbook, ws As Worksheet, hdr As Long, h As Object
    Dim kind As String, fw As String, key As String, why As String, reached As Boolean

    If Not IsLocalPath(Path) Then
        LogIt V_CHECK, "Load", "That path is a web address, not a file on this machine. " & _
              "Open the folder in File Explorer and use the local copy.", FileLeaf(Path)
        Exit Function
    End If

    On Error GoTo Done
    Set wb = Workbooks.Open(Path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    On Error GoTo CloseAndDone

    Set ws = PickSheet(wb, hdr, h, kind)
    If ws Is Nothing Then
        why = "no sheet in it carries the columns of an ALM output or a control report"
        GoTo CloseAndDone
    End If

    If kind = "OUTPUT" Then
        fw = FrameworkFromOutput(ws, hdr, h)
        If Len(fw) = 0 And Len(forceKey) = 0 Then
            why = "it is an output but nothing in it says which framework"
            GoTo CloseAndDone
        End If
        why = "an ALM output" & IIf(Len(fw) > 0, " for " & FwLabel(fw) & ", from its own framework column.", ".")
    Else
        why = "a control report - it carries " & IIf(kind = "CTRL3", F_CTRL3_AMT, F_CTRL6_AMT) & "."
    End If

    key = kind & "|" & fw
    If Len(forceKey) > 0 Then
        If StrComp(key, forceKey, vbTextCompare) <> 0 Then
            why = "YOU CHOSE THIS FILE FOR THIS ROW. Its contents read as " & _
                  Replace(key, "|", " ") & ", which is not what this row is for. Check this is what you meant."
        Else
            why = "You chose this file for this row, and its contents agree: " & why
        End If
        key = forceKey
    End If
    reached = True
    AssignFile = RecordSlot(key, Path, ws.Name, hdr, why, Len(forceKey) > 0)

CloseAndDone:
    why = IIf(Err.Number <> 0, Err.Number & " " & Err.Description, why)
    On Error Resume Next
    If Not AssignFile And Not reached Then LogIt V_CHECK, "Load", "Not placed - " & why, FileLeaf(Path)
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Err.Clear
Done:
End Function

' The widest real header row in the first eight rows of any visible sheet, and
' what that header makes the sheet.
Private Function PickSheet(ByVal wb As Workbook, ByRef hdrOut As Long, ByRef hOut As Object, _
                           ByRef kindOut As String) As Worksheet
    Dim ws As Worksheet, r As Long, c As Long, h As Object, v As String, kind As String
    Dim block As Variant
    For Each ws In wb.Worksheets
        If ws.visible = xlSheetVisible Then
            ' The candidate header rows in ONE read. Eight rows of a hundred and
            ' twenty cells read one at a time is nine hundred round trips per
            ' file, before anything useful has happened.
            On Error Resume Next
            block = ws.Range(ws.Cells(1, 1), ws.Cells(8, 120)).Value2
            On Error GoTo 0
            If Not IsEmpty(block) Then
                For r = 1 To 8
                    Set h = NewMap()
                    For c = 1 To 120
                        v = SafeText(block(r, c))
                        If Len(v) > 0 Then If Not h.Exists(NormKey(v)) Then h(NormKey(v)) = c
                    Next c
                    kind = KindOf(h)
                    If Len(kind) > 0 Then
                        hdrOut = r: Set hOut = h: kindOut = kind
                        Set PickSheet = ws
                        Exit Function
                    End If
                Next r
            End If
        End If
    Next ws
End Function

' Each kind named by the columns this tool actually reads off it. A file missing
' one of them cannot be used for the job, so recognising it would only move the
' failure later.
Private Function KindOf(ByVal h As Object) As String
    If h.Exists(NormKey(F_RULE_NAME)) And h.Exists(NormKey(F_PRE_LCY)) Then KindOf = "OUTPUT": Exit Function
    If h.Exists(NormKey(F_CTRL6_AMT)) And h.Exists(NormKey(F_ACCOUNT)) Then KindOf = "CTRL6": Exit Function
    If h.Exists(NormKey(F_CTRL3_AMT)) And h.Exists(NormKey(F_COA_CODE)) Then KindOf = "CTRL3": Exit Function
End Function

Private Function FrameworkFromOutput(ByVal ws As Worksheet, ByVal hdr As Long, ByVal h As Object) As String
    Dim col As Long, r As Long, v As String
    If Not h.Exists(NormKey(F_FRAMEWORK)) Then Exit Function
    col = CLng(h(NormKey(F_FRAMEWORK)))
    For r = hdr + 1 To hdr + 200
        v = SafeUpper(ws.Cells(r, col).Value2)
        If Len(v) > 0 Then
            If InStr(v, "MATURITY") > 0 Then FrameworkFromOutput = FW_ML: Exit Function
            If InStr(v, "NSFR") > 0 Then FrameworkFromOutput = FW_NSFR: Exit Function
            If InStr(v, "LCR") > 0 Then FrameworkFromOutput = FW_LCR: Exit Function
        End If
    Next r
End Function

Public Function RecordSlot(ByVal key As String, ByVal Path As String, ByVal sheetName As String, _
                           ByVal hdr As Long, ByVal why As String, _
                           Optional ByVal replace_ As Boolean = False) As Boolean
    Dim ws As Worksheet, r As Long, had As String
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Function
    r = SlotRow(key)
    If r = 0 Then
        LogIt V_CHECK, "Load", "Recognised as " & Replace(key, "|", " ") & " but there is no row for it.", FileLeaf(Path)
        Exit Function
    End If
    had = SafeText(ws.Cells(r, S_FILE).Value2)
    If replace_ And Len(had) > 0 And StrComp(had, Path, vbTextCompare) <> 0 Then
        LogIt V_OK, "Load", "Replaced " & FileLeaf(had) & " because you chose this one.", FileLeaf(Path)
        had = ""
    End If
    ' Two files for one slot is a mistake, not a merge. The first wins and the
    ' second is reported, because quietly replacing it hides which one was used.
    If Len(had) > 0 And StrComp(had, Path, vbTextCompare) <> 0 Then
        LogIt V_CHECK, "Load", "Two files both look like " & Replace(key, "|", " ") & ". Keeping " & _
              FileLeaf(had) & "; this one was NOT used.", FileLeaf(Path)
        Refused.Add Array(Replace(key, "|", " "), FileLeaf(had), FileLeaf(Path))
        Exit Function
    End If
    ws.Cells(r, S_FILE).Value2 = Path
    ws.Cells(r, S_SHEET).Value2 = sheetName
    ws.Cells(r, S_HDRROW).Value2 = hdr
    ws.Cells(r, S_NOTE).Value2 = why
    LogIt V_OK, "Load", "Placed as " & Replace(key, "|", " ") & ". " & why, FileLeaf(Path)
    RecordSlot = True
End Function

Public Function SheetOfSlot(ByVal wb As Workbook, ByVal key As String, ByRef hdr As Long) As Worksheet
    Dim ws As Worksheet, src As Worksheet, r As Long, nm As String
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Function
    r = SlotRow(key)
    If r = 0 Then Exit Function
    nm = SafeText(ws.Cells(r, S_SHEET).Value2)
    hdr = CLng(SafeNum(ws.Cells(r, S_HDRROW).Value2))
    If hdr < 1 Then hdr = 1
    On Error Resume Next
    Set src = wb.Worksheets(nm)
    Err.Clear
    Set SheetOfSlot = src
End Function

Public Sub NoteAmountField(ByVal key As String, ByVal Txt As String)
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Sub
    r = SlotRow(key)
    If r = 0 Then Exit Sub
    ws.Cells(r, S_AMOUNT).Value2 = Txt
End Sub

Public Sub NoteRows(ByVal key As String, ByVal n As Double, ByVal asOf As String)
    Dim ws As Worksheet, r As Long
    Set ws = GetSheet(SH_SOURCES)
    If ws Is Nothing Then Exit Sub
    r = SlotRow(key)
    If r = 0 Then Exit Sub
    ws.Cells(r, S_ROWS).Value2 = n
    ws.Cells(r, S_ROWS).NumberFormat = NUM_FMT
    If Len(asOf) > 0 Then ws.Cells(r, S_ASOF).Value2 = asOf
End Sub
