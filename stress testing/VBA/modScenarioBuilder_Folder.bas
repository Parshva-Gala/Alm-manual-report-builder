Attribute VB_Name = "modScenarioBuilder_Folder"
Option Explicit

' ============================================================================
'  Loading a whole folder of inputs.
'
'  Today there are two extracts. There will be more - max ECL, the capital
'  component report, an ALM file per liquidity framework - and asking someone to
'  find the right picker for each of them does not scale.
'
'  So: point at a folder. Each workbook is identified by WHAT IS IN IT rather
'  than by its name, because a file called "ECL output" is not always one, and a
'  file that is one is not always called that.
'
'  It takes minutes on real extracts, so it says what it is doing and roughly how
'  much longer, in a window that keeps drawing while Excel is busy.
' ============================================================================

Private Const MAX_DEPTH As Long = 2
' What the last folder load did, for a harness or a diagnostic run that has no
' message box to read it from.
Private mLastFolderLoad As String

Public Sub UploadInputFolder()
    Dim folder As String, ok As Boolean
    If JKB_Busy Then Exit Sub
    If Not ShowFolderGuide() Then Exit Sub
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Choose the folder holding the stress-testing inputs"
        If .Show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With
    UploadFolderHeadless folder
End Sub

' The same work with no picker and no guide, for a harness. Every file it leaves
' out goes to Build_Log with the reason, and the whole result is kept for
' LastFolderLoadSummary, because a headless run shows no message box.
Public Sub UploadFolderHeadless(ByVal folder As String)
    Dim files As Collection, i As Long, state As Object, taken As Object, er As String
    Dim loaded As Long, skipped As Long, notes As String, f As Variant
    Dim kind As String, t0 As Single, el As Single, what As String

    On Error GoTo Failed
    mLastFolderLoad = ""
    Set files = WorkbooksUnder(folder, 0)
    If files.count = 0 Then
        mLastFolderLoad = "Nothing loaded: there is no Excel workbook in " & folder & "."
        LogIssue LOG_LEVEL_WARN, "Upload", mLastFolderLoad, folder
        If Application.Visible Then UiNotice "Load a folder", "There is no Excel workbook in that folder.", folder
        Exit Sub
    End If

    Set state = CaptureState(): JKB_Busy = True
    HardenExcelForBulkRead
    t0 = Timer
    Set taken = NewMap()

    ' Work is measured in megabytes, not files: a 150 MB extract and a 300 KB one
    ' are not the same amount of work, and an estimate that pretends they are is
    ' worse than none.
    ProgressStart "Loading stress-testing inputs", _
                  "Reading " & files.count & " workbook(s) from " & FolderLeaf(folder), _
                  TotalMegabytes(files)

    For i = 1 To files.count
        Set f = files(i)
        ProgressStep "Reading " & f("name") & "   (" & i & " of " & files.count & ")", f("mb")
        kind = IdentifyAndLoad(CStr(f("path")), notes, taken)
        If Len(kind) > 0 Then
            loaded = loaded + 1
        Else
            skipped = skipped + 1
        End If
    Next i

    ProgressStep "Caching values for the condition builder", 0
    RefreshPreShockAvailability
    ' Timer resets at midnight, and a long load started at 23:58 would otherwise
    ' report a negative duration.
    el = Timer - t0
    If el < 0 Then el = el + 86400
    ProgressDone loaded & " file(s) loaded in " & format$(el, "0") & " seconds."

    what = loaded & " of " & files.count & " workbook(s) loaded"
    If taken.count > 0 Then what = what & ", as " & Join(taken.keys, ", ")
    what = what & "."
    If loaded > taken.count Then what = what & " " & (loaded - taken.count) & " replaced an earlier file for the same input."
    If skipped > 0 Then what = what & " " & skipped & " left out."
    mLastFolderLoad = what & IIf(Len(notes) > 0, vbCrLf & notes, "")
    LogIssue IIf(skipped > 0, LOG_LEVEL_WARN, LOG_LEVEL_INFO), "Upload", "Load a folder: " & what, folder

    StylePreShockWorkspace
    RestoreState state: JKB_Busy = False
    GoPsSheet PS_SOURCES_SHEET

    If Not Application.Visible Then Exit Sub
    UiNotice "Load a folder", what & IIf(Len(notes) > 0, " The details are below.", ""), notes, _
             "Click Run on the home screen.", (loaded = 0)
    Exit Sub

Failed:
    ' Kept before anything else runs: RestoreState clears the error.
    er = Err.description
    mLastFolderLoad = "Stopped: " & er
    ProgressFailed "Stopped: " & er
    RestoreState state: JKB_Busy = False
    LogIssue LOG_LEVEL_ERROR, "Upload", er, folder
    If Application.Visible Then UiProblem "Load a folder", "The folder could not be loaded.", er
End Sub

' What the last folder load did: how many files loaded, as which inputs, and
' every file left out with the reason. Empty before the first folder load.
Public Function LastFolderLoadSummary() As String
    LastFolderLoadSummary = mLastFolderLoad
End Function

' ===================== the guide ============================================

Private Function ShowFolderGuide() As Boolean
    Dim s As String, a As VbMsgBoxResult
    s = "LOADING A FOLDER OF INPUTS" & vbCrLf & vbCrLf
    s = s & "Point at a folder. Every Excel workbook in it, and in one level of" & vbCrLf
    s = s & "subfolders, is examined and loaded if it is one of the inputs." & vbCrLf & vbCrLf
    s = s & "Files are identified by WHAT IS IN THEM, not by their names, so" & vbCrLf
    s = s & "you do not have to rename or rearrange anything." & vbCrLf & vbCrLf
    s = s & "What it looks for (the same check as Load inputs):" & vbCrLf & vbCrLf
    s = s & "   ECL output" & vbTab & "a stage column with an outstanding" & vbCrLf
    s = s & "   " & vbTab & vbTab & "or ECL column" & vbCrLf
    s = s & "   CAPRWA output" & vbTab & "OUTSTANDING_FOR_RWA or RWA_LCY" & vbCrLf
    s = s & "   LL, LCR, NSFR" & vbTab & "ALM cashflow columns; which one is" & vbCrLf
    s = s & "   " & vbTab & vbTab & "read from ALM_FRAMEWORK_NAME" & vbCrLf
    s = s & "   Capital" & vbTab & vbTab & "CAPITAL_ELEMENT and REPORT_BALANCE" & vbCrLf & vbCrLf
    s = s & "A workbook with several sheets is fine - each is scored and the" & vbCrLf
    s = s & "best one is used. Anything it does not recognise is left alone and" & vbCrLf
    s = s & "reported with the reason, never treated as an error." & vbCrLf & vbCrLf
    s = s & "HOW LONG:  roughly 30 to 60 seconds per 100 MB. A progress window" & vbCrLf
    s = s & "shows what it is reading and about how much longer it has." & vbCrLf & vbCrLf
    s = s & "Continue?"
    a = MsgBox(s, vbQuestion + vbOKCancel, "Before you choose a folder")
    ShowFolderGuide = (a = vbOK)
End Function

' ===================== finding and identifying ==============================

Private Function WorkbooksUnder(ByVal folder As String, ByVal depth As Long) As Collection
    Dim fso As Object, f As Object, sub_ As Object, out As New Collection, ext As String, d As Object
    Dim inner As Collection, i As Long
    Set WorkbooksUnder = out
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function
    For Each f In fso.GetFolder(folder).files
        ext = LCase$(fso.GetExtensionName(f.name))
        If (ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb") And Left$(f.name, 1) <> "~" Then
            Set d = NewMap()
            d("path") = f.path
            d("name") = f.name
            d("mb") = WorksheetFunction.Max(f.Size / 1048576#, 0.1)
            out.Add d
        End If
    Next f
    If depth < MAX_DEPTH Then
        For Each sub_ In fso.GetFolder(folder).SubFolders
            Set inner = WorkbooksUnder(CStr(sub_.path), depth + 1)
            For i = 1 To inner.count
                out.Add inner(i)
            Next i
        Next sub_
    End If
End Function

Private Function TotalMegabytes(ByVal files As Collection) As Double
    Dim i As Long
    For i = 1 To files.count
        TotalMegabytes = TotalMegabytes + CDbl(files(i)("mb"))
    Next i
    If TotalMegabytes <= 0 Then TotalMegabytes = 1
End Function

' Opens the workbook once, decides what it is with the same check Load inputs
' uses (PickInputSheet), and routes it to the same loader. A file that is not
' loaded is always said, with the reason, in the notes and in Build_Log.
' taken holds the inputs already loaded in this folder load, so a second file
' for one of them is said too.
Private Function IdentifyAndLoad(ByVal path As String, ByRef notes As String, Optional ByVal taken As Object) As String
    Dim wb As Workbook, pick As Object, best As Worksheet, bestHr As Long, kind As String
    Dim oldSec As MsoAutomationSecurity, loadErr As Long, loadWhy As String

    On Error GoTo Done
    oldSec = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wb = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    On Error GoTo CloseAndDone

    Set pick = PickInputSheet(wb)
    kind = CStr(pick("Key"))
    If Len(kind) = 0 Then
        LeftOut notes, path, CStr(pick("Why"))
        GoTo CloseAndDone
    End If
    Set best = pick("Sheet"): bestHr = CLng(pick("HeaderRow"))
    If Len(CStr(pick("Warn"))) > 0 Then
        LogIssue LOG_LEVEL_WARN, "Upload", "Sheet '" & best.name & "' was read as the " & kind & " extract, but " & _
                 CStr(pick("Warn")) & ".", FileNameOnly2(path)
    End If

    ProgressDetail "reading " & best.name & " as the " & kind & " extract"
    LoadSheetForUpload best, bestHr, wb.date1904, path
    StoreActiveSourceFor kind
    RecordSourceRow kind, path, best.name
    BuildValueCacheFor kind
    IdentifyAndLoad = kind
    LogIssue LOG_LEVEL_INFO, "Upload", "Loaded as " & kind & " from sheet '" & best.name & "'.", FileNameOnly2(path)
    ' Two files that are the same input: the later one is used, as it always
    ' was, but that is now said.
    If Not taken Is Nothing Then
        If taken.Exists(kind) Then
            AddNote notes, FileNameOnly2(path) & " - loaded as " & kind & " in place of " & FileNameOnly2(CStr(taken(kind))) & _
                    ", which looked like the same input. Take the one you do not want out of the folder."
            LogIssue LOG_LEVEL_WARN, "Upload", "Loaded as " & kind & " in place of " & FileNameOnly2(CStr(taken(kind))) & _
                     ", which looked like the same input.", FileNameOnly2(path)
        End If
        taken(kind) = path
    End If

CloseAndDone:
    loadErr = Err.Number: loadWhy = Err.description
    On Error Resume Next
    If loadErr <> 0 Then
        ' Said, not skipped: an unreadable file is not an unrecognised one.
        AddNote notes, FileNameOnly2(path) & " - could not be read: " & loadWhy
        LogIssue LOG_LEVEL_ERROR, "Upload", "Could not read this file: " & loadWhy, FileNameOnly2(path)
    End If
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.AutomationSecurity = oldSec
    Err.Clear
    Exit Function
Done:
    loadWhy = Err.description
    On Error Resume Next
    Application.AutomationSecurity = oldSec
    AddNote notes, FileNameOnly2(path) & " - could not be opened: " & loadWhy
    LogIssue LOG_LEVEL_ERROR, "Upload", "Could not open this file: " & loadWhy, FileNameOnly2(path)
    Err.Clear
End Function

' One file, read exactly as a folder load reads it. For the harness, which loads
' file by file to show which one a problem is on.
Public Function LoadOneInputQuiet(ByVal path As String) As String
    Dim notes As String, kind As String
    kind = IdentifyAndLoad(path, notes)
    If Len(kind) > 0 Then LoadOneInputQuiet = "Loaded as " & kind Else LoadOneInputQuiet = "Not loaded: " & notes
End Function

' A file left out is said twice: in the notes the message box shows, and in
' Build_Log, which is the only place a headless run can say it.
Private Sub LeftOut(ByRef notes As String, ByVal path As String, ByVal why As String)
    AddNote notes, FileNameOnly2(path) & " - not loaded: " & why
    LogIssue LOG_LEVEL_WARN, "Upload", "Not loaded: " & why, FileNameOnly2(path)
End Sub

Private Sub AddNote(ByRef notes As String, ByVal s As String)
    If Len(notes) > 0 Then notes = notes & vbCrLf
    notes = notes & s
End Sub

Private Function FileNameOnly2(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then FileNameOnly2 = Mid$(p, i + 1) Else FileNameOnly2 = p
End Function

Private Function FolderLeaf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 And i < Len(p) Then FolderLeaf = Mid$(p, i + 1) Else FolderLeaf = p
End Function

' Excel opens a large workbook markedly faster when it is not trying to
' recalculate, repaint, follow links or ask questions along the way.
Private Sub HardenExcelForBulkRead()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    Err.Clear
End Sub
