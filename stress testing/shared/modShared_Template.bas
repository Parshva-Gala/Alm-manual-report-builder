Attribute VB_Name = "modShared_Template"
Option Explicit

' ============================================================================
'  The manual-report template, carried inside the workbook.
'
'  A generated report workbook is born from a template that carries its button
'  strip and the Workbook_Open that freezes its control band. That template used
'  to be a second file you had to remember to send, and "remember to send the
'  other file" is not a thing a tool should ask of anyone.
'
'  So it travels INSIDE the workbook, base64 on a very hidden sheet, and is
'  written out to the temporary folder the first time a report is built. Exactly
'  how the console, the condition builder and the progress window already travel:
'  one idea, used a fourth time, rather than a fourth mechanism.
'
'  A copy on disk beside the workbook still wins, so a newer template can be
'  dropped in without rebuilding anything.
' ============================================================================

Public Const TEMPLATE_SHEET As String = "_Template_Src"
Private Const TEMPLATE_LEAF As String = "Manual_Report_Template.xltm"

' Extracting it costs a moment, so a session does it once.
Private mPath As String

' The template to build a report workbook from, or "" if there is none.
Public Function TemplatePath() As String
    Dim p As String
    p = TemplateOnDisk()
    If Len(p) > 0 Then TemplatePath = p: Exit Function
    If Len(mPath) > 0 Then
        If Len(Dir$(mPath)) > 0 Then TemplatePath = mPath: Exit Function
    End If
    mPath = ExtractTemplate()
    TemplatePath = mPath
End Function

' Beside the workbook, one level up, or in an ALM_V10 folder below - which covers
' the shipping layout, a build running out of a seed folder, and a user who keeps
' the tool in a subfolder of the package.
Private Function TemplateOnDisk() As String
    Dim cands As Variant, c As Variant, p As String, here As String, up As String, i As Long
    here = ThisWorkbook.path
    If Len(here) = 0 Then Exit Function
    ' A workbook opened from a synced OneDrive or SharePoint folder reports its
    ' path as an https:// URL. Dir$ does not fail cleanly on one - it matches
    ' something and the path is then handed to Workbooks.Add, which cannot open
    ' it and raises 1004 in the middle of generating a report. The copy carried
    ' inside the workbook exists for exactly this, so use it.
    If Not IsLocalPath(here) Then Exit Function
    i = InStrRev(here, "\")
    If i > 1 Then up = Left$(here, i - 1) Else up = here
    cands = Array(here & "\" & TEMPLATE_LEAF, _
                  here & "\ALM_V10\" & TEMPLATE_LEAF, _
                  up & "\" & TEMPLATE_LEAF, _
                  up & "\ALM_V10\" & TEMPLATE_LEAF)
    For Each c In cands
        p = CStr(c)
        On Error Resume Next
        If Len(Dir$(p)) > 0 Then TemplateOnDisk = p
        Err.Clear
        On Error GoTo 0
        If Len(TemplateOnDisk) > 0 Then Exit Function
    Next c
End Function

' A real path on this machine: a drive letter, or a UNC share. Anything else -
' an http URL above all - is somewhere Dir$ and Workbooks.Add disagree about.
Private Function IsLocalPath(ByVal p As String) As Boolean
    If Len(p) < 3 Then Exit Function
    If Left$(p, 2) = "\\" Then IsLocalPath = True: Exit Function
    If Mid$(p, 2, 2) = ":\" Then IsLocalPath = True
End Function

' ===================== the embedded copy ====================================

Private Function ExtractTemplate() As String
    Dim ws As Worksheet, lastR As Long, r As Long, b64 As String, bytes() As Byte
    Dim fso As Object, folder As String, path As String, ff As Integer

    On Error GoTo Failed
    Set ws = SheetSafe(TEMPLATE_SHEET)
    If ws Is Nothing Then Exit Function
    lastR = ws.Cells(ws.rows.count, 1).End(xlUp).row
    If lastR < 1 Then Exit Function

    ' A cell holds 32,767 characters, so the encoding is split across rows and
    ' joined back here. Concatenating in a loop is fine at this size; the whole
    ' thing is a few dozen kilobytes.
    For r = 1 To lastR
        b64 = b64 & CStr(ws.Cells(r, 1).Value2)
    Next r
    If Len(b64) = 0 Then Exit Function

    bytes = DecodeBase64(b64)

    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetSpecialFolder(2)
    CleanOldTemplates fso, folder
    path = fso.BuildPath(folder, "JKB_ReportTemplate_" & format$(Now, "yyyymmddhhnnss") & ".xltm")

    ff = FreeFile
    Open path For Binary Access Write As #ff
    Put #ff, 1, bytes
    Close #ff

    If Len(Dir$(path)) = 0 Then Exit Function
    ExtractTemplate = path
    Exit Function
Failed:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    Err.Clear
End Function

' MSXML's bin.base64 element type, which is the decoder every Windows machine
' already has. No ADODB.Stream: it is the component locked-down estates disable
' first, and this has to work on a reviewer's machine, not just on a developer's.
Private Function DecodeBase64(ByVal s As String) As Byte()
    Dim dom As Object, el As Object
    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("b64")
    el.DataType = "bin.base64"
    el.text = s
    DecodeBase64 = el.nodeTypedValue
End Function

Private Sub CleanOldTemplates(ByVal fso As Object, ByVal folder As String)
    Dim f As Object
    On Error Resume Next
    For Each f In fso.GetFolder(folder).files
        If Left$(f.name, 18) = "JKB_ReportTemplate" Then
            If DateDiff("h", f.DateLastModified, Now) > 24 Then f.Delete True
        End If
    Next f
    Err.Clear
End Sub

Private Function SheetSafe(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetSafe = ThisWorkbook.Worksheets(nm)
    Err.Clear
    On Error GoTo 0
End Function

' Says where the template came from, for the log and for the Index of a report.
Public Function TemplateOrigin() As String
    If Len(TemplateOnDisk()) > 0 Then
        TemplateOrigin = "the copy beside the workbook"
    ElseIf Len(mPath) > 0 Then
        TemplateOrigin = "the copy carried inside the workbook"
    Else
        TemplateOrigin = "(none found)"
    End If
End Function
