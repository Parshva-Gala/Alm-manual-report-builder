Option Explicit

' The workbook's own events, and nothing else: each hands straight to the
' module that owns the behaviour, so what happens on open, on switching
' windows and on saving can be read in one place per concern.

Private Sub Workbook_Open()
    Dim wasSaved As Boolean
    On Error Resume Next
    wasSaved = ThisWorkbook.Saved
    If ThisWorkbook.Worksheets.count <= 1 Or GetSheet(SH_HOME) Is Nothing Then
        modPD_Run.PD_Setup
    Else
        modPD_Run.Restyle
        modPD_Files.RefreshStatuses
    End If
    modPD_Desk.Opened
    ' Repainting what was already true is not a change: closing straight
    ' away should not ask to save.
    If wasSaved Then ThisWorkbook.Saved = True
    Err.Clear
End Sub

' Excel's ribbon and formula bar belong to the person, not to this workbook.
' Whenever it stops being the one in front, they go back.
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    modPD_Desk.AppEnter
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    modPD_Desk.AppLeave False
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    modPD_Desk.FitIfDesk
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    modPD_Desk.SheetChrome Sh
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    modPD_Desk.AppLeave True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    modPD_Desk.BeforeSave
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    modPD_Desk.AfterSave
End Sub
