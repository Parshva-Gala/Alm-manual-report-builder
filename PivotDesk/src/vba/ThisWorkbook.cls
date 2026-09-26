Option Explicit

' The sheets are built on open rather than shipped in the seed, so the workbook
' in source control is one empty sheet and the whole layout lives in code where
' it can be read and changed.
Private Sub Workbook_Open()
    On Error Resume Next
    If ThisWorkbook.Worksheets.count <= 1 Or GetSheet(SH_HOME) Is Nothing Then
        modPD_Run.PD_Setup
    Else
        modPD_Theme.RailEverywhere
        modPD_Files.RefreshStatuses
        modPD_Theme.GoTo_ SH_HOME
    End If
    Err.Clear
End Sub
