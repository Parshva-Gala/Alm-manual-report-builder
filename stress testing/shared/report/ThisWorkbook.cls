VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' The workbook freezes its own control band the first time a person opens it.
'
' It cannot be done when the file is written: the generating Excel is invisible
' and has no window, and FreezePanes needs one. Doing it here means the band is
' pinned however the workbook was made, so a slicer or a button is never more
' than where it already is - no scrolling up to reach one and back down to see
' what it did.
Private Sub Workbook_Open()
    On Error Resume Next
    modReport.RptFreezeAll
    Err.Clear
End Sub
