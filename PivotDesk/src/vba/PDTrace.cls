Option Explicit

' ============================================================================
'  Where the run got to, written to disk as it goes.
'
'  Opened, written and closed per line. VBA buffers Print # until Close, so a
'  log held open reflects a buffer boundary rather than progress - which once
'  made a run that died during the load look like one that died after it, and
'  sent the search for the fault to the wrong module.
'
'  The desk never creates one; PD_Trace stays Nothing and Step_ only moves the
'  status bar.
' ============================================================================

Public Path As String

Public Sub Note(ByVal what As String)
    Dim fh As Integer
    On Error Resume Next
    If Len(Path) = 0 Then Exit Sub
    fh = FreeFile
    Open Path For Append As #fh
    Print #fh, Format$(Now, "hh:nn:ss") & "     . " & what
    Close #fh
    Err.Clear
End Sub
