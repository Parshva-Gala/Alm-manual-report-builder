Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    Dim dash As Worksheet, valid As Boolean, prefix As String
    On Error GoTo Failed
    Set dash = Me.Parent.Worksheets("Dashboard")
    prefix = "'" & Replace(Me.Parent.name, "'", "''") & "'!" & dash.CodeName & "."
    valid = CBool(Application.Run(prefix & "IsFormulaCell", target))
    If valid Then
        Cancel = True
        target.Select
        Application.Run prefix & "OpenFormulaEditor"
    End If
    Exit Sub
Failed:
    MsgBox "Could not open formula suggestions." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub

