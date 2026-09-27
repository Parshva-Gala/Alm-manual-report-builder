Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private mBoundName As String

Private Sub Workbook_Open()
    mBoundName = Me.name
    InitializeJKBTool
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    If JKB_Busy Then Exit Sub
    If TypeOf sh Is Worksheet Then
        On Error GoTo Failed
        ' Dashboard draws its own toolbar in its Worksheet_Activate.
        If Not sh Is Sheet1 Then Sheet1.EnsureRibbon sh
        If sh.name = SHEET_SELECTION Then RebindSelectionActions
        KeepBackstageHidden sh
        ' The four data sheets are dressed after any rebuild (late bound: this
        ' module must not name the master-only modules directly).
        On Error Resume Next
        Application.Run "'" & Replace(Me.name, "'", "''") & "'!DressWorkSheetByName", sh.name
        Err.Clear
    End If
    Exit Sub
Failed:
    MsgBox "Could not refresh sheet buttons." & vbCrLf & vbCrLf & "Reason: " & Err.description, vbExclamation, "JKB Stress Testing"
End Sub
Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal target As Range, Cancel As Boolean)
    If JKB_Busy Or sh.name <> SHEET_SELECTION Then Exit Sub
    If target.CountLarge <> 1 Or target.row < 7 Then Exit Sub
    Select Case target.Column
        Case 1, 4, 7, 10
            If SafeUpperText(target.Value2) = "YES" Or SafeUpperText(target.Value2) = "NO" Then
                Cancel = True
                target.Value2 = IIf(SafeUpperText(target.Value2) = "YES", "No", "Yes")
            End If
    End Select
End Sub
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim oldEvents As Boolean
    If Not Success Or JKB_Busy Then Exit Sub
    If mBoundName = Me.name Then Exit Sub
    On Error GoTo Done
    oldEvents = Application.EnableEvents
    mBoundName = Me.name
    Sheet1.RebindAll
    RebindSelectionActions
    EnsureReconWorkbench
    ' Rebinding only rewrites button targets, which Workbook_Open does again on every
    ' open, so there is nothing worth a second save.
    Me.Saved = True
Done:
    Application.EnableEvents = oldEvents
End Sub

