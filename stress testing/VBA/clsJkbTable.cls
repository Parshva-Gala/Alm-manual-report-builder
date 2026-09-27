VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJkbTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
'  One loaded input table, held for the pre-shock engine.
'
'  The engine works on a single module variable, its working table. A switch between
'  tables used to COPY the stored table into it, so the table in use was held
'  twice. ExchangeAt instead swaps the table between this object and that
'  variable: the VARIANT headers trade places and the rows never move. The
'  values are the same values - nothing is converted or re-read.
' ============================================================================

#If VBA7 Then
Private Declare PtrSafe Sub JkbMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As LongPtr, ByVal src As LongPtr, ByVal bytes As LongPtr)
#Else
Private Declare Sub JkbMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As Long, ByVal src As Long, ByVal bytes As Long)
#End If

#If Win64 Then
Private Const VARIANT_BYTES As Long = 24
#Else
Private Const VARIANT_BYTES As Long = 16
#End If

Private mData As Variant

' Swaps this table with the Variant at address p (VarPtr of the caller's variable).
Public Sub ExchangeAt(ByVal p As Variant)
    Dim tmp(0 To 23) As Byte
    If p = 0 Then Err.Raise 5, "clsJkbTable.ExchangeAt", "No variable to exchange with."
    JkbMoveMemory VarPtr(tmp(0)), p, VARIANT_BYTES
    JkbMoveMemory p, VarPtr(mData), VARIANT_BYTES
    JkbMoveMemory VarPtr(mData), VarPtr(tmp(0)), VARIANT_BYTES
End Sub

Public Sub SetCopy(ByRef v As Variant)
    mData = v
End Sub

Public Function CopyOut() As Variant
    CopyOut = mData
End Function

Public Function Cell(ByVal r As Long, ByVal c As Long) As Variant
    Cell = mData(r, c)
End Function

Public Function Holding() As Boolean
    Holding = IsArray(mData)
End Function
