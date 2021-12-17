VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum DetailFormState
    NotYetInitialized
    NoEntries
    NewEntry
    ViewSingle
    ViewXY
    EditXY
    EditXYSaving
End Enum

Private Type TDetailForm
    frm As Form
    lhs As New Collection
    rhs As New Collection
    State As DetailFormState
    entityID As Integer
    HasExistingRecords As Boolean
    TableName As String
End Type
Private This As TDetailForm

' Properties
Public Property Set Form(frm As Form)
    Debug.Assert Not frm Is Nothing
    Dim ctl As control
    Set This.frm = frm
    For Each ctl In frm.Controls
        If ctl.name Like "???LHS*" Then
            This.lhs.Add ctl
        ElseIf ctl.name Like "???RHS*" Then
            This.rhs.Add ctl
        End If
    Next
    This.TableName = Replace(frm.name, "sfrm", "tbl")
    
End Property

Public Property Get Form() As Form
    Debug.Assert Not This.frm Is Nothing
    Set Form = This.frm
End Property

Public Property Let State(vNewValue As DetailFormState)
    This.State = vNewValue
End Property

Public Property Get State() As DetailFormState
    State = This.State
End Property

Public Property Let Entity(vNewValue As Integer)
     This.entityID = vNewValue
     This.HasExistingRecords = CheckForExistingRecords()
End Property

Public Property Get Entity() As Integer
    Entity = This.entityID
End Property

' Ctor
Private Sub Class_Initialize()
    This.State = NotYetInitialized
End Sub

' Public Sub
Public Function ToString() As String
    ToString = "State is " & CStr(This.State)
End Function

Public Function TryChangeState(newState As DetailFormState) As Boolean
    TryChangeState = False
End Function

Public Sub ShowRHS()
    Dim ctl As control
    For Each ctl In This.rhs
        ctl.Visible = True
    Next ctl
End Sub

Public Sub HideRHS()
    Dim ctl As control
    For Each ctl In This.rhs
        ctl.Visible = False
    Next ctl
End Sub

Public Sub LockLHS()
    Dim ctl As control
    For Each ctl In This.lhs
        ctl.Locked = True
    Next ctl
End Sub

Public Sub UnlockLHS()
    Dim ctl As control
    For Each ctl In This.lhs
        ctl.Locked = False
    Next ctl
End Sub

Public Sub CopyRHStoLHS()
    Dim lhsCtl As control, rhsCtl As control
    
    For Each rhsCtl In This.rhs
        Set lhsCtl = GetControlByName(This.lhs, Replace(rhsCtl.name, "RHS", "LHS"))
        lhsCtl.Value = rhsCtl.Value
    Next rhsCtl
End Sub

Public Sub CopyLHStoRHS()
    Dim lhsCtl As control, rhsCtl As control
    
    For Each lhsCtl In This.lhs
        Set rhsCtl = GetControlByName(This.rhs, Replace(lhsCtl.name, "LHS", "RHS"))
        rhsCtl.Value = lhsCtl.Value
        lhsCtl.Locked = True
        'rhsCtl.Properties.Item("BackColor") = RAGColors.yellow
    Next lhsCtl
End Sub

Public Sub ResetLHSandRHS()
    Dim ctl As control
    For Each ctl In This.rhs
        ctl.Value = ""
        'Ctl.Properties.Item("BackColor") = RAGColors.yellow
    Next
    UnlockLHS
End Sub

Public Function ShowFirstValid(idx As Integer)
    Call This.frm.Recordset.FindFirst("EntityFK = " & idx & " AND ValidUntil = #9999/12/31#")
End Function

' Private Methods
Private Function GetControlByName(coll As Collection, name As String)
    Debug.Assert Not coll Is Nothing
    Dim ctl As control
    For Each ctl In coll
        If ctl.name = name Then
            Set GetControlByName = ctl
            Exit Function
        End If
    Next ctl
    Err.Raise 5, , "Cannot find GetControlByName '" & name & "'"
End Function

Private Function CheckForExistingRecords() As Boolean
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT Count(*) AS TotalCount FROM " & This.TableName & " WHERE EntityFK = " & This.entityID & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    CheckForExistingRecords = rs!TotalCount > 0
End Function
