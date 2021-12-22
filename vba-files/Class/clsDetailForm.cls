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

Private WithEvents frmEvents As Form
Attribute frmEvents.VB_VarHelpID = -1

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
    LHS As New Collection
    RHS As New Collection
    State As DetailFormState
    EntityID As Integer
    HasExistingRecords As Boolean
    tableName As String
    IsEditing As Boolean
End Type
Private This As TDetailForm

' Properties
Public Property Set Form(frm As Form)
    Debug.Assert Not frm Is Nothing
    Dim ctl As control
    Set This.frm = frm
    For Each ctl In frm.controls
        If ctl.name Like "???LHS*" Then
            This.LHS.Add ctl
        ElseIf ctl.name Like "???RHS*" Then
            This.RHS.Add ctl
        End If
    Next
    This.tableName = Replace(frm.name, "sfrm", "tbl")
    Set frmEvents = frm
    frm.AfterInsert = "[Event Procedure]"
    frm.AfterUpdate = "[Event Procedure]"
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
     This.EntityID = vNewValue
     This.HasExistingRecords = CheckForExistingRecords()
End Property

Public Property Get Entity() As Integer
    Entity = This.EntityID
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
    Select Case newState
        Case NewEntry
            TryChangeState = TryChangeStateNewEntry
    End Select
    
End Function

Public Sub ShowRHS()
    Dim ctl As control
    For Each ctl In This.RHS
        ctl.Visible = True
    Next ctl
End Sub

Public Sub HideRHS()
    Dim ctl As control
    For Each ctl In This.RHS
        ctl.Visible = False
    Next ctl
End Sub

Public Sub LockLHS()
    Dim ctl As control
    For Each ctl In This.LHS
        ctl.Locked = True
    Next ctl
End Sub

Public Sub UnlockLHS()
    Dim ctl As control
    For Each ctl In This.LHS
        ctl.Locked = False
    Next ctl
End Sub

Public Sub CopyRHStoLHS()
    Dim lhsCtl As control, rhsCtl As control
    
    For Each rhsCtl In This.RHS
        Set lhsCtl = GetControlByName(This.LHS, Replace(rhsCtl.name, "RHS", "LHS"))
        lhsCtl.Value = rhsCtl.Value
    Next rhsCtl
End Sub

Public Sub CopyLHStoRHS()
    Dim lhsCtl As control, rhsCtl As control
    
    For Each lhsCtl In This.LHS
        Set rhsCtl = GetControlByName(This.RHS, Replace(lhsCtl.name, "LHS", "RHS"))
        rhsCtl.Value = lhsCtl.Value
        lhsCtl.Locked = True
        'rhsCtl.Properties.Item("BackColor") = RAGColors.yellow
    Next lhsCtl
End Sub

Public Sub ResetLHSandRHS()
    Dim ctl As control
    For Each ctl In This.RHS
        ctl.Value = ""
        'Ctl.Properties.Item("BackColor") = RAGColors.yellow
    Next
    UnlockLHS
End Sub

Public Function ShowFirstValid(idx As Integer)
    Call This.frm.Recordset.FindFirst("EntityFK = " & idx & " AND ValidUntil = #9999/12/31#")
End Function

Public Sub notifychanged()
    MsgBox "thank you"
End Sub

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
    
    sql = "SELECT Count(*) AS TotalCount FROM " & This.tableName & " WHERE EntityFK = " & This.EntityID & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    'Debug.Print "Check for existing records"
    'Debug.Print "   " & sql
    'Debug.Print "    " & CStr(rs!TotalCount)
    
    CheckForExistingRecords = rs!TotalCount > 0
End Function

' State Changes

Private Function TryChangeStateNewEntry() As Boolean
    If This.IsEditing Then Exit Function
    If This.HasExistingRecords Then Exit Function
    If This.State = NewEntry Then Exit Function
    If This.EntityID = 0 Then Exit Function
    
    This.frm.dataentry = True

    'This.frm.Recordset.AddNew
    This.frm!EntityFK = This.EntityID
    This.frm!ValidFrom = Now()
    This.frm!CommitFK = 1
    
    'With This.frm
        '.DataEntry = True
        '.Recordset.AddNew
        '.Controls("EntityFK") = This.entityID
        '.Controls("ValidFrom") = Now
        '.Controls("CommitFK") = 1
    'End With
    
    ResetLHSandRHS
    HideRHS
    UnlockLHS
End Function

Private Sub frmEvents_AfterInsert()
    'MsgBox "AfterInsert"
End Sub

Private Sub frmEvents_AfterUpdate()
    'MsgBox "AfterUpdate"
End Sub

