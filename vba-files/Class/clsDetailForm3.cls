VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetailForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event StateChanged(newState As DetailState3)
Public Event EditingStarted()
Public Event EditingStopped()
Public Event EditingCancelled()

Private WithEvents frmForEvents As Access.Form
Attribute frmForEvents.VB_VarHelpID = -1

Public Enum DetailState3
    Invalid
    Ready
    NewMode
    EditMode
    ViewMode
End Enum

Private Type DetailForm3State
    frm As Form
    TableName As String
    SFrmName As String
    EntityID As Integer
    State As DetailState3
    IsEditing As Boolean
End Type
Private This As DetailForm3State


' Properties
Public Property Set Form(RHS As Form)
    Set This.frm = RHS
    This.SFrmName = RHS.name
    This.TableName = Replace(RHS.name, "sfrm", "tbl")
    Set frmForEvents = RHS
    'MsgBox "Ready"
    'Log "Property Set Form"
End Property

'Private Property Let TableName(RHS As String)
'    If This.TableName = RHS Then Exit Property
'    This.TableName = RHS
'    This.SFrmName = Replace(RHS, "tbl", "sfrm")
'    Log "Changed table name to '" & This.TableName & "'"
'End Property

Public Property Let Entity(RHS As Integer)
    If This.EntityID = RHS Then
        Exit Property
    End If
    This.EntityID = RHS
    TrySetState Ready
    'Log "Changed entity to '" & This.EntityID & "'"
End Property

Public Property Get TableName() As String
    TableName = This.TableName
End Property

Public Property Get SubFormName() As String
    SubFormName = This.SFrmName
End Property

Public Property Get Entity() As Integer
    Entity = This.EntityID
End Property

Public Property Get IsEditing() As Boolean
    IsEditing = This.IsEditing
End Property

' Ctor
Private Sub Class_Initialize()
    DoChangeState (Invalid)
End Sub

' Methods

' Helpers
Private Function GotoEntity()
    This.frm.Filter = "EntityFK = " & Entity & " AND ValidUntil = #9999/12/31#"
    This.frm.FilterOn = True
    Call This.frm.Recordset.FindLast("EntityFK = " & Entity & " AND ValidUntil = #9999/12/31#")
    
    'Log "GotoEntity #" & Entity
End Function

Private Function GetControlByName(controls As controls, name As String)
    Dim ctl As control
    For Each ctl In controls
        If ctl.name = name Then
            Set GetControlByName = ctl
            Exit Function
        End If
    Next ctl
    Err.Raise 5, , "Cannot find GetControlByName '" & name & "'"
End Function

Private Function CheckForExistingRecords() As Boolean
    If TableName = "" Then Exit Function
    If Entity = 0 Then Exit Function
    
    Dim results As Integer
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT Count(*) AS TotalCount FROM " & TableName & " WHERE EntityFK = " & Entity & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    results = rs!TotalCount
    
    'Log "zzCheckForExistingRecords = " & CStr(results)
    
    CheckForExistingRecords = results > 0
End Function

' Mode functions
Private Function ChangeToNewMode()
    With This.frm
        .AllowAdditions = True
        .AllowEdits = True
        .AllowDeletions = False
        .DataEntry = True
    End With
    
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = False
            ctl.Properties("BackColor") = 13431551
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
         ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394
        End If
    Next ctl
    
    ' Do not use - this will throw Runtime 2449
    'This.frm.SetFocus
    
    StartEditing
    
    'This.frm!EntityFK = Entity
    'This.frm!ValidFrom = Now()
    'This.frm!CommitFK = 1
    
    'Log "Now in New Mode"
End Function

Private Function ChangeToViewMode()
    With This.frm
        .AllowAdditions = False
        .AllowEdits = False
        .AllowDeletions = False
        .DataEntry = False
    End With
    
    GotoEntity
        
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = True
            ctl.Properties("BackColor") = 16777215
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394 '4394 or 6804
        End If
    Next ctl
    
    StopEditing
    'Log "Now in View Mode"
End Function

Private Function ChangeToEditMode()
    With This.frm
        .AllowAdditions = False
        .AllowEdits = True
        .AllowDeletions = False
        .DataEntry = False
    End With
    
    GotoEntity
        
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = True
            ctl.Properties("BackColor") = 16777215
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = True
            ctl.Properties("BackColor") = 13431551 '16777215
            ctl = GetControlByName(This.frm.controls, Replace(ctl.name, "RHS", "LHS"))
        ElseIf ctl.name Like "lblSuffix*" Then ctl.Properties("Left") = 6804
            '4394 or 6804
        End If
    Next ctl
    
    StartEditing
    
    'Log "Now in Edit Mode"
End Function

' More State stuff
Public Function TrySetState(newState As DetailState3) As Boolean
    If This.State = newState Then Exit Function
    
    Select Case newState
        Case Ready
            If CheckForExistingRecords Then
                TrySetState ViewMode
            Else
                TrySetState NewMode
            End If
            Exit Function
            
        Case NewMode
            If Not CheckForExistingRecords Then
                TrySetState = True
                ChangeToNewMode
                DoChangeState NewMode
                Exit Function
            End If
            
        Case ViewMode
            If CheckForExistingRecords Then
                TrySetState = True
                ChangeToViewMode
                DoChangeState ViewMode
                Exit Function
            End If
            
        Case EditMode
            If This.State = ViewMode Then
                TrySetState = True
                ChangeToEditMode
                DoChangeState EditMode
                Exit Function
            End If
    End Select
    
    DoChangeState Invalid
End Function

Private Function DoChangeState(newState As DetailState3)
    This.State = newState
    RaiseEvent StateChanged(newState)
End Function

' Editing
Public Function StartEditing()
    If This.IsEditing Then Exit Function
    This.IsEditing = True
    RaiseEvent EditingStarted
End Function

Public Function StopEditing()
    If This.IsEditing = False Then Exit Function
    This.IsEditing = False
    RaiseEvent EditingStopped
End Function

Public Function CancelEditing()
    If This.IsEditing = False Then Exit Function
    This.IsEditing = False
    This.frm.Undo
    Call TrySetState(Ready)
    RaiseEvent EditingCancelled
End Function


' SCD2 Saving
Public Function SaveEdits()
    If This.State <> EditMode Then Exit Function
    
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            ctl = GetControlByName(This.frm.controls, Replace(ctl.name, "LHS", "RHS"))
        End If
    Next ctl
    
    DoCmd.Save
    
    'Call TrySetState(ViewMode)
    StopEditing
End Function

Private Sub frmForEvents_AfterInsert()
    'Debug.Print "frmForEvents_AfterInsert()"
End Sub

Private Sub frmForEvents_AfterUpdate()
    'Debug.Print "frmForEvents_AfterUpdate()"
    Debug.Print "frmForEvents_AfterUpdate() & state = " & This.State
    If This.State = EditMode Then
        StopEditing
    End If
End Sub

' IAction stuff
Public Function GetButtonState(actionName As String) As Boolean
    Select Case actionName
        Case "New"
            'GetButtonState = True
            'Exit Function
            
        Case "SaveNew"
            GetButtonState = (This.State = NewMode)
            
        Case "CancelNew"
            GetButtonState = (This.State = NewMode)
            
        Case "View"
            'if (this.IsEditing = False) and (this.
            GetButtonState = True
            If This.IsEditing = True Then GetButtonState = False
            If This.State = ViewMode Then GetButtonState = False
            
        Case "Edit"
            GetButtonState = (This.State = ViewMode)
            
        Case "SaveEdits"
            GetButtonState = (This.State = EditMode)
            
        Case "CancelEdit"
            GetButtonState = (This.State = EditMode)
    End Select
End Function

Private Sub frmForEvents_GotFocus()
    Debug.Print "frmForEvents_GotFocus()"
    MsgBox "H"
    If This.State = NewMode And This.frm!ValidFrom = "" Then
        This.frm!EntityFK = Entity
        This.frm!ValidFrom = Now()
        This.frm!CommitFK = 1
    End If
End Sub
