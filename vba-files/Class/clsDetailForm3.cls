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
    ready
    ViewNew
    EditNew
    ViewExisting
    EditExisting
End Enum

Private Type DetailForm3State
    frm As Form
    TableName As String
    SFrmName As String
    EntityID As Integer
    state As DetailState3
    IsEditing As Boolean
End Type
Private This As DetailForm3State

' Properties
Public Property Set Form(RHS As Form)
    Set This.frm = RHS
    This.SFrmName = RHS.name
    This.TableName = Replace(RHS.name, "sfrm", "tbl")
    Set frmForEvents = RHS
    Log "clsDetailForm3_Form('" & RHS.name & "')"
End Property

Public Property Let Entity(RHS As Integer)
    If This.EntityID = RHS Then
        Exit Property
    End If
    This.EntityID = RHS
    TrySetState ready
    Log "Changed entity to '" & This.EntityID & "'"
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
    
    Log "GotoEntity #" & Entity
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
    
    Log "CheckForExistingRecords = " & CStr(results)
    
    CheckForExistingRecords = results > 0
End Function

' Mode functions
Private Function ChangeToViewNewMode()
    With This.frm
        .AllowAdditions = True
        .AllowEdits = False
        .AllowDeletions = False
        .DataEntry = True
    End With
    
    This.frm.controls("ValidFrom").SetFocus
    This.frm.controls("ValidFrom").Locked = True
    This.frm.controls("CommitFK").Locked = True
    
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = False
            ctl.Visible = False
            ctl.Properties("BackColor") = 13431551
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
         ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394
        End If
    Next ctl
    
    Log "Now in View New Mode"
End Function

Private Function ChangeToEditNewMode()
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
            ctl.Visible = True
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
    
    This.frm!EntityFK = Entity
    This.frm!ValidFrom = Now()
    This.frm!CommitFK = 1
    
    Log "Now in Edit New Mode"
End Function

Private Function ChangeToViewExistingMode()
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
            ctl.Visible = True
            ctl.Properties("BackColor") = 16777215
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394 '4394 or 6804
        End If
    Next ctl
    
    StopEditing
    
    Log "Now in View Existing Mode"
End Function

Private Function ChangeToEditExistingMode()
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
            ctl.Visible = True
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
    
    Log "Now in Edit Existing Mode"
End Function

' More State stuff
Public Function TrySetState(newState As DetailState3) As Boolean
    If This.state = newState Then Exit Function
    
    Select Case newState
        Case ready
            If CheckForExistingRecords Then
                TrySetState ViewExisting
            Else
                TrySetState ViewNew
            End If
            Exit Function
            
        Case ViewNew
            If Not CheckForExistingRecords Then
                TrySetState = True
                ChangeToViewNewMode
                DoChangeState ViewNew
                Exit Function
            End If
            
        Case EditNew
            If Not CheckForExistingRecords Then
                TrySetState = True
                ChangeToEditNewMode
                DoChangeState EditNew
                Exit Function
            End If
            
        Case ViewExisting
            If CheckForExistingRecords Then
                TrySetState = True
                ChangeToViewExistingMode
                DoChangeState ViewExisting
                Exit Function
            End If
            
        Case EditExisting
            If This.state = ViewExisting Then
                TrySetState = True
                ChangeToEditExistingMode
                DoChangeState EditExisting
                Exit Function
            End If
    End Select
    
    DoChangeState Invalid
End Function

Private Function DoChangeState(newState As DetailState3)
    This.state = newState
    RaiseEvent StateChanged(newState)
    Log "DoChangeState(" & EnumToState(newState) & ")"
End Function

' Editing
Public Function StartEditing()
    If This.IsEditing Then Exit Function
    This.IsEditing = True
    RaiseEvent EditingStarted
    Log "StartEditing()"
End Function

Public Function StopEditing()
    If This.IsEditing = False Then Exit Function
    This.IsEditing = False
    RaiseEvent EditingStopped
    Log "StopEditing() - CHK if saved"
End Function

Public Function CancelEditing()
    Dim trackFK As Integer
    Dim detailID As Integer
    
    If This.IsEditing = False Then Exit Function
    This.IsEditing = False
    
    If (This.state = EditNew) And (Not This.frm.NewRecord) Then
        trackFK = This.frm.Recordset!trackFK
        detailID = This.frm.Recordset.fields("tblDetail.ID")
        Call RemoveCancelledNewRecord(trackFK, detailID)
        'Call RemoveCancelledNewRecord(This.frm.Recordset!trackFK, This.frm.Recordset.fields("tblDetail.ID"))
    End If
    
    This.frm.Undo
    
    If This.state = EditExisting Then
        Call TrySetState(ViewExisting)
    ElseIf This.state = EditNew Then
        'This.frm.Requery
        Call TrySetState(ViewNew)
    End If
    
    'Call TrySetState(Ready)
    RaiseEvent EditingCancelled
    Log "CancelEditing()"
End Function

Private Function RemoveCancelledNewRecord(trackFK As Integer, detailID As Integer)
    Call CurrentDb.Execute("DELETE * FROM tblTrack WHERE ID = " & trackFK & ";")
    Call CurrentDb.Execute("DELETE * FROM " & This.TableName & " WHERE ID = " & detailID & ";")
End Function

' SCD2 Saving
Public Function SaveEdits()
    'If This.state <> EditExisting Then Exit Function
    Dim ctl As control
    
    If This.state = EditExisting Then
        SaveEditExisting
    ElseIf This.state = EditNew Then
        ' Can't do this here: 3332 cannot enter value into blank field on one side of outer join
        'This.frm!EntityFK = Entity
        'This.frm!ValidFrom = Now()
        'This.frm!CommitFK = 1
    End If
    
    DoCmd.Save
    
    'Call TrySetState(ViewExistingMode)
    StopEditing
    Log "SaveEdits()"
End Function

Private Function SaveEditExisting()
    Dim rs As Recordset
    Dim ctl As control
    
    This.frm.Recordset.Edit
    This.frm.Recordset!ValidUntil = Now()
    This.frm.Recordset.Update
    
    ChangeToEditNewMode
    
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            ctl = GetControlByName(This.frm.controls, Replace(ctl.name, "LHS", "RHS"))
        End If
    Next ctl
    
    
End Function

Private Sub frmForEvents_AfterInsert()
    'Debug.Print "frmForEvents_AfterInsert()"
End Sub

Private Sub frmForEvents_AfterUpdate()
    'Debug.Print "frmForEvents_AfterUpdate()"
    Debug.Print "frmForEvents_AfterUpdate() & state = " & This.state
    If This.state = EditExisting Then
        StopEditing
    End If
End Sub

' IAction stuff
Public Function GetButtonState(actionName As String) As Boolean
    Select Case actionName
        Case "New"
            'GetButtonState = True
            'Exit Function
        Case "AddNew"
            GetButtonState = (This.state = ViewNew)
            
        Case "SaveNew"
            GetButtonState = (This.state = EditNew)
            
        Case "CancelNew"
            GetButtonState = (This.state = EditNew)
            
        Case "View"
            'if (this.IsEditing = False) and (this.
            GetButtonState = True
            If This.state = EditNew Then GetButtonState = False
            If This.state = ViewNew Then GetButtonState = False
            If This.IsEditing = True Then GetButtonState = False
            If This.state = ViewExisting Then GetButtonState = False
            
        Case "Edit"
            GetButtonState = (This.state = ViewExisting)
            
        Case "SaveEdits"
            GetButtonState = (This.state = EditExisting)
            
        Case "CancelEdit"
            GetButtonState = (This.state = EditExisting)
    End Select
End Function

Private Sub frmForEvents_GotFocus()
    Debug.Assert False
    Log "frmForEvents_GotFocus()"
    If This.state = NewMode And This.frm!ValidFrom = "" Then
        This.frm!EntityFK = Entity
        This.frm!ValidFrom = Now()
        This.frm!CommitFK = 1
    End If
End Sub

Private Function Log(text As String)
    'Me.lblLog.Caption = "[" & Format(Now(), "hh:mm:ss") & "] " & text & vbCrLf & Me.lblLog.Caption
    Debug.Print "[" & Format(Now(), "hh:mm:ss") & "] " & text
End Function

Public Function EnumToState(state As DetailState3)
    Dim v As Variant
    v = Array("Invalid", "Ready", "ViewNew", "EditNew", "ViewExisting", "EditExisting")
    EnumToState = CStr(v(state))
End Function
