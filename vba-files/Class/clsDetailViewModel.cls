VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetailViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Event StateChanged(newState As DetailStateForSCD)
Public Event EditingStarted()
Public Event EditingStopped()

Private Const BACKCOLOR_YELLOW As Double = 13431551
Private Const BACKCOLOR_DEFAULT As Double = 16777215

Private WithEvents frmForEvents As Access.Form
Attribute frmForEvents.VB_VarHelpID = -1
Private Events As Collection

Public Enum DetailStateForSCD
    Invalid
    ready
    ViewNew
    EditNew
    ViewExisting
    EditExisting
    SaveNew
    CancelNew
    SaveExisting
    CancelExisting
    SaveAny
    CancelAny
End Enum

Private Type TDetailViewModel
    frm As Form
    sfrm As Subform
    tableName As String
    FrmName As String
    SFrmName As String
    EntityID As Integer
    CommitID As Integer
    State As DetailStateForSCD
    IsEditing As Boolean
    Detail As clsDetail
End Type
Private This As TDetailViewModel

' Properties
Public Property Set Detail(ByRef RHS As clsDetail)
    Debug.Assert Not RHS Is Nothing
    Log "(Property) Set Detail = '" & RHS.tableName & "'"
    
    Set This.Detail = RHS
    
    This.sfrm.SourceObject = Replace(RHS.tableName, "tbl", "sfrm")
    
    If This.EntityID > 0 Then TrySetState ready
End Property

Public Property Get Detail() As clsDetail
    Set Detail = This.Detail
End Property

' ---
Private Property Get DetailForm() As Form
    Debug.Assert Not This.sfrm Is Nothing
    Set DetailForm = This.sfrm.Form
End Property

Public Property Set Subform(RHS As Subform)
    Set This.sfrm = RHS
End Property

Public Property Set Form(RHS As Form)
    MsgBox "DEPREC"
    Exit Property
    Set This.frm = RHS
    This.SFrmName = RHS.name
    This.tableName = Replace(RHS.name, "sfrm", "tbl")
    Set frmForEvents = RHS
    Log "clsDetailViewModel_Form('" & RHS.name & "')"
    If This.EntityID > 0 Then TrySetState ready
End Property

' TODO Deprec this
Public Property Let tableName(tableName As String)
    'Err.Raise 5, , "DEPREC"
    If This.sfrm.SourceObject = "" Then Exit Property
    
    This.tableName = tableName
    This.SFrmName = Replace(tableName, "tbl", "sfrm")
    
    'If This.frm Is Nothing Then
    '    Set This.frm = Forms(This.FrmName) ' This is the detail form reference
    'End If
    'TODO FIX
    ' We are checkign if its loaded before we load it...
    'If Not CurrentProject.AllForms(This.SFrmName).IsLoaded Then
        'MsgBox "FAIL"
        'Exit Property
    'End If
    
    'Debug.Assert CurrentProject.AllForms(This.SFrmName).IsLoaded
    'If This.sfrm Is Nothing Then
        ' TODO FIX
        ' FIX This should fix it... we were trying to ref, but to load it we need to set .Form = formname as string
        ' Nope, this is still wrong - we lost the reference to the control
        'This.sfrm.Form = This.SFrmName
        'Set This.sfrm = Forms(This.SFrmName) ' This is the subform CONTROL reference
    'End If
    'If This.sfrm Is Nothing Then
        'Set This.sfrm = Forms!frmTest!sfrmDetail
    'End If
    
    If Not This.sfrm Is Nothing Then
        Set This.frm = This.sfrm.Form 'Forms(This.SFrmName)
        This.sfrm.SourceObject = This.SFrmName
    End If
    Set frmForEvents = This.frm
    Log "(Property) Let TableName = '" & tableName & "'"
    If This.EntityID > 0 Then TrySetState ready
End Property

Public Property Let formName(formName As String)
    This.FrmName = formName
End Property

Public Property Let Entity(RHS As Integer)
    If This.EntityID = RHS Then
        Exit Property
    End If
    This.EntityID = RHS
    If This.SFrmName <> "" Then TrySetState ready
    Log "(Property) Let Entity = '" & This.EntityID & "'"
End Property

Public Property Let CommitID(RHS As Integer)
    This.CommitID = RHS
End Property

Public Property Get tableName() As String
    tableName = This.tableName
End Property

Public Property Get Entity() As Integer
    Entity = This.EntityID
End Property

Public Property Get IsEditing() As Boolean
    IsEditing = This.IsEditing
End Property

' Ctor
Private Sub Class_Initialize()
    Set Events = New Collection
    This.CommitID = 1
End Sub

' Mode functions
Private Function ChangeToViewNewMode()
    SetFormEditState True, False, False, True
    
    With DetailForm
        .controls("ValidFrom").SetFocus
        .controls("ValidFrom").Locked = True
        .controls("CommitFK").Locked = True
    End With
    
    Dim ctl As control
    For Each ctl In DetailForm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = False
            ctl.Visible = False
            ctl.Properties("BackColor") = BACKCOLOR_YELLOW
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
         ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394
            ctl.Visible = False
        End If
    Next ctl
    
    Log "Now in View New Mode"
End Function

Private Function ChangeToEditNewMode()
    SetFormEditState True, True, False, True
    
    Dim ctl As control
    For Each ctl In DetailForm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = False
            ctl.Visible = True
            ctl.Properties("BackColor") = BACKCOLOR_YELLOW
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
         ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394
        End If
    Next ctl
    
    ' Do not use - this will throw Runtime 2449
    'This.frm.SetFocus
    
    StartEditing
    
    DetailForm!EntityFK = Entity
    DetailForm!ValidFrom = Now()
    DetailForm!CommitFK = This.CommitID
    
    Log "Now in Edit New Mode"
End Function

Private Function ChangeToViewExistingMode()
    SetFormEditState False, False, False, False
    
    GotoEntity
        
    Dim ctl As control
    For Each ctl In DetailForm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = True
            ctl.Visible = True
            ctl.Properties("BackColor") = BACKCOLOR_DEFAULT
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394 '4394 or 6804
            ctl.Visible = True
        End If
    Next ctl
    
    StopEditing
    
    Log "Now in View Existing Mode"
End Function

Private Function ChangeToEditExistingMode()
    SetFormEditState False, True, False, False

    GotoEntity
        
    Dim ctl As control
    For Each ctl In DetailForm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = True
            ctl.Visible = True
            ctl.Properties("BackColor") = BACKCOLOR_DEFAULT
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = True
            ctl.Properties("BackColor") = BACKCOLOR_YELLOW
            ctl = GetControlByName(DetailForm.controls, Replace(ctl.name, "RHS", "LHS"))
        ElseIf ctl.name Like "lblSuffix*" Then ctl.Properties("Left") = 6804
            '4394 or 6804
        End If
    Next ctl
    
    StartEditing
    
    Log "Now in Edit Existing Mode"
End Function

Private Function ChangeToSaveEdits()
    Dim rs As Recordset
    Dim ctl As control
    
    If This.State = EditExisting Then
        DetailForm.Recordset.Edit
        DetailForm.Recordset!ValidUntil = Now()
        DetailForm.Recordset.Update
        
        ChangeToEditNewMode
        
        For Each ctl In DetailForm.controls
            If ctl.name Like "???LHS*" Then
                ctl = GetControlByName(DetailForm.controls, Replace(ctl.name, "LHS", "RHS"))
            End If
        Next ctl
    End If
    
    DoCmd.Save
    
    StopEditing
    
    Log "ChangeToSaveEdits()"
End Function

Private Function ChangeToCancelEditing()
    Dim TrackFK As Integer
    Dim detailID As Integer
    
    If This.IsEditing = False Then Exit Function
    This.IsEditing = False
    
    If (This.State = EditNew) And (Not DetailForm.NewRecord) Then
        TrackFK = DetailForm.Recordset!TrackFK
        detailID = DetailForm.Recordset.fields("tblDetail.ID")
        Call RemoveCancelledNewRecord(TrackFK, detailID)
    End If
    
    DetailForm.Undo
    
    If This.State = EditExisting Then
        Call TrySetState(ViewExisting)
    ElseIf This.State = EditNew Then
        Call TrySetState(ViewNew)
    End If
    
    CancelEditing
    
    Log "ChangeToCancelEditing()"
End Function

' Giant Select Statement
Public Function TrySetState(newState As DetailStateForSCD) As Boolean
    Log "TrySetState(" & EnumToState(newState) & ")"
    Select Case newState
        Case ready
            If CheckForExistingRecords Then
                TrySetState ViewExisting
            Else
                TrySetState ViewNew
            End If
            Exit Function
            
        Case ViewNew
            TrySetState = True
            ChangeToViewNewMode
            DoChangeState ViewNew
            Exit Function
            
        Case EditNew
            TrySetState = True
            ChangeToEditNewMode
            DoChangeState EditNew
            Exit Function
            
        Case ViewExisting
            TrySetState = True
            ChangeToViewExistingMode
            DoChangeState ViewExisting
            Exit Function
            
        Case EditExisting
            TrySetState = True
            ChangeToEditExistingMode
            DoChangeState EditExisting
            Exit Function
        
        Case SaveExisting, SaveNew, SaveAny
            TrySetState = True
            ChangeToSaveEdits
            Exit Function
            
        Case CancelExisting, CancelNew, CancelAny
            TrySetState = True
            ChangeToCancelEditing
            Exit Function
        
    End Select
    
    DoChangeState Invalid
End Function

Private Function DoChangeState(newState As DetailStateForSCD)
    Log "DoChangeState from " & EnumToState(This.State) & " -> " & EnumToState(newState)
    This.State = newState
    RaiseEvent StateChanged(newState)
    UpdateButtonsState
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
    TrySetState ready
End Function

Public Function CancelEditing()
    RaiseEvent EditingStopped
    'MsgBox "NYI"
End Function

Private Function RemoveCancelledNewRecord(TrackFK As Integer, detailID As Integer)
    On Error GoTo ErrorHandling
    Call CurrentDb.Execute("DELETE * FROM tblTrack WHERE ID = " & TrackFK & ";")
    Call CurrentDb.Execute("DELETE * FROM " & This.tableName & " WHERE ID = " & detailID & ";")
    Exit Function
    
ErrorHandling:
    Err.Raise 5, , "Error in RemoveCancelledNewRecord()"
End Function


' Events for Handling
Private Sub frmForEvents_AfterUpdate()
    Log "frmForEvents_AfterUpdate() & State == " & EnumToState(This.State)
    If This.State = EditExisting Then
        ' FIX - use Save button, not Event?
        'StopEditing
    ElseIf This.State = DetailStateForSCD.EditNew Then
        'TODO Check if this is necessary
        'StopEditing
    End If
End Sub


Private Sub frmForEvents_BeforeUpdate(Cancel As Integer)
    ' TODO This won't work because Save/Cancel buttons are on the main form
    
    'Log "frmForEvents_BeforeUpdate and IsDirty == " & This.sfrm.Form.Dirty
    'If This.sfrm.Form.Dirty = True Then
    '    Cancel = True
    'End If
End Sub

' Helpers
Private Function GotoEntity()
    Debug.Assert Entity > 0
    DetailForm.Filter = "EntityFK = " & Entity & " AND ValidUntil = #9999/12/31#"
    DetailForm.FilterOn = True
    Call DetailForm.Recordset.FindLast("EntityFK = " & Entity & " AND ValidUntil = #9999/12/31#")
    
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

Private Function Log(text As String)
    Dim s As String
    s = "[" & format(Now(), "hh:mm:ss") & "] " & text
    Debug.Print s
End Function

Public Function EnumToState(State As DetailStateForSCD)
    Dim v As Variant
    v = Array("Invalid", "Ready", "ViewNew", "EditNew", "ViewExisting", "EditExisting", "SaveNew", "CancelNew", "SaveExisting", "CancelExisting", "SaveAny", "CancelAny")
    EnumToState = CStr(v(State))
End Function

Private Function SetFormEditState(additions As Boolean, edits As Boolean, deletions As Boolean, dataentry As Boolean)
    With DetailForm
        .AllowAdditions = additions
        .AllowEdits = edits
        .AllowDeletions = deletions
        .dataentry = dataentry
    End With
End Function

Private Function CheckForExistingRecords() As Boolean
    On Error GoTo ErrorHandling
    If tableName = "" Then Exit Function
    If Entity = 0 Then Exit Function
    
    Dim results As Integer
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT Count(*) AS TotalCount FROM " & tableName & " WHERE EntityFK = " & Entity & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    results = rs!TotalCount
    
    Log "CheckForExistingRecords = " & CStr(results)
    
    CheckForExistingRecords = results > 0
    Exit Function
    
ErrorHandling:
    Err.Raise 5, , "Something went wrong in CheckForExistingRecords()"
End Function

' Event handlers for CommandButtons
Public Function AddEvents(cb As CommandButton, State As DetailStateForSCD)
    With New clsEventCommandButton
        Set .CommandButton = cb
        Set .ViewModel = Me
        .ActionState = State
        Events.Add .Self
    End With
End Function

Private Sub UpdateButtonsState()
    Log "UpdateButtonState for " & CStr(Events.Count) & " buttons"
    Dim ecb As clsEventCommandButton
    For Each ecb In Events
        Select Case ecb.ActionState
            Case EditNew
                ecb.Enabled = (This.State = ViewNew)
            Case SaveNew
                ecb.Enabled = (This.State = EditNew)
            Case CancelNew
                ecb.Enabled = (This.State = EditNew)
            Case ViewExisting
                ecb.Enabled = True
                If This.sfrm.SourceObject = "" Then ecb.Enabled = False
                If This.IsEditing = True Then ecb.Enabled = False
                If This.State = EditNew Then ecb.Enabled = False
                If This.State = ViewNew Then ecb.Enabled = False
                If This.State = ViewExisting Then ecb.Enabled = False
            Case EditExisting
                ecb.Enabled = (This.State = ViewExisting)
            Case SaveExisting
                ecb.Enabled = (This.State = EditExisting)
            Case CancelExisting
                ecb.Enabled = (This.State = EditExisting)
            Case SaveAny
                ecb.Enabled = (This.State = EditNew) Or (This.State = EditExisting)
            Case CancelAny
                ecb.Enabled = (This.State = EditNew) Or (This.State = EditExisting)
        End Select
    Next ecb
End Sub

Public Function HandleEvent(State As DetailStateForSCD)
    TrySetState State
End Function