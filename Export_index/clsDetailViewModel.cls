VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetailViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder "ViewModels"
Option Compare Database
Option Explicit

Public Event StateChanged(newState As DetailStateForSCD)
Public Event EditingStarted()
Public Event EditingStopped()

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents frmForEvents As Access.Form
Attribute frmForEvents.VB_VarHelpID = -1
Private Events As Collection

Public Enum DetailStateForSCD
    invalid
    Ready
    Viewnew
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
    SubFormControl As SubForm
    EntityID As Integer
    CommitID As Integer
    State As DetailStateForSCD
    IsEditing As Boolean
    Detail As clsDetail
End Type
Private this As TDetailViewModel

' Properties
Public Property Set Detail(ByRef RHS As clsDetail)
    Debug.Assert Not RHS Is Nothing
    Log "(Property) Set Detail = '" & RHS.tableName & "'"
    
    Set this.Detail = RHS
    
    this.SubFormControl.SourceObject = Replace(RHS.tableName, "tbl", "sfrm")
    Log " SubFormControlSourceObject = " & Replace(RHS.tableName, "tbl", "sfrm")
    
    If this.EntityID > 0 Then TrySetState Ready
End Property

Public Property Get Detail() As clsDetail
    Set Detail = this.Detail
End Property

Public Property Set SubForm(ByRef RHS As SubForm)
    Set this.SubFormControl = RHS
End Property

Public Property Let EntityID(RHS As Integer)
    If this.EntityID = RHS Then
        Exit Property
    End If
    this.EntityID = RHS
    
    Log "(Property) Let Entity = '" & this.EntityID & "'"
    If Not this.Detail Is Nothing Then TrySetState Ready
End Property

Public Property Let CommitID(RHS As Integer)
    this.CommitID = RHS
End Property

Public Property Get CommitID() As Integer
    CommitID = this.CommitID
End Property

Public Property Get EntityID() As Integer
    EntityID = this.EntityID
End Property

Public Property Get IsEditing() As Boolean
    IsEditing = this.IsEditing
End Property

Private Property Get DetailForm() As Form
    'Debug.Assert Not This.Detail Is Nothing
    'Debug.Assert This.SubFormControl Is Nothing
    Set DetailForm = this.SubFormControl.Form
End Property

' Ctor
Private Sub Class_Initialize()
    Set Events = New Collection
    this.CommitID = 1
    this.State = invalid
End Sub

' Mode functions
Private Function ChangeToViewNewMode()
    SetFormEditState True, False, False, True
    
    With DetailForm
        .controls(TRACK_VALIDFROM_FIELDNAME).SetFocus
        .controls(TRACK_VALIDFROM_FIELDNAME).Locked = True
        .controls(TRACK_COMMITFK_FIELDNAME).Locked = True
    End With
    
    Dim ctl As control
    For Each ctl In DetailForm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Locked = False
            ctl.Visible = False
            ctl.Properties("BackColor") = RAGColors.yellow
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
            ctl.Properties("BackColor") = RAGColors.yellow
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
         ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394
        End If
    Next ctl
    
    StartEditing
    
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
            ctl.Properties("BackColor") = RAGColors.Default
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Properties("Left") = 4394 '4394 or 6804
            ctl.Visible = True
        End If
    Next ctl
    
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
            ctl.Properties("BackColor") = RAGColors.Default
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = True
            ctl.Properties("BackColor") = RAGColors.yellow
            Set ctl = GetControlByName(DetailForm.controls, Replace(ctl.name, "RHS", "LHS"))
        ElseIf ctl.name Like "lblSuffix*" Then ctl.Properties("Left") = 6804
            '4394 or 6804
        End If
    Next ctl
    
    StartEditing
    
    Log "Now in Edit Existing Mode"
End Function

Private Function ChangeToSaveEdits()
    Dim fieldName As String
    Dim rs As Recordset
    Dim ctl As control
    
    Set rs = DetailForm.Recordset
    Debug.Assert Not rs.BOF
    Debug.Assert Not rs.EOF
    
    If this.State = EditExisting Then
        rs.Edit
        rs!ValidUntil = Now()
        rs.Update
        
        SetFormEditState True, True, False, True
        ' Needs to be done again after changing SetFormEditState
        Set rs = DetailForm.Recordset
        
        rs.AddNew
        
        For Each ctl In DetailForm.controls
            If ctl.name Like "???RHS*" Then
                fieldName = Right$(ctl.name, Len(ctl.name) - 6)
                 rs.fields(fieldName) = ctl
            End If
        Next ctl
    ElseIf this.State = EditNew Then
        rs.Edit
    End If
    
    rs!EntityFK = EntityID
    rs!TrackFK = CreateNewTrackRecord(CommitID)
    rs.Update
    
    StopEditing
End Function

Private Function ChangeToCancelEditing()
    Dim TrackFK As Integer
    Dim detailID As Integer
    
    If this.IsEditing = False Then Exit Function
    this.IsEditing = False
    
    If (this.State = EditNew) And (Not DetailForm.NewRecord) Then
        TrackFK = DetailForm.Recordset!TrackFK
        detailID = DetailForm.Recordset.fields("tblDetail.ID")
        RemoveCancelledNewRecord TrackFK, detailID
    End If
    
    DetailForm.Undo
    
    If this.State = EditExisting Then
        TrySetState ViewExisting
    ElseIf this.State = EditNew Then
        TrySetState Viewnew
    End If
    
    CancelEditing
    
    Log "ChangeToCancelEditing()"
End Function

' Giant Select Statement
Public Function TrySetState(newState As DetailStateForSCD) As Boolean
    Log "TrySetState(" & EnumToState(newState) & ")"
    Select Case newState
        Case Ready
            DoChangeState Ready
            If CheckForExistingRecords Then
                TrySetState ViewExisting
            Else
                TrySetState Viewnew
            End If
            Exit Function
            
        Case Viewnew
            TrySetState = True
            ChangeToViewNewMode
            DoChangeState Viewnew
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
    
    DoChangeState invalid
End Function

Private Function DoChangeState(newState As DetailStateForSCD)
    Log "DoChangeState from " & EnumToState(this.State) & " -> " & EnumToState(newState)
    this.State = newState
    RaiseEvent StateChanged(newState)
    UpdateButtonsState
End Function

' Editing
Public Function StartEditing()
    If this.IsEditing Then Exit Function
    this.IsEditing = True
    RaiseEvent EditingStarted
    Log "StartEditing()"
End Function

Public Function StopEditing()
    If this.IsEditing = False Then Exit Function
    this.IsEditing = False
    RaiseEvent EditingStopped
    Log "StopEditing()"
    TrySetState Ready
End Function

Public Function CancelEditing()
    RaiseEvent EditingStopped
    'MsgBox "NYI"
End Function

Private Function RemoveCancelledNewRecord(TrackFK As Integer, detailID As Integer)
    On Error GoTo ErrorHandling
    CurrentDb.Execute "DELETE * FROM " & TRACKS_TABLE & " WHERE ID = " & TrackFK & ";"
    CurrentDb.Execute "DELETE * FROM " & Detail.tableName & " WHERE ID = " & detailID & ";"
    Exit Function
    
ErrorHandling:
    Err.Raise 5, , "Error in RemoveCancelledNewRecord()"
End Function

Private Function CreateNewTrackRecord(CommitFK As Double) As Integer
    Dim db As Database
    Dim rs As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & TRACKS_TABLE, dbOpenDynaset, dbSeeChanges)
    
    rs.AddNew
    CreateNewTrackRecord = rs!ID
    rs.fields(TRACK_COMMITFK_FIELDNAME) = CommitFK
    rs.fields(TRACK_VALIDFROM_FIELDNAME) = Now()
    rs.fields(TRACK_VALIDUNTIL_FIELDNAME) = #12/31/9999#
    rs.Update
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Form Events

' Helpers
Private Function GotoEntity()
    Dim recordSource As String
    Dim frm As Form
    Dim rs As Recordset
    
    Debug.Assert EntityID > 0
    
    Set frm = DetailForm
    recordSource = frm.recordSource
    frm.recordSource = vbNullString
    
    frm.Filter = "EntityFK = " & EntityID & " AND " & TRACK_VALIDUNTIL_FIELDNAME & " = #9999/12/31#"
    frm.FilterOn = True
    frm.recordSource = recordSource
    
    Set rs = frm.RecordsetClone
    rs.FindLast "EntityFK = " & EntityID & " AND " & TRACK_VALIDUNTIL_FIELDNAME & " = #9999/12/31#"
    
    frm.Bookmark = rs.Bookmark
    
    Log "GotoEntity #" & EntityID
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
    If Detail Is Nothing Then Exit Function
    If EntityID = 0 Then Exit Function
    
    Dim results As Integer
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT Count(*) AS TotalCount FROM " & Detail.tableName & " WHERE EntityFK = " & EntityID & ";"
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
    Log "UpdateButtonState for " & CStr(Events.count) & " buttons"
    Dim ecb As clsEventCommandButton
    For Each ecb In Events
        Select Case ecb.ActionState
            Case EditNew
                ecb.Enabled = (this.State = Viewnew)
            Case SaveNew
                ecb.Enabled = (this.State = EditNew)
            Case CancelNew
                ecb.Enabled = (this.State = EditNew)
            Case ViewExisting
                ecb.Enabled = True
                If Detail Is Nothing Then ecb.Enabled = False
                If this.IsEditing = True Then ecb.Enabled = False
                If this.State = EditNew Then ecb.Enabled = False
                If this.State = Viewnew Then ecb.Enabled = False
                If this.State = ViewExisting Then ecb.Enabled = False
            Case EditExisting
                ecb.Enabled = (this.State = ViewExisting)
            Case SaveExisting
                ecb.Enabled = (this.State = EditExisting)
            Case CancelExisting
                ecb.Enabled = (this.State = EditExisting)
            Case SaveAny
                ecb.Enabled = (this.State = EditNew) Or (this.State = EditExisting)
            Case CancelAny
                ecb.Enabled = (this.State = EditNew) Or (this.State = EditExisting)
        End Select
    Next ecb
End Sub

Public Function HandleEvent(State As DetailStateForSCD)
    TrySetState State
End Function

