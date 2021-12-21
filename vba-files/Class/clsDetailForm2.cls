VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDetailForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Enum DetailForm2State
    Invalid
    ready
    NewMode
    ViewMode
    EditMode
End Enum

Private Type TDetailForm2
    LHS As New Collection
    RHS As New Collection
    state As DetailForm2State
    EntityID As Integer
    CommityID As Integer
    frm As Form
    TableName As String
    IsEditing As Boolean
    HasExistingRecords As Boolean
End Type
Private This As TDetailForm2

' Properties
Public Property Get IsEditing() As Boolean
    IsEditing = This.IsEditing
End Property

Public Property Get HasExistingRecords() As Boolean
    HasExistingRecords = This.HasExistingRecords
End Property

Public Property Get state() As DetailForm2State
    state = This.state
End Property

Public Property Set Form(frm As Form)
    Debug.Assert Not frm Is Nothing
    Set This.frm = frm
    LoadLHSRHSControls
    AddTableName
End Property

Public Property Let Entity(EntityID As Integer)
    Debug.Assert This.TableName <> ""
    This.EntityID = EntityID
    This.HasExistingRecords = CheckForExistingRecords
End Property

Public Property Get Entity() As Integer
    Entity = This.EntityID
End Property

Public Property Get Form() As Form
    Debug.Assert Not This.frm Is Nothing
    Set Form = This.frm
End Property

' Ctor
Private Sub Class_Initialize()
    This.state = Invalid
End Sub

' Public Methods
Public Sub ResetLHSandRHS()
    Debug.Print "NYI"
End Sub

Public Sub HideRHS()
    Dim ctl As control
    For Each ctl In This.RHS
        ctl.Visible = False
    Next ctl
End Sub

Public Sub UnlockLHS()
    Dim ctl As control
    For Each ctl In This.LHS
        ctl.Locked = False
    Next ctl
End Sub
    
Public Function ChangeState(newState As DetailForm2State) As Boolean
    Select Case This.state
        Case Invalid
            Select Case newState
                Case Invalid
                    Exit Function
                Case NewMode
                    ' try if no exist
                Case ViewMode
                    ' try if has exist
                Case EditMode
                    Exit Function
            End Select
            
        Case NewMode
            Select Case newState
                Case Invalid
                Case NewMode
                    Exit Function
                Case ViewMode
                    ' only after saving
                Case EditMode
                    ' no
            End Select
            
        Case ViewMode
            Select Case newState
                Case Invalid
                Case NewMode
                    Exit Function
                Case ViewMode
                    Exit Function
                Case EditMode
                    ' try
            End Select
            
        Case EditMode
            Select Case newState
                Case Invalid
                Case NewMode
                Case ViewMode
                    ' cancel and try
                Case EditMode
                    Exit Function
            End Select
    End Select
End Function

' Private Methods
Private Function TryChangeStateNewMode() As Boolean
    '
End Function

Private Sub LoadLHSRHSControls()
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name Like "???LHS*" Then
            This.LHS.Add ctl
        ElseIf ctl.name Like "???RHS*" Then
            This.RHS.Add ctl
        End If
    Next
End Sub

Private Sub AddTableName()
    Debug.Assert Not This.frm Is Nothing
    This.TableName = Replace(This.frm.name, "sfrm", "tbl")
End Sub

Private Function CheckForExistingRecords() As Boolean
    Dim results As Integer
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT Count(*) AS TotalCount FROM " & This.TableName & " WHERE EntityFK = " & This.EntityID & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    results = rs!TotalCount
    
    CheckForExistingRecords = results > 0
End Function

Private Function zzGetControlByName(name As String)
    Debug.Assert Not This.frm Is Nothing
    Dim ctl As control
    For Each ctl In This.frm.controls
        If ctl.name = name Then
            Set zzGetControlByName = ctl
            Exit Function
        End If
    Next ctl
    Err.Raise 5, , "Cannot find GetControlByName '" & name & "'"
End Function

