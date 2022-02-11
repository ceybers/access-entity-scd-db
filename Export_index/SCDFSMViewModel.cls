VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCDFSMViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "SCDFSM"
Option Compare Database
Option Explicit
Implements IViewModel

Public Event PropertyChanged(ByVal PropertyName As String)
Public Event PleaseDebug()

Public Enum State
    Uninitialized
    ViewNone
    ViewOne
    ViewMany
    Add
    SaveAdd
    CancelAdd
    Edit
    SaveEdit
    CancelEdit
End Enum

Private Type TState
    CurrentState As State
    Edges As collection
    IsDirty As Boolean
    HasExistingRecords As Boolean
    Actions As collection
End Type
Private This As TState

' Properties
Public Property Get HasExistingRecords() As Boolean
    HasExistingRecords = This.HasExistingRecords
End Property

Public Property Let HasExistingRecords(rhs As Boolean)
    If This.IsDirty = True Then
        Err.Raise vbObjectError + 2, , "Cannot change HasExistingRecords or the form while editing (IsDirty = True)"
    End If
    
    If This.HasExistingRecords = rhs Then Exit Property
    This.HasExistingRecords = rhs
    
    If rhs = True And This.CurrentState = ViewNone Then
        HandleEvent "ViewOne", True
    ElseIf rhs = False And This.CurrentState = ViewOne Then
        HandleEvent "ViewNone", True
    ElseIf rhs = False And This.CurrentState = ViewMany Then
        HandleEvent "ViewNone", True
    End If
    
    RaiseEvent PropertyChanged("HasExistingRecords")
End Property

Public Property Get IsDirty() As Boolean
    IsDirty = This.IsDirty
End Property

Public Property Let IsDirty(rhs As Boolean)
    If This.IsDirty = rhs Then Exit Property
    This.IsDirty = rhs
    RaiseEvent PropertyChanged("IsDirty")
End Property

Public Property Get CurrentStateAsString() As String
    CurrentStateAsString = EnumToString(This.CurrentState)
End Property

Public Property Let CurrentState(rhs As State)
    If This.CurrentState = rhs Then Exit Property
    This.CurrentState = rhs
    UpdateActions
    RaiseEvent PropertyChanged("CurrentState")
End Property

Public Property Get CurrentState() As State
    CurrentState = This.CurrentState
End Property

' Ctor
Private Sub Class_Initialize()
    This.CurrentState = Uninitialized
    Set This.Actions = New collection
    Set This.Edges = New collection
    BuildEdges
End Sub

' Public Methods
Public Sub HandleEvent(ByVal EventName As String, Optional ByVal Enabled As Boolean = False)
    'MsgBox "EventName: " & EventName & " (" & Enabled & ")"
    Select Case EventName
        Case "Add"
            IsDirty = True
            CurrentState = Add
        Case "Edit"
            IsDirty = True
            CurrentState = Edit
        Case "Save"
            IsDirty = False
            If CurrentState = Add Then
                HasExistingRecords = True
                'CurrentState = SaveAdd
                CurrentState = ViewOne
            ElseIf CurrentState = Edit Then
                'CurrentState = SaveEdit
                CurrentState = ViewOne
            End If
        Case "Cancel"
            IsDirty = False
            If CurrentState = Add Then
                'CurrentState = CancelAdd
                CurrentState = ViewNone
            ElseIf CurrentState = Edit Then
                'CurrentState = CancelEdit
                CurrentState = ViewOne
            End If
        Case "ViewNone"
            CurrentState = ViewNone
        Case "ViewOne"
            CurrentState = ViewOne
        Case "ViewMany"
            CurrentState = ViewMany
            
        Case "IsEditing"
            IsDirty = Enabled
        Case "ExistingRecords"
            HasExistingRecords = Enabled
    End Select
    
    RaiseEvent PleaseDebug
End Sub

Public Sub AddEvent(ByVal EventName As String, ByVal Control As Object)
    Dim action As FSMAction
    Set action = New FSMAction
    
    action.Name = EventName
    Set action.Context = Me
    action.SetControl Control
    
    This.Actions.Add Item:=action, key:=EventName
End Sub

Public Function GetDEBUGText() As String
    Dim result As String
    
    result = result & "CurrentState: " & EnumToString(This.CurrentState) & vbCrLf
    result = result & "IsDirty: " & This.IsDirty & vbCrLf
    result = result & "HasExistingRecords: " & This.HasExistingRecords & vbCrLf
    result = result & "Actions count: " & This.Actions.count & vbCrLf
    
    GetDEBUGText = result
End Function

' Private Methods
Private Sub UpdateActions()
    Dim action As FSMAction
    For Each action In This.Actions
        action.Enabled = CheckEdge(This.CurrentState, action.Name)
        
        If action.Name = "ExistingRecords" Then
            action.Enabled = Not This.IsDirty
        End If
    Next action
End Sub

Private Function CheckEdge(ByVal LHS As State, ByVal rhs As String) As Boolean
    Dim thisEdge As String
    Dim checked As Variant
    thisEdge = LHS & "-->" & rhs
    
    On Error Resume Next
    checked = This.Edges(thisEdge)
    On Error GoTo 0
    
    CheckEdge = Not IsEmpty(checked)
End Function

Private Sub AddEdge(ByVal LHS As State, ByVal rhs As String)
    Dim thisEdge As String
    thisEdge = LHS & "-->" & rhs
    This.Edges.Add Item:=thisEdge, key:=thisEdge
End Sub

Private Sub BuildEdges()
    AddEdge ViewNone, "Add"
    AddEdge Add, "Save"
    AddEdge Add, "Cancel"
    AddEdge Edit, "Save"
    AddEdge Edit, "Cancel"
    AddEdge SaveAdd, "ViewOne"
    AddEdge SaveAdd, "ViewMany"
    AddEdge CancelAdd, "Add"
    AddEdge SaveEdit, "ViewOne"
    AddEdge SaveEdit, "ViewMany"
    AddEdge CancelEdit, "Edit"
    AddEdge CancelEdit, "ViewOne"
    AddEdge CancelEdit, "ViewMany"
    AddEdge ViewOne, "Edit"
    AddEdge ViewOne, "ViewMany"
    AddEdge ViewMany, "ViewOne"
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = False
End Function

Private Function EnumToString(ByVal mState As State) As String
    EnumToString = Array("Uninitialized", "ViewNone", "ViewOne", "ViewMany", "Add", "SaveAdd", "CancelAdd", "Edit", "SaveEdit", "CancelEdit")(mState)
End Function

