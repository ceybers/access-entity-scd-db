VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FSMAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "SCDFSM"
Option Compare Database
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ctlCommandButton As Access.CommandButton
Attribute ctlCommandButton.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ctlOptionButton As Access.OptionButton
Attribute ctlOptionButton.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ctlCheckBox As Access.CheckBox
Attribute ctlCheckBox.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents ctlOptionGroup As Access.OptionGroup
Attribute ctlOptionGroup.VB_VarHelpID = -1

Private Type TState
    Context As SCDFSMViewModel
    Name As String
    Enabled As Boolean
End Type
Private This As TState

Public Property Get Enabled() As Boolean
    Enabled = This.Enabled
End Property

Public Property Let Enabled(ByVal rhs As Boolean)
    'If This.Enabled = RHS Then Exit Property
    This.Enabled = rhs
    If Not ctlCommandButton Is Nothing Then
        ctlCommandButton.Enabled = rhs
    End If
    If Not ctlOptionGroup Is Nothing Then
        ctlOptionGroup.Enabled = rhs
    End If
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal rhs As String)
    This.Name = rhs
End Property

Public Property Get Context() As SCDFSMViewModel
    Set Context = This.Context
End Property

Public Property Set Context(ByVal rhs As SCDFSMViewModel)
    Set This.Context = rhs
End Property

Public Property Get CommandButton() As Access.CommandButton
    Debug.Assert Not ctlCommandButton Is Nothing
    Set CommandButton = ctlCommandButton
End Property

Public Property Set CommandButton(ByVal CommandButton As Access.CommandButton)
    Set ctlCommandButton = CommandButton
    ctlCommandButton.OnClick = "[Event Procedure]"
End Property

Public Property Get OptionButton() As Access.OptionButton
    Debug.Assert Not ctlOptionButton Is Nothing
    Set OptionButton = ctlOptionButton
End Property

Public Property Set OptionButton(ByVal OptionButton As Access.OptionButton)
    Set ctlOptionButton = OptionButton
    ctlOptionButton.OnClick = "[Event Procedure]"
End Property

Public Property Get CheckBox() As Access.CheckBox
    Debug.Assert Not ctlCheckBox Is Nothing
    Set CheckBox = ctlCheckBox
End Property

Public Property Set CheckBox(ByVal CheckBox As Access.CheckBox)
    Set ctlCheckBox = CheckBox
    ctlCheckBox.OnClick = "[Event Procedure]"
End Property

Public Property Get OptionGroup() As Access.OptionGroup
    Debug.Assert Not ctlOptionGroup Is Nothing
    Set OptionGroup = ctlOptionGroup
End Property

Public Property Set OptionGroup(ByVal OptionGroup As Access.OptionGroup)
    Set ctlOptionGroup = OptionGroup
    ctlOptionGroup.OnClick = "[Event Procedure]"
End Property

Private Sub Class_Initialize()
    This.Enabled = True
End Sub

Private Sub ctlCheckBox_Click()
    This.Context.HandleEvent This.Name, ctlCheckBox.Value
End Sub

Private Sub ctlCommandButton_Click()
    This.Context.HandleEvent This.Name
End Sub

Private Sub ctlOptionButton_Click()
    This.Context.HandleEvent This.Name, ctlOptionButton.Value
End Sub

Private Sub ctlOptionGroup_Click()
    This.Context.HandleEvent This.Name, ctlOptionGroup.Value
End Sub

Public Sub SetControl(ByVal Control As Access.Control)
    If TypeOf Control Is Access.CommandButton Then
        Set CommandButton = Control
    ElseIf TypeOf Control Is Access.OptionButton Then
        Set OptionButton = Control
    ElseIf TypeOf Control Is Access.CheckBox Then
        Set CheckBox = Control
    ElseIf TypeOf Control Is Access.OptionGroup Then
        Set OptionGroup = Control
    End If
End Sub
