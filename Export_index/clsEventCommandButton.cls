VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Explorer1"
Option Compare Database
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents cmdBtn As Access.CommandButton
Attribute cmdBtn.VB_VarHelpID = -1

Private Type TEventCommandButton
    State As DetailStateForSCD
    Context As clsDetailViewModel
End Type
Private This As TEventCommandButton

Public Property Set ViewModel(Context As clsDetailViewModel)
    Set This.Context = Context
End Property

Public Property Set CommandButton(textbox As Access.CommandButton)
    Set cmdBtn = textbox
    cmdBtn.OnClick = "[Event Procedure]"
End Property

Public Property Let ActionState(State As DetailStateForSCD)
    This.State = State
End Property

Public Property Get ActionState() As DetailStateForSCD
    ActionState = This.State
End Property

Public Property Let Enabled(State As Boolean)
    Debug.Assert Not cmdBtn Is Nothing
    cmdBtn.Enabled = State
End Property

Public Property Get Enabled() As Boolean
    Debug.Assert Not cmdBtn Is Nothing
    Enabled = cmdBtn.Enabled
End Property

Public Property Get Self() As clsEventCommandButton
    Set Self = Me
End Property

Private Sub cmdBtn_Click()
    This.Context.HandleEvent This.State
End Sub

