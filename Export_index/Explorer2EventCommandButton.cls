VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Explorer2EventCommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Explorer2")
Option Compare Database
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents cmdBtn As Access.CommandButton
Attribute cmdBtn.VB_VarHelpID = -1

Private Type TExplorer2EventCommandButton
    Context As ViewModelExplorer2
    State As String
End Type

Private this As TExplorer2EventCommandButton

Public Property Get Context() As ViewModelExplorer2
    Set Context = this.Context
End Property

Public Property Set Context(ByVal RHS As ViewModelExplorer2)
    Set this.Context = RHS
End Property

Public Property Get State() As String
    State = this.State
End Property

Public Property Let State(ByVal RHS As String)
    this.State = RHS
End Property

Public Property Get Self() As Explorer2EventCommandButton
    Set Self = Me
End Property

Private Sub cmdBtn_Click()
    this.Context.HandleEvent this.State
End Sub

Public Property Set CommandButton(textbox As Access.CommandButton)
    Set cmdBtn = textbox
    cmdBtn.OnClick = "[Event Procedure]"
End Property
