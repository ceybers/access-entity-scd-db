VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStateNewMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IState

Private Type TState
    Context As clsDetailForm2
End Type
Private This As TState

Private Function IState_CanChange(vCurrentState As Variant) As Boolean
    'This.Context.State would be used instead of vCurrentState
    ' then compare it against State.NewMode
    ' would also check This.HasExistingRecords = false
    ' and .isediting = false
    If This.Context.HasExistingRecords Then Exit Function
    If This.Context.IsEditing Then Exit Function
    If This.Context.State = NewMode Then Exit Function
    If This.Context.State = Invalid Then Exit Function
    IState_CanChange = True
End Function

Private Function IState_TryChange() As Boolean
    Dim frm As Form
    Set frm = This.Context.Form
    
    frm.DataEntry = True
    frm!EntityFK = This.Context.Entity
    frm!ValidFrom = Now()
    frm!CommitFK = 1

    This.Context.ResetLHSandRHS
    This.Context.HideRHS
    This.Context.UnlockLHS
End Function

Private Sub DoChange()
    
End Sub

Public Function Create(ByRef Context As Object) As IState
    Dim result As New clsStateNewMode
    result.Context = Context
    Set Create = result
End Function

Public Property Get Context() As Object
    Set Context = This.Context
End Property

Public Property Set Context(ByVal RHS As Object)
    Set This.Context = RHS
End Property
