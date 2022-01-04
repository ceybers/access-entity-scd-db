VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents tb As Access.textbox
Attribute tb.VB_VarHelpID = -1

Public Property Set SetTextbox(textbox As Access.textbox)
    Set tb = textbox
    tb.OnChange = vbNullString '"[Event Procedure]"
    tb.OnExit = "[Event Procedure]"
End Property

Private Sub tb_Change()
    Debug.Print tb.Parent.name & "!" & tb.name & "_Change()"
    Dim tbBefore As textbox
    Set tbBefore = tb.Parent.controls(Replace(tb.name, "txtAfter", "txtBefore"))
    If CStr(tbBefore) <> CStr(tb) Then
        Debug.Print "Differs"
        Debug.Print "   " & CStr(tbBefore)
        Debug.Print "   " & CStr(tb)
        
    Else
        Debug.Print "Same"
    End If
End Sub


Private Sub tb_Exit(Cancel As Integer)
    Debug.Print tb.Parent.name & "!" & tb.name & "_Exit()"
    Dim tbBefore As textbox
    Set tbBefore = tb.Parent.controls(Replace(tb.name, "txtRHS", "txtLHS"))
    If CStr(tbBefore) <> CStr(Nz(tb)) Then
        tb.BackColor = RAGColors.Red
    Else
        tb.BackColor = RAGColors.Green
    End If
End Sub
