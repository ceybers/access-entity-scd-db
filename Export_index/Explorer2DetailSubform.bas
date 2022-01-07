Attribute VB_Name = "Explorer2DetailSubform"
'@Folder("Explorer2")
Option Compare Database
Option Explicit

Public Sub EditNew(ByRef sfrm As Subform)
    If sfrm.SourceObject = vbNullString Then
        Exit Sub
    End If
    
    Dim frm As Form
    Set frm = sfrm.Form
    
    If Not frm.Recordset.BOF Then Exit Sub
    If Not frm.Recordset.EOF Then Exit Sub
    
    sfrm.SetFocus
    LeftOnlyReadWrite frm
End Sub

Public Sub ViewOne(ByRef sfrm As Subform)
    If sfrm.SourceObject = vbNullString Then
        Exit Sub
    End If
    
    Dim frm As Form
    Set frm = sfrm.Form
    
    If frm.Recordset.BOF Or frm.Recordset.EOF Then
        sfrm.Parent.controls("txtSearch").SetFocus
        LeftOnlyBlank frm
        Exit Sub
    End If
    
    sfrm.SetFocus
    HideSCD frm
    LeftOnlyReadOnly frm
End Sub

Public Sub ViewMany(ByRef sfrm As Subform)
    If sfrm.SourceObject = vbNullString Then
        Exit Sub
    End If
    
    Dim frm As Form
    Set frm = sfrm.Form
    
    If frm.Recordset.BOF Then Exit Sub
    If frm.Recordset.EOF Then Exit Sub
    
    sfrm.SetFocus
    DatasheetHistory frm
End Sub

Private Sub HideSCD(ByRef frm As Form)
    Dim ctl As control
    
    For Each ctl In frm.controls
        If ctl.name Like "*TrackFK" Or ctl.name Like "*CommitFK" Then
            ctl.Visible = False
        ElseIf ctl.name Like "txtLHS*" Then
            ctl.SetFocus
        End If
    Next ctl
    
    frm.controls("ValidFrom").Locked = True
    frm.controls("ValidUntil").Locked = True
    frm.controls("ValidFrom").Properties("BackColor") = RAGColors.Grey
    frm.controls("ValidUntil").Properties("BackColor") = RAGColors.Grey
End Sub

Private Sub LeftOnlyReadWrite(ByRef frm As Form)
    DoCmd.RunCommand acCmdSubformFormView
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Visible = True
            ctl.Properties("Backcolor") = RAGColors.Yellow
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Visible = True
        End If
    Next ctl
    frm.AllowAdditions = True
    frm.dataentry = False
End Sub

Private Sub LeftOnlyReadOnly(ByRef frm As Form)
    DoCmd.RunCommand acCmdSubformFormView
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Visible = True
            ctl.Locked = True
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Visible = True
        End If
    Next ctl
    frm.AllowAdditions = False
    frm.dataentry = False
End Sub

Private Sub LeftOnlyBlank(ByRef frm As Form)
    'DoCmd.RunCommand acCmdSubformFormView
    Dim ctl As control
    frm.AllowAdditions = False
    frm.dataentry = True
    ' AllowAdd = false and DataEntry = true will blank the subform OK
    Exit Sub
    
    For Each ctl In frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Visible = False
            ctl.Properties("Backcolor") = RAGColors.Red
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = False
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Visible = False
        End If
    Next ctl
    
End Sub

Private Sub LeftRightReadWrite(ByRef frm As Form)
    DoCmd.RunCommand acCmdSubformFormView
    Dim ctl As control
    For Each ctl In frm.controls
        If ctl.name Like "???LHS*" Then
            ctl.Visible = True
        ElseIf ctl.name Like "???RHS*" Then
            ctl.Visible = True
            ctl.Properties("BackColor") = RAGColors.Yellow
        ElseIf ctl.name Like "lblSuffix*" Then
            ctl.Visible = False
        End If
    Next ctl
End Sub

Private Sub DatasheetHistory(ByRef frm As Form, Optional ByRef sfrm As Subform)
    DoCmd.RunCommand acCmdSubformDatasheetView
    Dim ctl As control
    
    frm.AllowAdditions = False
    frm.dataentry = False
    
    For Each ctl In frm.controls
        If ctl.name Like "txtRHS*" Or ctl.name = "TrackFK" Or ctl.name = "CommitFK" Then
            ctl.ColumnHidden = True
        ElseIf ctl.name Like "Valid*" Then
            ctl.ColumnWidth = 4 * CM_TO_TWIP
        End If
    Next ctl
End Sub

