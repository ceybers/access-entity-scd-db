Attribute VB_Name = "modFormHelpers"
'@Folder "Helpers"
Option Compare Database
Option Explicit

Public Function DoesFormExist(ByVal formName As String) As Boolean
    Dim frm As Form
    For Each frm In Application.CurrentProject.AllForms
        If frm.Name = formName Then
            DoesFormExist = True
            Exit Function
        End If
    Next frm
End Function

Public Sub OpenFormInDesignMode(ByVal formName As String)
    DoCmd.OpenForm formName:=formName, View:=acDesign
End Sub

Public Sub CloseFormInDesignMode(ByVal formName As String)
    DoCmd.Close acForm, formName, acSaveYes
End Sub
