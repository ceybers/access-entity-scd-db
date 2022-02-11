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

Public Sub CreateBlankForm(formName As String)
    Dim oldName As String
    Dim frm As Form
    Set frm = CreateForm()
    oldName = frm.Name
    DoCmd.Close acForm, oldName, acSaveYes
    DoCmd.Rename formName, acForm, oldName
End Sub

Public Function DeleteExistingForm(formName As String) As Boolean
    Dim frm As Object
    For Each frm In CurrentProject.AllForms
        If frm.Name = formName Then
            DoCmd.DeleteObject acForm, formName
            DeleteExistingForm = True
            Exit Function
        End If
    Next frm
End Function

Public Sub RemoveAllControlsFromForm(ByRef frm As Form)
    Dim i As Long
    For i = frm.controls.count To 1 Step -1
        DeleteControl frm.Name, frm.controls(i - 1).Name
    Next i
End Sub

Public Sub OpenFormInDesignMode(ByVal formName As String)
    DoCmd.OpenForm formName:=formName, view:=acDesign
End Sub

Public Sub CloseFormInDesignMode(ByVal formName As String)
    DoCmd.Close acForm, formName, acSaveYes
End Sub

Public Sub HideForm(ByVal formName As String)
    Application.SetHiddenAttribute acForm, formName, True
End Sub
