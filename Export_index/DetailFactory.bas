Attribute VB_Name = "DetailFactory"
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, name As String, TableName As String) As clsDetail
    With New clsDetail
        .ID = ID
        .name = name
        .TableName = TableName
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsDetail
    With New clsDetail
        .ID = rs!ID
        .name = rs!DetailTable
        .TableName = rs!TableName
        Set CreateFromRecordset = .Self
    End With
End Function
