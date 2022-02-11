Attribute VB_Name = "DetailFactory"
'@IgnoreModule
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, Name As String, tableName As String) As clsDetail
    With New clsDetail
        .ID = ID
        .Name = Name
        .tableName = tableName
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsDetail
    With New clsDetail
        .ID = rs!ID
        .Name = rs!DetailTable
        .tableName = rs!tableName
        Set CreateFromRecordset = .Self
    End With
End Function
