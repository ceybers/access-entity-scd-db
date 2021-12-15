Attribute VB_Name = "DetailFactory"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, Name As String, TableName As String) As clsDetail
    With New clsDetail
        .ID = ID
        .Name = Name
        .TableName = TableName
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsDetail
    With New clsDetail
        .ID = rs!ID
        .Name = rs!DetailTable
        .TableName = rs!TableName
        Set CreateFromRecordset = .Self
    End With
End Function
