Attribute VB_Name = "EntityFactory"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, Name As String) As clsEntity
    With New clsEntity
        .ID = ID
        .Name = Name
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsEntity
    With New clsEntity
        .ID = rs!ID
        .Name = rs!Entity
        Set CreateFromRecordset = .Self
    End With
End Function
