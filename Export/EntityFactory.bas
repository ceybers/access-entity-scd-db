Attribute VB_Name = "EntityFactory"
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, name As String, EntityType As Double) As clsEntity
    With New clsEntity
        .ID = ID
        .name = name
        .EntityType = EntityType
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsEntity
    With New clsEntity
        .ID = rs!ID
        .name = rs!Entity
        .EntityType = rs!EntityType
        Set CreateFromRecordset = .Self
    End With
End Function
