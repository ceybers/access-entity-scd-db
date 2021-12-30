Attribute VB_Name = "EntityTypeFactory"
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, name As String) As clsEntityType
    With New clsEntityType
        .ID = ID
        .name = name
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsEntityType
    With New clsEntityType
        .ID = rs!ID
        .name = rs!EntityType
        Set CreateFromRecordset = .Self
    End With
End Function
