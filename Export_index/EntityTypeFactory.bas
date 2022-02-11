Attribute VB_Name = "EntityTypeFactory"
'@IgnoreModule
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, Name As String) As clsEntityType
    With New clsEntityType
        .ID = ID
        .Name = Name
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsEntityType
    With New clsEntityType
        .ID = rs!ID
        .Name = rs!EntityType
        Set CreateFromRecordset = .Self
    End With
End Function
