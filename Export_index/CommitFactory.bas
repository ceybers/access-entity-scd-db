Attribute VB_Name = "CommitFactory"
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, name As String) As clsCommit
    With New clsCommit
        .ID = ID
        .name = name
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsCommit
    With New clsCommit
        .ID = rs!ID
        .name = rs!Title
        Set CreateFromRecordset = .Self
    End With
End Function

