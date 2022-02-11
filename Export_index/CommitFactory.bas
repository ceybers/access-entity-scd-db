Attribute VB_Name = "CommitFactory"
'@IgnoreModule
'@Folder "Factories"
Option Compare Database
Option Explicit

Public Function Create(ID As Double, Name As String) As clsCommit
    With New clsCommit
        .ID = ID
        .Name = Name
        Set Create = .Self
    End With
End Function

Public Function CreateFromRecordset(ByRef rs As DAO.Recordset) As clsCommit
    With New clsCommit
        .ID = rs!ID
        .Name = rs!Title
        Set CreateFromRecordset = .Self
    End With
End Function

