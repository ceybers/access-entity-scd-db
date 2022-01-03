Attribute VB_Name = "modLinkToBackEnd"
'@Folder "Main"
Option Compare Database
Option Explicit

Public Sub Main()
    LinkToBackEnd
End Sub

Public Sub LinkToBackEnd()
    Dim db As Database
    Set db = CreateBackEndConnection()
    
    LinkTable ENTITYTYPES_TABLE, db
    LinkTable ENTITIES_TABLE, db
    LinkTable "Not exist", db
    
    db.Close
    Set db = Nothing
End Sub
