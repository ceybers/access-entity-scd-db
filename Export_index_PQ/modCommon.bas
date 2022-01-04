Attribute VB_Name = "modCommon"
'@Folder "Common"
Option Compare Database
Option Explicit

Public Function CreateBackEndConnection() As Database
    Set CreateBackEndConnection = OpenDatabase(modConstants.BE_DATABASE_FILENAME, dbOpenSnapshot, dbReadOnly)
End Function

Public Function ClearCollection(ByRef coll As Collection) As Boolean
    If coll Is Nothing Then Exit Function
    If coll.Count = 0 Then Exit Function
    
    Do While coll.Count > 0
        coll.Remove 0
    Loop
    
    ClearCollection = True
End Function
