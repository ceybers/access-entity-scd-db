Attribute VB_Name = "modMain"
'@Folder "Main"
Option Compare Database
Option Explicit

Public Sub Main()
    'LinkToBackEnd
    
    Dim ORM As ORM
    Set ORM = New ORM
    Debug.Print "ORM"
    Debug.Print "---"
    
    Dim i As Double
    For i = 1 To ORM.EntityTypes.Count
        Debug.Print ORM.EntityTypes(i).ID & "# " & ORM.EntityTypes(i).Name
    Next i
    
    Dim et As EntityType
    For Each et In ORM.EntityTypes
        Debug.Print et.ID & "# " & et.Name
    Next et
    
    Debug.Print vbNullString
End Sub

Public Sub LinkToBackEnd()
    LinkTable ENTITYTYPES_TABLE
    LinkTable ENTITIES_TABLE
End Sub
