Attribute VB_Name = "modMain"
'@Folder "Main"
Option Compare Database
Option Explicit

Public Sub Main()
    Dim ORM As ORM
    Dim et As EntityType
    Dim ent As Entity
    
    Debug.Print "ORM"
    Debug.Print "---"
    
    Set ORM = New ORM
    'Set et = ORM.EntityTypes.GetByName("Depot")
    Set et = ORM.EntityTypes.GetByID(3)
    
    Debug.Print "Entity Type = " & et.ToString

    For Each ent In et.Entities
        Debug.Print "   " & ent.ToString
    Next ent
    
    Debug.Print "."
End Sub
