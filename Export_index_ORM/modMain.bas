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
    
    'TestEntityType ORM
    TestCommits ORM
    Debug.Print "."
End Sub

Private Sub TestCommits(ByRef ORM As ORM)
    Debug.Print ORM.Commits(1).Name
End Sub

Private Sub TestEntityType(ByRef ORM As ORM)
    'Set et = ORM.EntityTypes.GetByName("Depot")
    Set et = ORM.EntityTypes.GetByID(3)
    
    Debug.Print "Entity Type = " & et.ToString

    For Each ent In et.Entities
        Debug.Print "   " & ent.ToString
    Next ent
End Sub
