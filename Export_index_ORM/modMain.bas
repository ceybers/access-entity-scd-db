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
    
    'Dim i As Double
    'For i = 1 To ORM.EntityTypes.Count
    '    Debug.Print ORM.EntityTypes(i).ID & "# " & ORM.EntityTypes(i).Name
    'Next i
    
    'Debug.Print " "
     
    Dim ent As Entity
    For Each ent In ORM.Entities
        Debug.Print ent.ToString
    Next ent
    
    Debug.Print " "
End Sub
