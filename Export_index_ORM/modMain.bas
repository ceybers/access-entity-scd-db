Attribute VB_Name = "modMain"
'@Folder "Main"
Option Compare Database
Option Explicit

Public Sub Main()
    Dim ORM As ORM
    
    Debug.Print "ORM"
    Debug.Print "---"
    
    Set ORM = New ORM
    
    'TestEntityType ORM
    'TestCommits ORM
    TestLookups ORM
    Debug.Print "."
End Sub

Private Sub TestLookups(ByRef ORM As ORM)
    Dim lkpt As LookupTable
    Dim lkpv As LookupValue
    
    Debug.Print "List all lookup tables:"
    For Each lkpt In ORM.LookupTables
        Debug.Print "   " & lkpt.ToString
    Next lkpt
    Debug.Print " "
    
    Set lkpt = ORM.LookupTables.GetByName("lkpService")
    'Debug.Print "Null check: " & (Not lkpt.LookupValues Is Nothing)
    Debug.Print "LookupTable Name: " & lkpt.Name
    Debug.Print "LookupTable Name: " & lkpt.ToString
    Debug.Print "Values:"
    For Each lkpv In lkpt.LookupValues
        Debug.Print "   " & lkpv.ToString
    Next lkpv
End Sub

Private Sub TestCommits(ByRef ORM As ORM)
    Dim trk As Track
    Dim cmt As Commit
    
    For Each cmt In ORM.Commits
        Debug.Print cmt.ToString
    Next cmt
    
    'For Each trk In ORM.Tracks
    '    Debug.Print trk.ToString
    'Next trk
End Sub

Private Sub TestEntityType(ByRef ORM As ORM)
    'Set et = ORM.EntityTypes.GetByName("Depot")
    Dim et As EntityType
    Dim ent As Entity
    
    Set et = ORM.EntityTypes.GetByID(3)
    
    Debug.Print "Entity Type = " & et.ToString

    For Each ent In et.Entities
        Debug.Print "   " & ent.ToString
    Next ent
End Sub

