Attribute VB_Name = "modMain"
'@Folder "Main"
'@IgnoreModule

Option Compare Database
Option Explicit

Public Function GetORM(Optional Force As Boolean = False) As ORM
    Static ORM As ORM
    If Force = True Then
        Set ORM = New ORM
    ElseIf ORM Is Nothing Then
        Set ORM = New ORM
    End If
    Set GetORM = ORM
End Function

Public Sub Main()
    Dim ORM As ORM
    Set ORM = GetORM
    
    Debug.Print "ORM"
    Debug.Print "---"
        
    TestEntityType ORM
    TestCommits ORM
    TestLookups ORM
    TestDetails ORM
    TestDetailFields ORM
    Debug.Print "."
End Sub

Private Sub TestDetailFields(ByRef ORM As ORM)
    Dim detFld As DetailField
    For Each detFld In ORM.DetailFields
        Debug.Print detFld.ToString
    Next detFld
End Sub

Private Sub TestDetails(ByRef ORM As ORM)
    Dim detTbl As DetailTable
    Dim detVal As DetailValue
    
    Debug.Print "Listing Detail Tables:"
    For Each detTbl In ORM.DetailTables
        Debug.Print "   " & detTbl.ToString
    Next detTbl
    Debug.Print " "
    
    Set detTbl = ORM.DetailTables.GetByName("tblDetailDimensions")
    
    Debug.Print "Detail Table Name: " & detTbl.Name
    Debug.Print detTbl.ToString
    For Each detVal In detTbl.DetailValues
        Debug.Print "   " & detVal.ToString
    Next detVal
    Debug.Print " "
    
    Set detVal = detTbl.DetailValues(1)
    Debug.Print detVal.ToString
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

