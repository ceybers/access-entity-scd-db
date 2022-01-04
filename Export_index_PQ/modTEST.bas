Attribute VB_Name = "modTEST"
'@Folder("index_PQ")
Option Compare Database
Option Explicit

Private Sub TEST_GetLinkedTables()
    Debug.Print "TEST_GetLinkedTables()"
    Dim tables As Variant
    
    Set tables = GetListOfTablesInLinkedDatabase
    
    'PrintCollection tables
    LinkTables tables
    
    PrintOpenDatabases
End Sub

Private Sub LinkTables(tables As Variant)
    Dim tbl As Variant
    For Each tbl In tables
        If tbl Like "tbl*" Or tbl Like "lkp*" Then
            LinkTable tbl
        End If
    Next tbl
    
    LinkTable SCHEMA_TABLE
    LinkTable QUERY_TRACK_LATEST
End Sub

Private Sub PrintCollection(ByRef coll As Variant)
    Dim v As Variant
    For Each v In coll
        Debug.Print v
    Next v
End Sub

Private Function GetTestListOfTables() As Variant
    Set GetTestListOfTables = New Collection
    With GetTestListOfTables
        .Add "tblEntities"
        .Add "tblTestTable"
        .Add "tblTrack"
    End With
End Function

Public Function GetListOfTablesInLinkedDatabase() As Variant
    Dim db As Database
    Set db = CreateBackEndConnection
    Set GetListOfTablesInLinkedDatabase = GetListOfTables(db)
    db.Close
    Set db = Nothing
End Function

Public Function GetListOfTables(Optional ByRef db As Database) As Variant
    Dim result As Collection
    Dim tbl As TableDef
    
    Set result = New Collection
    If db Is Nothing Then
        Set db = CurrentDb
    End If
    
    For Each tbl In db.TableDefs
        If Not tbl.Name Like "MSys*" Then
            result.Add tbl.Name
        End If
    Next tbl
    
    Set GetListOfTables = result
End Function
