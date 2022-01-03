Attribute VB_Name = "modDBHelpers"
'@Folder("index_PQ")
Option Compare Database
Option Explicit

Public Function CreateBackEndConnection() As Database
    Set CreateBackEndConnection = OpenDatabase(modConstants.BE_DATABASE_FILENAME, dbOpenSnapshot, dbReadOnly)
End Function

Public Function CreateQuery(ByVal queryName As String, ByVal sql As String) As Boolean
    Dim db As Database
    Dim qdf As QueryDef
    
    Set db = CurrentDb
    
    If DoesQueryExist(queryName, db) Then Exit Function
    
    Set qdf = db.CreateQueryDef(queryName, sql)
    
    'db.QueryDefs.Append qdf
    db.QueryDefs.Refresh
    
    Set db = Nothing
    
    CreateQuery = True
End Function

Public Function LinkTable(ByVal tableName As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    Dim fld As field
    
    Set db = CurrentDb
    
    If DoesTableExist(tableName, db) Then Exit Function
    
    Set tbl = db.CreateTableDef(tableName)
    
    tbl.Connect = LINKED_DB_CONNECT & BE_DATABASE_FILENAME
    tbl.SourceTableName = tableName

    db.TableDefs.Append tbl
    db.TableDefs.Refresh
    
    Set db = Nothing
    
    LinkTable = True
End Function

Public Function DoesQueryExist(ByVal queryName As String, Optional ByRef db As Database) As Boolean
    Dim qry As QueryDef
    If db Is Nothing Then Set db = CurrentDb
    For Each qry In db.QueryDefs
        If qry.Name = queryName Then
            DoesQueryExist = True
            Exit Function
        End If
    Next qry
End Function

Public Function DoesTableExist(ByVal tableName As String, Optional ByRef db As Database) As Boolean
    Dim tbl As TableDef
    If db Is Nothing Then Set db = CurrentDb
    For Each tbl In db.TableDefs
        If tbl.Name = tableName Then
            DoesTableExist = True
            Exit Function
        End If
    Next tbl
End Function

Public Sub PrintOpenDatabases()
    Dim c As Integer
    Dim i As Integer
    c = DBEngine(0).Databases.Count
    Debug.Print "There are " & c & "x open database(s)"
    For i = 0 To (c - 1)
        Debug.Print "#" & i & " " & DBEngine(0).Databases(i).Name
    Next i
    Debug.Print
End Sub
