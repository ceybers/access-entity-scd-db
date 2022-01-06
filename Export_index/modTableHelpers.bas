Attribute VB_Name = "modTableHelpers"
'@IgnoreModule ProcedureCanBeWrittenAsFunction
'@Folder "Helpers"
Option Compare Database
Option Explicit

Public Function DropTables(ByVal tables As Collection, Optional ByRef db As DAO.Database) As Double
    Dim tbl As Variant
    
    If db Is Nothing Then Set db = CurrentDb
    For Each tbl In tables
       If DropTable(tbl, db) Then
            DropTables = DropTables + 1
       End If
    Next tbl
End Function

Public Function DropTable(ByVal tableName As String, Optional ByRef db As DAO.Database) As Boolean
    Dim sql As String
    If db Is Nothing Then Set db = CurrentDb
    If DoesTableExist(tableName, db) Then
        sql = "DROP TABLE " & tableName
        db.Execute sql, dbFailOnError
        DropTable = True
    End If
End Function

Public Function IsTableEmpty(ByVal tableName As String, Optional ByRef db As DAO.Database) As Boolean
    Dim result As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    If db Is Nothing Then Set db = CurrentDb
    sql = "SELECT Count(*) AS TotalCount FROM " & tableName & ";"
    Set rs = db.OpenRecordset(sql)
    result = rs.fields("TotalCount").Value
    
    rs.Close
    Set rs = Nothing
    
    IsTableEmpty = (result = 0)
End Function

Public Sub DEBUG_PrintTables(Optional ByRef db As DAO.Database)
    Dim tdf As TableDef
    
    If db Is Nothing Then Set db = CurrentDb
    
    Debug.Print "DEBUG_PrintTables()"
    
    For Each tdf In db.TableDefs
        Debug.Print tdf.name
    Next tdf
    
    Debug.Print vbNullString
End Sub

Public Function DoesTableExist(ByVal tableName As String, Optional ByRef db As DAO.Database) As Boolean
    Dim tdf As TableDef
    
    If db Is Nothing Then Set db = CurrentDb
    
    For Each tdf In db.TableDefs
        If tdf.name = tableName Then
            DoesTableExist = True
            Exit Function
        End If
    Next tdf
End Function

Public Function LinkTable(ByVal tableName As String) As Boolean
    Dim db As Database
    Dim tdf As TableDef
    On Error GoTo Catch
    
'Try
    Set db = CurrentDb
    
    If DoesTableExist(tableName, db) Then Exit Function
    
    Set tdf = db.CreateTableDef(tableName)
    
    tdf.Connect = LINKED_DB_CONNECT & BE_DATABASE_FILENAME
    tdf.SourceTableName = tableName

    db.TableDefs.Append tdf
    db.TableDefs.Refresh

    LinkTable = True
    GoTo Finally
    
Catch:
    If Err.Number = 3012 Then
        Debug.Print "LinkTable() failed - Table already exists in CurrentDB"
    ElseIf Err.Number = 3011 Then
        ' Could not find tableName in BackEnd DB
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
Finally:
End Function
