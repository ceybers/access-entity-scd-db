Attribute VB_Name = "modCreateTableHelpers"
'@Folder("Common")
Option Compare Database
Option Explicit

Public Function CreateGenericField(tbl As TableDef, fieldName As String, Optional fieldType As Integer = dbText)
    Dim fld As Field
    Set fld = tbl.CreateField(fieldName, fieldType)
    tbl.Fields.Append fld
End Function

Public Function CreateIDField(tbl As TableDef, Optional fieldName As String = "ID")
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    Set fld = CreateAutoNumberField(tbl, fieldName)
    tbl.Fields.Append fld
    
    Set idx = tbl.CreateIndex
    
    With idx
        .Name = "Primary Key"
        .Fields.Append .CreateField(fieldName)
        .Unique = True
        .Primary = True
    End With
    tbl.Indexes.Append idx
End Function

Public Function CreateAutoNumberField(tbl As TableDef, fieldName As String) As Field
    Set CreateAutoNumberField = tbl.CreateField(fieldName, dbLong, 4)
    With CreateAutoNumberField
         .Attributes = dbAutoIncrField
    End With
End Function

Public Function IsTableEmpty(db, tableName As String) As Boolean
    Dim result As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    sql = "SELECT Count(*) AS TotalCount FROM " & tableName & ";"
    Set rs = db.OpenRecordset(sql)
    result = rs!TotalCount
    
    rs.Close
    Set rs = Nothing
    
    IsTableEmpty = (result = 0)
End Function

Public Function DoesTableExist(ByRef db As Database, ByVal tableName As String) As Boolean
    Dim tbl As TableDef
    For Each tbl In db.TableDefs
        If tbl.Name = tableName Then
            DoesTableExist = True
            Exit Function
        End If
    Next tbl
End Function

Public Function LinkTable(tableName As String, LINKED_DB_CONNECT As String, BE_DATABASE_FILENAME As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    Dim fld As Field
    
    Set db = CurrentDb
    
    If DoesTableExist(db, tableName) Then Exit Function
    
    Set tbl = db.CreateTableDef(tableName)
    
    tbl.Connect = LINKED_DB_CONNECT & BE_DATABASE_FILENAME
    tbl.SourceTableName = tableName

    db.TableDefs.Append tbl
    db.TableDefs.Refresh
    
    Set db = Nothing
    
    LinkTable = True
    
    Debug.Print "Linked table '" & tableName & "' from '" & BE_DATABASE_FILENAME & "'"
End Function
