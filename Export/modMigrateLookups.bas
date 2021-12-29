Attribute VB_Name = "modMigrateLookups"
'@Folder "MigrateLookups"
Option Compare Database
Option Explicit

Private Const BE_DATABASE_FILENAME As String = "C:\Users\User\Documents\access-entity-scd-db\index_BE.accdb"
Private Const LINKED_DB_CONNECT As String = ";DATABASE="

Public Sub MigrateLookups()
    Dim filename As String
    Dim dataline As String
    
    filename = "C:\Users\User\Documents\access-entity-scd-db\MigrateLookups.txt"
    
    Open filename For Input As #1
    Do While Not EOF(1)
        Line Input #1, dataline
        MigrateLookupTables dataline
    Loop
    Close #1
End Sub

Private Function MigrateLookupTables(instruction As String)
    Dim arr As Variant
    Dim srcTable As String
    Dim dstTable As String
    Dim fieldName As String
    
    arr = Split(instruction, ";")
        
    srcTable = arr(0)
    dstTable = arr(1)
    fieldName = arr(2)
    
    Debug.Print "Migrating Lookup Table '" & srcTable & "' -> '" & dstTable & "'"
    AddMigrationFields srcTable
    CreateLookupTable dstTable, fieldName
    LinkTable dstTable, LINKED_DB_CONNECT, BE_DATABASE_FILENAME
    MigrateLookupTable srcTable, dstTable, fieldName
End Function

Private Function MigrateLookupTable(ByVal srcTable As String, ByVal dstTable As String, ByVal fieldName As String) As Boolean

    Dim srcRS As Recordset
    Dim sql As String
    Dim fp As clsFieldPair
    
    Set fp = New clsFieldPair
    fp.Source = fieldName
    fp.Destination = fieldName
    sql = "SELECT * FROM " & srcTable & " WHERE MigrationID IS NULL"
    Set srcRS = CurrentDb.OpenRecordset(sql)
    
    If Not srcRS.BOF And Not srcRS.EOF Then
        Do While Not srcRS.EOF
            srcRS.Edit
            srcRS.Fields("MigrationID") = GUID.CreateGUID
            srcRS.Fields("newID") = MigrateRecord(srcRS, dstTable, fp)
            srcRS.Update
            srcRS.MoveNext
        Loop
    End If
    
    srcRS.Close
    Set srcRS = Nothing
    
    MigrateLookupTable = True
End Function

Private Function MigrateRecord(ByRef srcRS As Recordset, ByVal dstTable As String, ByVal fp As clsFieldPair) As Double
    Dim dstRS As Recordset
    Dim sql As String
    sql = "SELECT * FROM " & dstTable
    Set dstRS = CurrentDb.OpenRecordset(sql)
    dstRS.AddNew

    dstRS.Fields(fp.Destination) = srcRS.Fields(fp.Source)
    MigrateRecord = dstRS.Fields("ID")
    
    dstRS.Update
    dstRS.Close
    Set dstRS = Nothing
End Function

Private Function CreateLookupTable(tableName As String, fieldName As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    If DoesTableExist(db, tableName) Then
        Debug.Print "Table '" & tableName & "' already exists in BE db"
    Else
        Set tbl = db.CreateTableDef(tableName)
    
        Call CreateIDField(tbl)
        
        CreateGenericField tbl, fieldName, dbText
        
        db.TableDefs.Append tbl
        db.TableDefs.Refresh
        CreateLookupTable = True
    End If
    
    db.Close
    Set db = Nothing
End Function


