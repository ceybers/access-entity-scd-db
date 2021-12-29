Attribute VB_Name = "modMain"
Option Compare Database
Option Explicit

Public Sub Main()
    'Debug.Print Len(GUID.CreateGUID())
    'AddMigrationFields "tblDivision"
    'MigrateEntity "tblDivision", "tblEntities", 1, "divisionName"
    'MigrateEntity "tblBusStream", "tblEntities", 2, "streamName", "tblDivision", "divisionFK"
    
    Dim migSrc As clsMigrationSource
    Set migSrc = New clsMigrationSource
    migSrc.SetValues "tblBusStream", "streamID", "streamName", "divisionFK", "tblDivision", 2
    MigrateEntityTable migSrc
End Sub

Private Function MigrateEntityTable(migSrc As clsMigrationSource) As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    AddMigrationFields migSrc.tableName
    
    sql = "SELECT * FROM " & migSrc.tableName & " WHERE migrationID IS NULL"

    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            rs.Edit
            rs.Fields("newID") = AddEntity(rs.Fields(migSrc.nameField), 0, migSrc.entityTypeID)
            rs.Fields("MigrationID") = GUID.CreateGUID()
            rs.Update
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    MigrateEntityTable = True
End Function

'Private Function GetNewID(tableName as String, )
' Each table has different ID field name...

Private Function AddEntity(ByVal entityName As String, ByVal parentFK As Double, ByVal entityTypeID As Double) As Double
    Dim sql As String
    Dim rs As Recordset
    Dim db As Database
    
    Set db = CurrentDb
    sql = "INSERT INTO tblEntities (Entity, ParentFK, EntityType) VALUES ('" & entityName & "', 0, " & entityTypeID & ");"
    db.Execute sql
    
    Set rs = db.OpenRecordset("SELECT @@IDENTITY AS LastID;")
    AddEntity = rs.Fields("lastID")
    
    rs.Close
    Set rs = Nothing
End Function


Private Function AddMigrationFields(ByVal tableName As String) As Boolean
    Dim db As Database
    Dim tdf As TableDef
    Dim fld As Field
    
    If HasField(tableName, "MigrationID") Then Exit Function
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    
    Set fld = tdf.CreateField("MigrationID", dbText, 38)
    tdf.Fields.Append fld
    
    Set fld = tdf.CreateField("newID", dbLong)
    tdf.Fields.Append fld
    
    'MsgBox "Added Migration fields to '" & tableName & "'"
    
    Set tdf = Nothing
    Set db = Nothing
    AddMigrationFields = True
End Function


Private Function HasField(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Dim db As Database
    Dim tdf As TableDef
    Dim fld As Field
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    For Each fld In tdf.Fields
        If fld.Name = fieldName Then
            HasField = True
            Exit Function
        End If
    Next fld
    
    Set tdf = Nothing
    Set db = Nothing
End Function
