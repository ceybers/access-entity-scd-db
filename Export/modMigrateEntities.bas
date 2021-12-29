Attribute VB_Name = "modMigrateEntities"
Option Compare Database
Option Explicit

Public Sub MigrateEntities()
    With New clsMigrationSource
        .SetValues "tblDivision", "divisionID", "divisionName", 1
        MigrateEntityTable .Self
    End With
    
    With New clsMigrationSource
        .SetValues "tblBusStream", "streamID", "streamName", 2, "divisionFK", "tblDivision", "divisionID"
        MigrateEntityTable .Self
    End With
    
    With New clsMigrationSource
        .SetValues "tblDepot", "depotID", "depotName", 3, "busStreamFK", "tblBusStream", "streamID"
        MigrateEntityTable .Self
    End With
    
    With New clsMigrationSource
        .SetValues "tblTankID", "tankID", "tankCode", 4, "depotFK", "tblDepot", "depotID"
        MigrateEntityTable .Self
    End With
End Sub

Private Function MigrateEntityTable(migSrc As clsMigrationSource) As Boolean
    Dim sql As String
    Dim rs As Recordset
    Dim parentID As Double
    
    AddMigrationFields migSrc.tableName
    
    sql = "SELECT * FROM " & migSrc.tableName ' & " WHERE migrationID IS NULL"

    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            rs.Edit
            If migSrc.parentField = "" Then
                parentID = 0
            Else
                parentID = GetParentNewID(migSrc, rs.Fields(migSrc.parentField))
            End If
            rs.Fields("newID") = AddEntity(rs.Fields(migSrc.nameField), parentID, migSrc.entityTypeID)
            rs.Fields("MigrationID") = GUID.CreateGUID()
            rs.Update
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    MigrateEntityTable = True
End Function

Private Function GetParentNewID(migSrc As clsMigrationSource, someID As Double) As Double
    Dim sql As String
    Dim rs As Recordset
   
    sql = "SELECT * FROM " & migSrc.parentTableName & " WHERE " & migSrc.parentPK & " = " & someID
     
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        GetParentNewID = rs.Fields("newID")
    End If
    
    rs.Close
    Set rs = Nothing
End Function
'Private Function GetNewID(tableName as String, )
' Each table has different ID field name...

Private Function AddEntity(ByVal entityName As String, ByVal parentFK As Double, ByVal entityTypeID As Double) As Double
    Dim sql As String
    Dim rs As Recordset
    Dim db As Database
    
    Set db = CurrentDb
    sql = "INSERT INTO tblEntities (Entity, ParentFK, EntityType) VALUES ('" & entityName & "', " & parentFK & " , " & entityTypeID & ");"
    db.Execute sql
    
    Set rs = db.OpenRecordset("SELECT @@IDENTITY AS LastID;")
    AddEntity = rs.Fields("lastID")
    
    rs.Close
    Set rs = Nothing
End Function
