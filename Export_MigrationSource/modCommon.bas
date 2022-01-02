Attribute VB_Name = "modCommon"
'@Folder "Common"
Option Compare Database
Option Explicit

Public Function AddMigrationFields(ByVal tableName As String) As Boolean
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

Public Function HasField(ByVal tableName As String, ByVal fieldName As String) As Boolean
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
