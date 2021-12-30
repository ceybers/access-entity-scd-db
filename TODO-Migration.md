# Migration TODO

1. Entity Types (manual?)
2. Entities
3. tblUpdRef -> tblCommits
4. tblTracking -> tblTrack
5. lkp* -> lkp*
6. tbl* -> tblDetail*

## Patterns
### Notes
Should declare dstRS outside of the loop, and pass it ByRef, instead of continuously opening and closing the recordset.

### Iter
``` 
Dim db as Database
Dim rs As Recordset

sql = "SELECT * FROM table WHERE migrationID IS NULL"

Set db = CurrentDb
Set rs = db.OpenRecordset(sql)

If Not rs.BOF And Not rs.EOF Then
    Do While Not rs.EOF
        rs.Edit
        rs.Fields("MigrationID") = GUID.CreateGUID
        rs.Fields("newID") = MigrateRecord(payload, rs)
        rs.Update
        rs.MoveNext
    Loop
End If

rs.Close
Set rs = Nothing
Set db = Nothing
```

### Map
```
Dim db as Database 
Dim dstRS As Recordset
Dim fp As clsFieldPair

Set db = CurrentDb
'Set dstRS = db.OpenRecordset("SELECT * FROM " & tableName)
Set dstRS = db.OpenRecordset(tableName, dbOpenTable, dbAppendOnly)
dstRS.AddNew

For Each fp In payload.Fields
    dstRS.Fields(fp.Destination) = srcRS.Fields(fp.Source)
Next fp

MigrateRecord = dstRS.Fields(payload.DestinationIDField)

dstRS.Update
dstRS.Close

Set dstRS = Nothing
Set db = Nothing
```