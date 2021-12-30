Attribute VB_Name = "modMigrateCommits"
'@Folder "MigrateCommits"
Option Compare Database
Option Explicit

Public Sub MigrateCommits()
    Dim migCom As clsMigrateCommits
    Set migCom = LoadMigrateCommit(MIGRATE_COMMITS_FN)
    MigrateCommitTable migCom
End Sub

Private Function MigrateCommitTable(ByRef migCom As clsMigrateCommits) As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    AddMigrationFields migCom.SourceTableName
    
    sql = "SELECT * FROM " & migCom.SourceTableName & "  WHERE migrationID IS NULL"

    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            rs.Edit
            rs.Fields("MigrationID") = GUID.CreateGUID
            rs.Fields("newID") = MigrateRecord(migCom, rs)
            rs.Update
            rs.MoveNext
        Loop
    End If
    
    Set rs = Nothing
    MigrateCommitTable = True
End Function

Private Function MigrateRecord(ByRef migCom As clsMigrateCommits, ByRef srcRS As Recordset) As Double
    Dim dstRS As Recordset
    Dim fp As clsFieldPair
    Set dstRS = CurrentDb.OpenRecordset("SELECT * FROM " & migCom.DestinationTableName)
    
    dstRS.AddNew
    
    For Each fp In migCom.Fields
        dstRS.Fields(fp.Destination) = srcRS.Fields(fp.Source)
    Next fp
    
    MigrateRecord = dstRS.Fields(migCom.DestinationIDField)

    dstRS.Update
    dstRS.Close
    Set dstRS = Nothing
End Function

Private Function LoadMigrateCommit(filename As String) As clsMigrateCommits
    Dim migCom As clsMigrateCommits
    Set migCom = New clsMigrateCommits
    
    Dim dataline As String
    Open filename For Input As #1
    
    Line Input #1, dataline
    Debug.Assert Split(dataline, " ")(0) = "SOURCE"
    migCom.SourceTableName = Split(dataline, " ")(1)
    migCom.SourceIDField = Split(dataline, " ")(2)
    
    Line Input #1, dataline
    Debug.Assert Split(dataline, " ")(0) = "DESTINATION"
    migCom.DestinationTableName = Split(dataline, " ")(1)
    migCom.DestinationIDField = Split(dataline, " ")(2)
    
    Line Input #1, dataline
    Debug.Assert Split(dataline, " ")(0) = "FIELDS"
    Set migCom.Fields = New Collection
    
    Do While Not EOF(1)
        Line Input #1, dataline
        migCom.Fields.Add modFieldPairFactory.CreateFieldPair(Trim(CStr(Split(dataline, ";")(0))), CStr(Split(dataline, ";")(1)))
    Loop
    
    Close #1
    
    Set LoadMigrateCommit = migCom
End Function
