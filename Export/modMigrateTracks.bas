Attribute VB_Name = "modMigrateTracks"
'@Folder("MigrateCommits")
Option Compare Database
Option Explicit

Public Sub MigrateTracks()
    Dim migCom As clsMigrateCommits
    Set migCom = LoadMigrateCommit("C:\Users\User\Documents\access-entity-scd-db\MigrateTracks.txt")
    MigrateCommitTable migCom
End Sub

Private Function GetParentNewID(ByRef migCom As clsMigrateCommits, parentFK As Double) As Double
    Dim sql As String
    Dim rs As Recordset
   
    ' TODO
    ' Include this in clsMigrateCommits
    sql = "SELECT * FROM tblUpdRef WHERE ID = " & parentFK
     
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        GetParentNewID = rs.Fields("newID")
    End If
    
    rs.Close
    Set rs = Nothing
End Function

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

    dstRS.Fields("CommitFK") = GetParentNewID(migCom, srcRS.Fields("updRefFK"))

    dstRS.Update
    dstRS.Close
    Set dstRS = Nothing
End Function

Private Function LoadMigrateCommit(filename As String) As clsMigrateCommits
    Dim migCom As clsMigrateCommits
    Set migCom = New clsMigrateCommits
    
    Dim dataLine As String
    Open filename For Input As #1
    
    Line Input #1, dataLine
    Debug.Assert Split(dataLine, " ")(0) = "SOURCE"
    migCom.SourceTableName = Split(dataLine, " ")(1)
    migCom.SourceIDField = Split(dataLine, " ")(2)
    
    Line Input #1, dataLine
    Debug.Assert Split(dataLine, " ")(0) = "DESTINATION"
    migCom.DestinationTableName = Split(dataLine, " ")(1)
    migCom.DestinationIDField = Split(dataLine, " ")(2)
    
    Line Input #1, dataLine
    Debug.Assert Split(dataLine, " ")(0) = "FIELDS"
    Set migCom.Fields = New Collection
    
    Do While Not EOF(1)
        Line Input #1, dataLine
        migCom.Fields.Add modFieldPairFactory.CreateFieldPair(Trim(CStr(Split(dataLine, ";")(0))), CStr(Split(dataLine, ";")(1)))
    Loop
    
    Close #1
    
    Set LoadMigrateCommit = migCom
End Function

