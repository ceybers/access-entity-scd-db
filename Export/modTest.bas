Attribute VB_Name = "modTest"
Option Compare Database
Option Explicit

Private Sub aaa()
    Dim migCom As clsMigrateCommits
    Set migCom = LoadMigrateCommit(MIGRATE_COMMITS_FN)
    Call PrintMigrateCommit(migCom)
End Sub

Private Function PrintMigrateCommit(migCom As clsMigrateCommits)
    Dim fld As clsFieldPair
    Debug.Print "MigrateCommits:"
    Debug.Print " Source:"
    Debug.Print "  Table: " & migCom.SourceTableName
    Debug.Print "  ID: " & migCom.SourceIDField
    Debug.Print " Destination:"
    Debug.Print "  Table: " & migCom.DestinationTableName
    Debug.Print "  ID: " & migCom.DestinationIDField
    Debug.Print " Fields:"
    For Each fld In migCom.Fields
        Debug.Print fld.Source & " -> " & fld.Destination
    Next fld
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
        migCom.Fields.Add modFieldPairFactory.CreateFieldPair(CStr(Split(dataline, ";")(0)), CStr(Split(dataline, ";")(1)))
    Loop
    
    Close #1
    
    Set LoadMigrateCommit = migCom
End Function
