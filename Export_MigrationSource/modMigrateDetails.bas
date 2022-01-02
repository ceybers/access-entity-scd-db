Attribute VB_Name = "modMigrateDetails"
'@Folder "MigrateDetails"
Option Compare Database
Option Explicit

Public Sub MigrateDetails()
    Dim tables As Collection
    Dim migDet As clsMigrateDetailTable
    
    Set tables = LoadMigrateDetailTables(MIGRATE_DETAILS_FN)
    
    For Each migDet In tables
        'PrintMigrateDetailTable migDet
        ExecuteMigrateDetailTable migDet
    Next migDet
End Sub

Private Function ExecuteMigrateDetailTable(ByRef migDet As clsMigrateDetailTable)
    Debug.Print "Does table exist? '" & migDet.Source.tableName & "' = " & DoesTableExist(CurrentDb, migDet.Source.tableName)
    Debug.Print "Does table exist? '" & migDet.Destination.tableName & "' = " & DoesTableExist(CurrentDb, migDet.Destination.tableName)
    
    If DoesTableExist(CurrentDb, migDet.Destination.tableName) = False Then
        Exit Function
    End If
    
    AddMigrationFields (migDet.Source.tableName)
    
    Dim rs As Recordset
    Dim sql As String
    sql = "SELECT * FROM " & migDet.Source.tableName & " WHERE MigrationID IS NULL"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            rs.Edit
            rs.Fields("MigrationID") = GUID.CreateGUID
            rs.Fields("newID") = MigrateRecordset(migDet, rs)
            rs.Update
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function MigrateRecordset(ByRef migDet As clsMigrateDetailTable, ByRef srcRS As Recordset)
    Dim dstRS As Recordset
    Dim db As Database
    Dim sql As String
    Dim fp As clsFieldPair
    
    'Set db = OpenDatabase(FilePaths.BE_DATABASE_FILENAME, False, False)
    Set db = CurrentDb
    sql = "SELECT * FROM " & migDet.Destination.tableName & " WHERE 1 = 0"
    Set dstRS = db.OpenRecordset(sql)
    
    dstRS.AddNew
    
    dstRS.Fields(migDet.Destination.TrackFK) = GetNewID("tblTracking", srcRS.Fields(migDet.Source.TrackFK))
    dstRS.Fields(migDet.Destination.ID) = GetNewID("tblTankID", srcRS.Fields(migDet.Source.ID), "tankID")
    
    For Each fp In migDet.Fields
        If fp.Lookup = "" Then
            dstRS.Fields(fp.Destination) = srcRS.Fields(fp.Source)
        Else
            dstRS.Fields(fp.Destination) = TranslateLookupValues(fp, srcRS.Fields(fp.Source))
        End If
    Next fp

    dstRS.Update
    
    dstRS.Close
    Set dstRS = Nothing

    MigrateRecordset = db.OpenRecordset("SELECT @@IDENTITY")(0)
End Function

Private Function GetNewID(ByVal tableName As String, oldID As Double, Optional idFieldName As String = "ID") As Double
    Dim sql As String
    Dim rs As Recordset
    Dim db As Database
    
    sql = "SELECT newID FROM " & tableName & " WHERE " & idFieldName & " = " & oldID
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql)
    GetNewID = rs.Fields("newID")
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Private Function TranslateLookupValues(ByRef fp As clsFieldPair, oldLookupValue As Double) As Double
    Dim rs As Recordset
    Dim db As Database
    Dim sql As String
    
    sql = "SELECT newID FROM " & fp.Lookup & " WHERE ID = " & oldLookupValue
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    TranslateLookupValues = rs.Fields("newID")
    
    rs.Close
    Set rs = Nothing
    
    'TranslateLookupValues = 42
End Function

Private Function PrintMigrateDetailTable(migDet As clsMigrateDetailTable)
    Debug.Print "PrintMigrateDetailTable"
    Debug.Print " " & migDet.Source.tableName & " -> " & migDet.Destination.tableName
    Debug.Print "  " & migDet.Source.ID & " -> " & migDet.Destination.ID
    Debug.Print "  " & migDet.Source.TrackFK & " -> " & migDet.Destination.TrackFK
    Dim fp As clsFieldPair
    For Each fp In migDet.Fields
        Debug.Print "  " & fp.Source & " -> " & fp.Destination & " (" & fp.Lookup & ")"
    Next fp
    Debug.Print ""
End Function

Private Function LoadMigrateDetailTables(filename) As Variant
    Dim dataline As String
    Dim arr As Variant
    Dim migDet As clsMigrateDetailTable
    Dim fp As clsFieldPair
    Dim tables As Collection
    
    Set tables = New Collection
    
    Open filename For Input As #1
    Do While Not EOF(1)
        Line Input #1, dataline
        arr = Split(dataline, " ")
        
        If arr(0) = "DETAIL" Then
            If Not migDet Is Nothing Then
                tables.Add migDet
            End If
            Set migDet = New clsMigrateDetailTable
            Set migDet.Source = New clsDetailTableStub
            Set migDet.Destination = New clsDetailTableStub
            Set migDet.Fields = New Collection
        ElseIf arr(0) = "END" Then
            tables.Add migDet
        Else
            Select Case arr(1)
                Case "SOURCE"
                    migDet.Source.tableName = arr(2)
                Case "DESTINATION"
                    migDet.Destination.tableName = arr(2)
                Case "ID"
                    migDet.Source.ID = Split(arr(2), ";")(0)
                    migDet.Destination.ID = Split(arr(2), ";")(1)
                Case "TRACK"
                    migDet.Source.TrackFK = Split(arr(2), ";")(0)
                    migDet.Destination.TrackFK = Split(arr(2), ";")(1)
                Case "FIELDS"
                Case Else
                    Set fp = CreateFieldPairFromArray(Split(Trim(dataline), ";"))
                    
                    migDet.Fields.Add fp
                
            End Select
        End If
    Loop
    Close #1
    Set LoadMigrateDetailTables = tables
End Function
