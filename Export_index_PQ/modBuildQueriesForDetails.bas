Attribute VB_Name = "modBuildQueriesForDetails"
'@Folder("index_PQ")
Option Compare Database
Option Explicit

Public Sub BuildQueriesForDetailTables()
    Dim tables As Variant, table As Variant
    Dim queryName As String
    Dim sql As String
    
    Set tables = GetListOfTablesInLinkedDatabase
    
    For Each table In tables
        If table Like "tblDetail*" Then
            queryName = Replace(table, "tbl", "qry")
            sql = GenerateSQLforDetailTable(table)
            Call CreateQuery(queryName, sql)
        End If
    Next table
End Sub

Private Function GetFieldsFromSchema(ByVal tableName As String) As Variant
    Dim fields As Collection
    Dim rs As Recordset
    Dim sql As String
    
    Set fields = New Collection
    fields.Add "TrackFK," & QUERY_TRACK_LATEST
    fields.Add "EntityFK,"
    
    sql = "SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & tableName & "';"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            fields.Add CStr(rs.fields("FieldName") & "," & rs.fields("LookupTable"))
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    
    Set GetFieldsFromSchema = fields
End Function

Private Function GenerateSQLforDetailTable(ByVal tableName As String) As String
    Dim sql As String
    Dim joins As String
    Dim fields As Collection
    Dim field As Variant
    
    Set fields = GetFieldsFromSchema(tableName)
    
    sql = "SELECT"
    
    For Each field In fields
        Debug.Print field
        If Right$(field, 1) = "," Then
            sql = sql & " " & tableName & "." & field
        Else
            ' TODO FIX
            If field <> "TrackFK,qryTrack_Latest" Then
                sql = sql & " " & GetValueFieldFromLookupTable(Split(field, ",")(1)) & ","
            End If
        End If
    Next field
    sql = Left$(sql, Len(sql) - 1)
    
    sql = sql & " FROM"
    
    joins = tableName
    
    For Each field In fields
        joins = ConcatenateJoin(joins, tableName, field)
    Next field
    
    sql = sql & joins
    
    GenerateSQLforDetailTable = sql
End Function

Private Function GetValueFieldFromLookupTable(ByVal lookupTableName As String) As String
    On Error GoTo HandleError
    If lookupTableName = QUERY_TRACK_LATEST Then
        GetValueFieldFromLookupTable = QUERY_TRACK_LATEST & ".ID"
        Exit Function
    End If
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset(lookupTableName, dbOpenSnapshot, dbReadOnly)
    GetValueFieldFromLookupTable = lookupTableName & "." & rs.fields(1).Name
    rs.Close
    
ExitHere:
    Set rs = Nothing
    Exit Function

HandleError:
    If Err.Number = 3078 Then
        GetValueFieldFromLookupTable = lookupTableName & "." & "ID"
        GoTo ExitHere
    End If
End Function

Private Function ConcatenateJoin(ByVal previous As String, ByVal tableName As String, ByVal payload As String) As String
    Dim s As String
    Dim JoinTable As String
    If Right$(payload, 1) = "," Then
        ConcatenateJoin = previous
        Exit Function
    End If
    payload = tableName & "." & Replace(payload, ",", " = ") & ".ID"
    JoinTable = Split(Replace(payload, " ", "."), ".")(3)
    s = " (" & previous & " INNER JOIN " & JoinTable & " ON " & payload & ")"
    ConcatenateJoin = s
End Function

