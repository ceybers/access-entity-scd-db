Attribute VB_Name = "modBuildQueriesForDetails"
'@Folder("Provisioning")
Option Compare Database
Option Explicit

Public Sub AAA_TEST()
    Dim sql As String
    sql = GenerateSQLforDetailTable("tblDetailMaintPlan")
    Debug.Print "[" & format(Now(), "hh:mm:ss") & "] " & sql
    'Debug.Print sql
End Sub

Private Function GetFieldsFromSchema(ByVal tableName As String) As Variant
    Dim fields As Collection
    Dim rs As Recordset
    Dim sql As String
    
    Set fields = New Collection
    fields.Add "TrackFK," & QUERY_TRACK_LATEST
    
    sql = "SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & tableName & "';"
    Set rs = CurrentDB.OpenRecordset(sql)
    
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
        If Right$(field, 1) = "," Then
            sql = sql & " " & tableName & "." & field
        Else
            sql = sql & " " & GetValueFieldFromLookupTable(split(field, ",")(1)) & ","
        End If
    Next field
    sql = left$(sql, Len(sql) - 1)
    
    sql = sql & " FROM"
    
    joins = tableName
    
    For Each field In fields
        joins = ConcatenateJoin(joins, tableName, field)
    Next field
    
    sql = sql & joins
    
    GenerateSQLforDetailTable = sql
End Function

Private Function GetValueFieldFromLookupTable(ByVal lookupTableName As String) As String
    If lookupTableName = QUERY_TRACK_LATEST Then
        GetValueFieldFromLookupTable = QUERY_TRACK_LATEST & ".ID"
        Exit Function
    End If
    Dim rs As Recordset
    Set rs = CurrentDB.OpenRecordset(lookupTableName, dbOpenSnapshot, dbReadOnly)
    GetValueFieldFromLookupTable = lookupTableName & "." & rs.fields(1).name
    rs.Close
    Set rs = Nothing
End Function

Private Function ConcatenateJoin(ByVal previous As String, ByVal tableName As String, ByVal payload As String) As String
    Dim s As String
    Dim joinTable As String
    If Right$(payload, 1) = "," Then
        ConcatenateJoin = previous
        Exit Function
    End If
    payload = tableName & "." & Replace(payload, ",", " = ") & ".ID"
    joinTable = split(Replace(payload, " ", "."), ".")(3)
    s = " (" & previous & " INNER JOIN " & joinTable & " ON " & payload & ")"
    ConcatenateJoin = s
End Function
