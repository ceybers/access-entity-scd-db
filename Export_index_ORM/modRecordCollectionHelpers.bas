Attribute VB_Name = "modRecordCollectionHelpers"
'@Folder "ORM"
Option Compare Database
Option Explicit

Public Sub LoadFromRecordset(ByRef coll As Collection, ByVal TableName As String, ByRef recordClass As IRecord)
    Dim db As Database
    Dim rs As Recordset
    Dim key As String
    On Error GoTo Catch
    
'Try
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TableName, dbOpenSnapshot, dbReadOnly)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            key = COLLECTION_INDEX_PREFIX & rs.Fields("ID")
            coll.Add recordClass.Create(rs), key
            rs.MoveNext
        Loop
    End If
    GoTo Finally
    
Catch:
    Err.Raise Err.Number, Err.Source, Err.Description
    
Finally:
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' TODO Refactor this; we can't run single queries for each iteration of the calling for-each loop
Public Function GetFieldValue(ByVal TableName As String, ByVal ID As Double, ByVal fieldName As String) As Double
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    On Error GoTo Catch
    
'Try
    GetFieldValue = -1
    sql = "SELECT " & fieldName & " FROM " & TableName & " WHERE ID = " & ID
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)

    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            GetFieldValue = Nz(rs.Fields(fieldName).Value, -1)
            rs.MoveNext
        Loop
    End If
    GoTo Finally
    
Catch:
    Err.Raise Err.Number, , Err.Description
    
Finally:
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function