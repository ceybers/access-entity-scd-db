Attribute VB_Name = "modRecordCollectionHelpers"
'@Folder "ORM"
Option Compare Database
Option Explicit

Public Sub LoadFromRecordset(ByRef coll As Collection, ByVal tableName As String, ByRef recordClass As IRecord)
    Dim db As Database
    Dim rs As Recordset
    On Error GoTo Catch
    
'Try
    Set db = CurrentDb
    Set rs = db.OpenRecordset(tableName, dbOpenSnapshot, dbReadOnly)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            coll.Add recordClass.Create(rs)
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

Public Function GetFieldValue(tableName As String, ID As Double, fieldName As String) As Double
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    On Error GoTo Catch
    
'Try
    sql = "SELECT " & fieldName & " FROM " & tableName & " WHERE ID = " & ID
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)

    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            GetFieldValue = rs.Fields(fieldName).Value
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
End Function
