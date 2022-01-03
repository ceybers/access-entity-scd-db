Attribute VB_Name = "modRecordCollectionHelpers"
'@Folder("index_ORM")
Option Compare Database
Option Explicit

Public Sub LoadFromRecordset(ByRef coll As Collection, ByVal TableName As String, ByRef recordClass As IRecord)
    Dim db As Database
    Dim rs As Recordset
    On Error GoTo Catch
    
'Try
    Set db = CurrentDb
    Set rs = db.OpenRecordset(TableName, dbOpenSnapshot, dbReadOnly)
    
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

