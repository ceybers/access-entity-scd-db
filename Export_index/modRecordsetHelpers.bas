Attribute VB_Name = "modRecordsetHelpers"
'@Folder("Helpers")
Option Compare Database
Option Explicit

Public Function RecordsetToCollection(ByVal sql As String, ByRef recordsetCollector As IRecordsetCollector) As collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim coll As collection
    
    Set coll = New collection
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            recordsetCollector.AddRecord rs, coll
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    Set RecordsetToCollection = coll
End Function
