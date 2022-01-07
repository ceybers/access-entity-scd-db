Attribute VB_Name = "modTest"
'@Folder("Main")
Option Compare Database
Option Explicit

Public Sub TEST()
    Dim controlSets As collection
    Dim sql As String
    
    sql = "SELECT * FROM metaSchema WHERE TableName = 'tblDetailA';"
    Set controlSets = modRecordsetHelpers.RecordsetToCollection(sql, SubformControlSetRSCollector)
    
    Debug.Print controlSets.count
    Dim cs As subformcontrolset
    For Each cs In controlSets
        Debug.Print cs.FieldName
    Next
End Sub

Private Function GetFields(tableName As String) As Variant ' TControlSet()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim results() As TControlSet
    Dim i As Long
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & tableName & "';")
    i = 1
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            ReDim Preserve results(1 To i)
            results(i) = RecordToControlSet(rs)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    GetFields = results
End Function
