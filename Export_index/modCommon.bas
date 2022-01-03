Attribute VB_Name = "modCommon"
'@Folder "index"
Option Compare Database
Option Explicit

Private Function abc()
    Debug.Print "HI"
End Function

Public Function DropTables(tables As Collection, Optional ByRef db As Database) As Integer
    Dim tbl As Variant
    DropTables = 0
    
    If db Is Nothing Then Set db = CurrentDB
    For Each tbl In tables
        Debug.Print " Deleting table '" & tbl & "'"
        If DoesTableExist(tbl, db) Then
            db.Execute "DROP TABLE " & tbl, dbFailOnError
            DropTables = DropTables + 1
        End If
    Next tbl
    
    db.Close
    Set db = Nothing
End Function

Public Function IsTableEmpty(tableName As String, Optional ByRef db As Database) As Boolean
    Dim result As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    If db Is Nothing Then Set db = CurrentDB
    sql = "SELECT Count(*) AS TotalCount FROM " & tableName & ";"
    Set rs = db.OpenRecordset(sql)
    result = rs!TotalCount
    
    rs.Close
    Set rs = Nothing
    
    IsTableEmpty = (result = 0)
End Function

Public Function DoesTableExist(ByVal tableName As String, Optional ByRef db As Database) As Boolean
    Dim tbl As TableDef
    If db Is Nothing Then Set db = CurrentDB
    For Each tbl In db.TableDefs
        If tbl.name = tableName Then
            DoesTableExist = True
            Exit Function
        End If
    Next tbl
End Function

Public Function DoesFormExist(formName As String) As Boolean
    Dim frm As Form
    For Each frm In Application.CurrentProject.AllForms
        If frm.name = formName Then
            DoesFormExist = True
            Exit Function
        End If
    Next frm
End Function

Public Sub OpenFormInDesignMode(formName As String)
    DoCmd.OpenForm formName:=formName, View:=acDesign
End Sub

Public Sub CloseFormInDesignMode(formName As String)
    DoCmd.Close acForm, formName, acSaveYes
End Sub
