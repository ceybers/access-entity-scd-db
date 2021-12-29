Attribute VB_Name = "modBuildTablesForDetails"
Option Compare Database
Option Explicit

Private Const BE_DATABASE_FILENAME As String = "C:\Users\User\Documents\xvba-access-test\index_BE.accdb"
Private Const LINKED_DB_CONNECT As String = ";DATABASE="
Dim FORM_NAME As String
Dim TABLE_NAME As String

Private Type TControlSet
    fieldName As String
    caption As String
    width As String
    lookupTable As String
    suffix As String
    format As String
    textalign As String
End Type

Public Sub BuildTablesForDetails()
    'If MsgBox("Build tables?", vbYesNo + vbDefaultButton2) = vbNo Then
    '    Exit Sub
    'End If
    
    Dim tables As Collection, tables2 As Collection
    
    Debug.Print "Getting list of detail tables from metaSchema..."
    Set tables = GetListOfTablesFromSchema
    Debug.Print " " & tables.count & " table(s) found"
    Debug.Print
    
    Debug.Print "Filtering to include only empty tables..."
    Set tables2 = FilterEmptyTablesOnly(tables)
    Debug.Print " " & tables2.count & " table(s) found"
    Debug.Print
    
    Dim dropResult As Integer
    Debug.Print "Removing tables with 0 records..."
    dropResult = DropTables(tables2)
    Debug.Print " " & dropResult & " table(s) dropped"
    Debug.Print
        
    Dim createResult As Integer
    Debug.Print "Creating tables..."
    createResult = CreateTables(tables2)
    Debug.Print " " & createResult & " table(s) created and linked"
    Debug.Print
        
    Dim formResult As Integer
    If MsgBox("Build forms?", vbYesNo + vbDefaultButton2 + vbInformation) = vbYes Then
        Debug.Print "Creating forms..."
        formResult = CreateForms(tables2)
        Debug.Print " " & formResult & " forms(s) built"
        Debug.Print
    End If
    
    Debug.Print "END"
End Sub

Private Function CreateForms(tables As Collection) As Integer
    Dim tbl As Variant
    For Each tbl In tables
        If BuildFormForDetail(Replace(tbl, "tblDetail", "")) Then
            CreateForms = CreateForms + 1
        End If
    Next tbl
End Function

Private Function CreateTables(tables As Collection) As Integer
    Dim tbl As Variant
    
    For Each tbl In tables
        If CreateTable(CStr(tbl)) Then
            LinkTable (CStr(tbl))
            CreateTables = CreateTables + 1
        End If
    Next tbl
End Function

Private Function CreateTable(tableName As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    Set tbl = db.CreateTableDef(tableName)

    Call CreateIDField(tbl)
    
    CreateGenericField tbl, "EntityFK", dbLong
    CreateGenericField tbl, "TrackFK", dbLong
    
    AddFieldsToTableDefFromMetaSchema tbl, tableName
    
    db.TableDefs.Append tbl
    db.TableDefs.Refresh
    
    db.Close
    Set db = Nothing
    
    CreateTable = True
End Function

Private Sub AddFieldsToTableDefFromMetaSchema(tblDef As TableDef, tableName As String)
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT * FROM metaSchema WHERE TableName = '" & tableName & "';"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            Call AddFieldToTableDefFromMetaSchema(tblDef, rs)
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

Private Function AddFieldToTableDefFromMetaSchema(tblDef As TableDef, rs As Recordset)
    Dim fldType As Integer
    Dim prop As DAO.Property
    Dim fld As Field
    
    fldType = dbText
    Select Case rs!fieldType
        Case "Number"
            fldType = dbLong
        Case "Date/Time"
            fldType = dbDate
        Case "Double"
            fldType = dbDouble
    End Select
    
    Set fld = tblDef.CreateField(rs!fieldName, fldType)
    
    If Nz(rs!format) <> "" Then
        Set prop = fld.CreateProperty("Format", dbText)
        prop.Value = rs!format
        'fld.Properties.Append prop
    End If
    
    If Nz(rs!defaultValue) <> "" Then
        fld.defaultValue = rs!defaultValue
    End If
    
    tblDef.fields.Append fld
End Function

Private Function CreateGenericField(tbl As TableDef, fieldName As String, Optional fieldType As Integer = dbText)
    Dim fld As Field
    Set fld = tbl.CreateField(fieldName, fieldType)
    tbl.fields.Append fld
End Function

Private Function CreateIDField(tbl As TableDef, Optional fieldName As String = "ID")
    Dim fld As DAO.Field
    Dim idx As DAO.index
    
    Set fld = CreateAutoNumberField(tbl, fieldName)
    tbl.fields.Append fld
    
    Set idx = tbl.CreateIndex
    
    With idx
        .name = "Primary Key"
        .fields.Append .CreateField(fieldName)
        .Unique = True
        .Primary = True
    End With
    tbl.Indexes.Append idx
End Function

Private Function CreateAutoNumberField(tbl As TableDef, fieldName As String) As Field
    Set CreateAutoNumberField = tbl.CreateField(fieldName, dbLong, 4)
    With CreateAutoNumberField
         .Attributes = dbAutoIncrField
    End With
End Function

Private Function LinkTable(tableName As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    Dim fld As Field
    
    Set db = CurrentDb
    
    If DoesTableExist(db, tableName) Then Exit Function
    
    Set tbl = db.CreateTableDef(tableName)
    
    tbl.Connect = LINKED_DB_CONNECT & BE_DATABASE_FILENAME
    tbl.SourceTableName = tableName

    db.TableDefs.Append tbl
    db.TableDefs.Refresh
    
    Set db = Nothing
    
    LinkTable = True
End Function

Private Function DropTables(tables As Collection) As Integer
    Dim tbl As Variant
    Dim db As Database
    DropTables = 0
    
    'Set db = CurrentDb
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    For Each tbl In tables
        Debug.Print " Deleting table '" & tbl & "'"
        If DoesTableExist(db, CStr(tbl)) Then
            db.Execute "DROP TABLE " & CStr(tbl), dbFailOnError
            DropTables = DropTables + 1
        End If
    Next tbl
    
    db.Close
    Set db = Nothing
End Function

Private Function GetListOfTablesFromSchema() As Collection
    Dim rs As Recordset
    Dim sql As String
    
    Set GetListOfTablesFromSchema = New Collection
    sql = "SELECT DISTINCT TableName FROM metaSchema;"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            GetListOfTablesFromSchema.Add CStr(rs!tableName)
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function FilterEmptyTablesOnly(tables As Collection) As Collection
    Dim tableName As String
    Dim tbl As Variant
    Dim db As Database
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    Set FilterEmptyTablesOnly = New Collection
    
    For Each tbl In tables
        tableName = CStr(tbl)
        If DoesTableExist(db, tableName) Then
            If IsTableEmpty(db, tableName) Then
                FilterEmptyTablesOnly.Add tbl
            End If
        Else
            FilterEmptyTablesOnly.Add tbl
        End If
    Next
    
    db.Close
    Set db = Nothing
End Function

Private Function IsTableEmpty(db, tableName As String) As Boolean
    Dim result As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    sql = "SELECT Count(*) AS TotalCount FROM " & tableName & ";"
    Set rs = db.OpenRecordset(sql)
    result = rs!TotalCount
    
    rs.Close
    Set rs = Nothing
    
    IsTableEmpty = (result = 0)
End Function

Private Function DoesTableExist(ByRef db As Database, ByVal tableName As String) As Boolean
    Dim tbl As TableDef
    For Each tbl In db.TableDefs
        If tbl.name = tableName Then
            DoesTableExist = True
            Exit Function
        End If
    Next tbl
End Function

Private Function DoesFormExist(formName As String) As Boolean
    Dim frm As Form
    For Each frm In Application.CurrentProject.AllForms
        If frm.name = formName Then
            DoesFormExist = True
            Exit Function
        End If
    Next frm
End Function
