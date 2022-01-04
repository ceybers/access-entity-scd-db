Attribute VB_Name = "modBuildTablesForDetails"

'@Folder "Provisioning"
Option Compare Database
Option Explicit

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
        If BuildFormForDetail(Replace(tbl, "tblDetail", vbNullString)) Then ' TODO Const this
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

Private Function CreateTable(TableName As String) As Boolean
    Dim db As Database
    Dim tbl As TableDef
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    Set tbl = db.CreateTableDef(TableName)

    CreateIDField tbl
    
    CreateGenericField tbl, "EntityFK", dbLong ' TODO Const this
    CreateGenericField tbl, "TrackFK", dbLong ' TODO Const this
    
    AddFieldsToTableDefFromMetaSchema tbl, TableName
    
    db.TableDefs.Append tbl
    db.TableDefs.Refresh
    
    db.Close
    Set db = Nothing
    
    CreateTable = True
End Function

Private Sub AddFieldsToTableDefFromMetaSchema(tblDef As TableDef, TableName As String)
    Dim db As Database
    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & SCHEMA_TABLE & " WHERE TableName = '" & TableName & "';"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            AddFieldToTableDefFromMetaSchema tblDef, rs
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
    
    If Nz(rs!format) <> vbNullString Then
        Set prop = fld.CreateProperty("Format", dbText)
        prop.Value = rs!format
        'fld.Properties.Append prop
    End If
    
    If Nz(rs!defaultValue) <> vbNullString Then
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

Private Function GetListOfTablesFromSchema() As Collection
    Dim rs As Recordset
    Dim sql As String
    
    Set GetListOfTablesFromSchema = New Collection
    sql = "SELECT DISTINCT TableName FROM " & SCHEMA_TABLE & ";"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            GetListOfTablesFromSchema.Add CStr(rs!TableName)
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
End Function

Private Function FilterEmptyTablesOnly(tables As Collection) As Collection
    Dim TableName As String
    Dim tbl As Variant
    Dim db As Database
    
    Set db = OpenDatabase(BE_DATABASE_FILENAME, False, False)
    
    Set FilterEmptyTablesOnly = New Collection
    
    For Each tbl In tables
        TableName = CStr(tbl)
        If DoesTableExist(TableName, db) Then
            If IsTableEmpty(TableName, db) Then
                FilterEmptyTablesOnly.Add tbl
            End If
        Else
            FilterEmptyTablesOnly.Add tbl
        End If
    Next
    
    db.Close
    Set db = Nothing
End Function
